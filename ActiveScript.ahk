/*
 *  ActiveScript for AutoHotkey v2.0-a128
 *
 *  Provides an interface to Active Scripting languages like VBScript and JScript,
 *  without relying on Microsoft's ScriptControl, which is not available to 64-bit
 *  programs.
 *
 *  License: Use, modify and redistribute without limitation, but at your own risk.
 */
class ActiveScript
{
    __New(Language)
    {
        try {
            if this._script := ComObject(Language, ActiveScript.IID)
                this._scriptParse := ComObjQuery(this._script, ActiveScript.IID_Parse)
        } catch {
            throw ValueError("Invalid language", -1, Language)
        }
        this._site := ActiveScriptSite(this)
        this._SetScriptSite(this._site.ptr)
        this._InitNew()
        this._objects := Map()
        this._objects.CaseSense := false ; Legacy behaviour.
        this.Error := ""
        this._dsp := this._GetScriptDispatch()
        try
            if this.ScriptEngine() = "JScript"
                this.SetJScript58()
        for m in ['__Get', '__Call', '__Set']  ; Must be done last.
            this.%m% := ActiveScript.Meta.Prototype.%m%
    }

    SetJScript58()
    {
        static IID_IActiveScriptProperty := "{4954E0D0-FBC7-11D1-8410-006008C3FBFC}"
        prop := ComObjQuery(this._script, IID_IActiveScriptProperty)
        NumPut 'int64', 3, 'int64', 2, var := Buffer(24)
        ComCall 4, prop, "uint", 0x4000, "ptr", 0, "ptr", &var
    }
    
    Eval(Code)
    {
        ref := ComValue(0x400C, (var := Buffer(24, 0)).ptr)
        this._ParseScriptText(Code, 0x20, var)  ; SCRIPTTEXT_ISEXPRESSION := 0x20
        return (val := ref[], ref[] := 0, val)
    }
    
    Exec(Code)
    {
        this._ParseScriptText(Code, 0x42, 0)  ; SCRIPTTEXT_ISVISIBLE := 2, SCRIPTTEXT_ISPERSISTENT := 0x40
        this._SetScriptState(2)  ; SCRIPTSTATE_CONNECTED := 2
    }
    
    AddObject(Name, DispObj, AddMembers := false)
    {
        this._objects[Name] := DispObj
        this._AddNamedItem(Name, AddMembers ? 8 : 2)  ; SCRIPTITEM_ISVISIBLE := 2, SCRIPTITEM_GLOBALMEMBERS := 8
    }
    
    _GetObjectUnk(Name)
    {
        return !IsObject(dsp := this._objects[Name]) ? dsp  ; Pointer
            : ComObjType(dsp) ? ComObjValue(dsp)  ; ComObject
            : ObjPtr(dsp)  ; AutoHotkey object
    }
    
    class Meta
    {
        __Call(Method, Params)
        {
            try
                return this._dsp.%Method%(Params*)
            catch as e
                throw Error(e.Message, -1, e.Extra)
        }
        
        __Get(Property, Params)
        {
            try
                return this._dsp.%Property%[Params*]
            catch as e
                throw Error(e.Message, -1, e.Extra)
        }
        
        __Set(Property, Params, Value)
        {
            try
                return this._dsp.%Property%[Params*] := Value
            catch as e
                throw Error(e.Message, -1, e.Extra)
        }
    }
    
    _SetScriptSite(Site)
    {
        ; IActiveScript::SetScriptSite
        ComCall 3, this._script, "ptr", Site
    }
    
    _SetScriptState(State)
    {
        try ; IActiveScript::SetScriptState
            ComCall 5, this._script, "int", State
        catch as err
            this._Rethrow err
    }
    
    _AddNamedItem(Name, Flags)
    {
        ; IActiveScript::AddNamedItem
        ComCall 8, this._script, "wstr", Name, "uint", Flags
    }
    
    _GetScriptDispatch()
    {
        ; IActiveScript::GetScriptDispatch
        ComCall 10, this._script, "ptr", 0, "ptr*", &pdsp := 0
        return ComObjFromPtr(pdsp)
    }
    
    _InitNew()
    {
        ; IActiveScriptParse::InitNew
        ComCall 3, this._scriptParse
    }
    
    _ParseScriptText(Code, Flags, pvarResult)
    {
        try ; IActiveScriptParse::ParseScriptText
            ComCall 5, this._scriptParse
                , "wstr", Code, "ptr", 0, "ptr", 0, "ptr", 0, "uptr", 0, "uint", 1
                , "uint", Flags, "ptr", pvarResult, "ptr", 0
        catch as err
            this._Rethrow err
    }
    
    _Rethrow(err)
    {
        ; If _OnScriptError was called, the error information was stored in this.Error.
        throw this.HasOwnProp('Error') ? this.DeleteProp('Error') : err
    }
    
    _OnScriptError(err) ; IActiveScriptError err
    {
        excp := Buffer(8 * A_PtrSize, 0)
        ComCall 3, err, "ptr", excp ; GetExceptionInfo
        ComCall 4, err, "uint*", &srcctx := 0, "uint*", &srcline := 0, "int*", &srccol := 0 ; GetSourcePosition
        ; Seems to always throw "unspecified error":
        ; ComCall 5, err, "ptr*", &pbstrcode := 0 ; GetSourceLineText
        ; code := StrGet(pbstrcode, "UTF-16"), DllCall("OleAut32\SysFreeString", "ptr", pbstrcode)
        if fn := NumGet(excp, 6 * A_PtrSize, "ptr") ; pfnDeferredFillIn
            DllCall fn, "ptr", excp
        wcode := NumGet(excp, 0, "ushort")
        hr := wcode ? 0x80040200 + wcode : NumGet(excp, 7 * A_PtrSize, "uint")
        this.Error := e := (ActiveScript.Error)()
        for field in ['What', 'Message', 'File']
            if pbstr := NumGet(excp, A_Index * A_PtrSize, "ptr")
                e.%field% := StrGet(pbstr, "UTF-16"), DllCall("OleAut32\SysFreeString", "ptr", pbstr)
            else
                e.%field% := ""
        switch e.File
        {
        case "", A_LineFile:
            e.File := "<Eval>", e.Line := srcline ; Won't affect error dialogs, but might be used by script.
        default:
            e.Line := NumGet(excp, 4 * A_PtrSize, "uint") ; dwHelpContext is set by built-in IDispatch support.
        }
        e.Message := Format("`nError code:`t0x{:x}`nSource:`t`t{}`nDescription:`t{}`nLine:`t`t{}`nColumn:`t`t{}"
            , hr, e.What, e.Message, srcline, srccol)
        ; Returning any failure code results in 0x80020009 (DISP_E_EXCEPTION),
        ; whereas returning 0 (S_OK) results in 0x80020101 (SCRIPT_E_REPORTED)
        ; for _ParseScriptText and 0 (S_OK) for _SetScriptState.  Return failure
        ; so we don't need to check for this.Error after successful execution.
        return 0x80020009
    }
    
    __Delete()
    {
        if this._script
            ComCall 7, this._script  ; Close
    }
    
    class Error extends Error
    {
        __New() => this  ; Override the default constructor.
    }
    
    static IID := "{BB1A2AE1-A4F9-11cf-8F20-00805F2CD064}"
    static IID_Parse := A_PtrSize=8 ? "{C7EF7658-E1EE-480E-97EA-D52CB4D76D17}" : "{BB1A2AE2-A4F9-11cf-8F20-00805F2CD064}"
}

class ActiveScriptSite
{
    __New(Script)
    {
        _vftable(PrmCounts, EIBase)
        {
            buf := Buffer(StrLen(PrmCounts) * A_PtrSize)
            Loop Parse PrmCounts
            {
                cb := CallbackCreate(_ActiveScriptSite.Bind(A_Index + EIBase), "F", A_LoopField)
                NumPut 'ptr', cb, buf, (A_Index-1) * A_PtrSize
            }
            return buf
        }
        
        static vft := _vftable("31125232211", 0)
        static vft_w := _vftable("31122", 0x100)
        
        NumPut 'ptr', vft.ptr, 'ptr', vft_w.ptr, 'ptr', ObjPtr(Script)
            , this.ptr := Buffer(3 * A_PtrSize)
    }
}

_ActiveScriptSite(index, this, a1:=0, a2:=0, a3:=0, a4:=0, a5:=0)
{
    if index >= 0x100  ; IActiveScriptSiteWindow
    {
        index -= 0x100
        switch index
        {
        case 4:  ; GetWindow
            NumPut 'ptr', 0, a1 ; *phwnd := 0
            return 0 ; S_OK
        case 5:  ; EnableModeless
            return 0 ; S_OK
        }
        this -= A_PtrSize     ; Cast to IActiveScriptSite
    }
    ;else: IActiveScriptSite
    switch index
    {
    case 1:  ; QueryInterface
        iid := _AS_GUIDToString(a1)
        if (iid = "{00000000-0000-0000-C000-000000000046}"  ; IUnknown
         || iid = "{DB01A1E3-A42B-11cf-8F20-00805F2CD064}") ; IActiveScriptSite
        {
            NumPut 'ptr', this, a2
            return 0 ; S_OK
        }
        if (iid = "{D10F6761-83E9-11cf-8F20-00805F2CD064}") ; IActiveScriptSiteWindow
        {
            NumPut 'ptr', this + A_PtrSize, a2
            return 0 ; S_OK
        }
        NumPut 'ptr', 0, a2
        return 0x80004002 ; E_NOINTERFACE
    case 5:  ; GetItemInfo
        a1 := StrGet(a1, "UTF-16")
        , (a3 && NumPut('ptr', 0, a3))  ; *ppiunkItem := NULL
        , (a4 && NumPut('ptr', 0, a4))  ; *ppti := NULL
        if (a2 & 1) ; SCRIPTINFO_IUNKNOWN
        {
            if !(unk := ObjFromPtrAddRef(NumGet(this + A_PtrSize*2, 'ptr'))._GetObjectUnk(a1))
                return 0x8002802B ; TYPE_E_ELEMENTNOTFOUND
            ObjAddRef(unk), NumPut('ptr', unk, a3)
        }
        return 0 ; S_OK
    case 9:  ; OnScriptError
        return ObjFromPtrAddRef(NumGet(this + A_PtrSize*2, 'ptr'))._OnScriptError(a1)
    }
    
    ; AddRef and Release don't do anything because we want to avoid circular references.
    ; The site and IActiveScript are both released when the AHK script releases its last
    ; reference to the ActiveScript object.
    
    ; All of the other methods don't require implementations.
    return 0x80004001 ; E_NOTIMPL
}

_AS_GUIDToString(pGUID)
{
    VarSetStrCapacity(&sGuid, 38)
    DllCall("ole32\StringFromGUID2", "ptr", pGUID, "str", &sGuid, "int", 39)
    return sGuid
}
