/*
 *  JsRT for AutoHotkey v2.0-a128
 *
 *  Utilizes the JavaScript engine that comes with IE11 or legacy Edge.
 *
 *  License: Use, modify and redistribute without limitation, but at your own risk.
 */
class JsRT
{
    __New()
    {
        throw Error("This class is abstract. Use JsRT.IE or JSRT.Edge instead.", -1)
    }
    
    class IE extends JsRT
    {
        __New()
        {
            if !this._hmod := DllCall("LoadLibrary", "str", "jscript9", "ptr")
                throw Error("Failed to load jscript9.dll", -1)
            if DllCall("jscript9\JsCreateRuntime", "int", 0, "int", -1
                , "ptr", 0, "ptr*", &runtime:=0) != 0
                throw Error("Failed to initialize JsRT", -1)
            DllCall("jscript9\JsCreateContext", "ptr", runtime, "ptr", 0, "ptr*", &context:=0)
            this._Initialize("jscript9", runtime, context)
        }
    }
    
    class Edge extends JsRT
    {
        __New()
        {
            if !this._hmod := DllCall("LoadLibrary", "str", "chakra", "ptr")
                throw Error("Failed to load chakra.dll", -1)
            if DllCall("chakra\JsCreateRuntime", "int", 0
                , "ptr", 0, "ptr*", &runtime:=0) != 0
                throw Error("Failed to initialize JsRT", -1)
            DllCall("chakra\JsCreateContext", "ptr", runtime, "ptr*", &context:=0)
            this._Initialize("chakra", runtime, context)
        }
        
        ProjectWinRTNamespace(namespace)
        {
            return DllCall("chakra\JsProjectWinRTNamespace", "wstr", namespace)
        }
    }
    
    _Initialize(dll, runtime, context)
    {
        this._dll := dll
        this._runtime := runtime
        this._context := context
        DllCall(dll "\JsSetCurrentContext", "ptr", context)
        DllCall(dll "\JsGetGlobalObject", "ptr*", &globalObject:=0)
        this._dsp := this._JsToVt(globalObject)
        for m in ['__Get', '__Call', '__Set']  ; Must be done last.
            this.%m% := JsRT.Meta.Prototype.%m%
    }
    
    __Delete()
    {
        this._dsp := ""
        if dll := this._dll
        {
            DllCall(dll "\JsSetCurrentContext", "ptr", 0)
            DllCall(dll "\JsDisposeRuntime", "ptr", this._runtime)
        }
        DllCall("FreeLibrary", "ptr", this._hmod)
    }
    
    _JsToVt(valref)
    {
        ref := ComValue(0x400C, (var := Buffer(24, 0)).ptr)
        DllCall(this._dll "\JsValueToVariant", "ptr", valref, "ptr", var)
        return (val := ref[], ref[] := 0, val)
    }
    
    _ToJs(val)
    {
        ref := ComValue(0x400C, (var := Buffer(24, 0)).ptr)
        ref[] := val
        DllCall(this._dll "\JsVariantToValue", "ptr", var, "ptr*", &valref:=0)
        ref[] := 0
        return valref
    }
    
    _JsEval(code)
    {
        e := DllCall(this._dll "\JsRunScript", "wstr", code, "uptr", 0, "wstr", "source.js"
            , "ptr*", &result:=0)
        if e
        {
            if DllCall(this._dll "\JsGetAndClearException", "ptr*", &excp:=0) = 0
                throw this._JsToVt(excp)
            throw Error("JsRT error", -2, format("0x{:X}", e))
        }
        return result
    }
    
    Exec(code)
    {
        this._JsEval(code)
    }
    
    Eval(code)
    {
        return this._JsToVt(this._JsEval(code))
    }
    
    AddObject(name, obj, addMembers := false)
    {
        if addMembers
            throw Error("AddMembers=true is not supported", -1)
        this._dsp.%name% := obj
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
}
