# ActiveScript for AutoHotkey v2.0-beta.1

Scripts for hosting other scripting languages; specifically:
  - Active Scripting languages such as VBScript and JScript (without relying on Microsoft's ScriptControl, which is not available to 64-bit programs).
  - JavaScript as implemented in IE11 and Edge.

This branch contains scripts which are intended to mimic Microsoft's ScriptControl. As such, the feature set might be more limited than what the underlying API actually allows. An alternative wrapper for JsRT can be found in  [AutoHotkey-jk](https://github.com/Lexikos/AutoHotkey-jk).

**License:** Use, modify and redistribute without limitation, but at your own risk.

This branch is compatible with AutoHotkey v2.0-beta.1 and probably later versions. For AutoHotkey v1.1, get the `for-v1` branch.

## Usage

Save `ActiveScript.ahk` and `JsRT.ahk` (if needed) in a [Lib folder](http://ahkscript.org/docs/Functions.htm#lib).

### ActiveScript

Supports JScript, VBScript and possibly other scripting engines which are registered with COM and implement the IActiveScript interface.

```AutoHotkey
#Include <ActiveScript>

script := ActiveScript("JScript")
script := ActiveScript("VBScript")
```

More examples are included in the Example\*.ahk files.

### JsRT

Supports JavaScript as implemented in IE11 or Edge (Windows 10).

```AutoHotkey
#Include <JsRT>

script := JsRT.IE()  ; IE11 feature set.
script := JsRT.Edge()  ; Edge feature set.
```

Use either `IE` or `Edge`. Loading both runtimes into the same process is unsupported by Microsoft, and attempting it generally causes the process to crash. WebBrowser ActiveX controls and MSHTML use the IE runtime, and therefore must not be used in the same process as `JsRT.Edge()`.

This version of the library is self-contained within JsRT.ahk; it does not require ActiveScript.ahk.

More examples are included in Example\_JsRT.ahk.

## Methods

### Eval

Evaluate an expression and return the result.

```AutoHotkey
Result := script.Eval(Code)
```

### Exec

Execute script code.

```AutoHotkey
script.Exec(Code)
```

### AddObject

Add an object to the global namespace of the script.

```AutoHotkey
script.AddObject(Name, DispObj, AddMembers := false)
```

*Name* is required and must be unique.

If *AddMembers* is true, the object's methods and properties will be added to the script's global namespace instead of the object itself. If omitted, it defaults to *false*.

*DispObj* must be either an object or an interface pointer for an object which implements IDispatch. Can be an AutoHotkey object (reference or address). If it is a ComObject, the interface pointer it contains is used.

Evaluating code with *Eval* or *Exec* may also add global variables and functions.

`script[Name] := DispObj` will usually have the same effect if *AddMembers* is false or omitted.

**JsRT:** *AddMembers* must be false. *DispObj* can be any value, and will be added as is. Do not pass a pointer, or it will be added as a number.


### ProjectWinRTNamespace

**JsRT.Edge only:** "Project" a Windows Runtime (WinRT) namespace -- make it accessible to JavaScript.

```AutoHotkey
script.ProjectWinRTNamespace(Namespace)
```

For example, the following is sufficient to make most of the WinRT available to the script:

```AutoHotkey
script.ProjectWinRTNamespace("Windows")
```


### Anything else

To call functions or retrieve or set variables defined in the script,  use normal object notation on the ActiveScript object.  For example:

```AutoHotkey
result := script.MyFunc()
value := script.globalvar
script.globalvar := value
```

New VBScript variables cannot be created this way. New JScript variables can be created this way only on AutoHotkey v1.1.18 and later.

New variables can be created by declaring them in script with Exec() or Eval().

The hosted script can be given access to AutoHotkey functions by assigning them to global variables:

```AutoHotkey
script.alert := alert
alert(message) {
	MsgBox message, "Message from script", "Icon!"
}
```


## Error Handling

AutoHotkey has very limited support for propagating exceptions thrown by a JavaScript method; generally the value/object which was thrown is not accessible.  The message format is slightly different depending on whether the method was called directly or via Eval/Exec.

**JsRT:** Exceptions thrown by JavaScript can be caught by AutoHotkey script if the JavaScript was called via Eval/Exec.  Try..catch can also be used to handle compiler/syntax errors.  However, since AutoHotkey doesn't understand JavaScript Error objects, it will display a generic error message if the exception isn't handled.  If a string is thrown from JavaScript, it will be shown as the error message.