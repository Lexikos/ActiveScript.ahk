# ActiveScript for AutoHotkey v1.1

Provides an interface to Active Scripting languages like VBScript and JScript, without relying on Microsoft's ScriptControl, which is not available to 64-bit programs.

**License:** Use, modify and redistribute without limitation, but at your own risk.

## Usage

Save `ActiveScript.ahk` in a [Lib folder](http://ahkscript.org/docs/Functions.htm#lib).

    #Include <ActiveScript>

    script := new ActiveScript("JScript")
    script := new ActiveScript("VBScript")

More examples are included in the *Example\*.ahk* files.

## Methods

### Eval

Evaluate an expression and return the result.

    Result := script.Eval(Code)

### Exec

Execute script code.

    script.Exec(Code)

### AddObject

Add an object to the global namespace of the script.

    script.AddObject(Name, DispObj, AddMembers := false)

*Name* is required and must be unique.

If *AddMembers* is true, the object's methods and properties will be added to the script's global namespace instead of the object itself. If omitted, it defaults to *false*.

*DispObj* must be either an object which implements the IDispatch interface, passed either via a ComObject wrapper or by address. Can be an AutoHotkey object if running on AutoHotkey v1.1.17 or later.   

### Anything else

To call functions or retrieve or set variables defined in the script,  use normal object notation on the ActiveScript object.  For example:

    result := script.MyFunc()
    value := script.globalvar
    script.globalvar := value

New VBScript variables cannot be created this way. New JScript variables can be created this way only on AutoHotkey v1.1.18 and later.

New variables can be created by declaring them in script with Exec() or Eval().