#Include JsRT.ahk
/* Preferred usage:
      ; Put JsRT.ahk in a Lib folder, then:
      #Include <JsRT>
 */

if MsgBox("Use Edge version of JsRT?",, "y/n") = "yes"
    js := (JsRT.Edge)()
else
    js := (JsRT.IE)()

js.Exec '
(
    function ToJS(v) {
        MsgBox("JScript says foo is " + v.foo);
        ToAHK(v);
        return v;
    }
)'

; Add functions callable by name:
js.AddObject "MsgBox", MyMsgBox
js.AddObject "ToAHK", ToAHK

; Pass an AutoHotkey object to a JScript function:
theirObj := js.ToJS(myObj := {foo: "bar"})

; ...and check the object it returned to us.
MsgBox "ToJS returned " (myObj=theirObj ? "the original":"a different") " object and foo is " theirObj.foo


ToAHK(v) {
    MsgBox "ToAHK got " (myObj=v ? "the original":"a different") " object and foo is " v.foo, "ToAHK"
}

MyMsgBox(s) {
    MsgBox s, "JsRT"
}
