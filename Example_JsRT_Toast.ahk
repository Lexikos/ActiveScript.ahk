#NoEnv
#Include ActiveScript.ahk
#Include JsRT.ahk

; Toast notifications from desktop apps can only use local image files.
if !FileExist("sample.png")
    URLDownloadToFile https://autohotkey.com/boards/styles/simplicity/theme/images/announce_unread.png
        , % A_ScriptDir "\sample.png"
; The templates are described here:
;  http://msdn.com/library/windows/apps/windows.ui.notifications.toasttemplatetype.aspx
toast_template := "toastImageAndText02"
; Image path/URL must be absolute, not relative.
toast_image := A_ScriptDir "\sample.png"
; Text is an array because some templates have multiple text elements.
toast_text := ["Hello, world!", "This is the sub-text."]

; Only the Edge version of JsRT supports WinRT.
js := new JsRT.Edge
js.AddObject("alert", Func("alert"))
alert(s) {
    MsgBox % s
}

; Enable use of WinRT.  "Windows.UI" or "Windows" would also work.
js.ProjectWinRTNamespace("Windows.UI.Notifications")
code =
(
    function toast(template, image, text, app) {
        // Alias for convenience.
        var N = Windows.UI.Notifications;
        // Get the template XML as an XmlDocument.
        var toastXml = N.ToastNotificationManager
            .getTemplateContent(N.ToastTemplateType[template]);
        // Insert our content.
        var i = 0;
        for (let el of toastXml.getElementsByTagName("text")) {
            if (typeof text == 'string') {
                el.innerText = text;
                break;
            }
            el.innerText = text[++i];
        }
        toastXml.getElementsByTagName("image")[0]
            .setAttribute("src", image);
        // Show the notification.
        var toastNotifier = N.ToastNotificationManager
            .createToastNotifier(app || "AutoHotkey");
        toastNotifier.show(
            new N.ToastNotification(toastXml));
    }
)
try {
    ; Define the toast function.
    js.Exec(code)
    ; Show a toast notification.
    js.toast(toast_template, toast_image, toast_text)
    ; Wait (optional).
    Sleep 5000
}
catch ex {
    try MsgBox % ex.stack
    catch
        MsgBox % ex.message
}
ExitApp