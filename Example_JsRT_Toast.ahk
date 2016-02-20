#NoEnv
#Include ActiveScript.ahk
#Include JsRT.ahk

; Unlike TrayTip, this API does not require a tray icon:
#NoTrayIcon

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
js.AddObject("yesno", Func("yesno"))
yesno(s) {
    MsgBox 4,, %s%
    IfMsgBox Yes
        return true
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
        var notification = new N.ToastNotification(toastXml);
        toastNotifier.show(notification);
        // Unlike TrayTip, this API lets us hide the notification:
        if (yesno("Hide the notification?")) {
            toastNotifier.hide(notification);
        }
    }
)
try {
    ; Define the toast function.
    js.Exec(code)
    ; Show a toast notification.
    js.toast(toast_template, toast_image, toast_text)
}
catch ex {
    try errmsg := ex.stack
    if !errmsg
        errmsg := "Error: " ex.message
    MsgBox % errmsg
}
; Note: If the notification wasn't hidden, it will remain after we exit.
ExitApp