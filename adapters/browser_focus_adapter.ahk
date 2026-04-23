FocusEdge() {
    if WinExist("ahk_exe msedge.exe") {
        WinActivate
        WinWaitActive "ahk_exe msedge.exe",, 2
        return true
    }
    return false
}

FocusChrome() {
    if WinExist("ahk_exe chrome.exe") {
        WinActivate
        WinWaitActive "ahk_exe chrome.exe",, 2
        return true
    }
    return false
}

FocusWorkBrowser() {
    if FocusChrome()
        return true
    if FocusEdge()
        return true
    return false
}
