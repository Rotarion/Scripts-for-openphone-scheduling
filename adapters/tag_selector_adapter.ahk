QuoBuildTagSelectorJS() {
    global tagSelectorJsFile
    if !FileExist(tagSelectorJsFile) {
        MsgBox("Missing file: " tagSelectorJsFile "`n`nPlace 'tag_selector.js' under assets\js.")
        return ""
    }

    js := Trim(FileRead(tagSelectorJsFile, "UTF-8"))
    if (js = "")
        return ""

    if RegExMatch(js, "i)^\s*copy\s*\(")
        return js

    return "copy(String(" . js . "))"
}

RunQuoTagSelector() {
    js := QuoBuildTagSelectorJS()
    if (js = "")
        return ""
    return RunDevToolsJSGetResult(js)
}

QuoDeleteAddTagConfirm(tagText) {
    global BATCH_BEFORE_TAG_PASTE, BATCH_AFTER_TAG_PASTE, BATCH_AFTER_ENTER

    if StopRequested()
        return false
    if !FocusWorkBrowser()
        return false

    if !SafeSleep(120)
        return false
    if StopRequested()
        return false
    Send "{Backspace}"
    if !SafeSleep(120)
        return false

    if !SafeSleep(BATCH_BEFORE_TAG_PASTE)
        return false
    if !PasteValueRaw(tagText)
        return false
    if !SafeSleep(BATCH_AFTER_TAG_PASTE)
        return false

    if StopRequested()
        return false
    Send "{Enter}"
    if !SafeSleep(BATCH_AFTER_ENTER)
        return false
    if StopRequested()
        return false
    Send "{Enter}"
    if !SafeSleep(BATCH_AFTER_ENTER)
        return false

    FocusSlateComposer()
    if !SafeSleep(200)
        return false

    return true
}

QuoPrepareFieldThenAddTag(tagText) {
    if StopRequested()
        return false
    if !FocusWorkBrowser()
        return false

    if StopRequested()
        return false
    Send "{Tab}"
    if !SafeSleep(220)
        return false
    if StopRequested()
        return false
    Send "{Tab}"
    if !SafeSleep(220)
        return false
    if StopRequested()
        return false
    Send "{Enter}"
    if !SafeSleep(300)
        return false

    return QuoDeleteAddTagConfirm(tagText)
}

HandleQuoTagSelectorResult(result, tagText) {
    if !FocusWorkBrowser()
        return "FAILED - Browser lost focus"

    if (result = "PLUS_FALLBACK") {
        if !QuoDeleteAddTagConfirm(tagText)
            return "FAILED - PLUS_FALLBACK sequence failed"
        return "OK"
    }

    if (result = "HITTEST_TARGET" || result = "STRUCTURE_TARGET" || result = "ANCESTOR_TARGET" || result = "LOCAL_TARGET") {
        if !QuoPrepareFieldThenAddTag(tagText)
            return "FAILED - " result " sequence failed"
        return "OK"
    }

    if (result = "NO_TARGET")
        return "PAUSED - No tag target"

    if (result = "")
        return "FAILED - Blank selector result"

    return "FAILED - Unhandled selector result: " result
}

ApplyQuoTag(tagText) {
    tagText := Trim(tagText)
    if (tagText = "")
        return "FAILED - Blank tag value"
    if StopRequested()
        return "STOPPED"

    result := RunQuoTagSelector()
    if StopRequested()
        return "STOPPED"
    if (result = "")
        return "FAILED - Blank selector result"

    return HandleQuoTagSelectorResult(result, tagText)
}
