RunNationalGeneralProspectFillFromClipboard() {
    global PROSPECT_TOOLTIP_DELAY

    raw := Trim(A_Clipboard)
    if (raw = "") {
        MsgBox("Clipboard empty. Copy the raw lead first.")
        return
    }

    fields := NormalizeProspectInput(raw)

    if !FocusWorkBrowser() {
        MsgBox("Browser not found. Open National General first.")
        return
    }

    BeginAutomationRun()
    ToolTip("Click FIRST NAME now (2s)")
    Sleep PROSPECT_TOOLTIP_DELAY
    ToolTip()

    FillNationalGeneralForm(fields)
}

RunNewProspectFillFromClipboard() {
    global PROSPECT_TOOLTIP_DELAY

    raw := Trim(A_Clipboard)
    if (raw = "") {
        MsgBox("Clipboard empty. Copy the raw lead or FORMMAP block first.")
        return
    }

    fields := NormalizeProspectInput(raw)

    if !FocusEdge() {
        MsgBox("Microsoft Edge not found. Open the target prospect page first.")
        return
    }

    BeginAutomationRun()
    ToolTip("Click FIRST NAME now (2s)")
    Sleep PROSPECT_TOOLTIP_DELAY
    ToolTip()

    FillNewProspectForm(fields)

    email := Trim(fields["EMAIL"])
    if IsEmailToken(email) && !SetClip(email)
        MsgBox("Failed to copy the lead email to the clipboard.")
}
