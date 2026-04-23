CsvEscape(value) {
    text := value ?? ""
    text := StrReplace(text, '"', '""')
    return '"' . text . '"'
}

EnsureBatchLogHeader() {
    global batchLogFile
    if !FileExist(batchLogFile)
        FileAppend("Timestamp,LeadName,Phone,CarCount,Status`n", batchLogFile, "UTF-8")
}

AppendBatchLog(lead, status) {
    global batchLogFile

    timestamp := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
    line := CsvEscape(timestamp) ","
        . CsvEscape(lead["FULL_NAME"]) ","
        . CsvEscape(lead["PHONE"]) ","
        . CsvEscape(lead["VEHICLE_COUNT"]) ","
        . CsvEscape(status) "`n"

    FileAppend(line, batchLogFile, "UTF-8")
}

BatchResumeModeToken(mode) {
    return (mode = "fast") ? "N" : "B"
}

BatchRunTitle(mode, pausedAt := 0) {
    if (mode = "fast")
        return pausedAt ? "Batch Run Log (Fast Paused)" : "Batch Run Log (Fast)"
    return pausedAt ? "Batch Run Log (Paused)" : "Batch Run Log"
}

RunBatchFromClipboard(mode := "stable") {
    global batchLeadHolder
    global batchLogFile
    global batchResumeIndex, batchResumeMode, batchResumeRaw

    raw := Trim(A_Clipboard)
    if (raw = "") {
        MsgBox("Clipboard empty. Copy the batch first.")
        return
    }

    batchLeadHolder := BuildBatchLeadHolder(raw)
    if (batchLeadHolder.Length = 0) {
        MsgBox("No lead rows detected in clipboard.")
        return
    }

    if !FocusWorkBrowser() {
        MsgBox("Browser not detected. Open the CRM page first.")
        return
    }

    BeginAutomationRun()
    if !SafeSleep(300)
        return

    log := []
    okCount := 0
    failCount := 0
    pausedAt := 0
    stopped := false

    try {
        EnsureBatchLogHeader()
    } catch Error as err {
        MsgBox("Cannot write to batch log file:`n" batchLogFile "`n`nClose Excel or any app using it, then try again.`n`nDetails: " err.Message)
        return
    }

    modeToken := BatchResumeModeToken(mode)
    startIndex := 1
    if (batchResumeMode = modeToken && batchResumeRaw = raw && batchResumeIndex >= 1 && batchResumeIndex <= batchLeadHolder.Length)
        startIndex := batchResumeIndex
    else
        ClearBatchResumeState("batch-resume-reset")

    Loop batchLeadHolder.Length - startIndex + 1 {
        if StopRequested() {
            stopped := true
            break
        }

        i := startIndex + A_Index - 1
        lead := batchLeadHolder[i]
        status := RunBatchLeadFlow(lead, mode)

        try {
            AppendBatchLog(lead, status)
        } catch Error as err {
            MsgBox("Batch ran, but logging failed for:`n" lead["FULL_NAME"] "`n`nClose the CSV if it's open.`n`nDetails: " err.Message)
            return
        }

        log.Push(i . ". " . lead["FULL_NAME"] . " -> " . status)

        if (status = "OK")
            okCount += 1
        else
            failCount += 1

        if (status = "STOPPED") {
            stopped := true
            break
        }

        if (SubStr(status, 1, 8) = "PAUSED -") {
            pausedAt := i
            if (i < batchLeadHolder.Length)
                SetBatchResumeState(i + 1, modeToken, raw)
            else
                ClearBatchResumeState("batch-resume-finished")
            break
        }
    }

    if (!pausedAt && !stopped)
        ClearBatchResumeState("batch-complete")

    msg := stopped
        ? "Batch stopped by ESC.`n`n"
        : pausedAt
        ? "Batch paused at lead " pausedAt ".`nRe-run " (mode = "fast" ? "Ctrl+Alt+N" : "Ctrl+Alt+B") " to continue with the next lead.`n`n"
        : (mode = "fast" ? "Batch complete (FAST).`n`n" : "Batch complete.`n`n")

    msg .= "Success: " okCount "`n"
        . "Failed/Skipped: " failCount "`n`n"

    for _, line in log
        msg .= line "`n"

    MsgBox(msg, BatchRunTitle(mode, pausedAt))
}

RunBatchLeadFlow(lead, mode := "stable") {
    global BATCH_AFTER_SCHEDULE, BATCH_AFTER_TAG_COMPLETE

    if StopRequested()
        return "STOPPED"
    if !FocusWorkBrowser()
        return "FAILED - Browser lost focus"

    if !SafeSleep(150)
        return "STOPPED"

    if (lead["PHONE"] = "" || lead["FULL_NAME"] = "")
        return "SKIPPED - Missing phone or name"

    offset := NextRotationOffset()

    if !QuoPrimeNewConversation(lead["PHONE"], mode)
        return StopRequested() ? "STOPPED" : "FAILED - Could not open a new Quo conversation"

    if !ScheduleBuilderForLead(lead, offset, mode)
        return StopRequested() ? "STOPPED" : "FAILED - Builder scheduling failed"
    if !SafeSleep(BATCH_AFTER_SCHEDULE)
        return "STOPPED"

    if !ScheduleRegularFollowupsForLead(lead, offset, mode)
        return StopRequested() ? "STOPPED" : "FAILED - Follow-up scheduling failed"
    if !SafeSleep(BATCH_AFTER_SCHEDULE)
        return "STOPPED"

    if !QuoSelectLeadHolder(lead["HOLDER_NAME"])
        return StopRequested() ? "STOPPED" : "FAILED - Could not select lead holder"

    if !SafeSleep(BATCH_AFTER_TAG_COMPLETE)
        return "STOPPED"

    tagStatus := ApplyQuoTag(lead["TAG_VALUE"])
    if (tagStatus != "OK")
        return tagStatus

    return "OK"
}

TraceBatchLeadPlan(lead, mode := "stable") {
    builder := GetInitialQuoteDateTime(0)
    followups := BuildFollowupQueue(lead["FULL_NAME"], 0)
    steps := [
        "Focus browser",
        "Open new Quo conversation",
        "Paste participant phone",
        "Move from participant field to composer",
        "Schedule builder quote (" . mode . ")",
        "Schedule " . followups.Length . " follow-up messages (" . mode . ")",
        "Select lead holder",
        "Apply tag " . lead["TAG_VALUE"]
    ]

    return Map(
        "lead", lead["FULL_NAME"],
        "phone", lead["PHONE"],
        "mode", mode,
        "builderDate", builder["date"],
        "builderTime", builder["time"],
        "followupCount", followups.Length,
        "steps", steps
    )
}
