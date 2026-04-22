#Requires AutoHotkey v2.0
#SingleInstance Force
SendMode "Input"

global projectRoot := A_ScriptDir
global configRoot := projectRoot "\config"
global logsRoot := projectRoot "\logs"
global assetsRoot := projectRoot "\assets"
global settingsFile := configRoot "\settings.ini"
global timingsFile := configRoot "\timings.ini"
global templatesFile := configRoot "\templates.ini"
global holidaysFile := configRoot "\holidays_2026.ini"
global batchLogFile := logsRoot "\batch_lead_log.csv"
global runStateFile := logsRoot "\run_state.json"
global tagSelectorJsFile := assetsRoot "\js\tag_selector.js"
global participantInputJsFile := assetsRoot "\js\participant_input_focus.js"
global legacyMonolithFile := projectRoot "\..\Final_V5.5.ahk"

global holidays := []
global batchLeadHolder := []
global running := false
global StopFlag := false
global batchResumeIndex := 1
global batchResumeMode := ""
global batchResumeRaw := ""

InitializeApplication()
PersistRunState("startup")

#Include domain\lead_normalizer.ahk
#Include domain\lead_parser.ahk
#Include domain\pricing_rules.ahk
#Include domain\date_rules.ahk
#Include domain\batch_rules.ahk
#Include domain\message_templates.ahk

#Include adapters\clipboard_adapter.ahk
#Include adapters\browser_focus_adapter.ahk
#Include adapters\devtools_bridge.ahk
#Include adapters\quo_adapter.ahk
#Include adapters\crm_adapter.ahk
#Include adapters\tag_selector_adapter.ahk

#Include workflows\single_lead_create.ahk
#Include workflows\batch_run.ahk
#Include workflows\message_schedule.ahk
#Include workflows\prospect_fill.ahk
#Include workflows\crm_activity.ahk

#Include hotkeys\lead_hotkeys.ahk
#Include hotkeys\schedule_hotkeys.ahk
#Include hotkeys\crm_hotkeys.ahk
#Include hotkeys\debug_hotkeys.ahk

Return

InitializeApplication() {
    global settingsFile, timingsFile, holidaysFile
    global holidays
    global agentName, agentEmail, tagSymbol, configDays
    global batchMinVehicles, batchMaxVehicles, dobDefaultDay, batchTabsChatToName
    global priceOldCar, priceOneCar, priceOneCar2020Plus, priceTwoCars
    global priceThreeCars, priceFourCars, priceFiveCars, singleCarModernYearCutoff
    global SLOW_ACTIVATE_DELAY, SLOW_AFTER_MSG, SLOW_AFTER_SCHED, SLOW_AFTER_DT_PASTE, SLOW_AFTER_ENTER
    global BATCH_AFTER_ALTN, BATCH_AFTER_PHONE, BATCH_AFTER_TAB, BATCH_AFTER_SCHEDULE
    global BATCH_AFTER_ENTER, BATCH_AFTER_NAME_PICK, BATCH_AFTER_TAG_PICK
    global BATCH_BEFORE_TAG_PASTE, BATCH_AFTER_TAG_PASTE, BATCH_POST_PARTICIPANT_READY_STABLE
    global BATCH_POST_PARTICIPANT_READY_FAST, BATCH_AFTER_PARTICIPANT_TO_COMPOSER, BATCH_AFTER_TAG_COMPLETE
    global PROSPECT_TOOLTIP_DELAY, FORM_FIELD_DELAY, FORM_TAB_DELAY, FORM_PASTE_DELAY
    global FORM_PASTE_TAB_DELAY, FORM_CITY_TAB_DELAY
    global CRM_ACTION_FOCUS_DELAY, CRM_KEYSTEP_DELAY, CRM_SHORT_DELAY, CRM_MEDIUM_DELAY
    global CRM_QUOTE_SHIFT_TAB_DELAY, CRM_SAVE_HISTORY_DELAY, CRM_ADD_APPOINTMENT_DELAY
    global CRM_FOCUS_DATE_DELAY, CRM_FINAL_SAVE_DELAY

    agentName := IniRead(settingsFile, "Agent", "Name", "Pablo Cabrera")
    agentEmail := IniRead(settingsFile, "Agent", "Email", "pablocabrera@allstate.com")
    tagSymbol := Trim(IniRead(settingsFile, "Agent", "TagSymbol", "+"))
    if (tagSymbol = "")
        tagSymbol := "+"

    daysStr := IniRead(settingsFile, "Schedule", "Days", "1,2,4,5")
    configDays := ParseDays(daysStr)
    if (configDays.Length != 4) {
        configDays := [1, 2, 4, 5]
        IniWrite("1,2,4,5", settingsFile, "Schedule", "Days")
    }

    batchMinVehicles := Integer(IniRead(settingsFile, "Batch", "MinVehicles", "0"))
    batchMaxVehicles := Integer(IniRead(settingsFile, "Batch", "MaxVehicles", "99"))
    dobDefaultDay := Integer(IniRead(settingsFile, "DOB", "DefaultDay", "16"))
    batchTabsChatToName := Integer(IniRead(settingsFile, "UI", "BatchTabsChatToName", "7"))

    priceOldCar := Integer(IniRead(settingsFile, "Pricing", "OldCar", "98"))
    priceOneCar := Integer(IniRead(settingsFile, "Pricing", "OneCar", "117"))
    priceOneCar2020Plus := Integer(IniRead(settingsFile, "Pricing", "OneCar2020Plus", "167"))
    priceTwoCars := Integer(IniRead(settingsFile, "Pricing", "TwoCars", "176"))
    priceThreeCars := Integer(IniRead(settingsFile, "Pricing", "ThreeCars", "284"))
    priceFourCars := Integer(IniRead(settingsFile, "Pricing", "FourCars", "397"))
    priceFiveCars := Integer(IniRead(settingsFile, "Pricing", "FiveCars", "397"))
    singleCarModernYearCutoff := Integer(IniRead(settingsFile, "Pricing", "SingleCarModernYearCutoff", "2020"))

    SLOW_ACTIVATE_DELAY := Integer(IniRead(timingsFile, "SlowMode", "ActivateDelay", "250"))
    SLOW_AFTER_MSG := Integer(IniRead(timingsFile, "SlowMode", "AfterMessage", "550"))
    SLOW_AFTER_SCHED := Integer(IniRead(timingsFile, "SlowMode", "AfterSchedule", "650"))
    SLOW_AFTER_DT_PASTE := Integer(IniRead(timingsFile, "SlowMode", "AfterDatePaste", "650"))
    SLOW_AFTER_ENTER := Integer(IniRead(timingsFile, "SlowMode", "AfterEnter", "300"))

    BATCH_AFTER_ALTN := Integer(IniRead(timingsFile, "Batch", "AfterAltN", "5000"))
    BATCH_AFTER_PHONE := Integer(IniRead(timingsFile, "Batch", "AfterPhone", "650"))
    BATCH_AFTER_TAB := Integer(IniRead(timingsFile, "Batch", "AfterTab", "150"))
    BATCH_AFTER_SCHEDULE := Integer(IniRead(timingsFile, "Batch", "AfterSchedule", "600"))
    BATCH_AFTER_ENTER := Integer(IniRead(timingsFile, "Batch", "AfterEnter", "150"))
    BATCH_AFTER_NAME_PICK := Integer(IniRead(timingsFile, "Batch", "AfterNamePick", "250"))
    BATCH_AFTER_TAG_PICK := Integer(IniRead(timingsFile, "Batch", "AfterTagPick", "250"))
    BATCH_BEFORE_TAG_PASTE := Integer(IniRead(timingsFile, "Batch", "BeforeTagPaste", "500"))
    BATCH_AFTER_TAG_PASTE := Integer(IniRead(timingsFile, "Batch", "AfterTagPaste", "700"))
    BATCH_POST_PARTICIPANT_READY_STABLE := Integer(IniRead(timingsFile, "Batch", "PostParticipantReadyStable", "150"))
    BATCH_POST_PARTICIPANT_READY_FAST := Integer(IniRead(timingsFile, "Batch", "PostParticipantReadyFast", "0"))
    BATCH_AFTER_PARTICIPANT_TO_COMPOSER := Integer(IniRead(timingsFile, "Batch", "AfterParticipantToComposer", "1000"))
    BATCH_AFTER_TAG_COMPLETE := Integer(IniRead(timingsFile, "Batch", "AfterTagComplete", "300"))

    PROSPECT_TOOLTIP_DELAY := Integer(IniRead(timingsFile, "ProspectFill", "TooltipLeadIn", "500"))
    FORM_FIELD_DELAY := Integer(IniRead(timingsFile, "ProspectFill", "FieldDelay", "30"))
    FORM_TAB_DELAY := Integer(IniRead(timingsFile, "ProspectFill", "TabDelay", "30"))
    FORM_PASTE_DELAY := Integer(IniRead(timingsFile, "ProspectFill", "PasteDelay", "120"))
    FORM_PASTE_TAB_DELAY := Integer(IniRead(timingsFile, "ProspectFill", "PasteTabDelay", "80"))
    FORM_CITY_TAB_DELAY := Integer(IniRead(timingsFile, "ProspectFill", "CityTabDelay", "50"))

    CRM_ACTION_FOCUS_DELAY := Integer(IniRead(timingsFile, "CrmActivity", "ActionFocusDelay", "500"))
    CRM_KEYSTEP_DELAY := Integer(IniRead(timingsFile, "CrmActivity", "KeyStepDelay", "150"))
    CRM_SHORT_DELAY := Integer(IniRead(timingsFile, "CrmActivity", "ShortDelay", "200"))
    CRM_MEDIUM_DELAY := Integer(IniRead(timingsFile, "CrmActivity", "MediumDelay", "250"))
    CRM_QUOTE_SHIFT_TAB_DELAY := Integer(IniRead(timingsFile, "CrmActivity", "QuoteShiftTabDelay", "3050"))
    CRM_SAVE_HISTORY_DELAY := Integer(IniRead(timingsFile, "CrmActivity", "SaveHistoryDelay", "800"))
    CRM_ADD_APPOINTMENT_DELAY := Integer(IniRead(timingsFile, "CrmActivity", "AddAppointmentDelay", "800"))
    CRM_FOCUS_DATE_DELAY := Integer(IniRead(timingsFile, "CrmActivity", "FocusDateDelay", "300"))
    CRM_FINAL_SAVE_DELAY := Integer(IniRead(timingsFile, "CrmActivity", "FinalSaveDelay", "400"))

    holidays := LoadHolidayList(holidaysFile)
    EnsureBatchLogHeader()
}

UpdateAgentConfiguration(newName, newEmail, newTagSymbol) {
    global settingsFile
    global agentName, agentEmail, tagSymbol

    IniWrite(newName, settingsFile, "Agent", "Name")
    IniWrite(newEmail, settingsFile, "Agent", "Email")
    IniWrite(newTagSymbol, settingsFile, "Agent", "TagSymbol")

    agentName := newName
    agentEmail := newEmail
    tagSymbol := newTagSymbol
    PersistRunState("agent-config-updated")
}

ResetRotationOffset() {
    global settingsFile
    IniWrite(0, settingsFile, "Times", "Offset")
    PersistRunState("rotation-reset")
}

SetBatchResumeState(nextIndex, mode, raw) {
    global batchResumeIndex, batchResumeMode, batchResumeRaw

    batchResumeIndex := nextIndex
    batchResumeMode := mode
    batchResumeRaw := raw
    PersistRunState("batch-resume-set")
}

ClearBatchResumeState(lastAction := "batch-resume-cleared") {
    global batchResumeIndex, batchResumeMode, batchResumeRaw

    batchResumeIndex := 1
    batchResumeMode := ""
    batchResumeRaw := ""
    PersistRunState(lastAction)
}

PersistRunState(lastAction := "") {
    global settingsFile, runStateFile
    global batchResumeIndex, batchResumeMode, batchResumeRaw
    global running, StopFlag

    offset := Integer(IniRead(settingsFile, "Times", "Offset", "0"))
    updatedAt := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
    json := "{`n"
        . '  "offset": ' offset ",`n"
        . '  "batchResumeIndex": ' Integer(batchResumeIndex) ",`n"
        . '  "batchResumeMode": "' JsonEscape(batchResumeMode) '",`n'
        . '  "batchResumeRawLength": ' StrLen(batchResumeRaw) ",`n"
        . '  "running": ' (running ? "true" : "false") ",`n"
        . '  "stopFlag": ' (StopFlag ? "true" : "false") ",`n"
        . '  "lastAction": "' JsonEscape(lastAction) '",`n'
        . '  "updatedAt": "' JsonEscape(updatedAt) '"`n'
        . "}`n"

    WriteUtf8File(runStateFile, json)
}

JsonEscape(value) {
    text := value ?? ""
    text := StrReplace(text, "\", "\\")
    text := StrReplace(text, '"', '\"')
    text := StrReplace(text, "`r", "\r")
    text := StrReplace(text, "`n", "\n")
    return text
}

WriteUtf8File(path, text) {
    try FileDelete(path)
    FileAppend(text, path, "UTF-8")
}
