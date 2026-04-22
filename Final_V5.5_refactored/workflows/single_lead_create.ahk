RunQuickLeadCreateAndTag(lead) {
    if StopRequested()
        return "STOPPED"
    if !FocusWorkBrowser()
        return "FAILED - Browser lost focus"

    if !SafeSleep(150)
        return "STOPPED"

    if (lead["PHONE"] = "" || lead["FULL_NAME"] = "")
        return "FAILED - Missing phone or name"

    if !QuoPrimeNewConversation(lead["PHONE"], "fast")
        return StopRequested() ? "STOPPED" : "FAILED - Could not open a new Quo conversation"

    if !SafeSleep(250)
        return "STOPPED"

    if !QuoSelectLeadHolder(lead["HOLDER_NAME"])
        return StopRequested() ? "STOPPED" : "FAILED - Could not select lead holder"

    tagStatus := ApplyQuoTag(lead["TAG_VALUE"])
    if (tagStatus != "OK")
        return tagStatus

    return "OK"
}
