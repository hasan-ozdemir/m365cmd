# Core: Mail
# Purpose: Mail shared utilities.
function New-ContentBody {
    param(
        [string]$Content,
        [string]$ContentType
    )
    if (-not $ContentType) { $ContentType = "Text" }
    return @{
        contentType = $ContentType
        content     = $Content
    }
}



function Build-RecipientList {
    param([string]$Value)
    $list = @()
    foreach ($addr in (Parse-CommaList $Value)) {
        $list += @{ emailAddress = @{ address = $addr } }
    }
    return $list
}



function Resolve-CalendarRange {
    param(
        [string]$DateValue,
        [string]$Range
    )
    $date = if ($DateValue) { [datetime]::Parse($DateValue) } else { Get-Date }
    $r = if ($Range) { $Range.ToLowerInvariant() } else { "day" }
    switch ($r) {
        "week" {
            $start = $date.Date
            $end = $start.AddDays(7)
        }
        "month" {
            $start = New-Object datetime($date.Year, $date.Month, 1)
            $end = $start.AddMonths(1)
        }
        "year" {
            $start = New-Object datetime($date.Year, 1, 1)
            $end = $start.AddYears(1)
        }
        default {
            $start = $date.Date
            $end = $start.AddDays(1)
        }
    }
    return [pscustomobject]@{
        Start = $start
        End   = $end
    }
}



function Build-MailMessage {
    param([hashtable]$Map)
    $message = @{}
    $subject = Get-ArgValue $Map "subject"
    $bodyText = Get-ArgValue $Map "body"
    $contentType = Get-ArgValue $Map "contentType"
    $to = Build-RecipientList (Get-ArgValue $Map "to")
    $cc = Build-RecipientList (Get-ArgValue $Map "cc")
    $bcc = Build-RecipientList (Get-ArgValue $Map "bcc")
    if ($subject) { $message.subject = $subject }
    if ($bodyText) { $message.body = New-ContentBody $bodyText $contentType }
    if ($to.Count -gt 0) { $message.toRecipients = $to }
    if ($cc.Count -gt 0) { $message.ccRecipients = $cc }
    if ($bcc.Count -gt 0) { $message.bccRecipients = $bcc }
    return $message
}



