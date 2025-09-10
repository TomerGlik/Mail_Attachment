# ================== CONFIG ==================
$SavePath   = "C:\Path\To\Save"     # path to save attachments
$Mailbox    = "The mail box"                 # the mailbox name as shown in Outlook
$SenderNeedle = "Sender_Email"     # string to match in sender / on behalf of
$StateFile = Join-Path $SavePath "processed_entryids.txt"   # text file to track processed emails
$LogFile   = Join-Path $SavePath "download_log.txt"         # log file path
$MaxItemsToScan = 800               # limit to avoid scanning too many items
# ============================================

New-Item -ItemType Directory -Force -Path $SavePath | Out-Null
if (-not (Test-Path $StateFile)) { New-Item -ItemType File -Path $StateFile | Out-Null }
$Processed = Get-Content $StateFile -ErrorAction SilentlyContinue | Where-Object { $_ } | Select-Object -Unique

function Write-Log($m){ "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $m" | Tee-Object -FilePath $LogFile -Append }

# MAPI properties for sender details
$PR_SENDER_NAME                        = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F"
$PR_SENDER_EMAIL_ADDRESS               = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F"
$PR_SENT_REPRESENTING_NAME             = "http://schemas.microsoft.com/mapi/proptag/0x0042001F"
$PR_SENT_REPRESENTING_EMAIL_ADDRESS    = "http://schemas.microsoft.com/mapi/proptag/0x0065001F"


function Matches-Sender($mail, $needle){
    $pa = $mail.PropertyAccessor
    $vals = @(
        [string]$mail.SentOnBehalfOfName,
        [string]$mail.SenderName,
        [string]$mail.SenderEmailAddress,
        $(try { [string]$pa.GetProperty($PR_SENDER_NAME) } catch { "" }),
        $(try { [string]$pa.GetProperty($PR_SENDER_EMAIL_ADDRESS) } catch { "" }),
        $(try { [string]$pa.GetProperty($PR_SENT_REPRESENTING_NAME) } catch { "" }),
        $(try { [string]$pa.GetProperty($PR_SENT_REPRESENTING_EMAIL_ADDRESS) } catch { "" })
    ) | Where-Object { $_ -and $_.Trim() -ne "" }

    $needleLC = $needle.ToLowerInvariant()
    return $vals | ForEach-Object { $_.ToLowerInvariant() } | Where-Object { $_ -like "*$needleLC*" } | Measure-Object | Select-Object -ExpandProperty Count
}
# Main
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")

    $store = $ns.Stores | Where-Object { $_.DisplayName -ieq $Mailbox }
    if (-not $store) { throw "Mailbox '$Mailbox' not found. Stores: " + (($ns.Stores | Select-Object -Expand DisplayName) -join ', ') }

    $root  = $store.GetRootFolder()
    $inbox = $root.Folders.Item("Inbox")
    if (-not $inbox) { throw "Inbox not found under '$Mailbox'." }

    Write-Log "Using Inbox: $($inbox.FolderPath)"

    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)

    $today = (Get-Date).Date
    $saved = 0
    $limit = [Math]::Min($MaxItemsToScan, $items.Count)

    for ($i = 1; $i -le $limit; $i++) {
        $m = $items.Item($i)
        if (-not $m) { continue }
        if ($m.Class -ne 43) { continue }  # רק MailItem

        # only mails from today
        if ($m.ReceivedTime.Date -ne $today) { continue }

        # Sender check
        if ((Matches-Sender $m $SenderNeedle) -eq 0) { continue }

        # skip if already processed
        if ($Processed -contains $m.EntryID) {
            Write-Log "Skip (already processed): $($m.Subject)"
            continue
        }
        # mark as processed now to avoid re processing if script fails later
        if ($m.Attachments.Count -gt 0) {
            for ($a=1; $a -le $m.Attachments.Count; $a++) {
                $att = $m.Attachments.Item($a)
                if ($att.FileName -match '\.csv$') {
                    $ts  = Get-Date -Format "yyyyMMdd_HHmmss"
                    $safe = ($att.FileName -replace '[\\/:*?"<>|]', '_')
                    $out  = Join-Path $SavePath ("{0}_{1}" -f $ts, $safe)
                    $att.SaveAsFile($out)
                    Write-Log "Saved CSV: $out (From: $($m.SenderName) | OnBehalfOf: $($m.SentOnBehalfOfName) | Subject: '$($m.Subject)')"
                    $saved++
                }
            }
        } else {
            Write-Log "No attachments in: '$($m.Subject)'"
        }

        Add-Content -Path $StateFile -Value $m.EntryID
    }

    Write-Log "Done. Total saved this run: $saved"
}
catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    throw
}
