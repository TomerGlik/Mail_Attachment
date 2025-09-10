# 📥 Outlook SOC CSV Downloader

This PowerShell script automatically downloads **CSV attachments** from a shared Outlook mailbox.  
It was designed to handle a SOC workflow where daily reports are sent *on behalf of* `NT SOC MOT-IL`.

---

## ✨ Features
- Connects to a **shared mailbox** in Outlook (`SOC` in this example).
- Filters incoming emails by **sender / on behalf of** (e.g. `NT SOC MOT-IL`).
- Downloads **only `.csv` attachments**.
- Restricts downloads to **today’s date only** – ensures you only process fresh reports.
- Keeps a local state file to **avoid duplicates** across runs.
- Logs every run (`download_log.txt`) for auditing.

---

## 📂 File Structure
- `Save-Attachments-BySender.ps1` → Main script
- `processed_entryids.txt` → Stores which emails were already processed
- `download_log.txt` → Log file with full run history
- `C:\Users\<username>\Desktop\Test` → Default output directory (can be changed in `$SavePath`)

---

## ⚙️ How to Use
1. Clone or download this repository.
2. Edit the script variables:
   ```powershell
   $SavePath      = "C:\Path\To\Save"
   $Mailbox       = "SOC"                # Shared mailbox display name
   $SenderNeedle  = "NT SOC MOT-IL"      # Sender / on behalf of filter
  '''
3. Open PowerShell and run:
          Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
          .\Mail_Attachment.ps1
4. All matching CSV attachments from today’s emails will be saved in the output folder.

## ⏰ Automation (Dedicated)

You can set this script to run automatically:

Use Windows Task Scheduler to execute it every day at a fixed time.

## 🔒 Notes

Requires Outlook installed and the shared mailbox (SOC) added to your profile.

Make sure you have access permissions to the shared mailbox.

Logs and state files will grow over time; clean them periodically if needed.
