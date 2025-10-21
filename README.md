# Start-DirectSendAudit.ps1

## Overview
`Start-DirectSendAudit.ps1` is a PowerShell script designed to help Microsoft 365 administrators **detect potential Direct Send usage** within their tenant. Direct Send occurs when internal systems or devices send emails directly to Exchange Online without using an authenticated connector, which can pose security and compliance risks.

This script automates the process of:
- Connecting to **Exchange Online**.
- Initiating a **Historical Search** (Connector Report / NoConnector) for a specified date range.
- Guiding you to **download the report from Microsoft Purview**.
- Filtering the CSV to include only messages where the **sender belongs to your tenant domains**.
- Producing a **clean, timestamped CSV** and a **summary of message counts per domain**.

---

## Key Features
- ✅ Clears existing Exchange Online sessions before starting.
- ✅ Prompts for output folder, date range (1–90 days), and notification email.
- ✅ Detects tenant domains automatically.
- ✅ Starts a Historical Search for **NoConnector** messages (indicative of Direct Send).
- ✅ Guides you through downloading the report from Purview.
- ✅ Filters and sanitises the CSV for tenant senders.
- ✅ Generates a summary of messages per domain and a grand total.
- ✅ Optionally opens the filtered CSV and disconnects from Exchange Online.

---

## Prerequisites
- **PowerShell 5.x or later** (or PowerShell 7).
- **Exchange Online PowerShell module**:

  ```powershell
  Install-Module ExchangeOnlineManagement

Here’s a **README section** you can include to explain how to run the script and execute the function:

## How to interpret the results

- The Historical Search returns **all inbound messages with no connector**.
- Filtering to **your own domains in the sender** helps isolate **Direct Send from your internal devices/apps** (e.g., printers, apps, or on-prem systems sending through MX directly to Exchange Online without SMTP AUTH).
- It’s a **signal**, not definitive proof—some additional investigation may still be needed.

***


## **How to Run the Script**

### **1. Save the Script**

*   Copy the full script into a file named:
        Start-DirectSendAudit.ps1

### **2. Open PowerShell**

*   Launch **PowerShell** (preferably as Administrator).
*   Ensure you have the **Exchange Online Management module** installed:
    ```powershell
    Install-Module ExchangeOnlineManagement
    ```

### **3. Set Execution Policy (if needed)**

If scripts are blocked, allow them for the current session:

```powershell
Set-ExecutionPolicy RemoteSigned -Scope Process
```

### **4. Navigate to the Script Location**

```powershell
cd "C:\Path\To\Your\Script"
```

### **5. Load the Function (Dot-Source the Script)**

To make the function available in your session:

```powershell
. .\Start-DirectSendAudit.ps1
```

*(Note the dot and space before the script path.)*

### **6. Run the Function**

```powershell
Start-DirectSendAudit
```

The script will:

- Ask for output folder (press Enter for default C:\Temp).
- Ask for number of days (press Enter for default 7).
- Ask for notification email (press Enter for default).
- Then it will connect to Exchange Online and guide you through downloading the report.


✅ Tip: Keep the PowerShell window open until the report is ready and downloaded.

***

### **Quick One-Liner**

If you want to **load and run in one go**:

```powershell
. .\Start-DirectSendAudit.ps1; Start-DirectSendAudit
```

***

✅ **Tip:** If you use this often, add the function to your **PowerShell profile** so it’s always available.

