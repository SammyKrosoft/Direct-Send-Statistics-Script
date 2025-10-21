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
