# 📧 Content Search & Purge Script (Universal)

**Version:** 3.1.0  
**Author:** Mohsen Heidari  
**Contact:** abarkakia@gmail.com  

## 📌 Purpose
This PowerShell script automates **Microsoft Purview (Security & Compliance)** content searches in **Exchange Online** and optionally purges matching items. It is designed for admins who need to quickly locate and remove malicious or unwanted emails across mailboxes.

---

## ✅ Features
- Works on **Windows PowerShell 5.1** and **PowerShell 7+**.
- **No forced elevation** (runs as admin first time to install modules and then can run as current user).
- Auto-installs or updates **ExchangeOnlineManagement** module:
  - Admin → installs for **AllUsers**
  - Non-admin → installs for **CurrentUser**
- Adaptive sign-in:
  - Prefers **system browser** (`-DisableWAM`)
  - Uses **search-only session** if available
  - Falls back to **secure credential prompt**
- Detects **IE Enhanced Security Configuration** and warns.
- Emoji-friendly console output.
- Creates **compliance search** for last 7 days based on:
  - Sender email
  - Optional subject
- Supports **soft-delete purge** after review or immediate action.

---

## ✅ Prerequisites
- **PowerShell 5.1 or later**
- **Microsoft Purview admin account** with:
  - **Search and Purge** role assigned
- Internet access to connect to Microsoft 365 services
- **Microsoft Edge WebView2 Runtime** (recommended for sign-in)

---

## ✅ Installation
1. Save the script as **UTF-8 with BOM** for emoji compatibility in PS 5.1.
2. Ensure execution policy allows script execution:
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   ```
3. Run the script in **PowerShell 5.1** or **PowerShell 7+**.

---

## ✅ Usage
1. Launch PowerShell and First run as admin:
   ```powershell
   .\AutoPurge.ps1
   ```
2. Sign in with your **Purview-enabled admin account**.
3. Provide inputs when prompted:
   - **Sender email address** (required)
   - **Email subject** (optional)
   - **Search name or ticket number** (required)
   - **Quick purge? (Y/N)**

---

## ✅ Workflow Overview
1. **Environment setup** (TLS, encoding, execution policy)
2. **Module check & install**
3. **Sign-in to Microsoft Purview**
4. **Role validation**
5. **Collect user inputs**
6. **Build KQL query** for last 7 days
7. **Create and start compliance search**
8. **Wait for completion**
9. **Display results**
10. **Optional purge (SoftDelete)**
11. **Summary report**

---

## ✅ Troubleshooting
- **Sign-in issues**:
  - Install/update **Microsoft Edge WebView2 Runtime**
  - Retry with `pwsh -NoProfile`
  - Use `-DisableWAM` for system browser flow
- **Role errors**:
  - Ensure your account has **Search and Purge** role
- **IE ESC warning**:
  - Disable IE Enhanced Security in **Server Manager**

---

## ✅ Version History
- **3.1.0** – Added adaptive sign-in, IE ESC detection, emoji support, improved purge tracking.

---

## ✅ Author
Mohsen Heidari  
📧 **abarkakia@gmail.com**

---

## ⚠️ Disclaimer
Use this script responsibly. Purging emails is irreversible. Test in a non-production environment before deploying.
