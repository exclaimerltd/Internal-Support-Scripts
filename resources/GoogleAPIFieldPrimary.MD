# Google Workspace User Attribute Repair (Exclaimer fix)

This script will set the `primary` flag on a user's `organizations` entry in Google Directory so it is synced with Exclaimer.

## Prerequisites

Before running this script, complete the following steps:

1. Go to https://console.cloud.google.com/ and sign in with your Super Admin account.

2. Create a new project by clicking ***“Create Project”***.

<img src="img/GoogleAPIFieldPrimary/image002.png" alt="image002" style="border:1px solid #d0d7de; padding:4px; border-radius:4px;" />

If you already have a project and it has automatically been opened, click on ***“Select a Project”*** at the top of the page next to the Google Cloud logo

<img src="img/GoogleAPIFieldPrimary/image003.png" alt="image003" style="border:1px solid #d0d7de; padding:4px; border-radius:4px;" />

In the new window, click on ***“New Project”***

<img src="img/GoogleAPIFieldPrimary/image004.jpg" alt="image004" style="border:1px solid #d0d7de; padding:4px; border-radius:4px;" />

3. Add a descriptive name and select your organization.

4. With the Project selected, navigate to ***“IAM & Admin > Service Accounts”***.

5. Click on ***“Create service account”***.

<img src="img/GoogleAPIFieldPrimary/image005.png" alt="image005" style="border:1px solid #d0d7de; padding:4px; border-radius:4px;" />

6. Enter a name for the service account.

7. Click on ***“Done”***.

8. Note down the ***“OAuth 2 Client ID”*** for this account

9. Click on the 3-dots in the Actions column and select ***“Manage Keys”***

<img src="img/GoogleAPIFieldPrimary/image006.png" alt="image006" style="border:1px solid #d0d7de; padding:4px; border-radius:4px;" />

10. Click on ***“Add Key”*** followed by ***“Create new key”***

<img src="img/GoogleAPIFieldPrimary/image007.png" alt="image007" style="border:1px solid #d0d7de; padding:4px; border-radius:4px;" />

11. Select JSON and click ***“Create”***

<img src="img/GoogleAPIFieldPrimary/image008.png" alt="image008" style="border:1px solid #d0d7de; padding:4px; border-radius:4px;" />

12. A JSON file will be downloaded. This will be required for the script.

---

13. Click on the ***Navigation Menu***

<img src="img/GoogleAPIFieldPrimary/image009.png" alt="image009" style="border:1px solid #d0d7de; padding:4px; border-radius:4px;" />

14. Go to ***“APIs & Services > Enabled APIs & Services”***

Click on ***“Enable APIs and services”***

<img src="img/GoogleAPIFieldPrimary/image010.png" alt="image010" style="border:1px solid #d0d7de; padding:4px; border-radius:4px;" />

Search for ***“Admin SDK API”***

<img src="img/GoogleAPIFieldPrimary/image011.png" alt="image011" style="border:1px solid #d0d7de; padding:4px; border-radius:4px;" />

Select ***“Admin SDK API”*** and then click on ***“Enable”***

<img src="img/GoogleAPIFieldPrimary/image012.png" alt="image012" style="border:1px solid #d0d7de; padding:4px; border-radius:4px;" />

15. Or, you can use the link ***“https://console.developers.google.com/apis/api/admin.googleapis.com/overview”***

---

16. Sign into the Google Admin Console in a new tab (https://admin.google.com)

17. Navigate to ***Security > Access and data control > API controls***

<img src="img/GoogleAPIFieldPrimary/image013.png" alt="image013" style="border:1px solid #d0d7de; padding:4px; border-radius:4px;" />

18. Click on ***“Manage Domain Wide Delegation”***

19. Click on ***“Add New”***

20. For Client ID, paste in the ***“OAuth 2 Client ID”*** for the service account

21. For OAuth scopes, enter in `https://www.googleapis.com/auth/admin.directory.user`

22. Click on ***“Authorize”***

<img src="img/GoogleAPIFieldPrimary/image014.png" alt="image014" style="border:1px solid #d0d7de; padding:4px; border-radius:4px;" />

---

### PowerShell Requirement

PowerShell 7 is required to run this script. To install this, please go to the following site:

Install PowerShell on Windows - PowerShell | Microsoft Learn  
https://learn.microsoft.com/en-gb/powershell/scripting/install/install-powershell-on-windows?view=powershell-7.5

---

## To run the script

Please execute the PowerShell script provided and complete the following steps:

1. Open PowerShell 7 (Press Win+R and type in ***“pwsh”*** and run)

The command window should show PowerShell 7.x.x

<img src="img/GoogleAPIFieldPrimary/image015.jpg" alt="image015" style="border:1px solid #d0d7de; padding:4px; border-radius:4px;" />

2. Run the command:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass; Invoke-WebRequest -Uri "https://raw.githubusercontent.com/exclaimerltd/Internal-Support-Scripts/master/GoogleAPIFieldPrimary.ps1" -OutFile "$env:TEMP\GoogleAPIFieldPrimary.ps1"; & "$env:TEMP\GoogleAPIFieldPrimary.ps1"
