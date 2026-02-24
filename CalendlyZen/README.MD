# CalendlyZen Extension – Quick Guide

## What It Does

CalendlyZen automatically updates Calendly links copied to your clipboard so that when pasted in Zendesk, they include the ticket requester’s name, email, and ticket ID.

**Example 1**

Before:   `https://calendly.com/adele.vance/45min`

After:    https://calendly.com/adele.vance/45min/?name={{ticket.requester.name}}&email={{ticket.requester.email}}&a1={{ticket.id}}

Rendered: https://calendly.com/adele.vance/45min/?name=adelev&email=adelev@1025.nc-portal.uk&a1=665067

**Example 2**
Before:   https://calendly.com/d/cx94-dcb-zmj
After:    https://calendly.com/d/cx93-78k-xx4/?name={{ticket.requester.name}}&email={{ticket.requester.email}}&a1={{ticket.id}}
Rendered: https://calendly.com/d/cx93-78k-xx4/?name=adelev&email=adelev@1025.nc-portal.uk&a1=665067

When rendered in Zendesk, it automatically fills in actual values.

## Google Chrome Installation (CalendlyZen_Chrome.zip)

> **Developer mode must be enabled in Chrome before installing.**

### Steps

1. Save **CalendlyZen_Chrome.zip** to your computer.  
2. Right-click the file and select **Extract All**.  
3. Move the extracted folder to a permanent location such as: C:\Extensions\CalendlyZen


Do not leave it in your Downloads folder.  

4. Open Chrome.  
5. In the address bar, go to: chrome://extensions


6. Enable **Developer mode** using the toggle in the top-right corner.  
7. Drag and drop the extracted folder onto the Extensions page.  

The extension should now appear and be enabled.

### Important Notes

- Refresh your Browser and test
- Do not delete the extracted folder after installing.  
- If the folder is removed, Chrome will disable the extension.  
- If Chrome is reset or your profile is cleared, the extension may need to be reinstalled.  

---

## Mozilla Firefox Installation (CalendlyZen_Firefox.xpi)

### Steps

1. Save **CalendlyZen_Firefox.xpi** to your computer.  
2. Open Firefox.  
3. In the address bar, go to: about:addons


4. Click the **settings icon** (gear icon).  
5. Select **Install Add-on From File**.  
6. Browse to the downloaded **CalendlyZen_Firefox.xpi** file.  
7. Select it and confirm installation.  

The extension will now be installed and active.

### Important Notes

- Refresh your Browser and test
- The `.xpi` file does not need to remain in a specific folder after installation.  
- If Firefox blocks the install, confirm any security prompts shown.