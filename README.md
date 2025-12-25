# Purge Pro v5
A specialized PowerShell GUI tool for searching and purging emails across Microsoft 365 mailboxes via the Microsoft Graph API.  

## **🚀 Features**  
• Search Filters: Subject, sender, and specific date ranges.  
• Scoped Scanning: Search Inbox, Sent Items, Junk, Deleted Items, or All Folders.  
• Bulk Actions: Preview, export results to CSV, or download as .eml.  
• Secure Purge: Permanently delete selected emails from target mailboxes.  
• Tenant-Wide: Search specific users or scan all mailboxes in the directory.  

## **🛠️ Requirements**   
• PowerShell: 5.1 or higher.  
• M365 Permissions: Azure App Registration with Mail.ReadWrite (Application Permission).  
• Credentials: Tenant ID, Client ID, and Client Secret.  

## **📖 Quick Start**    
• Run the script: Right-click Purge Pro.ps1 and select Run with PowerShell.  
• Authenticate: Enter your App Registration credentials and click Test Connection.  
• Target: Input target user UPNs (comma-separated) or leave blank for the whole tenant.  
• Action: Define your search criteria, click Start Search, and use the results window to preview or delete.  
