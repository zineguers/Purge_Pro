<#
======================================================================
 Purge Pro - Command Center (v5.2)
======================================================================
 Author          : Jonathan Morelli (JoMo)
 Design          : Zine Guers
 Refactored By   : AI Commander
 Company         : S3 Technologies
 Purpose         : Search/Purge emails in M365 via Graph API.
 Requirements    : PowerShell 5.1+, Graph API Access (Mail.ReadWrite)
======================================================================
#>

# Force clean run
[System.GC]::Collect()

# Load Assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Configuration & Persistence ---
$ConfigFile = Join-Path $env:APPDATA "O365SearchTool_Config.xml"

function Get-Config {
    if (Test-Path $ConfigFile) {
        try { return Import-Clixml $ConfigFile } catch { return $null }
    }
    return $null
}

function Save-Config {
    param($Data)
    try { $Data | Export-Clixml $ConfigFile } catch { [System.Windows.Forms.MessageBox]::Show("Failed to save config: $_") }
}

# --- Base64 Logo Asset ---
# --- Load Logo from Imgur ---
$imgurUrl = "https://i.imgur.com/PBBbl67.png"

$wc = New-Object System.Net.WebClient
$wc.Headers.Add("User-Agent", "PowerShell")
try {
    $imageBytes = $wc.DownloadData($imgurUrl)
    $ms = New-Object System.IO.MemoryStream(,$imageBytes)
    $logoImage = [System.Drawing.Image]::FromStream($ms)
    $ms.Dispose()
} catch {
    $logoImage = $null
}

# --- UI Setup ---
$form = New-Object Windows.Forms.Form
$form.Text = "Purge Pro v5.2 - Command Center"
$form.Size = New-Object Drawing.Size(600, 750)
$form.MinimumSize = New-Object Drawing.Size(600, 750)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object Drawing.Font("Segoe UI", 9)

# Define ToolTip control
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 5000
$toolTip.InitialDelay = 1000
$toolTip.ReshowDelay = 500

# Logo
if ($logoImage) {
    $logoBox = New-Object Windows.Forms.PictureBox
    $logoBox.Image = $logoImage
    $logoBox.SizeMode = "Zoom"
    $logoBox.Size = New-Object Drawing.Size(550, 60)
    $logoBox.Location = New-Object Drawing.Point(15, 10)
    $logoBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $form.Controls.Add($logoBox)
}

# --- Group: Connection Settings ---
$grpConn = New-Object Windows.Forms.GroupBox
$grpConn.Text = "Connection Settings"
$grpConn.Location = New-Object Drawing.Point(15, 80)
$grpConn.Size = New-Object Drawing.Size(550, 150)
$grpConn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($grpConn)

# Tenant ID
$lblTenant = New-Object Windows.Forms.Label
$lblTenant.Text = "Tenant ID:"
$lblTenant.Location = New-Object Drawing.Point(15, 30)
$lblTenant.AutoSize = $true
$grpConn.Controls.Add($lblTenant)

$txtTenant = New-Object Windows.Forms.TextBox
$txtTenant.Location = New-Object Drawing.Point(120, 27)
$txtTenant.Width = 410
$txtTenant.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$grpConn.Controls.Add($txtTenant)
$toolTip.SetToolTip($txtTenant, "The Directory (Tenant) ID from Azure AD.")

# Client ID
$lblClient = New-Object Windows.Forms.Label
$lblClient.Text = "Client ID:"
$lblClient.Location = New-Object Drawing.Point(15, 60)
$lblClient.AutoSize = $true
$grpConn.Controls.Add($lblClient)

$txtClient = New-Object Windows.Forms.TextBox
$txtClient.Location = New-Object Drawing.Point(120, 57)
$txtClient.Width = 410
$txtClient.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$grpConn.Controls.Add($txtClient)
$toolTip.SetToolTip($txtClient, "The Application (Client) ID of your App Registration.")

# Client Secret
$lblSecret = New-Object Windows.Forms.Label
$lblSecret.Text = "Client Secret:"
$lblSecret.Location = New-Object Drawing.Point(15, 90)
$lblSecret.AutoSize = $true
$grpConn.Controls.Add($lblSecret)

$txtSecret = New-Object Windows.Forms.TextBox
$txtSecret.Location = New-Object Drawing.Point(120, 87)
$txtSecret.Width = 410
$txtSecret.UseSystemPasswordChar = $true
$txtSecret.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$grpConn.Controls.Add($txtSecret)
$toolTip.SetToolTip($txtSecret, "The Client Secret value. Hidden for security.")

# Clear Config Button
$btnClearConfig = New-Object Windows.Forms.Button
$btnClearConfig.Text = "Clear"
$btnClearConfig.Location = New-Object Drawing.Point(120, 115)
$btnClearConfig.Size = New-Object Drawing.Size(100, 25)
$btnClearConfig.BackColor = [System.Drawing.Color]::WhiteSmoke
$grpConn.Controls.Add($btnClearConfig)

$btnClearConfig.Add_Click({
    $txtTenant.Text = ""
    $txtClient.Text = ""
    $txtSecret.Text = ""
})

# --- Test Connection Button ---
$btnTestConn = New-Object Windows.Forms.Button
$btnTestConn.Text = "Test Connection"
$btnTestConn.Location = New-Object Drawing.Point(230, 115)
$btnTestConn.Size = New-Object Drawing.Size(150, 25)
$btnTestConn.BackColor = [System.Drawing.Color]::LightBlue
$grpConn.Controls.Add($btnTestConn)

$btnTestConn.Add_Click({
    # 1. Inputs
    $tenantId = $txtTenant.Text.Trim()
    $clientId = $txtClient.Text.Trim()
    $clientSecret = $txtSecret.Text.Trim()
    
    # 2. Validation
    if ([string]::IsNullOrWhiteSpace($tenantId) -or [string]::IsNullOrWhiteSpace($clientId) -or [string]::IsNullOrWhiteSpace($clientSecret)) {
        [System.Windows.Forms.MessageBox]::Show("Please fill in all Connection Settings.", "Missing Credentials", "OK", "Error")
        return
    }

    $statusLabel.Text = "Testing authentication..."
    Pump-UI 

    # 3. Auth Test
    try {
        $body = @{
            grant_type    = "client_credentials"
            scope         = "https://graph.microsoft.com/.default"
            client_id     = $clientId
            client_secret = $clientSecret
        }
        $tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $body -ErrorAction Stop
        
        # Test connection by calling a simple Graph endpoint
        Invoke-RestMethod -Headers @{ Authorization = "Bearer $($tokenResponse.access_token)" } -Uri "https://graph.microsoft.com/v1.0/users?`$top=1&`$select=id" -ErrorAction Stop | Out-Null
        
        $statusLabel.Text = "Connection SUCCESS: Token retrieved and Graph API accessible."
        [System.Windows.Forms.MessageBox]::Show("Connection Successful!", "Success", "OK", "Information")
    } catch {
        $statusLabel.Text = "Connection FAILED."
        [System.Windows.Forms.MessageBox]::Show("Connection Test failed. Check credentials and App Registration permissions (Mail.ReadWrite).`r`nError: $($_.Exception.Message)", "Test Error", "OK", "Error")
    }
})

# --- Group: Search Criteria ---
$grpSearch = New-Object Windows.Forms.GroupBox
$grpSearch.Text = "Search Criteria"
$grpSearch.Location = New-Object Drawing.Point(15, 240)
$grpSearch.Size = New-Object Drawing.Size(550, 300)
$grpSearch.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$form.Controls.Add($grpSearch)

# Mailboxes
$lblMailbox = New-Object Windows.Forms.Label
$lblMailbox.Text = "Target Mailboxes:"
$lblMailbox.Location = New-Object Drawing.Point(15, 30)
$lblMailbox.AutoSize = $true
$grpSearch.Controls.Add($lblMailbox)

$txtMailbox = New-Object Windows.Forms.TextBox
$txtMailbox.Location = New-Object Drawing.Point(120, 27)
$txtMailbox.Width = 410
$txtMailbox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$grpSearch.Controls.Add($txtMailbox)
$toolTip.SetToolTip($txtMailbox, "Comma-separated list of UserPrincipalNames. Leave empty to search ALL users (Warning: Slow).")

# Partial Subject
$lblSubject = New-Object Windows.Forms.Label 
$lblSubject.Text = "Partial Subject:"
$lblSubject.Location = New-Object Drawing.Point(15, 65)
$lblSubject.AutoSize = $true
$grpSearch.Controls.Add($lblSubject)

$txtSubject = New-Object Windows.Forms.TextBox
$txtSubject.Location = New-Object Drawing.Point(120, 62)
$txtSubject.Width = 410
$txtSubject.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$grpSearch.Controls.Add($txtSubject)

# Senders
$lblSender = New-Object Windows.Forms.Label
$lblSender.Text = "Sender Email(s):"
$lblSender.Location = New-Object Drawing.Point(15, 100)
$lblSender.AutoSize = $true
$grpSearch.Controls.Add($lblSender)

$txtSender = New-Object Windows.Forms.TextBox
$txtSender.Location = New-Object Drawing.Point(120, 97)
$txtSender.Width = 410
$txtSender.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$grpSearch.Controls.Add($txtSender)

# Folder Scope
$lblScope = New-Object Windows.Forms.Label
$lblScope.Text = "Folder Scope:"
$lblScope.Location = New-Object Drawing.Point(15, 135)
$lblScope.AutoSize = $true
$grpSearch.Controls.Add($lblScope)

$cbScope = New-Object Windows.Forms.ComboBox
$cbScope.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cbScope.Items.AddRange(@("Inbox (default)", "All folders", "Sent Items", "Junk Email", "Deleted Items", "Drafts"))
$cbScope.SelectedIndex = 0
$cbScope.Location = New-Object Drawing.Point(120, 132)
$cbScope.Width = 200
$grpSearch.Controls.Add($cbScope)

# Dates
$chkDates = New-Object Windows.Forms.CheckBox
$chkDates.Text = "Filter by Date Range"
$chkDates.Location = New-Object Drawing.Point(120, 170) 
$chkDates.AutoSize = $true
$grpSearch.Controls.Add($chkDates)

$lblStart = New-Object Windows.Forms.Label
$lblStart.Text = "Start Date:"
$lblStart.Location = New-Object Drawing.Point(15, 200) 
$lblStart.AutoSize = $true
$grpSearch.Controls.Add($lblStart)

$dtStart = New-Object Windows.Forms.DateTimePicker
$dtStart.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$dtStart.Location = New-Object Drawing.Point(120, 197) 
$dtStart.Width = 120
$dtStart.Enabled = $false
$grpSearch.Controls.Add($dtStart)

$lblEnd = New-Object Windows.Forms.Label
$lblEnd.Text = "End Date:"
$lblEnd.Location = New-Object Drawing.Point(260, 200) 
$lblEnd.AutoSize = $true
$grpSearch.Controls.Add($lblEnd)

$dtEnd = New-Object Windows.Forms.DateTimePicker
$dtEnd.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$dtEnd.Location = New-Object Drawing.Point(330, 197) 
$dtEnd.Width = 120
$dtEnd.Enabled = $false
$grpSearch.Controls.Add($dtEnd)

# Toggle dates logic
$chkDates.Add_CheckedChanged({
    $dtStart.Enabled = $chkDates.Checked
    $dtEnd.Enabled = $chkDates.Checked
})

# --- Status & Actions ---
$progressBar = New-Object Windows.Forms.ProgressBar
$progressBar.Location = New-Object Drawing.Point(15, 555)
$progressBar.Size = New-Object Drawing.Size(550, 20)
$progressBar.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($progressBar)

# Status Strip
$statusStrip = New-Object Windows.Forms.StatusStrip
$statusLabel = New-Object Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready"
$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

# Buttons
$btnRun = New-Object Windows.Forms.Button
$btnRun.Text = "Start Search"
$btnRun.Location = New-Object Drawing.Point(15, 590)
$btnRun.Size = New-Object Drawing.Size(550, 40)
$btnRun.Font = New-Object Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$btnRun.BackColor = [System.Drawing.Color]::LightCoral
$btnRun.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($btnRun)

$btnClear = New-Object Windows.Forms.Button
$btnClear.Text = "Reset Form"
$btnClear.Location = New-Object Drawing.Point(15, 640)
$btnClear.Size = New-Object Drawing.Size(100, 30)
$btnClear.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($btnClear)

$btnStop = New-Object Windows.Forms.Button
$btnStop.Text = "Stop Search"
$btnStop.Location = New-Object Drawing.Point(125, 640)
$btnStop.Size = New-Object Drawing.Size(100, 30)
$btnStop.Enabled = $false
$btnStop.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($btnStop)

# Footer
$labelCopyright = New-Object Windows.Forms.Label
$labelCopyright.Text = "Created by JoMo`r`nDeveloped by ZG"
$labelCopyright.AutoSize = $false
$labelCopyright.Size = New-Object Drawing.Size(200, 32)
$labelCopyright.TextAlign = [System.Drawing.ContentAlignment]::BottomRight
$labelCopyright.Font = New-Object Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$labelCopyright.ForeColor = [System.Drawing.Color]::Gray
$labelCopyright.Location = New-Object Drawing.Point(365, 638)
$labelCopyright.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($labelCopyright)

# --- Logic ---

$btnClear.Add_Click({
    $txtMailbox.Text = ""
    $txtSubject.Text = ""
    $txtSender.Text = ""
    $progressBar.Value = 0
    $statusLabel.Text = "Form reset."
})

$global:StopSearch = $false
$btnStop.Add_Click({
    $global:StopSearch = $true
    $statusLabel.Text = "Stopping..."
})

function Pump-UI { [System.Windows.Forms.Application]::DoEvents() }

$btnRun.Add_Click({
    # 1. Inputs
    $tenantId = $txtTenant.Text.Trim()
    $clientId = $txtClient.Text.Trim()
    $clientSecret = $txtSecret.Text.Trim()
    $mailboxes = ($txtMailbox.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }) -join ','
    
    # 2. Validation
    if ([string]::IsNullOrWhiteSpace($tenantId) -or [string]::IsNullOrWhiteSpace($clientId) -or [string]::IsNullOrWhiteSpace($clientSecret)) {
        [System.Windows.Forms.MessageBox]::Show("Please fill in all Connection Settings.", "Missing Credentials", "OK", "Error")
        return
    }

    # 3. UI State
    $btnRun.Enabled = $false
    $btnStop.Enabled = $true
    $global:StopSearch = $false
    $statusLabel.Text = "Authenticating..."
    Pump-UI

    # 4. Auth
    try {
        $body = @{
            grant_type    = "client_credentials"
            scope         = "https://graph.microsoft.com/.default"
            client_id     = $clientId
            client_secret = $clientSecret
        }
        $tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $body -ErrorAction Stop
        $accessToken = $tokenResponse.access_token
    } catch {
        $statusLabel.Text = "Authentication Failed."
        [System.Windows.Forms.MessageBox]::Show("Authentication error:`r`n$($_.Exception.Message)", "Auth Error", "OK", "Error")
        $btnRun.Enabled = $true
        $btnStop.Enabled = $false
        return
    }

    # 5. Build Target User List
    $users = @()
    if ($mailboxes) {
        $users = $mailboxes.Split(',') | ForEach-Object { [PSCustomObject]@{ userPrincipalName = $_ } }
    } else {
        # Fetch all users (paged)
        $statusLabel.Text = "Fetching all users from directory..."
        Pump-UI
        $url = "https://graph.microsoft.com/v1.0/users?`$select=userPrincipalName"
        do {
            try {
                $response = Invoke-RestMethod -Headers @{ Authorization = "Bearer $accessToken" } -Uri $url -ErrorAction Stop
                foreach ($user in $response.value) { if ($user.userPrincipalName) { $users += [PSCustomObject]@{ userPrincipalName = $user.userPrincipalName } } }
                $url = $response.'@odata.nextLink'
            } catch {
                Write-Warning "Error fetching users: $_"; break
            }
            if ($global:StopSearch) { break }
        } while ($url)
    }

    if ($users.Count -eq 0) {
        $statusLabel.Text = "No users found."
        $btnRun.Enabled = $true; $btnStop.Enabled = $false; return
    }

    # 6. Search Loop
    $results = @()
    $counter = 0
    $total = $users.Count
    $statusLabel.Text = "Starting search across $total mailbox(es)..."

    # Construct Filter
    $filterParts = @()
    
    # Subject
    if ($txtSubject.Text) {
        $s = $txtSubject.Text.Replace("'","''")
        $filterParts += "contains(subject,'$s')"
    }
    # Senders
    $senders = ($txtSender.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
    if ($senders.Count -gt 0) {
        $senderClauses = $senders | ForEach-Object { "(from/emailAddress/address eq '$($_ -replace "'","''")')" }
        $filterParts += "(" + ($senderClauses -join " or ") + ")"
    }

    # Dates
    if ($chkDates.Checked) {
        $dStart = $dtStart.Value.Date.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $dEnd = $dtEnd.Value.Date.AddDays(1).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $filterParts += "receivedDateTime ge $dStart and receivedDateTime lt $dEnd"
    }

    $filterQuery = $filterParts -join " and "
    $scopeFolder = $cbScope.SelectedItem
    
    foreach ($user in $users) {
        Pump-UI
        if ($global:StopSearch) { break }
        
        $counter++
        $pct = if ($total -gt 0) { [math]::Round(($counter / $total) * 100) } else { 100 }
        $progressBar.Value = [math]::Min($pct, 100)
        $statusLabel.Text = "Scanning ($counter/$total): $($user.userPrincipalName)"

        # Determine Endpoint based on scope
        $baseUri = switch($scopeFolder) {
            "All folders"   { "https://graph.microsoft.com/v1.0/users/$($user.userPrincipalName)/messages" }
            "Sent Items"    { "https://graph.microsoft.com/v1.0/users/$($user.userPrincipalName)/mailFolders/SentItems/messages" }
            "Junk Email"    { "https://graph.microsoft.com/v1.0/users/$($user.userPrincipalName)/mailFolders/JunkEmail/messages" }
            "Deleted Items" { "https://graph.microsoft.com/v1.0/users/$($user.userPrincipalName)/mailFolders/DeletedItems/messages" }
            "Drafts"        { "https://graph.microsoft.com/v1.0/users/$($user.userPrincipalName)/mailFolders/Drafts/messages" }
            default         { "https://graph.microsoft.com/v1.0/users/$($user.userPrincipalName)/mailFolders/Inbox/messages" }
        }
        
        # Build URL
        try {
            $ub = [System.UriBuilder]$baseUri
            $sel = 'id,subject,receivedDateTime,from,hasAttachments,isRead'
            if ($filterQuery) {
                $enc = [System.Uri]::EscapeDataString($filterQuery)
                $ub.Query = "`$filter=$enc&`$select=$sel"
            } else {
                $ub.Query = "`$select=$sel"
            }
            $uri = $ub.Uri.AbsoluteUri
        } catch { continue }

        # Fetch Messages
        do {
            Pump-UI
            if ($global:StopSearch) { break }
            try {
                $res = Invoke-RestMethod -Headers @{ Authorization = "Bearer $accessToken" } -Uri $uri
                foreach ($msg in $res.value) {
                    
                    # [FIX] Handle Exchange Legacy DN (X500) vs SMTP in search list
                    $sAddr = $msg.from.emailAddress.address
                    if ($sAddr -and $sAddr.StartsWith("/")) {
                        $sAddr = $msg.from.emailAddress.name
                    }

                    $results += [PSCustomObject]@{
                        User        = $user.userPrincipalName
                        Subject     = $msg.subject
                        Sender      = $sAddr
                        Date        = $msg.receivedDateTime
                        MsgID       = $msg.id
                        Attachments = $msg.hasAttachments
                        ReadStatus  = $msg.isRead
                    }
                }
                $uri = $res.'@odata.nextLink'
            } catch {
                break 
            }
        } while ($uri)
    }

    # 7. Show Results
    if ($results.Count -gt 0) {
        $statusLabel.Text = "Found $($results.Count) messages."
        Show-ResultsWindow -Results $results -AccessToken $accessToken
    } else {
        $statusLabel.Text = "Search complete. No messages found."
        [System.Windows.Forms.MessageBox]::Show("No messages found matching criteria.", "Result", "OK", "Information")
    }

    $btnRun.Enabled = $true
    $btnStop.Enabled = $false
})

function Show-ResultsWindow {
    param($Results, $AccessToken)

    $resForm = New-Object Windows.Forms.Form
    $resForm.Text = "Search Results - $($Results.Count) items"
    $resForm.Size = New-Object Drawing.Size(900, 600)
    $resForm.StartPosition = "CenterParent"
    
    $initialResults = $Results | Select-Object *

    # --- MenuStrip for Top Actions ---
    $menuBar = New-Object System.Windows.Forms.MenuStrip
    $resForm.Controls.Add($menuBar)

    # --- Helper function to refresh the ListView with data ---
    function Refresh-ListView($listView, $data) {
        $listView.Items.Clear()
        $data | ForEach-Object {
            try {
                $displayDate = ([datetime]$_.Date).ToString("yyyy-MM-dd HH:mm:ss")
            } catch {
                $displayDate = [string]$_.Date
            }
            
            $lvi = New-Object System.Windows.Forms.ListViewItem($displayDate)
            $lvi.SubItems.Add([string]$_.User) | Out-Null
            $lvi.SubItems.Add([string]$_.Sender) | Out-Null
            $lvi.SubItems.Add([string]$_.Subject) | Out-Null
            
            $attText = if ($_.Attachments) { "Yes" } else { "No" }
            $lvi.SubItems.Add($attText) | Out-Null

            $readText = if ($_.ReadStatus) { "Read" } else { "Unread" }
            $lvi.SubItems.Add($readText) | Out-Null
            
            $lvi.Tag = $_
            $listView.Items.Add($lvi) | Out-Null
        }
    }

    # --- 1. File Menu ---
    $fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $fileMenu.Text = "&File"
    $menuBar.Items.Add($fileMenu) | Out-Null

    # Help Menu Item
    $helpMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $helpMenuItem.Text = "&Help"
    $helpMenuItem.Add_Click({
        $helpText = @"
About Purge Pro
Purge Pro is a powerful email management tool designed to help professionals efficiently search, filter, and manage emails.

Developed by: Zine Guers for S3 Technologies
Version: 5.2

For support or inquiries, please contact Zine Guers.
"@
        [System.Windows.Forms.MessageBox]::Show($helpText, "About Purge Pro", "OK", "Information")
    })
    $fileMenu.DropDownItems.Add($helpMenuItem) | Out-Null
    
    $fileMenu.DropDownItems.Add("-") | Out-Null

    # Close App Menu Item
    $closeMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $closeMenuItem.Text = "&Close App"
    $closeMenuItem.Add_Click({
        $form.Close()
    })
    $fileMenu.DropDownItems.Add($closeMenuItem) | Out-Null
    
    # --- 2. Select All ---
    $selectMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $selectMenuItem.Text = "&Select All"
    $selectMenuItem.Add_Click({ 
        foreach($item in $resList.Items) { $item.Checked = $true } 
    })
    $menuBar.Items.Add($selectMenuItem) | Out-Null

    # --- 3. Deselect All ---
    $deselectMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $deselectMenuItem.Text = "&Deselect All"
    $deselectMenuItem.Add_Click({ 
        foreach($item in $resList.Items) { $item.Checked = $false } 
    })
    $menuBar.Items.Add($deselectMenuItem) | Out-Null

    # --- 4. View Menu ---
    $viewMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $viewMenu.Text = "&View"
    $menuBar.Items.Add($viewMenu) | Out-Null
    
    # Sort Recent to Oldest
    $sortRecentMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $sortRecentMenuItem.Text = "Sort by Date: &Recent to Oldest"
    $sortRecentMenuItem.Add_Click({
        $sorted = $Results | Sort-Object { 
            try { [datetime]$_.Date } catch { [datetime]::MinValue }
        } -Descending
        Refresh-ListView $resList $sorted
    })
    $viewMenu.DropDownItems.Add($sortRecentMenuItem) | Out-Null

    $viewMenu.DropDownItems.Add("-") | Out-Null
    
    # Reset View
    $resetViewMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $resetViewMenuItem.Text = "&Reset View"
    $resetViewMenuItem.Add_Click({
        Refresh-ListView $resList $initialResults
    })
    $viewMenu.DropDownItems.Add($resetViewMenuItem) | Out-Null

    # --- ListView ---
    $resList = New-Object System.Windows.Forms.ListView
    $resList.View = [System.Windows.Forms.View]::Details
    $resList.FullRowSelect = $true
    $resList.GridLines = $true
    $resList.CheckBoxes = $true
    $resList.MultiSelect = $true
    $resList.Location = New-Object Drawing.Point(0, 25)
    $resList.Size = New-Object Drawing.Size(880, 480)
    $resList.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    
    $resList.Columns.Add("Date", 150) | Out-Null
    $resList.Columns.Add("Mailbox", 200) | Out-Null
    $resList.Columns.Add("Sender", 200) | Out-Null
    $resList.Columns.Add("Subject", 250) | Out-Null
    $resList.Columns.Add("Att.", 40) | Out-Null
    $resList.Columns.Add("Status", 60) | Out-Null

    # Populate
    Refresh-ListView $resList $Results
    
    $resForm.Controls.Add($resList)

    # --- Bottom Actions ---
    $pnlBottom = New-Object Windows.Forms.Panel
    $pnlBottom.Height = 50
    $pnlBottom.Dock = "Bottom"
    $resForm.Controls.Add($pnlBottom)

    # Export CSV
    $btnCsv = New-Object Windows.Forms.Button
    $btnCsv.Text = "Export All to CSV"
    $btnCsv.Location = New-Object Drawing.Point(10, 10)
    $btnCsv.Size = New-Object Drawing.Size(120, 30)
    $btnCsv.Add_Click({
        $sfd = New-Object Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV (*.csv)|*.csv"; $sfd.FileName = "O365_Results.csv"
        if ($sfd.ShowDialog() -eq "OK") {
            $Results | Export-Csv -Path $sfd.FileName -NoTypeInformation
            [System.Windows.Forms.MessageBox]::Show("Export Successful.")
        }
    })
    $pnlBottom.Controls.Add($btnCsv)

    # Download
    $btnDownload = New-Object Windows.Forms.Button
    $btnDownload.Text = "Download Selected"
    $btnDownload.Location = New-Object Drawing.Point(140, 10)
    $btnDownload.Size = New-Object Drawing.Size(120, 30)
    $btnDownload.Add_Click({
        if ($resList.CheckedItems.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Select items first."); return }
        
        $fbd = New-Object Windows.Forms.FolderBrowserDialog
        if ($fbd.ShowDialog() -ne "OK") { return }
        
        $count = 0
        foreach ($item in $resList.CheckedItems) {
            $msgData = $item.Tag
            try {
                $fName = "{0}_{1}.eml" -f ($msgData.User -split "@")[0], $msgData.MsgID.Substring(0,8)
                $fPath = Join-Path $fbd.SelectedPath $fName
                $uri = "https://graph.microsoft.com/v1.0/users/$($msgData.User)/messages/$($msgData.MsgID)/`$value"
                Invoke-WebRequest -Uri $uri -Headers @{ Authorization = "Bearer $AccessToken" } -OutFile $fPath
                $count++
            } catch {}
        }
        [System.Windows.Forms.MessageBox]::Show("Downloaded $count messages.")
    })
    $pnlBottom.Controls.Add($btnDownload)
    
    # --- Helper Function for Preview Form ---
    function Add-HeaderField {
        param(
            [System.Windows.Forms.Panel]$panel, 
            [string]$label, 
            [string]$value, 
            [ref]$yPosition
        )
        
        # [FIX] Force integer cast to avoid op_Subtraction error
        $currentY = [int]$yPosition.Value
        
        $lbl = New-Object Windows.Forms.Label
        $lbl.Text = $label
        $lbl.Location = New-Object Drawing.Point(10, $currentY)
        $lbl.AutoSize = $true
        $lbl.Font = New-Object Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        [void]$panel.Controls.Add($lbl)
        
        $txt = New-Object Windows.Forms.TextBox
        $txt.Text = $value
        
        $txtTop = $currentY - 3
        $txt.Location = New-Object Drawing.Point(100, $txtTop)
        
        $txtWidth = $panel.Width - 115
        $txt.Width = $txtWidth
        
        $txt.ReadOnly = $true
        $txt.BorderStyle = [System.Windows.Forms.BorderStyle]::None
        $txt.BackColor = $panel.BackColor
        $txt.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
        [void]$panel.Controls.Add($txt)
        
        # Update the reference variable explicitly as an int
        $yPosition.Value = [int]($currentY + 25)
    }

    $btnPreview = New-Object Windows.Forms.Button
    $btnPreview.Text = "Preview"
    $btnPreview.Location = New-Object Drawing.Point(270, 10)
    $btnPreview.Size = New-Object Drawing.Size(120, 30)
    $btnPreview.Add_Click({
        if ($resList.SelectedItems.Count -ne 1) {
            [System.Windows.Forms.MessageBox]::Show("Please select exactly one message to preview.", "Selection Error", "OK", "Error")
            return
        }
        $item = $resList.SelectedItems[0]
        $msgData = $item.Tag
        
        try {
            # 1. Fetch comprehensive message data
            $uri = "https://graph.microsoft.com/v1.0/users/$($msgData.User)/messages/$($msgData.MsgID)?`$select=subject,receivedDateTime,from,toRecipients,body"
            $msg = Invoke-RestMethod -Headers @{ Authorization = "Bearer $AccessToken" } -Uri $uri -ErrorAction Stop

            # 2. Setup Preview Form
            $prevForm = New-Object Windows.Forms.Form
            $prevForm.Text = "Message Preview"
            if ($msg.subject) { $prevForm.Text += " - $($msg.subject)" }
            $prevForm.Size = New-Object Drawing.Size(800, 600)
            $prevForm.StartPosition = "CenterParent"
            $prevForm.MinimizeBox = $false
            $prevForm.MaximizeBox = $false
            $prevForm.Font = New-Object Drawing.Font("Segoe UI", 9)

            # 3. Create Header Panel
            $pnlHeader = New-Object Windows.Forms.Panel
            $pnlHeader.Dock = "Top"
            $pnlHeader.Height = 110
            $pnlHeader.BackColor = [System.Drawing.Color]::WhiteSmoke
            $pnlHeader.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
            $prevForm.Controls.Add($pnlHeader)

            # Format Recipient List safely
            $toAddresses = ""
            if ($msg.toRecipients) {
                $toAddresses = ($msg.toRecipients | ForEach-Object { 
                    if ([string]::IsNullOrWhiteSpace($_.emailAddress.address)) { $_.emailAddress.name } else { $_.emailAddress.address }
                }) -join "; "
            }
            
            # [FIX] Handle internal display name if address starts with / in Preview
            $fromVal = $msg.from.emailAddress.address
            if ($fromVal -and $fromVal.StartsWith("/")) { $fromVal = $msg.from.emailAddress.name }

            # Add Header Fields using the hardened function
            $yPos = 5
            Add-HeaderField -panel $pnlHeader -label "Subject:" -value "$($msg.subject)" -yPosition ([ref]$yPos)
            Add-HeaderField -panel $pnlHeader -label "From:" -value "$fromVal" -yPosition ([ref]$yPos)
            Add-HeaderField -panel $pnlHeader -label "To:" -value "$toAddresses" -yPosition ([ref]$yPos)
            
            # Safe date conversion
            $dateStr = ""
            try { $dateStr = ([datetime]$msg.receivedDateTime).ToString("yyyy-MM-dd HH:mm:ss") } catch { $dateStr = "$($msg.receivedDateTime)" }
            Add-HeaderField -panel $pnlHeader -label "Date:" -value $dateStr -yPosition ([ref]$yPos)

            # 4. Create Body Panel
            $pnlBody = New-Object Windows.Forms.Panel
            $pnlBody.Dock = "Fill"
            $pnlBody.Padding = New-Object System.Windows.Forms.Padding(5)
            $prevForm.Controls.Add($pnlBody)
            $pnlBody.BringToFront()

            if ($msg.body.contentType -eq "html") {
                $web = New-Object System.Windows.Forms.WebBrowser
                $web.Dock = "Fill"
                $htmlContent = "<!DOCTYPE html><html><head><style>body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 10pt; }</style></head><body>" + $msg.body.content + "</body></html>"
                $web.DocumentText = $htmlContent
                $pnlBody.Controls.Add($web)
            } else {
                $txtPrev = New-Object Windows.Forms.TextBox
                $txtPrev.Multiline = $true
                $txtPrev.Dock = "Fill"
                $txtPrev.ScrollBars = "Vertical"
                $txtPrev.ReadOnly = $true
                $txtPrev.Text = $msg.body.content
                $txtPrev.Font = New-Object Drawing.Font("Consolas", 9)
                $pnlBody.Controls.Add($txtPrev)
            }
            
            [void]$prevForm.ShowDialog()
            
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to fetch message details for preview:`r`n$($_.Exception.Message)", "Error", "OK", "Error")
        }
    })
    $pnlBottom.Controls.Add($btnPreview)

    # Delete
    $btnDelete = New-Object Windows.Forms.Button
    $btnDelete.Text = "DELETE Selected"
    $btnDelete.Location = New-Object Drawing.Point(750, 10)
    $btnDelete.Size = New-Object Drawing.Size(120, 30)
    $btnDelete.ForeColor = "Red"
    $btnDelete.Anchor = "Right, Top"
    $btnDelete.Add_Click({
        $chkCount = $resList.CheckedItems.Count
        if ($chkCount -eq 0) { return }

        $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to PERMANENTLY DELETE $chkCount message(s)?", "Confirm Purge", "YesNo", "Warning")
        if ($confirm -eq "Yes") {
            $deleted = 0
            foreach ($item in $resList.CheckedItems) {
                $msgData = $item.Tag
                try {
                    $uri = "https://graph.microsoft.com/v1.0/users/$($msgData.User)/messages/$($msgData.MsgID)"
                    Invoke-RestMethod -Method Delete -Uri $uri -Headers @{ Authorization = "Bearer $AccessToken" }
                    $resList.Items.Remove($item)
                    $deleted++
                } catch { Write-Warning "Failed to delete message: $($_.Exception.Message)" }
            }
            [System.Windows.Forms.MessageBox]::Show("Deletion complete. Removed $deleted messages.")
        }
    })
    $pnlBottom.Controls.Add($btnDelete)

    $resForm.ShowDialog() | Out-Null
}

[void]$form.ShowDialog()