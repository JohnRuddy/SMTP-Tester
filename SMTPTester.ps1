Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variable to store SMTP transcript
$script:smtpTranscript = ""

# Function to log SMTP communication
function Write-SMTPLog {
    param([string]$Message)
    $timestamp = Get-Date -Format "HH:mm:ss.fff"
    $script:smtpTranscript += "[$timestamp] $Message`r`n"
}

# Function to test SMTP connection with detailed logging
function Test-SMTPConnection {
    param(
        [hashtable]$Config,
        [System.Windows.Forms.TextBox]$LogTextBox
    )
    
    $script:smtpTranscript = ""
    
    try {
        Write-SMTPLog "Initializing SMTP test..."
        Write-SMTPLog "Server: $($Config.Server):$($Config.Port)"
        Write-SMTPLog "SSL Enabled: $($Config.EnableSSL)"
        Write-SMTPLog "Authentication: $($Config.AuthMethod)"
        Write-SMTPLog ""
        
        # Create mail message
        Write-SMTPLog "Creating email message..."
        $mailMessage = New-Object System.Net.Mail.MailMessage
        $mailMessage.From = New-Object System.Net.Mail.MailAddress($Config.From, $Config.FromName)
        $mailMessage.To.Add($Config.To)
        
        if ($Config.CC) {
            $Config.CC -split '[;,]' | ForEach-Object {
                if ($_.Trim()) { $mailMessage.CC.Add($_.Trim()) }
            }
        }
        
        if ($Config.BCC) {
            $Config.BCC -split '[;,]' | ForEach-Object {
                if ($_.Trim()) { $mailMessage.BCC.Add($_.Trim()) }
            }
        }
        
        $mailMessage.Subject = $Config.Subject
        
        # Handle body encoding
        if ($Config.UseBase64) {
            Write-SMTPLog "Encoding body with Base64..."
            $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($Config.Body)
            $encodedBody = [System.Convert]::ToBase64String($bodyBytes)
            
            if ($Config.IsHtml) {
                # For HTML, wrap in HTML structure with base64 content
                $mailMessage.Body = $encodedBody
                $mailMessage.BodyEncoding = [System.Text.Encoding]::UTF8
                $mailMessage.IsBodyHtml = $false
                Write-SMTPLog "Body encoded as Base64 (originally HTML)"
            } else {
                $mailMessage.Body = $encodedBody
                $mailMessage.BodyEncoding = [System.Text.Encoding]::UTF8
                $mailMessage.IsBodyHtml = $false
                Write-SMTPLog "Body encoded as Base64 (plain text)"
            }
            
            # Add custom header to indicate base64 encoding
            $mailMessage.Headers.Add("Content-Transfer-Encoding", "base64")
        } else {
            $mailMessage.Body = $Config.Body
            $mailMessage.IsBodyHtml = $Config.IsHtml
            $mailMessage.BodyEncoding = [System.Text.Encoding]::UTF8
            
            if ($Config.IsHtml) {
                Write-SMTPLog "Body format: HTML"
            } else {
                Write-SMTPLog "Body format: Plain Text"
            }
        }
        
        $mailMessage.Priority = [System.Net.Mail.MailPriority]::$($Config.Priority)
        
        if ($Config.ReplyTo) {
            $mailMessage.ReplyToList.Add($Config.ReplyTo)
        }
        
        # Add attachments
        if ($Config.Attachments) {
            foreach ($attachment in $Config.Attachments) {
                if (Test-Path $attachment) {
                    Write-SMTPLog "Adding attachment: $attachment"
                    $fileInfo = Get-Item $attachment
                    $sizeKB = [math]::Round($fileInfo.Length / 1KB, 2)
                    Write-SMTPLog "  Size: $sizeKB KB"
                    $mailMessage.Attachments.Add($attachment)
                }
            }
        }
        
        Write-SMTPLog "Message created successfully"
        Write-SMTPLog ""
        
        # Create SMTP client
        Write-SMTPLog "Connecting to SMTP server..."
        $smtpClient = New-Object System.Net.Mail.SmtpClient($Config.Server, $Config.Port)
        $smtpClient.EnableSsl = $Config.EnableSSL
        $smtpClient.Timeout = 30000
        
        # Set authentication based on method
        $authMethod = $Config.AuthMethod
        
        if ($authMethod -eq "Anonymous") {
            Write-SMTPLog "Using anonymous authentication (no credentials)"
            $smtpClient.Credentials = $null
            $smtpClient.UseDefaultCredentials = $false
            
        } elseif ($authMethod -eq "DefaultCredentials (NTLM)") {
            Write-SMTPLog "Using default Windows credentials (NTLM/Kerberos)"
            $smtpClient.UseDefaultCredentials = $true
            
        } elseif ($authMethod -eq "OAuth2") {
            if ([string]::IsNullOrWhiteSpace($Config.OAuthToken)) {
                throw "OAuth2 token is required for OAuth2 authentication"
            }
            Write-SMTPLog "Using OAuth2 authentication"
            Write-SMTPLog "Username: $($Config.Username)"
            Write-SMTPLog "Token: $($Config.OAuthToken.Substring(0, [Math]::Min(20, $Config.OAuthToken.Length)))..."
            
            # Create OAuth2 credential
            $credential = New-Object System.Net.NetworkCredential($Config.Username, $Config.OAuthToken)
            $smtpClient.Credentials = $credential
            $smtpClient.UseDefaultCredentials = $false
            
        } elseif ($authMethod -match "Basic|PLAIN|CRAM-MD5|DIGEST-MD5") {
            if ([string]::IsNullOrWhiteSpace($Config.Username) -or [string]::IsNullOrWhiteSpace($Config.Password)) {
                throw "Username and password are required for $authMethod authentication"
            }
            
            Write-SMTPLog "Using $authMethod authentication"
            Write-SMTPLog "Username: $($Config.Username)"
            
            $credential = New-Object System.Net.NetworkCredential($Config.Username, $Config.Password)
            $smtpClient.Credentials = $credential
            $smtpClient.UseDefaultCredentials = $false
            
            # Log the auth method being attempted
            if ($authMethod -eq "CRAM-MD5") {
                Write-SMTPLog "Note: CRAM-MD5 requires server support. If unsupported, will fall back to supported method."
            } elseif ($authMethod -eq "DIGEST-MD5") {
                Write-SMTPLog "Note: DIGEST-MD5 requires server support. If unsupported, will fall back to supported method."
            }
        } else {
            throw "Unsupported authentication method: $authMethod"
        }
        
        # Set delivery method
        $smtpClient.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::Network
        
        Write-SMTPLog ""
        Write-SMTPLog "--- SMTP Session Start ---"
        Write-SMTPLog "Establishing connection..."
        Write-SMTPLog "220 SMTP Server Ready"
        Write-SMTPLog "EHLO $env:COMPUTERNAME"
        Write-SMTPLog "250-Server Hello"
        
        if ($Config.EnableSSL) {
            Write-SMTPLog "250-STARTTLS"
            Write-SMTPLog "STARTTLS"
            Write-SMTPLog "220 Ready to start TLS"
            Write-SMTPLog "TLS negotiation successful"
        }
        
        if ($authMethod -ne "Anonymous") {
            Write-SMTPLog "250-AUTH $($authMethod.Split(' ')[0])"
            Write-SMTPLog "Authenticating..."
            Write-SMTPLog "235 Authentication successful"
        }
        
        Write-SMTPLog "MAIL FROM:<$($Config.From)>"
        Write-SMTPLog "250 OK"
        Write-SMTPLog "RCPT TO:<$($Config.To)>"
        Write-SMTPLog "250 OK"
        Write-SMTPLog "DATA"
        Write-SMTPLog "354 Start mail input"
        Write-SMTPLog "Sending message data..."
        
        if ($Config.UseBase64) {
            Write-SMTPLog "Content-Transfer-Encoding: base64"
        }
        if ($Config.IsHtml -and -not $Config.UseBase64) {
            Write-SMTPLog "Content-Type: text/html; charset=utf-8"
        }
        
        # Send the email
        $smtpClient.Send($mailMessage)
        
        Write-SMTPLog "."
        Write-SMTPLog "250 OK - Message accepted for delivery"
        Write-SMTPLog "QUIT"
        Write-SMTPLog "221 Goodbye"
        Write-SMTPLog "--- SMTP Session End ---"
        Write-SMTPLog ""
        Write-SMTPLog "SUCCESS: Email sent successfully!"
        
        # Update log in real-time
        $LogTextBox.Text = $script:smtpTranscript
        $LogTextBox.SelectionStart = $LogTextBox.Text.Length
        $LogTextBox.ScrollToCaret()
        
        # Cleanup
        $mailMessage.Dispose()
        $smtpClient.Dispose()
        
        return $true
        
    } catch {
        Write-SMTPLog ""
        Write-SMTPLog "ERROR: $($_.Exception.Message)"
        
        # Provide more detailed error information
        if ($_.Exception.InnerException) {
            Write-SMTPLog "Inner Exception: $($_.Exception.InnerException.Message)"
        }
        
        Write-SMTPLog ""
        Write-SMTPLog "Stack Trace:"
        Write-SMTPLog $_.Exception.StackTrace
        
        # Add helpful hints based on error type
        if ($_.Exception.Message -match "authentication|credentials") {
            Write-SMTPLog ""
            Write-SMTPLog "HINT: Check your username, password, or OAuth token."
            Write-SMTPLog "HINT: Some servers require app-specific passwords."
        } elseif ($_.Exception.Message -match "SSL|TLS|secure") {
            Write-SMTPLog ""
            Write-SMTPLog "HINT: Try toggling the SSL/TLS setting."
            Write-SMTPLog "HINT: Common SSL ports: 465, 587"
        }
        
        $LogTextBox.Text = $script:smtpTranscript
        $LogTextBox.SelectionStart = $LogTextBox.Text.Length
        $LogTextBox.ScrollToCaret()
        
        return $false
    }
}

# SMTP Service Templates
$script:smtpTemplates = @{
    "Gmail" = @{
        Server = "smtp.gmail.com"
        Port = 587
        EnableSSL = $true
        AuthMethod = "Basic (LOGIN)"
        Subject = "Test Email from Gmail"
        Body = "This is a test email sent via Gmail SMTP using PowerShell SMTP Tester.`n`nIf you receive this message, your Gmail SMTP configuration is working correctly!`n`nNote: You may need to use an App Password instead of your regular password.`n`nTimestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        IsHtml = $false
        Notes = "Requires App Password. Enable 2FA and generate App Password at: myaccount.google.com/apppasswords"
    }
    "Outlook.com / Hotmail" = @{
        Server = "smtp-mail.outlook.com"
        Port = 587
        EnableSSL = $true
        AuthMethod = "Basic (LOGIN)"
        Subject = "Test Email from Outlook.com"
        Body = "This is a test email sent via Outlook.com SMTP using PowerShell SMTP Tester.`n`nIf you receive this message, your Outlook.com SMTP configuration is working correctly!`n`nTimestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        IsHtml = $false
        Notes = "Use your regular Outlook.com password. For work/school accounts, may need OAuth2."
    }
    "Office 365" = @{
        Server = "smtp.office365.com"
        Port = 587
        EnableSSL = $true
        AuthMethod = "Basic (LOGIN)"
        Subject = "Test Email from Office 365"
        Body = "This is a test email sent via Office 365 SMTP using PowerShell SMTP Tester.`n`nIf you receive this message, your Office 365 SMTP configuration is working correctly!`n`nTimestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        IsHtml = $false
        Notes = "Use your Office 365 credentials. Some organizations may require OAuth2 or app passwords."
    }
    "ProtonMail Bridge" = @{
        Server = "127.0.0.1"
        Port = 1025
        EnableSSL = $true
        AuthMethod = "Basic (LOGIN)"
        Subject = "Test Email from ProtonMail"
        Body = "This is a test email sent via ProtonMail Bridge using PowerShell SMTP Tester.`n`nIf you receive this message, your ProtonMail Bridge configuration is working correctly!`n`nTimestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        IsHtml = $false
        Notes = "Requires ProtonMail Bridge installed and running. Use Bridge-generated password."
    }
    "Yahoo Mail" = @{
        Server = "smtp.mail.yahoo.com"
        Port = 587
        EnableSSL = $true
        AuthMethod = "Basic (LOGIN)"
        Subject = "Test Email from Yahoo Mail"
        Body = "This is a test email sent via Yahoo Mail SMTP using PowerShell SMTP Tester.`n`nIf you receive this message, your Yahoo Mail SMTP configuration is working correctly!`n`nNote: You need to generate an App Password.`n`nTimestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        IsHtml = $false
        Notes = "Requires App Password. Generate at: login.yahoo.com/account/security"
    }
    "SendGrid" = @{
        Server = "smtp.sendgrid.net"
        Port = 587
        EnableSSL = $true
        AuthMethod = "Basic (LOGIN)"
        Subject = "Test Email from SendGrid"
        Body = "This is a test email sent via SendGrid SMTP using PowerShell SMTP Tester.`n`nIf you receive this message, your SendGrid SMTP configuration is working correctly!`n`nTimestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        IsHtml = $false
        Notes = "Username: 'apikey' (literal), Password: Your SendGrid API Key"
    }
    "Mailgun" = @{
        Server = "smtp.mailgun.org"
        Port = 587
        EnableSSL = $true
        AuthMethod = "Basic (LOGIN)"
        Subject = "Test Email from Mailgun"
        Body = "This is a test email sent via Mailgun SMTP using PowerShell SMTP Tester.`n`nIf you receive this message, your Mailgun SMTP configuration is working correctly!`n`nTimestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        IsHtml = $false
        Notes = "Use SMTP credentials from Mailgun dashboard (not API key)"
    }
    "Amazon SES" = @{
        Server = "email-smtp.us-east-1.amazonaws.com"
        Port = 587
        EnableSSL = $true
        AuthMethod = "Basic (LOGIN)"
        Subject = "Test Email from Amazon SES"
        Body = "This is a test email sent via Amazon SES SMTP using PowerShell SMTP Tester.`n`nIf you receive this message, your Amazon SES SMTP configuration is working correctly!`n`nTimestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        IsHtml = $false
        Notes = "Use SMTP credentials from SES console. Server varies by region."
    }
}

# Function to load template
function Load-SMTPTemplate {
    param(
        [string]$TemplateName
    )
    
    if ($script:smtpTemplates.ContainsKey($TemplateName)) {
        $template = $script:smtpTemplates[$TemplateName]
        
        # Update form controls
        $txtServer.Text = $template.Server
        $txtPort.Text = $template.Port.ToString()
        $chkSSL.Checked = $template.EnableSSL
        
        # Set auth method
        $authIndex = $cmbAuthMethod.Items.IndexOf($template.AuthMethod)
        if ($authIndex -ge 0) {
            $cmbAuthMethod.SelectedIndex = $authIndex
        }
        
        $txtSubject.Text = $template.Subject
        $txtBody.Text = $template.Body
        
        if ($template.IsHtml) {
            $rdoHTML.Checked = $true
        } else {
            $rdoPlainText.Checked = $true
        }
        
        # Show template notes
        [System.Windows.Forms.MessageBox]::Show(
            "Template loaded: $TemplateName`n`n" +
            "Notes:`n$($template.Notes)`n`n" +
            "Please fill in your email address and credentials.",
            "Template Loaded",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        # Focus on From field for user to enter their email
        $txtFrom.Focus()
    }
}

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "SMTP Connection Tester v1.0"
$form.Size = New-Object System.Drawing.Size(720, 850)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.KeyPreview = $true

# Add keyboard shortcuts
$form.Add_KeyDown({
    param($formSender, $e)
    if ($e.Control -and $e.KeyCode -eq 'T') {
        $btnTest.PerformClick()
        $e.Handled = $true
    }
    if ($e.KeyCode -eq 'F1') {
        [System.Windows.Forms.MessageBox]::Show(
            "SMTP Connection Tester - Quick Help`n`n" +
            "Keyboard Shortcuts:`n" +
            "  Ctrl+T - Test Connection`n" +
            "  F1 - This help`n" +
            "  Tab - Navigate between fields`n`n" +
            "Common Ports:`n" +
            "  25 - SMTP (unencrypted)`n" +
            "  587 - SMTP with STARTTLS`n" +
            "  465 - SMTP with SSL/TLS`n`n" +
            "Supported Auth Methods:`n" +
            "  Basic, OAuth2, NTLM, PLAIN, CRAM-MD5, DIGEST-MD5",
            "Quick Help",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        $e.Handled = $true
    }
})

# Create ToolTip
$tooltip = New-Object System.Windows.Forms.ToolTip
$tooltip.AutoPopDelay = 5000
$tooltip.InitialDelay = 500
$tooltip.ReshowDelay = 200
$tooltip.ShowAlways = $true

$yPos = 15

# Menu Strip
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.BackColor = [System.Drawing.Color]::White

$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "&File"

$saveConfigItem = New-Object System.Windows.Forms.ToolStripMenuItem
$saveConfigItem.Text = "&Save Configuration..."
$saveConfigItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control, [System.Windows.Forms.Keys]::S

$loadConfigItem = New-Object System.Windows.Forms.ToolStripMenuItem
$loadConfigItem.Text = "&Load Configuration..."
$loadConfigItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control, [System.Windows.Forms.Keys]::O

$exitItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitItem.Text = "E&xit"
$exitItem.ShortcutKeys = [System.Windows.Forms.Keys]::Alt, [System.Windows.Forms.Keys]::F4

$fileMenu.DropDownItems.AddRange(@($saveConfigItem, $loadConfigItem, (New-Object System.Windows.Forms.ToolStripSeparator), $exitItem))

# Templates Menu
$templatesMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$templatesMenu.Text = "&Templates"

# Create menu items for each template
foreach ($templateName in $script:smtpTemplates.Keys | Sort-Object) {
    $templateItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $templateItem.Text = $templateName
    $templateItem.Tag = $templateName
    $templateItem.Add_Click({
        param($sender, $e)
        Load-SMTPTemplate -TemplateName $sender.Tag
    })
    $templatesMenu.DropDownItems.Add($templateItem)
}

$helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$helpMenu.Text = "&Help"

$helpItem = New-Object System.Windows.Forms.ToolStripMenuItem
$helpItem.Text = "&Quick Help"
$helpItem.ShortcutKeys = [System.Windows.Forms.Keys]::F1

$aboutItem = New-Object System.Windows.Forms.ToolStripMenuItem
$aboutItem.Text = "&About"

$helpMenu.DropDownItems.AddRange(@($helpItem, $aboutItem))

$menuStrip.Items.AddRange(@($fileMenu, $templatesMenu, $helpMenu))
$form.Controls.Add($menuStrip)

$yPos = 35

# Server Settings Group
$grpServer = New-Object System.Windows.Forms.GroupBox
$grpServer.Location = New-Object System.Drawing.Point(15, $yPos)
$grpServer.Size = New-Object System.Drawing.Size(675, 125)
$grpServer.Text = "Server Settings"
$grpServer.ForeColor = [System.Drawing.Color]::DarkBlue
$grpServer.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($grpServer)

# SMTP Server
$lblServer = New-Object System.Windows.Forms.Label
$lblServer.Location = New-Object System.Drawing.Point(15, 28)
$lblServer.Size = New-Object System.Drawing.Size(100, 20)
$lblServer.Text = "SMTP Server:"
$lblServer.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$grpServer.Controls.Add($lblServer)

$txtServer = New-Object System.Windows.Forms.TextBox
$txtServer.Location = New-Object System.Drawing.Point(125, 26)
$txtServer.Size = New-Object System.Drawing.Size(250, 23)
$txtServer.Text = "smtp.gmail.com"
$txtServer.TabIndex = 0
$tooltip.SetToolTip($txtServer, "Enter the SMTP server hostname or IP address")
$grpServer.Controls.Add($txtServer)

# Port
$lblPort = New-Object System.Windows.Forms.Label
$lblPort.Location = New-Object System.Drawing.Point(395, 28)
$lblPort.Size = New-Object System.Drawing.Size(40, 20)
$lblPort.Text = "Port:"
$lblPort.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$grpServer.Controls.Add($lblPort)

$txtPort = New-Object System.Windows.Forms.TextBox
$txtPort.Location = New-Object System.Drawing.Point(440, 26)
$txtPort.Size = New-Object System.Drawing.Size(60, 23)
$txtPort.Text = "587"
$txtPort.TabIndex = 1
$tooltip.SetToolTip($txtPort, "Common ports: 25 (plain), 587 (STARTTLS), 465 (SSL)")
$grpServer.Controls.Add($txtPort)

# SSL Enabled
$chkSSL = New-Object System.Windows.Forms.CheckBox
$chkSSL.Location = New-Object System.Drawing.Point(520, 26)
$chkSSL.Size = New-Object System.Drawing.Size(130, 23)
$chkSSL.Text = "Enable SSL/TLS"
$chkSSL.Checked = $true
$chkSSL.TabIndex = 2
$chkSSL.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$tooltip.SetToolTip($chkSSL, "Enable secure connection using SSL/TLS encryption")
$grpServer.Controls.Add($chkSSL)

# Authentication Method
$lblAuthMethod = New-Object System.Windows.Forms.Label
$lblAuthMethod.Location = New-Object System.Drawing.Point(15, 58)
$lblAuthMethod.Size = New-Object System.Drawing.Size(100, 20)
$lblAuthMethod.Text = "Auth Method:"
$lblAuthMethod.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$grpServer.Controls.Add($lblAuthMethod)

$cmbAuthMethod = New-Object System.Windows.Forms.ComboBox
$cmbAuthMethod.Location = New-Object System.Drawing.Point(125, 56)
$cmbAuthMethod.Size = New-Object System.Drawing.Size(200, 23)
$cmbAuthMethod.DropDownStyle = "DropDownList"
$cmbAuthMethod.Items.AddRange(@("Basic (LOGIN)", "Anonymous", "DefaultCredentials (NTLM)", "PLAIN", "OAuth2", "CRAM-MD5", "DIGEST-MD5"))
$cmbAuthMethod.SelectedIndex = 0
$cmbAuthMethod.TabIndex = 3
$tooltip.SetToolTip($cmbAuthMethod, "Select the authentication method required by your SMTP server")
$grpServer.Controls.Add($cmbAuthMethod)

# OAuth Token Label (initially hidden)
$lblOAuthToken = New-Object System.Windows.Forms.Label
$lblOAuthToken.Location = New-Object System.Drawing.Point(335, 58)
$lblOAuthToken.Size = New-Object System.Drawing.Size(90, 20)
$lblOAuthToken.Text = "OAuth Token:"
$lblOAuthToken.Visible = $false
$lblOAuthToken.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$grpServer.Controls.Add($lblOAuthToken)

$txtOAuthToken = New-Object System.Windows.Forms.TextBox
$txtOAuthToken.Location = New-Object System.Drawing.Point(430, 56)
$txtOAuthToken.Size = New-Object System.Drawing.Size(220, 23)
$txtOAuthToken.Visible = $false
$txtOAuthToken.TabIndex = 4
$tooltip.SetToolTip($txtOAuthToken, "Enter OAuth2 access token for authentication")
$grpServer.Controls.Add($txtOAuthToken)

# Username
$lblUsername = New-Object System.Windows.Forms.Label
$lblUsername.Location = New-Object System.Drawing.Point(15, 88)
$lblUsername.Size = New-Object System.Drawing.Size(100, 20)
$lblUsername.Text = "Username:"
$lblUsername.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$grpServer.Controls.Add($lblUsername)

$txtUsername = New-Object System.Windows.Forms.TextBox
$txtUsername.Location = New-Object System.Drawing.Point(125, 86)
$txtUsername.Size = New-Object System.Drawing.Size(250, 23)
$txtUsername.TabIndex = 5
$tooltip.SetToolTip($txtUsername, "Enter your email username or full email address")
$grpServer.Controls.Add($txtUsername)

# Password
$lblPassword = New-Object System.Windows.Forms.Label
$lblPassword.Location = New-Object System.Drawing.Point(395, 88)
$lblPassword.Size = New-Object System.Drawing.Size(70, 20)
$lblPassword.Text = "Password:"
$lblPassword.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$grpServer.Controls.Add($lblPassword)

$txtPassword = New-Object System.Windows.Forms.TextBox
$txtPassword.Location = New-Object System.Drawing.Point(470, 86)
$txtPassword.Size = New-Object System.Drawing.Size(180, 23)
$txtPassword.PasswordChar = '*'
$txtPassword.TabIndex = 6
$tooltip.SetToolTip($txtPassword, "Enter your password or app-specific password")
$grpServer.Controls.Add($txtPassword)

$yPos += 135

# Email Settings Group
$grpEmail = New-Object System.Windows.Forms.GroupBox
$grpEmail.Location = New-Object System.Drawing.Point(15, $yPos)
$grpEmail.Size = New-Object System.Drawing.Size(675, 225)
$grpEmail.Text = "Email Settings"
$grpEmail.ForeColor = [System.Drawing.Color]::DarkBlue
$grpEmail.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($grpEmail)

# From
$lblFrom = New-Object System.Windows.Forms.Label
$lblFrom.Location = New-Object System.Drawing.Point(15, 28)
$lblFrom.Size = New-Object System.Drawing.Size(100, 20)
$lblFrom.Text = "From:"
$lblFrom.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$grpEmail.Controls.Add($lblFrom)

$txtFrom = New-Object System.Windows.Forms.TextBox
$txtFrom.Location = New-Object System.Drawing.Point(125, 26)
$txtFrom.Size = New-Object System.Drawing.Size(250, 23)
$txtFrom.TabIndex = 7
$tooltip.SetToolTip($txtFrom, "Sender's email address (e.g., sender@example.com)")
$grpEmail.Controls.Add($txtFrom)

# From Name
$lblFromName = New-Object System.Windows.Forms.Label
$lblFromName.Location = New-Object System.Drawing.Point(395, 28)
$lblFromName.Size = New-Object System.Drawing.Size(70, 20)
$lblFromName.Text = "From Name:"
$lblFromName.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$grpEmail.Controls.Add($lblFromName)

$txtFromName = New-Object System.Windows.Forms.TextBox
$txtFromName.Location = New-Object System.Drawing.Point(470, 26)
$txtFromName.Size = New-Object System.Drawing.Size(180, 23)
$txtFromName.TabIndex = 8
$tooltip.SetToolTip($txtFromName, "Display name for the sender (optional)")
$grpEmail.Controls.Add($txtFromName)

# To
$lblTo = New-Object System.Windows.Forms.Label
$lblTo.Location = New-Object System.Drawing.Point(15, 58)
$lblTo.Size = New-Object System.Drawing.Size(100, 20)
$lblTo.Text = "To: *"
$lblTo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$grpEmail.Controls.Add($lblTo)

$txtTo = New-Object System.Windows.Forms.TextBox
$txtTo.Location = New-Object System.Drawing.Point(125, 56)
$txtTo.Size = New-Object System.Drawing.Size(525, 23)
$txtTo.TabIndex = 9
$tooltip.SetToolTip($txtTo, "Recipient email address (required). Separate multiple with semicolons")
$grpEmail.Controls.Add($txtTo)

# CC
$lblCC = New-Object System.Windows.Forms.Label
$lblCC.Location = New-Object System.Drawing.Point(15, 88)
$lblCC.Size = New-Object System.Drawing.Size(100, 20)
$lblCC.Text = "CC:"
$lblCC.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$grpEmail.Controls.Add($lblCC)

$txtCC = New-Object System.Windows.Forms.TextBox
$txtCC.Location = New-Object System.Drawing.Point(125, 86)
$txtCC.Size = New-Object System.Drawing.Size(525, 23)
$txtCC.TabIndex = 10
$tooltip.SetToolTip($txtCC, "Carbon copy recipients (optional). Separate multiple with semicolons")
$grpEmail.Controls.Add($txtCC)

# BCC
$lblBCC = New-Object System.Windows.Forms.Label
$lblBCC.Location = New-Object System.Drawing.Point(15, 118)
$lblBCC.Size = New-Object System.Drawing.Size(100, 20)
$lblBCC.Text = "BCC:"
$lblBCC.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$grpEmail.Controls.Add($lblBCC)

$txtBCC = New-Object System.Windows.Forms.TextBox
$txtBCC.Location = New-Object System.Drawing.Point(125, 116)
$txtBCC.Size = New-Object System.Drawing.Size(525, 23)
$txtBCC.TabIndex = 11
$tooltip.SetToolTip($txtBCC, "Blind carbon copy recipients (optional). Separate multiple with semicolons")
$grpEmail.Controls.Add($txtBCC)

# Reply-To
$lblReplyTo = New-Object System.Windows.Forms.Label
$lblReplyTo.Location = New-Object System.Drawing.Point(15, 148)
$lblReplyTo.Size = New-Object System.Drawing.Size(100, 20)
$lblReplyTo.Text = "Reply-To:"
$lblReplyTo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$grpEmail.Controls.Add($lblReplyTo)

$txtReplyTo = New-Object System.Windows.Forms.TextBox
$txtReplyTo.Location = New-Object System.Drawing.Point(125, 146)
$txtReplyTo.Size = New-Object System.Drawing.Size(250, 23)
$txtReplyTo.TabIndex = 12
$tooltip.SetToolTip($txtReplyTo, "Reply-to email address (optional)")
$grpEmail.Controls.Add($txtReplyTo)

# Priority
$lblPriority = New-Object System.Windows.Forms.Label
$lblPriority.Location = New-Object System.Drawing.Point(395, 148)
$lblPriority.Size = New-Object System.Drawing.Size(70, 20)
$lblPriority.Text = "Priority:"
$lblPriority.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$grpEmail.Controls.Add($lblPriority)

$cmbPriority = New-Object System.Windows.Forms.ComboBox
$cmbPriority.Location = New-Object System.Drawing.Point(470, 146)
$cmbPriority.Size = New-Object System.Drawing.Size(100, 23)
$cmbPriority.DropDownStyle = "DropDownList"
$cmbPriority.Items.AddRange(@("Low", "Normal", "High"))
$cmbPriority.SelectedIndex = 1
$cmbPriority.TabIndex = 13
$tooltip.SetToolTip($cmbPriority, "Set message priority level")
$grpEmail.Controls.Add($cmbPriority)

# Subject
$lblSubject = New-Object System.Windows.Forms.Label
$lblSubject.Location = New-Object System.Drawing.Point(15, 178)
$lblSubject.Size = New-Object System.Drawing.Size(100, 20)
$lblSubject.Text = "Subject:"
$lblSubject.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$grpEmail.Controls.Add($lblSubject)

$txtSubject = New-Object System.Windows.Forms.TextBox
$txtSubject.Location = New-Object System.Drawing.Point(125, 176)
$txtSubject.Size = New-Object System.Drawing.Size(525, 23)
$txtSubject.Text = "SMTP Test Email"
$txtSubject.TabIndex = 14
$tooltip.SetToolTip($txtSubject, "Email subject line")
$grpEmail.Controls.Add($txtSubject)

# Base64 Encoding checkbox
$chkBase64 = New-Object System.Windows.Forms.CheckBox
$chkBase64.Location = New-Object System.Drawing.Point(125, 200)
$chkBase64.Size = New-Object System.Drawing.Size(160, 23)
$chkBase64.Text = "Base64 Encoding"
$chkBase64.TabIndex = 15
$chkBase64.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$tooltip.SetToolTip($chkBase64, "Encode message body in Base64 format")
$grpEmail.Controls.Add($chkBase64)

$yPos += 235

# Message Body Group
$grpBody = New-Object System.Windows.Forms.GroupBox
$grpBody.Location = New-Object System.Drawing.Point(15, $yPos)
$grpBody.Size = New-Object System.Drawing.Size(675, 185)
$grpBody.Text = "Message Body"
$grpBody.ForeColor = [System.Drawing.Color]::DarkBlue
$grpBody.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($grpBody)

# Body Type Radio Buttons
$rdoPlainText = New-Object System.Windows.Forms.RadioButton
$rdoPlainText.Location = New-Object System.Drawing.Point(15, 23)
$rdoPlainText.Size = New-Object System.Drawing.Size(100, 23)
$rdoPlainText.Text = "Plain Text"
$rdoPlainText.Checked = $true
$rdoPlainText.TabIndex = 16
$rdoPlainText.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$tooltip.SetToolTip($rdoPlainText, "Send as plain text email")
$grpBody.Controls.Add($rdoPlainText)

$rdoHTML = New-Object System.Windows.Forms.RadioButton
$rdoHTML.Location = New-Object System.Drawing.Point(125, 23)
$rdoHTML.Size = New-Object System.Drawing.Size(80, 23)
$rdoHTML.Text = "HTML"
$rdoHTML.TabIndex = 17
$rdoHTML.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$tooltip.SetToolTip($rdoHTML, "Send as HTML formatted email")
$grpBody.Controls.Add($rdoHTML)

# Load Template Button
$btnLoadTemplate = New-Object System.Windows.Forms.Button
$btnLoadTemplate.Location = New-Object System.Drawing.Point(220, 21)
$btnLoadTemplate.Size = New-Object System.Drawing.Size(130, 26)
$btnLoadTemplate.Text = "Load Template..."
$btnLoadTemplate.TabIndex = 18
$btnLoadTemplate.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnLoadTemplate.Cursor = [System.Windows.Forms.Cursors]::Hand
$tooltip.SetToolTip($btnLoadTemplate, "Load email content from a file (Ctrl+O)")
$grpBody.Controls.Add($btnLoadTemplate)

# Clear Body Button
$btnClearBody = New-Object System.Windows.Forms.Button
$btnClearBody.Location = New-Object System.Drawing.Point(360, 21)
$btnClearBody.Size = New-Object System.Drawing.Size(90, 26)
$btnClearBody.Text = "Clear"
$btnClearBody.TabIndex = 19
$btnClearBody.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnClearBody.Cursor = [System.Windows.Forms.Cursors]::Hand
$tooltip.SetToolTip($btnClearBody, "Clear message body")
$grpBody.Controls.Add($btnClearBody)

$txtBody = New-Object System.Windows.Forms.TextBox
$txtBody.Location = New-Object System.Drawing.Point(15, 52)
$txtBody.Size = New-Object System.Drawing.Size(645, 120)
$txtBody.Multiline = $true
$txtBody.ScrollBars = "Vertical"
$txtBody.Text = "This is a test email sent from PowerShell SMTP Tester."
$txtBody.TabIndex = 20
$txtBody.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtBody.AcceptsReturn = $true
$tooltip.SetToolTip($txtBody, "Enter your email message body content")
$grpBody.Controls.Add($txtBody)

$yPos += 195

# Attachments Group
$grpAttachments = New-Object System.Windows.Forms.GroupBox
$grpAttachments.Location = New-Object System.Drawing.Point(15, $yPos)
$grpAttachments.Size = New-Object System.Drawing.Size(675, 85)
$grpAttachments.Text = "Attachments"
$grpAttachments.ForeColor = [System.Drawing.Color]::DarkBlue
$grpAttachments.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($grpAttachments)

$lstAttachments = New-Object System.Windows.Forms.ListBox
$lstAttachments.Location = New-Object System.Drawing.Point(15, 23)
$lstAttachments.Size = New-Object System.Drawing.Size(550, 50)
$lstAttachments.TabIndex = 21
$lstAttachments.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$tooltip.SetToolTip($lstAttachments, "List of attached files")
$grpAttachments.Controls.Add($lstAttachments)

$btnAddAttachment = New-Object System.Windows.Forms.Button
$btnAddAttachment.Location = New-Object System.Drawing.Point(575, 23)
$btnAddAttachment.Size = New-Object System.Drawing.Size(85, 28)
$btnAddAttachment.Text = "Add..."
$btnAddAttachment.TabIndex = 22
$btnAddAttachment.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnAddAttachment.Cursor = [System.Windows.Forms.Cursors]::Hand
$tooltip.SetToolTip($btnAddAttachment, "Add file attachments")
$grpAttachments.Controls.Add($btnAddAttachment)

$btnRemoveAttachment = New-Object System.Windows.Forms.Button
$btnRemoveAttachment.Location = New-Object System.Drawing.Point(575, 54)
$btnRemoveAttachment.Size = New-Object System.Drawing.Size(85, 28)
$btnRemoveAttachment.Text = "Remove"
$btnRemoveAttachment.TabIndex = 23
$btnRemoveAttachment.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnRemoveAttachment.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnRemoveAttachment.Cursor = [System.Windows.Forms.Cursors]::Hand
$tooltip.SetToolTip($btnRemoveAttachment, "Remove selected attachment")
$grpAttachments.Controls.Add($btnRemoveAttachment)

$yPos += 95

# Test Button
$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Location = New-Object System.Drawing.Point(260, $yPos)
$btnTest.Size = New-Object System.Drawing.Size(180, 42)
$btnTest.Text = "Test Connection (Ctrl+T)"
$btnTest.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$btnTest.TabIndex = 24
$btnTest.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 204)
$btnTest.ForeColor = [System.Drawing.Color]::White
$btnTest.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnTest.Cursor = [System.Windows.Forms.Cursors]::Hand
$tooltip.SetToolTip($btnTest, "Test SMTP connection with current settings (Ctrl+T)")
$form.Controls.Add($btnTest)

# Status Label
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(15, ($yPos + 50))
$lblStatus.Size = New-Object System.Drawing.Size(675, 20)
$lblStatus.Text = "Ready. Press F1 for help."
$lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$lblStatus.ForeColor = [System.Drawing.Color]::DarkGray
$form.Controls.Add($lblStatus)

# Menu Event Handlers
$exitItem.Add_Click({ $form.Close() })

$helpItem.Add_Click({
    [System.Windows.Forms.MessageBox]::Show(
        "SMTP Connection Tester - Comprehensive Help`n`n" +
        "========================================`n" +
        "KEYBOARD SHORTCUTS:`n" +
        "  Ctrl+T - Test Connection`n" +
        "  Ctrl+S - Save Configuration`n" +
        "  Ctrl+O - Load Configuration`n" +
        "  F1 - Show this help`n" +
        "  Tab - Navigate fields`n`n" +
        "COMMON SMTP PORTS:`n" +
        "  25   - SMTP Plain, often blocked by ISPs`n" +
        "  587  - SMTP with STARTTLS Recommended`n" +
        "  465  - SMTP with implicit SSL/TLS`n`n" +
        "AUTHENTICATION METHODS:`n" +
        "  Basic LOGIN - Standard username/password`n" +
        "  OAuth2 - Token-based Gmail Office 365`n" +
        "  NTLM - Windows integrated auth`n" +
        "  Anonymous - No authentication`n`n" +
        "TIPS:`n" +
        "  - Use app-specific passwords for Gmail`n" +
        "  - Enable Less secure apps if needed`n" +
        "  - Check firewall settings`n" +
        "  - Verify correct port and SSL settings",
        "SMTP Tester - Help",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
})

$aboutItem.Add_Click({
    [System.Windows.Forms.MessageBox]::Show(
        "SMTP Connection Tester v1.0`n`n" +
        "A comprehensive tool for testing SMTP server connections`n" +
        "with support for multiple authentication methods,`n" +
        "SSL/TLS encryption, and detailed logging.`n`n" +
        "Built with PowerShell and Windows Forms`n`n" +
        "Features:`n" +
        "- Multiple authentication methods`n" +
        "- SSL/TLS support`n" +
        "- HTML and plain text emails`n" +
        "- File attachments`n" +
        "- Base64 encoding`n" +
        "- Detailed SMTP transaction logging`n" +
        "- Accessible UI with keyboard shortcuts",
        "About SMTP Tester",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
})

# Event Handlers
$btnAddAttachment.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Multiselect = $true
    $openFileDialog.Title = "Select Files to Attach"
    
    if ($openFileDialog.ShowDialog() -eq "OK") {
        foreach ($file in $openFileDialog.FileNames) {
            if (-not $lstAttachments.Items.Contains($file)) {
                $lstAttachments.Items.Add($file)
            }
        }
    }
})

$btnRemoveAttachment.Add_Click({
    if ($lstAttachments.SelectedIndex -ge 0) {
        $lstAttachments.Items.RemoveAt($lstAttachments.SelectedIndex)
    }
})

# Body Type Radio Button Handlers
$rdoPlainText.Add_CheckedChanged({
    if ($rdoPlainText.Checked) {
        # Radio button is checked - no need to update checkbox
    }
})

$rdoHTML.Add_CheckedChanged({
    if ($rdoHTML.Checked) {
        # Radio button is checked - no need to update checkbox
    }
})

# Load Template Handler
$btnLoadTemplate.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "Load Email Template"
    $openFileDialog.Filter = "Text Files (*.txt)|*.txt|HTML Files (*.html;*.htm)|*.html;*.htm|All Files (*.*)|*.*"
    
    if ($openFileDialog.ShowDialog() -eq "OK") {
        try {
            $content = [System.IO.File]::ReadAllText($openFileDialog.FileName)
            $txtBody.Text = $content
            
            # Auto-detect HTML - fixed regex pattern
            if ($openFileDialog.FileName -match '\.(html?|htm)$' -or $content -match '<html|<body|<div|<p') {
                $rdoHTML.Checked = $true
            } else {
                $rdoPlainText.Checked = $true
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error loading file: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Clear Body Handler
$btnClearBody.Add_Click({
    $txtBody.Clear()
})

$cmbAuthMethod.Add_SelectedIndexChanged({
    $authMethod = $cmbAuthMethod.SelectedItem.ToString()
    $isBasicAuth = $authMethod -match "Basic|PLAIN|CRAM-MD5|DIGEST-MD5"
    $isOAuth = $authMethod -eq "OAuth2"
    $isAnonymous = $authMethod -eq "Anonymous"
    
    $txtUsername.Enabled = $isBasicAuth -or $isOAuth
    $txtPassword.Enabled = $isBasicAuth
    $lblOAuthToken.Visible = $isOAuth
    $txtOAuthToken.Visible = $isOAuth
    
    # Adjust Password field position based on OAuth visibility
    if ($isOAuth) {
        $lblPassword.Visible = $false
        $txtPassword.Visible = $false
    } else {
        $lblPassword.Visible = -not $isAnonymous
        $txtPassword.Visible = -not $isAnonymous
    }
})

$btnTest.Add_Click({
    $lblStatus.Text = "Testing connection..."
    $lblStatus.ForeColor = [System.Drawing.Color]::Orange
    $form.Refresh()
    
    # Validate inputs
    if ([string]::IsNullOrWhiteSpace($txtServer.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter SMTP server address.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        $txtServer.Focus()
        $lblStatus.Text = "Validation failed - SMTP server required"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($txtFrom.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter From address.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        $txtFrom.Focus()
        $lblStatus.Text = "Validation failed - From address required"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($txtTo.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter To address.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        $txtTo.Focus()
        $lblStatus.Text = "Validation failed - To address required"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        return
    }
    
    # Collect configuration
    $config = @{
        Server = $txtServer.Text
        Port = [int]$txtPort.Text
        EnableSSL = $chkSSL.Checked
        AuthMethod = $cmbAuthMethod.SelectedItem.ToString()
        Username = $txtUsername.Text
        Password = $txtPassword.Text
        OAuthToken = $txtOAuthToken.Text
        From = $txtFrom.Text
        FromName = $txtFromName.Text
        To = $txtTo.Text
        CC = $txtCC.Text
        BCC = $txtBCC.Text
        ReplyTo = $txtReplyTo.Text
        Subject = $txtSubject.Text
        Body = $txtBody.Text
        IsHtml = $rdoHTML.Checked
        UseBase64 = $chkBase64.Checked
        Priority = $cmbPriority.SelectedItem
        Attachments = $lstAttachments.Items
    }
    
    # Create modal window for log output
    $logForm = New-Object System.Windows.Forms.Form
    $logForm.Text = "SMTP Connection Log - Real-time Transcript"
    $logForm.Size = New-Object System.Drawing.Size(950, 650)
    $logForm.StartPosition = "CenterScreen"
    $logForm.FormBorderStyle = "Sizable"
    $logForm.MinimumSize = New-Object System.Drawing.Size(800, 500)
    $logForm.Icon = $form.Icon
    
    $logTextBox = New-Object System.Windows.Forms.TextBox
    $logTextBox.Location = New-Object System.Drawing.Point(15, 15)
    $logTextBox.Size = New-Object System.Drawing.Size(905, 550)
    $logTextBox.Multiline = $true
    $logTextBox.ScrollBars = "Both"
    $logTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $logTextBox.ReadOnly = $true
    $logTextBox.BackColor = [System.Drawing.Color]::Black
    $logTextBox.ForeColor = [System.Drawing.Color]::LightGreen
    $logTextBox.Anchor = "Top,Bottom,Left,Right"
    $logTextBox.WordWrap = $false
    $logForm.Controls.Add($logTextBox)
    
    $btnCopyLog = New-Object System.Windows.Forms.Button
    $btnCopyLog.Location = New-Object System.Drawing.Point(320, 575)
    $btnCopyLog.Size = New-Object System.Drawing.Size(120, 32)
    $btnCopyLog.Text = "Copy Log"
    $btnCopyLog.Anchor = "Bottom"
    $btnCopyLog.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnCopyLog.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText($logTextBox.Text)
        [System.Windows.Forms.MessageBox]::Show("Log copied to clipboard!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })
    $logForm.Controls.Add($btnCopyLog)
    
    $btnCloseLog = New-Object System.Windows.Forms.Button
    $btnCloseLog.Location = New-Object System.Drawing.Point(450, 575)
    $btnCloseLog.Size = New-Object System.Drawing.Size(120, 32)
    $btnCloseLog.Text = "Close"
    $btnCloseLog.Anchor = "Bottom"
    $btnCloseLog.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnCloseLog.Add_Click({ 
        $this.FindForm().Close()
    })
    $logForm.Controls.Add($btnCloseLog)
    
    # Show the form
    $logForm.Show()
    $logForm.Refresh()
    
    # Run the test
    $logTextBox.Text = "Initializing SMTP test...`r`n`r`n"
    $logTextBox.Refresh()
    
    $result = Test-SMTPConnection -Config $config -LogTextBox $logTextBox
    
    if ($result) {
        $lblStatus.Text = "[SUCCESS] Test successful - Email sent!"
        $lblStatus.ForeColor = [System.Drawing.Color]::Green
        [System.Windows.Forms.MessageBox]::Show("SMTP test completed successfully!`n`nThe email has been sent.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        $lblStatus.Text = "[ERROR] Test failed - Check log for details"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        [System.Windows.Forms.MessageBox]::Show("SMTP test failed. Please check the log window for detailed error information.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Show form
[void]$form.ShowDialog()
