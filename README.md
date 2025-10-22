# SMTP Tester

A comprehensive PowerShell-based SMTP connection testing tool with a user-friendly GUI interface.

## Features

- **Pre-configured Templates**: Quick setup for Gmail, Outlook.com, Office 365, ProtonMail, Yahoo Mail, SendGrid, Mailgun, and Amazon SES
- **Multiple Authentication Methods**: Basic (LOGIN), OAuth2, NTLM, PLAIN, CRAM-MD5, DIGEST-MD5, and Anonymous
- **SSL/TLS Support**: Secure connections with configurable encryption
- **Email Formats**: Support for both HTML and plain text emails
- **File Attachments**: Attach multiple files to test emails
- **Base64 Encoding**: Optional Base64 encoding for message body
- **Detailed Logging**: Real-time SMTP transaction logging with detailed error information
- **Accessible UI**: Keyboard shortcuts and tooltips for enhanced usability
- **Configuration Management**: Save and load SMTP configurations

## Requirements

- Windows PowerShell 5.1 or later
- .NET Framework (included with Windows)

## Installation

1. Download `SMTPTester.ps1`
2. No installation required - just run the script!

## Usage

### Running the Script

```powershell
.\SMTPTester.ps1
```

Or right-click the file and select "Run with PowerShell"

### Keyboard Shortcuts

- **Ctrl+T** - Test Connection
- **Ctrl+S** - Save Configuration
- **Ctrl+O** - Load Configuration
- **F1** - Show Help
- **Tab** - Navigate between fields

## Using Service Templates

The application includes pre-configured templates for popular email services. Access them via the **Templates** menu:

### Available Templates

1. **Gmail** - Google's email service
2. **Outlook.com / Hotmail** - Microsoft consumer email
3. **Office 365** - Microsoft business email
4. **ProtonMail Bridge** - Secure email (requires Bridge app)
5. **Yahoo Mail** - Yahoo's email service
6. **SendGrid** - Email delivery service
7. **Mailgun** - Email API service
8. **Amazon SES** - AWS email service

### How to Use Templates

1. Click **Templates** in the menu bar
2. Select your email service
3. Read the template notes (important setup information)
4. Fill in your email address in the "From" field
5. Fill in your credentials (username/password)
6. Add a recipient in the "To" field
7. Click **Test Connection**

**Note:** Each template includes specific notes about authentication requirements (e.g., app passwords, API keys).

### Common SMTP Ports

- **25** - SMTP (Plain, often blocked by ISPs)
- **587** - SMTP with STARTTLS (Recommended)
- **465** - SMTP with implicit SSL/TLS

## Configuration Examples

### Gmail

- **Server**: smtp.gmail.com
- **Port**: 587
- **SSL/TLS**: Enabled
- **Auth Method**: Basic (LOGIN) or OAuth2
- **Username**: your-email@gmail.com
- **Password**: App-specific password (recommended)

### Office 365

- **Server**: smtp.office365.com
- **Port**: 587
- **SSL/TLS**: Enabled
- **Auth Method**: Basic (LOGIN) or OAuth2
- **Username**: your-email@company.com
- **Password**: Your password or app password

### Generic SMTP

- **Server**: Your SMTP server hostname
- **Port**: 587 or 465
- **SSL/TLS**: Check with your provider
- **Auth Method**: Usually Basic (LOGIN)
- **Username**: Your email or username
- **Password**: Your password

## Authentication Methods

### Basic (LOGIN)
Standard username and password authentication. Most commonly used.

### OAuth2
Token-based authentication. Required for Gmail and Office 365 with modern auth.

### NTLM (Default Credentials)
Uses your Windows credentials. Typically for internal mail servers.

### Anonymous
No authentication required. Rare in modern environments.

## Troubleshooting

### Authentication Errors
- Verify your username and password are correct
- For Gmail, use an app-specific password
- Check if "Less secure app access" needs to be enabled (not recommended)
- Try OAuth2 if Basic authentication fails

### Connection Errors
- Verify the SMTP server address is correct
- Check that you're using the correct port
- Try toggling the SSL/TLS setting
- Ensure your firewall allows outbound connections on the SMTP port

### SSL/TLS Errors
- Port 587 typically uses STARTTLS
- Port 465 typically uses implicit SSL
- Try both with SSL enabled

## Security Notes

- Never commit files containing real passwords
- Use app-specific passwords when possible
- Store sensitive configurations securely
- This tool is for testing purposes - use production credentials carefully

## License

This project is provided as-is for educational and testing purposes.

## Contributing

Feel free to submit issues and enhancement requests!

## Author

PowerShell SMTP Tester - A tool for network administrators and developers
