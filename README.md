# AI-Powered Email Automation Agent

A professional email automation tool designed for recruitment outreach, featuring AI-generated subject lines, personalized templates, and intelligent delivery management.

## Features

- **AI-Powered Subject Lines**: Leverages OpenAI's GPT-4 to generate contextually relevant, professional subject lines
- **Multi-Format Support**: Compatible with `.csv`, `.xlsx`, `.txt`, `.docx`, and `.html` files
- **HTML Email Support**: Send richly formatted HTML emails with inline logo embedding
- **Smart Personalization**: Dynamic placeholder replacement for names, job titles, and locations
- **Spam Prevention**: Built-in headers (Message-ID, Reply-To) to improve deliverability
- **Automatic Sent Folder Sync**: Saves sent emails to your IMAP Sent folder
- **Connection Resilience**: Automatic reconnection on SMTP disconnects
- **Rate Limiting**: Configurable delays (3-7 seconds) to protect sender reputation
- **Comprehensive Logging**: Real-time progress tracking with detailed status updates

## Prerequisites

- Python 3.7+
- OpenAI API key
- Email account with SMTP/IMAP access (GoDaddy configuration included)
- Properly configured SPF/DKIM records for your sending domain

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/email-automation-agent.git
cd email-automation-agent
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

3. Create a `.env` file in the project root:
```env
OPENAI_API_KEY=your_openai_api_key_here
EMAIL_USER=your_email@example.com
EMAIL_PASSWORD=your_email_password
SMTP_SERVER=smtpout.secureserver.net
SMTP_PORT=465
```

## Dependencies

Create a `requirements.txt` file with the following:

```
openai>=1.0.0
python-dotenv>=1.0.0
python-docx>=0.8.11
openpyxl>=3.1.0
```

## Usage

### 1. Prepare Your Contacts File

Create a contacts file (`.csv` or `.xlsx`) with the following columns:

| Column Name | Required | Description |
|------------|----------|-------------|
| `email` | Yes | Recipient email address |
| `name` | No | Recipient's name (defaults to "Hiring Manager") |
| `job title` or `position` | No | Job title being recruited for |
| `country` | No | Recipient's country (defaults to "US") |

**Example CSV:**
```csv
email,name,job title,country
recruiter@example.com,John Smith,Senior Software Engineer,UK
hiring@company.com,Jane Doe,Product Manager,US
```

### 2. Create Your Email Template

Create a template file (`.txt`, `.docx`, or `.html`) with placeholders:

**Available Placeholders:**
- `[Name]` or `[NAME]` - Recipient's name
- `[Job Title]`, `[Position]`, or `[JOB POSITION]` - Job position
- `[Country]` or `[COUNTRY]` - Recipient's country

**Example Template (template.txt):**
```
Hi [Name],

I noticed your opening for a [Job Title] position in [Country]. I specialize in connecting 
top-tier candidates with leading firms in the recruitment space.

I'd love to explore how we can support your hiring needs.

Best regards,
Zainab
```

**HTML Template Example (template.html):**
```html
<html>
<body>
    <img src="cid:company_logo" alt="Logo" width="150"><br><br>
    <p>Hi [Name],</p>
    <p>I noticed your opening for a <strong>[Job Title]</strong> position in [Country]...</p>
    <p>Best regards,<br>Zainab</p>
</body>
</html>
```

### 3. Run the Agent

```bash
python email_agent.py
```

Follow the interactive prompts:
1. Enter your contacts filename
2. Enter your template filename
3. (Optional) Enter logo filename for HTML emails
4. Confirm to start sending

## Configuration

### SMTP Settings

The default configuration uses GoDaddy's SMTP servers. To use a different provider, update your `.env`:

```env
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
```

### Delay Configuration

To modify the sending delay, edit the `main()` function:

```python
delay = random.uniform(3, 7)  # Adjust min/max seconds
```

## Best Practices

1. **Domain Authentication**: Ensure SPF, DKIM, and DMARC records are properly configured
2. **Warm-Up Period**: Start with small batches (10-20 emails) if using a new domain
3. **Monitor Deliverability**: Check bounce rates and spam reports regularly
4. **Personalization**: Always use real data in placeholders to avoid generic messages
5. **Rate Limiting**: Avoid sending more than 50-100 emails per hour on shared hosting

## Security Considerations

- Never commit your `.env` file to version control
- Use app-specific passwords where available
- Regularly rotate API keys and email credentials
- Review sent emails for data accuracy before mass deployment

## Troubleshooting

### Connection Issues
```
❌ SMTP Connection Failed: [Errno 11001] getaddrinfo failed
```
**Solution**: Verify SMTP_SERVER and SMTP_PORT in `.env`

### Authentication Errors
```
❌ SMTP Connection Failed: (535, b'Authentication failed')
```
**Solution**: Check EMAIL_USER and EMAIL_PASSWORD credentials

### Emails Going to Spam
**Solutions**:
- Verify SPF/DKIM records
- Reduce sending rate
- Improve email content quality
- Warm up your sending domain

### Invalid Email Column
```
❌ Error: CSV must contain an 'Email' column
```
**Solution**: Ensure your contacts file has a column named `email`, `Email`, or `Email ID`

## Limitations

- Maximum 5MB per email (including attachments)
- Recommended batch size: 100-200 emails per session
- IMAP Sent folder sync may fail on some providers

## Contributing

Contributions are welcome! Please follow these steps:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Disclaimer

This tool is intended for legitimate business communication only. Users are responsible for:
- Compliance with anti-spam laws (CAN-SPAM, GDPR, etc.)
- Obtaining proper consent for email outreach
- Adhering to email service provider terms of service

Misuse of this tool for spam or unsolicited bulk email is strictly prohibited.

## Support

For issues, questions, or feature requests, please open an issue on GitHub.

---

**Built with ❤️ for recruitment professionals**
