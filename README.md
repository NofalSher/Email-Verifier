# Email Validator

A Python tool for validating email addresses using SMTP verification. This tool checks email format, domain MX records, and verifies email existence by connecting to the mail server.

## Features

- ✅ Email format validation
- ✅ Domain MX record lookup
- ✅ SMTP server verification
- ✅ Batch processing from Excel files
- ✅ Rate limiting to avoid being blocked
- ✅ Comprehensive logging
- ✅ Results export to Excel

## Installation

### Prerequisites

- Python 3.7 or higher
- pip package manager

### Setup

1. Clone this repository:
```bash
git clone https://github.com/yourusername/email-validator.git
cd email-validator
```

2. Create a virtual environment (recommended):
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

1. Prepare your Excel file with email addresses in a column named 'Email'
2. Place the file in the project directory and name it `input_emails.xlsx`
3. Run the validator:

```bash
python email_validator.py
```

### Custom Configuration

You can modify the configuration in the `main()` function:

```python
# Configuration
INPUT_FILE = 'your_file.xlsx'        # Your input file name
EMAIL_COLUMN = 'Email_Address'       # Your email column name
OUTPUT_FILE = 'results.xlsx'         # Output file name
```

### Using as a Library

```python
from email_validator import EmailValidator, ExcelHandler

# Initialize validator
validator = EmailValidator()

# Validate a single email
is_valid, message = validator.verify_email("test@example.com")
print(f"Valid: {is_valid}, Message: {message}")

# Validate multiple emails
emails = ["test1@example.com", "test2@example.com"]
results = validator.process_email_list(emails)

# Work with Excel files
excel_handler = ExcelHandler()
emails, df = excel_handler.read_emails_from_excel("input.xlsx", "Email")
```

## File Structure

```
email-validator/
├── email_validator.py      # Main application code
├── requirements.txt        # Python dependencies
├── README.md              # This file
├── input_emails.xlsx      # Your input file (not included)
├── verified_emails.xlsx   # Output file (generated)
└── email_verification.log # Log file (generated)
```

## Input File Format

Your Excel file should have a column containing email addresses. Example:

| Email | Name | Company |
|-------|------|---------|
| john@example.com | John Doe | ABC Corp |
| jane@test.com | Jane Smith | XYZ Inc |

## Output

The tool generates:

1. **Excel file** (`verified_emails.xlsx`) with validation results:
   - Original data plus two new columns:
   - `Is_Valid`: Boolean indicating if email is valid
   - `Validation_Message`: Detailed validation message

2. **Log file** (`email_verification.log`) with detailed logs of the validation process

3. **Console output** showing real-time validation progress

## Rate Limiting

The tool includes built-in rate limiting (1-5 second delays between requests) to avoid being blocked by mail servers. You can adjust this in the `process_email_list` method.

## Error Handling

The tool handles various error scenarios:
- Invalid email formats
- Missing MX records
- SMTP connection errors
- Server timeouts
- File reading/writing errors

## Limitations

⚠️ **Important Notes:**

- Some mail servers may block or limit verification requests
- Results may vary depending on the target mail server's configuration
- This tool should be used responsibly and in compliance with anti-spam policies
- Consider the ethical implications of email verification

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Troubleshooting

### Common Issues

1. **"No emails to process"**
   - Check if your Excel file exists and has the correct column name
   - Ensure the email column contains valid data

2. **"SMTP connection error"**
   - Some mail servers block verification attempts
   - Try reducing the validation rate or using a different network

3. **"Permission denied" when writing files**
   - Ensure you have write permissions in the directory
   - Close any open Excel files that might be locking the output file

### Getting Help

If you encounter issues:
1. Check the log file (`email_verification.log`) for detailed error messages
2. Ensure all dependencies are installed correctly
3. Verify your input file format matches the expected structure

## Disclaimer

This tool is for educational and legitimate business purposes only. Users are responsible for complying with applicable laws and regulations regarding email verification and anti-spam policies.