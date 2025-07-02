# Automated Outlook Email Data Extraction System

An autonomous email processing system that continuously monitors Microsoft Outlook mailboxes and extracts structured business data using Google's Gemini AI.

## Features

- 🔄 Continuous monitoring of Outlook mailbox (hourly by default)
- 🤖 AI-powered data extraction using Gemini 1.5 Flash
- 📊 Structured JSON output with standardized formats
- 🔒 Secure OAuth 2.0 authentication
- 📈 Real-time processing statistics and monitoring
- 🚀 RESTful API for integration and control
- 📝 Comprehensive logging and error handling

## Prerequisites

- Python 3.8 or higher
- Microsoft Azure account with registered application
- Google Cloud Platform account with Vertex AI enabled
- Outlook mailbox with appropriate permissions

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd email-extraction-system
```

2. Create and activate a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Create a `.env` file with your configuration:
```env
# Microsoft Graph API Configuration
OUTLOOK_CLIENT_ID=your_app_registration_id
OUTLOOK_CLIENT_SECRET=your_client_secret
OUTLOOK_TENANT_ID=your_tenant_id

# Google Vertex AI Configuration
GEMINI_API_KEY=your_vertex_ai_key
GEMINI_PROJECT_ID=your_gcp_project

# Optional Settings
PROCESSING_INTERVAL=3600  # seconds (default: 1 hour)
EMAIL_FOLDER=Inbox       # Outlook folder to monitor
BATCH_SIZE=10            # emails per processing batch
LOG_LEVEL=INFO          # logging verbosity
```

## Usage

1. Start the server:
```bash
python main.py
```

2. The system will automatically:
   - Monitor your Outlook mailbox every hour
   - Process unread emails
   - Extract structured data using Gemini AI
   - Save results to JSON files in the `output` directory
   - Mark processed emails as read

3. Access the API endpoints:
   - `GET /`: System status and next run time
   - `GET /stats`: Processing statistics
   - `POST /process-now`: Trigger immediate processing
   - `GET /health`: Health check endpoint

## API Documentation

Once the server is running, visit:
```
http://localhost:8000/docs
```
for interactive API documentation.

## Output Format

The system generates JSON files with the following structure:
```json
{
  "customer_name": "string",
  "company_name": "string",
  "email": "email_format",
  "contact_number": "phone_format",
  "from_address": "address_string",
  "to_address": "address_string",
  "country": "country_code",
  "currency": "currency_code",
  "products": ["array_of_strings"],
  "quantity": "number",
  "mode_of_transportation": "Air|Water|Road",
  "expected_delivery_date": "YYYY-MM-DD",
  "confidence_score": "float_0_to_1",
  "processing_timestamp": "ISO_timestamp",
  "email_id": "outlook_message_id"
}
```

## Monitoring

- Logs are stored in `logs/email_processor.log`
- Processing statistics are available via the `/stats` endpoint
- Health checks can be performed via the `/health` endpoint

## Error Handling

The system includes:
- Automatic retry with exponential backoff
- Comprehensive error logging
- Dead letter queue for failed extractions
- Health monitoring and alerts

## Security

- OAuth 2.0 authentication for Outlook access
- Secure credential management
- CORS protection
- Rate limiting
- Input validation

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For support, please:
1. Check the documentation
2. Review the logs
3. Open an issue in the repository

## Acknowledgments

- Microsoft Graph API
- Google Vertex AI
- FastAPI
- Python community

# Email Extraction API

API for automated email data extraction using Gemini AI with support for single and multiple email accounts.

## Features

- ✅ **Single Account Processing**: Process emails from one account (backward compatible)
- ✅ **Multi-Account Processing**: Process emails from multiple accounts simultaneously
- ✅ **Parallel Processing**: Concurrent processing of emails within accounts and across accounts
- ✅ **Rate Limiting**: Built-in rate limiting to prevent API throttling
- ✅ **Comprehensive Logging**: Detailed logging for troubleshooting
- ✅ **RESTful API**: Easy-to-use REST endpoints
- ✅ **Real-time Monitoring**: Live stats and processing information

## Quick Start

### Single Account Setup (Existing Behavior)

1. Copy `env.example` to `.env`
2. Configure a single email account:
```bash
OUTLOOK_USER_ID=your.email@domain.com
```

### Multi-Account Setup (New Feature)

1. Copy `env.example` to `.env`
2. Configure multiple email accounts using JSON array format:
```bash
# Comment out or remove OUTLOOK_USER_ID
# OUTLOOK_USER_ID=your.email@domain.com

# Add multiple accounts
OUTLOOK_USER_IDS=["user1@domain.com", "user2@domain.com", "user3@domain.com"]
```

## Multi-Account Configuration

### Environment Variables

| Variable | Description | Example |
|----------|-------------|---------|
| `OUTLOOK_USER_IDS` | JSON array of email accounts | `["user1@domain.com", "user2@domain.com"]` |
| `BATCH_SIZE_PER_ACCOUNT` | Emails to process per account | `5` |
| `MAX_CONCURRENT_ACCOUNTS` | Max accounts processed simultaneously | `3` |

### Configuration Examples

#### Example 1: Small Team (2-3 accounts)
```bash
OUTLOOK_USER_IDS=["manager@company.com", "support@company.com", "sales@company.com"]
BATCH_SIZE_PER_ACCOUNT=10
MAX_CONCURRENT_ACCOUNTS=3
```

#### Example 2: Large Organization (5+ accounts)
```bash
OUTLOOK_USER_IDS=["dept1@company.com", "dept2@company.com", "dept3@company.com", "dept4@company.com", "dept5@company.com"]
BATCH_SIZE_PER_ACCOUNT=5
MAX_CONCURRENT_ACCOUNTS=3
```

## API Endpoints

### Multi-Account Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| `GET` | `/accounts` | List configured email accounts |
| `POST` | `/process-now` | Process all configured accounts |
| `POST` | `/process-account/{user_id}` | Process specific account |
| `GET` | `/stats` | Get processing statistics |

### Usage Examples

#### Get Account Configuration
```bash
curl http://localhost:8000/accounts
```

Response:
```json
{
  "accounts": ["user1@domain.com", "user2@domain.com"],
  "total_count": 2,
  "max_concurrent": 3,
  "batch_size_per_account": 5
}
```

#### Process All Accounts
```bash
curl -X POST http://localhost:8000/process-now
```

#### Process Specific Account
```bash
curl -X POST http://localhost:8000/process-account/user1@domain.com
```

Response:
```json
{
  "user_id": "user1@domain.com",
  "success": true,
  "processed_count": 3,
  "results": [...],
  "error": null
}
```

## Output Format

### Multi-Account Output Structure

The system saves results in enhanced format with account-level tracking:

```json
{
  "summary": {
    "timestamp": "20241201_143022",
    "total_accounts": 3,
    "successful_accounts": 2,
    "failed_accounts": 1,
    "total_emails_processed": 15,
    "average_confidence": 0.85
  },
  "account_results": [
    {
      "user_id": "user1@domain.com",
      "success": true,
      "processed_count": 8,
      "results": [...],
      "error": null
    },
    {
      "user_id": "user2@domain.com", 
      "success": true,
      "processed_count": 7,
      "results": [...],
      "error": null
    },
    {
      "user_id": "user3@domain.com",
      "success": false,
      "processed_count": 0,
      "results": [],
      "error": "Authentication failed"
    }
  ],
  "all_extractions": [...]
}
```

### Individual Email Results

Each processed email includes account information:

```json
{
  "user_id": "user1@domain.com",
  "email_id": "AAMkAGE...",
  "sender": "John Doe <john@example.com>",
  "subject": "Invoice Payment",
  "received_datetime": "2024-12-01T14:30:22Z",
  "processing_timestamp": "2024-12-01T14:35:45Z",
  "confidence_score": 0.92,
  // ... extracted data fields
}
```

## Performance Considerations

### Concurrent Processing

The system processes accounts concurrently with configurable limits:

- **Account-level concurrency**: `MAX_CONCURRENT_ACCOUNTS` accounts processed simultaneously
- **Email-level concurrency**: Emails within each account processed in parallel
- **Rate limiting**: Built-in delays to prevent API throttling

### Recommended Settings

| Organization Size | Accounts | Batch Size | Max Concurrent |
|-------------------|----------|------------|----------------|
| Small (1-5 accounts) | 1-5 | 10 | 3 |
| Medium (5-15 accounts) | 5-15 | 5 | 3 |
| Large (15+ accounts) | 15+ | 3 | 2 |

## Migration from Single Account

### Backward Compatibility

Existing single-account configurations continue to work without changes. The system automatically detects the configuration type:

1. If `OUTLOOK_USER_IDS` is set → Multi-account mode
2. If only `OUTLOOK_USER_ID` is set → Single-account mode (legacy)

### Migration Steps

1. **Backup your current `.env` file**
2. **Update configuration**:
   ```bash
   # Old (still works)
   OUTLOOK_USER_ID=your.email@domain.com
   
   # New (recommended)
   OUTLOOK_USER_IDS=["your.email@domain.com", "additional@domain.com"]
   ```
3. **Test with `/accounts` endpoint** to verify configuration
4. **Monitor processing** with enhanced statistics

## Troubleshooting

### Common Issues

1. **Authentication Failures**
   - Ensure all accounts have proper Microsoft Graph permissions
   - Check tenant ID matches for all accounts

2. **Rate Limiting**
   - Reduce `BATCH_SIZE_PER_ACCOUNT`
   - Decrease `MAX_CONCURRENT_ACCOUNTS`

3. **Memory Usage**
   - Lower concurrent processing limits for large account counts

### Monitoring

Monitor multi-account processing through:
- Enhanced `/stats` endpoint
- Detailed logs with account-specific information
- Output files with account-level summaries

## Security Notes

- All accounts must be within the same Microsoft tenant
- Each account requires individual authentication via the OAuth flow
- Token files are shared across accounts (same tenant)

## Support

For issues specific to multi-account processing, include:
1. Number of configured accounts
2. Concurrent processing settings
3. Account-specific error messages from logs 