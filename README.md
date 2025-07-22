# Cool Drive Audit

This tool performs comprehensive audits of Google Drive files shared publicly across your Google Workspace domain, with options for automated remediation and detailed reporting.

## Features

- **Domain-wide Audit**: Scan all users in your Google Workspace domain for publicly shared files
- **Multiple Output Formats**: Console output, HTML reports, and Google Sheets integration
- **Configurable Fields**: Display file names, sharing links, file IDs, and modification dates
- **Error Handling**: Robust error handling with detailed debug mode
- **API Validation**: Early detection of missing API permissions with direct console links
- **Automated Lockdown**: Remove public sharing from stale files while preserving access levels

## Scripts

### audit.py
Performs a comprehensive audit of all files stored in Google Drive that are shared as "anyone with the link can access", even outside of your domain.

**Basic Usage:**
```bash
python3 audit.py
```

**Advanced Usage:**
```bash
# Show specific fields
python3 audit.py -f name link id modified

# Generate Google Sheets report
python3 audit.py --sheets

# Audit shared drives only
python3 audit.py --shared-drives-only

# Console output only (no HTML)
python3 audit.py --no-html -f name link

# Debug mode with detailed error messages
python3 audit.py --debug --sheets
```

**Output Options:**
- **Console**: Real-time progress with customizable field display
- **HTML Reports**: Individual HTML files for each user (in `out-YYYYMMDD-HHMMSS/` directory)
- **Google Sheets**: Collaborative spreadsheet with separate tabs for each user

### lockdown.py
Automatically removes public sharing from files last modified more than the grace period (default: 180 days) that are owned by the specified user. It retains the share role (e.g. "reader", "editor") but restricts access to your domain only.

**Usage:**
```bash
python3 lockdown.py user@yourdomain.com
```

**Features:**
- **Grace Period**: Only processes files older than `LOCKDOWN_GRACE_DAYS` setting
- **Dry Run Mode**: Shows what would be changed without making actual changes
- **TSV Output**: Logs all changes to `out-ld-[email]-[timestamp].tsv`
- **Role Preservation**: Maintains original permission levels (viewer/editor/etc.)

## Installation

### Prerequisites
- Python 3.10+ recommended
- Google Workspace Admin access
- Google Cloud Project with appropriate APIs enabled

### Setup
```bash
# Set up virtual environment (recommended)
pyenv virtualenv 3.10.6 driveaudit
pyenv activate driveaudit

# Install dependencies
pip install -r requirements.txt

# Configure settings
cp settings.example settings.py
vim settings.py  # Update DOMAIN, ADMIN_USERNAME, SERVICE_ACCOUNT_FILE
```

## Google Cloud Setup

### 1. Enable Required APIs
The tool will automatically detect missing APIs and provide direct console links to enable them:

- **Admin SDK API**: `https://console.developers.google.com/apis/api/admin.googleapis.com/overview?project=YOUR_PROJECT_ID`
- **Google Drive API**: `https://console.developers.google.com/apis/api/drive.googleapis.com/overview?project=YOUR_PROJECT_ID`
- **Google Sheets API**: `https://console.developers.google.com/apis/api/sheets.googleapis.com/overview?project=YOUR_PROJECT_ID` (only if using `--sheets`)

### 2. Create Service Account
1. Go to [Google Cloud Console > Service Accounts](https://console.cloud.google.com/iam-admin/serviceaccounts)
2. Select your project
3. Click "Create Service Account"
4. Give it a descriptive name (e.g., `google-drive-auditor`)
5. Skip the optional steps
6. Click "Keys" tab → "Add Key" → Download JSON file
7. Place the JSON file in your project directory
8. Update `SERVICE_ACCOUNT_FILE` in `settings.py`

### 3. Domain-wide Delegation
Configure the service account to impersonate users across your domain:

1. Go to [Google Workspace Admin Console > Domain-wide Delegation](https://admin.google.com/ac/owl/domainwidedelegation)
2. Click "Add new"
3. Enter your service account's Client ID (from the JSON file)
4. Add these OAuth scopes:
   ```
   https://www.googleapis.com/auth/admin.directory.user.readonly,https://www.googleapis.com/auth/drive.metadata.readonly,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/drive.file
   ```

**Scope Purposes:**
- `admin.directory.user.readonly`: List all users in your domain
- `drive.metadata.readonly`: Read file metadata and sharing permissions
- `drive`: Modify file permissions (for lockdown.py)
- `spreadsheets`: Create and edit Google Sheets reports
- `drive.file`: Create files in Google Drive (for Sheets reports)

## Configuration

Edit `settings.py` with your domain details:

```python
DEBUG = False
DOMAIN = "yourdomain.com"
ADMIN_USERNAME = "admin@yourdomain.com"  # Must be a domain admin
SERVICE_ACCOUNT_FILE = "your-service-account-file.json"
LOCKDOWN_GRACE_DAYS = 180  # Days before files are eligible for lockdown
```

## Command Line Options

### audit.py Options
- `--fields`, `-f`: Output fields (name, link, id, modified)
- `--no-html`: Skip HTML report generation
- `--sheets`: Create Google Sheets report
- `--shared-drives-only`: Audit shared drives only, skip individual user files
- `--debug`: Enable debug mode with detailed error messages
- `--help`: Show help message with examples

### Examples
```bash
# Basic audit with default output
python3 audit.py

# Custom fields with Google Sheets
python3 audit.py -f name link modified --sheets

# Audit only shared drives
python3 audit.py --shared-drives-only

# Debug mode for troubleshooting
python3 audit.py --debug --sheets

# Console-only output
python3 audit.py --no-html -f name link
```

## Troubleshooting

### Common Issues

1. **API Not Enabled**: The tool will show direct console links to enable required APIs
2. **Permission Denied**: Ensure service account has domain-wide delegation with correct scopes
3. **Invalid Admin User**: Verify `ADMIN_USERNAME` is a valid domain administrator
4. **Authentication Errors**: Check that service account JSON file is valid and accessible

### Debug Mode
Use `--debug` flag for detailed error information:
```bash
python3 audit.py --debug
```

This provides:
- Full stack traces for all errors
- API call debugging information
- Step-by-step validation progress
- Detailed Google API error messages

## License

This project is licensed under the GNU Affero General Public License v3.0 (AGPL-3.0).

### Key License Terms:
- **Open Source**: Free to use, modify, and distribute
- **Copyleft**: Modifications must be released under the same license
- **Network Use**: If you run this as a web service, you must provide source code to users
- **No Commercial SaaS**: Cannot be used to create commercial or SaaS products without releasing full source code

See the [LICENSE](LICENSE) file for complete terms.

### Why AGPL-3.0?
This license ensures that improvements to the tool benefit the entire community while preventing commercial entities from creating proprietary SaaS offerings based on this code.

## Contributing

Contributions are welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

All contributions must be compatible with AGPL-3.0 licensing.

## Support

For issues and questions:
1. Check the troubleshooting section above
2. Use `--debug` mode for detailed error information
3. Review Google Cloud Console for API and permission issues
4. Open an issue on the project repository

## Security

This tool requires broad Google Workspace permissions. Ensure:
- Service account JSON files are kept secure
- Access is restricted to authorized administrators
- Regular security reviews of domain-wide delegation permissions
- Monitoring of audit tool usage in your environment
