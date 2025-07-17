# drive-audit

These scripts perform actions around public Google Drive files.

## Scripts

### audit.py
This performs an audit of all files stored in Google Drive that are shared as "anyone with the link can access", even outside of your domain.
It produces a directory called `out-[timestamp]` with HTML files for emails to be sent to all users with public files.
The template for these emails can be modified in `email_template.html`.
The output emails have the subject line and email address at the top, for easy copy-pasting into Gmail.

Usage:
```
python audit.py
```

### lockdown.py
This automatically removes public sharing of any files last modified more than a month ago that are owned by the specified user.
It retains the share role (e.g. "reader", "editor", etc.) but restricts it to the domain.

Usage:
```
python lockdown.py user@example.com
```

This script stores its output in an output file `out-ld-[email].txt` in a TSV format, in addition to outputting this to the screen as it goes.
It outputs as it goes, which may be useful if the script fails for whatever reason a thousand files into a 2000-file operation.

## Installation

Ensure you have a reasonable version of Python3 (e.g. 3.10.6), pyenv, and pyenv-virtualenv installed.

Set up your virtual environment:
```
$ pyenv virtualenv 3.10.6 driveaudit
$ pyenv activate driveaudit
$ pip install -r requirements.txt
```

## Credentials

### Configuration

Copy `settings.example` to `settings.py` and adjust the settings according to your needs:
```
$ cp settings.example settings.py
$ vim settings.py
```

Then set `DOMAIN` to whatever your domain is, and so on.


### Account Setup

The Google Cloud project you use here doesn't make much of a difference, so use whatever is most contextually appropriate in your organization.

Enable the relevant APIs in the Google Cloud project:
* Admin SDK API (to get the list of users)
* Google Drive API (for obvious reasons)

Create a service account in the Google Cloud project:
* Go to the Google Cloud Console [Service Accounts configuration](https://console.cloud.google.com/iam-admin/serviceaccounts)
* Select a relevant project
* Click "Create Service Account"
  * Give it a descriptive name, e.g. `google-drive-auditor`
  * Skip "Grant this service account access to project" and "Grant users access to this service account"
* Once the account is created, click the "Keys" tab
* Click "Add Key", and download the credentials JSON file
* Copy the file into this directory, and ensure the filename matches `SERVICE_ACCOUNT_FILE` in `settings.py`

### Domain-wide Delegation

We need to grant this service account certain scopes that it can use domain-wide.

* Go to the Google Admin Console [Domain-wide Delegation configuration](https://admin.google.com/ac/owl/domainwidedelegation)
* Click "Add new"
* Enter your service account's client ID
* Enter the following scopes:
  * `https://www.googleapis.com/auth/admin.directory.user.readonly`
    * this is needed to find all users in your domain, which we will iterate over
  * `https://www.googleapis.com/auth/drive.metadata.readonly`
    * this is needed to find publicly-shared files
  * `https://www.googleapis.com/auth/drive`
    * this is needed in order to modify permissions (see `lockdown.py`)

### Enable Google APIs

You will get two errors, one to enable Google SDK API and another to enable Google Drive API.

Click on the links (https://console.developers.google.com/apis/api/drive.googleapis.com/overview?project=<client ID>) to enable via console.


