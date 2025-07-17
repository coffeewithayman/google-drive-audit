# Google Drive Audit Tool - Common Functions
# Copyright (C) 2025 Ayman Elsawah
# 
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU Affero General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Affero General Public License for more details.
#
# You should have received a copy of the GNU Affero General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

from googleapiclient import discovery
from google.oauth2 import service_account

import settings

_credentials = {"_delegated": {}}

SCOPES = {
	"directory": ["https://www.googleapis.com/auth/admin.directory.user.readonly"],
	"audit": ["https://www.googleapis.com/auth/drive.metadata.readonly"],
	"lockdown": ["https://www.googleapis.com/auth/drive"],
	"sheets": ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.file"],
}

PUBLIC_PERMISSION_ID = "anyoneWithLink"
DISCOVERABLE_PERMISSION_ID = "anyoneCanFind"

def credentials(category):
	if category in _credentials:
		return _credentials[category]
	elif category in SCOPES:
		_credentials[category] = service_account.Credentials.from_service_account_file(
			settings.SERVICE_ACCOUNT_FILE,
			scopes=SCOPES[category])
		return _credentials[category]
	else:
		raise

def delegated_credentials(email, category):
	if category not in _credentials["_delegated"]:
		_credentials["_delegated"][category] = {}
	if email in _credentials["_delegated"][category]:
		return _credentials["_delegated"][category][email]
	_credentials["_delegated"][category][email] = credentials(category).with_subject(email)
	return _credentials["_delegated"][category][email]

def collect_paginated(engine, field, **kwargs):
	if settings.DEBUG:
		print(kwargs)
	page_size = 1000
	results = engine(**kwargs).execute()
	page_token = results.get("nextPageToken", None)
	items = results.get(field, [])
	while page_token is not None:
		if settings.DEBUG:
			print("  fetching another page...")
		results = engine(pageToken=page_token, **kwargs).execute()
		page_token = results.get("nextPageToken", None)
		items.extend(results.get(field, []))
	return items

def get_domain_users(fields_override=None):
	if fields_override:
		fields = fields_override
	else:
		fields = "nextPageToken, users(primaryEmail, name)"
	directory = discovery.build(
		"admin",
		"directory_v1",
		credentials=delegated_credentials(settings.ADMIN_USERNAME, "directory"))
	users = collect_paginated(
		directory.users().list,
		"users",
		fields=fields,
		domain=settings.DOMAIN)
	return users

def get_publicly_shared_files(email, fields_override=None, query_override=None):
	if fields_override:
		fields = fields_override
	else:
		fields = "nextPageToken, files(id, name, webViewLink, permissionIds, permissions, modifiedTime)"
	if query_override:
		query = query_override
	else:
		query = "'{email}' in owners and (visibility='{vis1}' or visibility='{vis2}')".format(
			email=email,
			vis1=PUBLIC_PERMISSION_ID,
			vis2=DISCOVERABLE_PERMISSION_ID)
	drive = discovery.build(
		"drive",
		"v3",
		credentials=delegated_credentials(email, "audit"))
	items = collect_paginated(
		drive.files().list,
		"files",
		fields=fields,
		q=query)
	return items

def get_public_role(file):
	public_permission = [p for p in file['permissions'] if p['id'] == 'anyoneWithLink'][0]
	return public_permission['role']

def filter_folders(files):
	return [f for f in files if "folders" in f.get("webViewLink", "")]

def filter_files_unmodified_since(files, cutoff):
	return [f for f in files if f.get("modifiedTime", "") < cutoff]
