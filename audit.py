# Cool Drive Audit
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

import argparse
import datetime
import os
import sys
import traceback

import common
import settings

settings.DEBUG = False


def log_error(error, context=""):
	"""Log errors to errors.txt file"""
	error_file = "errors.txt"
	
	error_entry = f"""
=== ERROR at {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")} ===
Context: {context}
Error: {str(error)}
Traceback:
{traceback.format_exc()}
==================================================

"""
	
	try:
		with open(error_file, 'a') as f:
			f.write(error_entry)
		print(f"Error logged to: {error_file}")
	except Exception as log_error:
		print(f"Failed to log error: {log_error}")

def format_file_output(file, fields):
	"""Format file output based on selected fields"""
	output_parts = []
	
	if 'name' in fields:
		name = file['name']
		if file.get('isSharedDriveFile'):
			name += f" (Shared Drive: {file['sharedDriveName']})"
		output_parts.append(name)
	if 'link' in fields:
		output_parts.append(file['webViewLink'])
	if 'id' in fields:
		output_parts.append(file['id'])
	if 'modified' in fields:
		output_parts.append(file.get('modifiedTime', 'N/A'))
	
	return " | ".join(output_parts)

def parse_api_error(error):
	"""Parse Google API error to extract console activation URL"""
	error_str = str(error)
	
	# Look for activation URL in error message
	if 'console.developers.google.com' in error_str:
		import re
		url_match = re.search(r'https://console\.developers\.google\.com[^\s\'"]*', error_str)
		if url_match:
			return url_match.group(0)
	
	return None

def test_admin_api():
	"""Test Admin SDK Directory API"""
	try:
		from googleapiclient import discovery
		if settings.DEBUG:
			print("DEBUG: Testing Admin SDK Directory API...")
		directory = discovery.build(
			"admin", "directory_v1", 
			credentials=common.delegated_credentials(settings.ADMIN_USERNAME, "directory"))
		# Minimal test - just check if we can access the API
		directory.users().list(domain=settings.DOMAIN, maxResults=1).execute()
		if settings.DEBUG:
			print("DEBUG: Admin SDK Directory API test passed")
		return True
	except Exception as e:
		log_error(e, "Admin SDK Directory API test")
		activation_url = parse_api_error(e)
		if activation_url:
			print("ERROR: Admin SDK Directory API is not enabled.")
			print("Enable it here: {}".format(activation_url))
		else:
			print("ERROR: Admin SDK Directory API test failed: {}".format(str(e)))
		if settings.DEBUG:
			import traceback
			print("DEBUG: Full traceback:")
			traceback.print_exc()
		return False

def test_drive_api():
	"""Test Google Drive API"""
	try:
		from googleapiclient import discovery
		if settings.DEBUG:
			print("DEBUG: Testing Google Drive API...")
		drive = discovery.build(
			"drive", "v3", 
			credentials=common.delegated_credentials(settings.ADMIN_USERNAME, "audit"))
		# Minimal test - just check if we can access the API
		drive.files().list(q="'me' in owners", pageSize=1).execute()
		if settings.DEBUG:
			print("DEBUG: Google Drive API test passed")
		return True
	except Exception as e:
		log_error(e, "Google Drive API test")
		activation_url = parse_api_error(e)
		if activation_url:
			print("ERROR: Google Drive API is not enabled.")
			print("Enable it here: {}".format(activation_url))
		else:
			print("ERROR: Google Drive API test failed: {}".format(str(e)))
		if settings.DEBUG:
			import traceback
			print("DEBUG: Full traceback:")
			traceback.print_exc()
		return False

def test_sheets_api():
	"""Test Google Sheets API"""
	try:
		from googleapiclient import discovery
		if settings.DEBUG:
			print("DEBUG: Testing Google Sheets API...")
		sheets_service = discovery.build(
			'sheets', 'v4', 
			credentials=common.delegated_credentials(settings.ADMIN_USERNAME, 'sheets'))
		# Minimal test - just check if we can access the API
		sheets_service.spreadsheets().create(body={'properties': {'title': 'API Test'}}).execute()
		if settings.DEBUG:
			print("DEBUG: Google Sheets API test passed")
		return True
	except Exception as e:
		log_error(e, "Google Sheets API test")
		activation_url = parse_api_error(e)
		if activation_url:
			print("ERROR: Google Sheets API is not enabled.")
			print("Enable it here: {}".format(activation_url))
		else:
			print("ERROR: Google Sheets API test failed: {}".format(str(e)))
		if settings.DEBUG:
			import traceback
			print("DEBUG: Full traceback:")
			traceback.print_exc()
		return False

def validate_apis(use_sheets=False):
	"""Validate all required APIs are enabled"""
	print("Validating Google APIs...")
	
	all_good = True
	
	# Test Admin SDK (always required)
	if not test_admin_api():
		all_good = False
	
	# Test Drive API (always required)
	if not test_drive_api():
		all_good = False
	
	# Test Sheets API (only if --sheets is used)
	if use_sheets and not test_sheets_api():
		all_good = False
	
	if not all_good:
		print("\nPlease enable the required APIs using the links above, then re-run the script.")
		return False
	
	print("All required APIs are enabled.\n")
	return True

def create_google_sheets_report(user_data, admin_email):
	"""Create a Google Sheets report with separate tabs for each user"""
	from googleapiclient import discovery
	
	# Create the spreadsheet
	sheets_service = discovery.build('sheets', 'v4', credentials=common.delegated_credentials(admin_email, 'sheets'))
	
	# Create spreadsheet
	spreadsheet_body = {
		'properties': {
			'title': 'Cool Drive Audit Report - {}'.format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
		}
	}
	
	spreadsheet = sheets_service.spreadsheets().create(body=spreadsheet_body).execute()
	spreadsheet_id = spreadsheet['spreadsheetId']
	
	# Rename default "Sheet1" to "Dashboard"
	sheets_service.spreadsheets().batchUpdate(
		spreadsheetId=spreadsheet_id,
		body={'requests': [{'updateSheetProperties': {
			'properties': {'sheetId': 0, 'title': 'Dashboard'},
			'fields': 'title'
		}}]}
	).execute()
	
	# Create tabs for each user with data
	requests = []
	sheet_data = []
	
	for user_email, files in user_data.items():
		if files:  # Only create tabs for users with files
			if user_email == '__SHARED_DRIVES__':
				sheet_title = 'Shared Drives'
			else:
				sheet_title = user_email.split('@')[0][:30]  # Truncate for sheet name limits
			
			# Add sheet creation request
			requests.append({
				'addSheet': {
					'properties': {
						'title': sheet_title
					}
				}
			})
			
			# Prepare data for this sheet - always include all fields in sheets
			if user_email == '__SHARED_DRIVES__':
				headers = ['File Name', 'Share Link', 'File ID', 'Modified Time', 'Shared Drive']
			else:
				headers = ['File Name', 'Share Link', 'File ID', 'Modified Time']
			rows = [headers]
			
			for file in files:
				if user_email == '__SHARED_DRIVES__':
					row = [
						file['name'],
						file['webViewLink'],
						file['id'],
						file.get('modifiedTime', 'N/A'),
						file.get('sharedDriveName', 'N/A')
					]
				else:
					row = [
						file['name'],
						file['webViewLink'],
						file['id'],
						file.get('modifiedTime', 'N/A')
					]
				rows.append(row)
			
			sheet_data.append((sheet_title, rows))
	
	# Create all sheets
	if requests:
		sheets_service.spreadsheets().batchUpdate(
			spreadsheetId=spreadsheet_id,
			body={'requests': requests}
		).execute()
	
	# Prepare all data for batch update
	total_files = sum(len(files) for files in user_data.values())
	users_with_files = len([email for email, files in user_data.items() if files and email != '__SHARED_DRIVES__'])
	shared_drive_files = len(user_data.get('__SHARED_DRIVES__', []))
	
	dashboard_data = [
		['Cool Drive Audit Report'],
		[''],
		['Audit Date:', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
		['Domain:', settings.DOMAIN],
		['Total Public Files:', str(total_files)],
		['Users with Public Files:', str(users_with_files)],
		['Shared Drive Public Files:', str(shared_drive_files)],
		[''],
		['Summary:'],
		['Source', 'Files Count', 'Sheet Tab']
	]
	
	# Add user summary rows
	for user_email, files in user_data.items():
		if files:
			if user_email == '__SHARED_DRIVES__':
				sheet_name = 'Shared Drives'
				source_name = 'Shared Drives'
			else:
				sheet_name = user_email.split('@')[0][:30]
				source_name = user_email
			dashboard_data.append([source_name, str(len(files)), sheet_name])
	
	# Prepare batch update data
	batch_data = [
		{
			'range': 'Dashboard!A1',
			'values': dashboard_data
		}
	]
	
	# Add user sheet data to batch
	for sheet_title, rows in sheet_data:
		batch_data.append({
			'range': f"'{sheet_title}'!A1",
			'values': rows
		})
	
	# Write all data in a single batch update
	if batch_data:
		sheets_service.spreadsheets().values().batchUpdate(
			spreadsheetId=spreadsheet_id,
			body={
				'valueInputOption': 'RAW',
				'data': batch_data
			}
		).execute()
	
	return spreadsheet['spreadsheetUrl']

if __name__ == "__main__":
	parser = argparse.ArgumentParser(
		description='Cool Drive Audit - Audit Google Drive files for public sharing across your domain',
		usage='%(prog)s [options]\n\nExamples:\n  %(prog)s\n  %(prog)s -f name link\n  %(prog)s -f name link id modified --no-html',
		formatter_class=argparse.RawDescriptionHelpFormatter
	)
	parser.add_argument('--fields', '-f', 
		choices=['name', 'link', 'id', 'modified'],
		nargs='+',
		default=['name'],
		metavar='FIELD',
		help='output fields: name (file name), link (sharing URL), id (file ID), modified (last modified date)')
	parser.add_argument('--no-html', action='store_true',
		help='console output only, skip HTML report generation')
	parser.add_argument('--sheets', action='store_true',
		help='create Google Sheets report with separate tabs for each user')
	parser.add_argument('--shared-drives-only', action='store_true',
		help='audit shared drives only, skip individual user files')
	parser.add_argument('--debug', action='store_true',
		help='enable debug mode with detailed error messages and API call logging')
	
	args = parser.parse_args()

	# Set debug mode
	settings.DEBUG = args.debug

	# Validate APIs before processing
	if not validate_apis(use_sheets=args.sheets):
		sys.exit(1)

	if not args.no_html:
		with open('email_template.html') as f:
			template = f.read()

		outdir = "out-{}".format(datetime.datetime.now().strftime("%Y%m%d-%H%M%S"))
		os.mkdir(outdir)

	users = common.get_domain_users()
	total_files = 0
	sheets_data = {}
	
	# Process shared drives first
	print("\nAuditing shared drives...")
	try:
		shared_drives = common.get_shared_drives()
		shared_drive_files = []
		
		for drive in shared_drives:
			print(f"\nShared Drive: {drive['name']}")
			try:
				drive_files = common.get_publicly_shared_files_from_shared_drive(drive['id'], drive['name'])
				if drive_files:
					shared_drive_files.extend(drive_files)
					for file in drive_files:
						print("    {}".format(format_file_output(file, args.fields)))
				else:
					print("    No publicly shared files found")
			except Exception as e:
				log_error(e, f"Accessing shared drive: {drive['name']}")
				print(f"    Error accessing shared drive: {str(e)}")
		
		if shared_drive_files:
			total_files += len(shared_drive_files)
			sheets_data['__SHARED_DRIVES__'] = shared_drive_files
			
			if not args.no_html:
				result_elem = ""
				for file in shared_drive_files:
					result_elem += "<li><a href=\"{link}\">{name}</a> (Shared Drive: {drive})</li>\n".format(
						link=file['webViewLink'], 
						name=file['name'],
						drive=file['sharedDriveName'])
				
				output = template.format(
					name="Shared Drives", 
					result_elem=result_elem, 
					email="shared-drives@" + settings.DOMAIN)
				with open("{}/shared-drives.html".format(outdir), "w") as outfile:
					outfile.write(output)
		
		print(f"\nTotal shared drive files: {len(shared_drive_files)}")
		
	except Exception as e:
		log_error(e, "Processing shared drives")
		print(f"Error processing shared drives: {str(e)}")
	
	if not args.shared_drives_only:
		print("\nAuditing user files...")
		for user in users:
			user_email = user['primaryEmail']
			user_name = user['name'].get('givenName', user_email)

			print("\n{}:".format(user_email))
			
			try:
				public_files = common.get_publicly_shared_files(user_email)
			except Exception as e:
				log_error(e, f"Accessing files for user: {user_email}")
				print("    Error accessing user's files: {}".format(str(e)))
				if settings.DEBUG:
					import traceback
					print("    DEBUG: Full traceback:")
					traceback.print_exc()
				continue

			if public_files:
				total_files += len(public_files)
				sheets_data[user_email] = public_files
				result_elem = ""
				for file in public_files:
					print("    {}".format(format_file_output(file, args.fields)))
					if not args.no_html:
						result_elem += "<li><a href=\"{link}\">{name}</a></li>\n".format(link=file['webViewLink'], name=file['name'])
				
				if not args.no_html:
					output = template.format(name=user_name, result_elem=result_elem, email=user_email)
					with open("{}/{}.html".format(outdir, user_email), "w") as outfile:
						outfile.write(output)
			else:
				print("    No publicly shared files found")
	else:
		print("\nSkipping user auditing (shared drives only mode)")
	
	if args.shared_drives_only:
		print("\nTotal shared drive files found: {}".format(total_files))
	else:
		print("\nTotal publicly shared files found: {}".format(total_files))
	
	if args.sheets and sheets_data:
		print("Creating Google Sheets report...")
		try:
			sheets_url = create_google_sheets_report(sheets_data, settings.ADMIN_USERNAME)
			print("Google Sheets report created: {}".format(sheets_url))
		except Exception as e:
			log_error(e, "Creating Google Sheets report")
			print("Error creating Google Sheets report: {}".format(str(e)))
			if settings.DEBUG:
				import traceback
				print("DEBUG: Full traceback:")
				traceback.print_exc()
	
	if not args.no_html:
		if args.shared_drives_only:
			print("HTML report for shared drives generated in: {}".format(outdir))
		else:
			print("HTML reports generated in: {}".format(outdir))
	print("done.")
