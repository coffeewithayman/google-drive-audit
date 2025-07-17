# Google Drive Audit Tool - Lockdown Script
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

import sys
from datetime import datetime, timedelta

from googleapiclient import discovery

import common
import settings

settings.DEBUG = False

def replace_public_share(email, files, commit=False):
	delegated_credentials = common.delegated_credentials(email, "lockdown")
	drive = discovery.build("drive", "v3", credentials=delegated_credentials)

	outfilename = "out-ld-{}-{}.tsv".format(
		email.replace("@", "-at-"),
		datetime.now().strftime("%Y%m%d-%H%M%S"))

	with open(outfilename, 'w') as outfile:

		for file in files:
			role = common.get_public_role(file)
			replacement_permission = {
				"role": role,
				"type": "domain",
				"domain": settings.DOMAIN
			}
			print("{}\t{}\t{}".format(role, file["webViewLink"], file["name"]))
			print("{}\t{}\t{}".format(role, file["webViewLink"], file["name"]), file=outfile)

			if commit:
				try:
					drive.permissions().delete(fileId=file["id"], permissionId=common.PUBLIC_PERMISSION_ID).execute()
					drive.permissions().create(fileId=file["id"], body=replacement_permission).execute()
				except:
					# I suspect this happens when a file was public because its parent folder was public
					pass
					# print("ABOVE PERMISSION MISSING".format(role, file["webViewLink"], file["name"]))
					# print("ABOVE PERMISSION MISSING".format(role, file["webViewLink"], file["name"]), file=outfile)


if __name__ == "__main__":
	try:
		email = sys.argv[1]
		if not email.endswith(settings.DOMAIN):
			raise()
		if len(sys.argv) != 2:
			raise()
	except:
		print("usage:\npython lockdown.py email@{}".format(settings.DOMAIN))
		sys.exit(1)

	cutoff = (datetime.now() - timedelta(days=settings.LOCKDOWN_GRACE_DAYS)).isoformat()

	print("Last modified cutoff will be: {}".format(cutoff))

	# TODO: first do folders, then do files

	print("Getting public files for {}...".format(email))
	public_files = common.get_publicly_shared_files(email)
	print("    {} public files found".format(len(public_files)))


	old_public_files = common.filter_files_unmodified_since(public_files, cutoff)
	print("    {} public files modified since cutoff".format(len(public_files) - len(old_public_files)))
	print("    {} remaining public files found".format(len(old_public_files)))

	dryrun_str = input("Dry run? [Y]/n ")
	commit = dryrun_str == "n"
	replace_public_share(email, old_public_files, commit=commit)
