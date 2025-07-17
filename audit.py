import argparse
import datetime
import os

import common
import settings

settings.DEBUG = False


def format_file_output(file, fields):
	"""Format file output based on selected fields"""
	output_parts = []
	
	if 'name' in fields:
		output_parts.append(file['name'])
	if 'link' in fields:
		output_parts.append(file['webViewLink'])
	if 'id' in fields:
		output_parts.append(file['id'])
	if 'modified' in fields:
		output_parts.append(file.get('modifiedTime', 'N/A'))
	
	return " | ".join(output_parts)

if __name__ == "__main__":
	parser = argparse.ArgumentParser(
		description='Audit Google Drive files for public sharing across your domain',
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
	
	args = parser.parse_args()

	if not args.no_html:
		with open('email_template.html') as f:
			template = f.read()

		outdir = "out-{}".format(datetime.datetime.now().strftime("%Y%m%d-%H%M%S"))
		os.mkdir(outdir)

	users = common.get_domain_users()
	total_files = 0
	
	for user in users:
		user_email = user['primaryEmail']
		user_name = user['name'].get('givenName', user_email)

		print("\n{}:".format(user_email))
		
		try:
			public_files = common.get_publicly_shared_files(user_email)
		except Exception as e:
			print("    Error accessing user's files: {}".format(str(e)))
			continue

		if public_files:
			total_files += len(public_files)
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
	
	print("\nTotal publicly shared files found: {}".format(total_files))
	if not args.no_html:
		print("HTML reports generated in: {}".format(outdir))
	print("done.")
