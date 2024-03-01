import os

# paths & directories
script_dir = os.path.dirname(os.path.abspath(__file__))
de_dir = os.path.join(script_dir, 'BKN_Dokumenten', 'de')
fr_dir = os.path.join(script_dir, 'BKN_Dokumenten', 'fr')
it_dir = os.path.join(script_dir, 'BKN_Dokumenten', 'it')
target_dir = os.path.join(script_dir, 'BKN_Dokumenten', '_drei sprachig')

class Folder:
    def __init__(self, full_name, name, lang, common):
        self.full_name = full_name
        self.name = name
        self.lang = lang
        self.common = common

    def __str__(self):
        return f'{self.full_name}, {self.name}, {self.lang}, {self.common}'

# Create an array to store all folder classes
folder_list = []
folders = ['BODLUV Br 33']

# Function to get list of folders in a directory
def get_folders_in_directory(directory):
    for folder_name in os.listdir(directory):
        if folder_name not in folders:
            continue
        folder_path = os.path.join(directory, folder_name)
        if os.path.isdir(folder_path):
            full_name = folder_name
            name = folder_name
            lang = directory.split(os.path.sep)[-1]
            folder_instance = Folder(full_name, name, lang, 0)
            folder_list.append(folder_instance)

    return folder_list

de_folders = get_folders_in_directory(de_dir)
fr_folders = get_folders_in_directory(fr_dir)
it_folders = get_folders_in_directory(it_dir)

# Create a dictionary to store name, language, and common of every class
combined_folders = {}

# Iterate through all folders in each language
for folder in folder_list:
    name = folder.name
    lang = folder.lang
    if name not in combined_folders:
        combined_folders[name] = {lang}
    else:
        combined_folders[name].add(lang)

# Update the common attribute for folders present in all three languages
for folder_name, languages in combined_folders.items():
    if len(languages) == 3:
        for folder in folder_list:
            if folder.name == folder_name:
                folder.common = 1

# Function to get all HTML files for each language
def get_html_files(lang):
    files = []
    for folder in folder_list:
        if folder.common == 1 and folder.lang == lang:
            file_path = os.path.join(script_dir, 'BKN_Dokumenten', lang, folder.full_name, 'HTML', 'TEST')
            for item in os.listdir(file_path):
                item_path = os.path.join(file_path, item)
                if os.path.isfile(item_path):
                    files.append(item_path)  # Append the full file path
    return files

de_files = get_html_files(lang='de')
fr_files = get_html_files(lang='fr')
it_files = get_html_files(lang='it')

# Create a dictionary to store files with the same name
combined_files = {}

for de, fr, it in zip(de_files, fr_files, it_files):
    # Extract the file name without extension
    file_name = os.path.splitext(os.path.basename(de))[0]

    if file_name not in combined_files:
        combined_files[file_name] = []

    combined_files[file_name].append(de)
    combined_files[file_name].append(fr)
    combined_files[file_name].append(it)

# Define the target directory for saving merged content
merged_dir_path = os.path.join(target_dir, 'merged_output')


# Create the target directory if it doesn't exist
os.makedirs(merged_dir_path, exist_ok=True)

# Last element in css: </style>
css = '''<!DOCTYPE html>
<html>

<head>
	<style>
		@font-face {
			font-family: SegoeUI;
			src:
				local("Segoe UI Light"),
				url(https://c.s-microsoft.com/static/fonts/segoe-ui/west-european/light/latest.woff2) format("woff2"),
				url(https://c.s-microsoft.com/static/fonts/segoe-ui/west-european/light/latest.woff) format("woff"),
				url(https://c.s-microsoft.com/static/fonts/segoe-ui/west-european/light/latest.ttf) format("truetype");
			font-weight: 100;
		}

		@font-face {
			font-family: SegoeUI;
			src:
				local("Segoe UI Semilight"),
				url(https://c.s-microsoft.com/static/fonts/segoe-ui/west-european/semilight/latest.woff2) format("woff2"),
				url(https://c.s-microsoft.com/static/fonts/segoe-ui/west-european/semilight/latest.woff) format("woff"),
				url(https://c.s-microsoft.com/static/fonts/segoe-ui/west-european/semilight/latest.ttf) format("truetype");
			font-weight: 200;
		}

		@font-face {
			font-family: SegoeUI;
			src:
				local("Segoe UI"),
				url(https://c.s-microsoft.com/static/fonts/segoe-ui/west-european/normal/latest.woff2) format("woff2"),
				url(https://c.s-microsoft.com/static/fonts/segoe-ui/west-european/normal/latest.woff) format("woff"),
				url(https://c.s-microsoft.com/static/fonts/segoe-ui/west-european/normal/latest.ttf) format("truetype");
			font-weight: 400;
		}

		@font-face {
			font-family: SegoeUI;
			src:
				local("Segoe UI Bold"),
				url(https://c.s-microsoft.com/static/fonts/segoe-ui/west-european/bold/latest.woff2) format("woff2"),
				url(https://c.s-microsoft.com/static/fonts/segoe-ui/west-european/bold/latest.woff) format("woff"),
				url(https://c.s-microsoft.com/static/fonts/segoe-ui/west-european/bold/latest.ttf) format("truetype");
			font-weight: 600;
		}

		@font-face {
			font-family: SegoeUI;
			src:
				local("Segoe UI Semibold"),
				url(https://c.s-microsoft.com/static/fonts/segoe-ui/west-european/semibold/latest.woff2) format("woff2"),
				url(https://c.s-microsoft.com/static/fonts/segoe-ui/west-european/semibold/latest.woff) format("woff"),
				url(https://c.s-microsoft.com/static/fonts/segoe-ui/west-european/semibold/latest.ttf) format("truetype");
			font-weight: 700;
		}

		@page {
			size: A4;
			margin: 0;
		}

		* {
			font-family: "SegoeUI";
			box-sizing: border-box;
  			-moz-box-sizing: border-box;
		}

		header {
			display: flex;
			justify-content: space-between;
			margin-bottom: 48px;
		}

		body {
			font-size: 12pt;
		}

		.page {
			padding: 20mm;
			page-break-after: always;
		}

		.page:last-of-type {
			page-break-after: avoid;
		}

		h1 {
			color: #00008A;
			margin-bottom: 48px;
			font-size: 30px;
		}

		table {
			border-collapse: collapse;
		}

		table.medium,
		p.medium {
			font-size: 10pt;
		}

		table.competences {
			margin-bottom: 12pt;
		}

		th {
			font-weight: bold;
			text-align: left;
			vertical-align: top;
		}

		table.competences th {
			padding: 10px 8px;
		}

		td {
			vertical-align: top;
		}

		table.competences td {
			border-bottom: 1px solid black;
			padding: 6px 8px;
		}

		p.small {
			font-size: 9pt;
		}

		.colored {
			background-color: #00008A;
			color: white;
		}
	</style>'''

# Merge and save files with the same name
for file_name, file_paths in combined_files.items():
    merged_content = []
    merged_content.append(css)

    for file_path in file_paths:
        with open(file_path, 'rb') as file:
            content = file.read().decode('utf-8', errors='replace')
            content_v2 = content.split('</style>')
            merged_content.append(content_v2[1])

    # Define the target file path for saving merged content
    merged_file_path = os.path.join(merged_dir_path, f'{file_name}_merged.html')

    # Save merged content to a new file
    with open(merged_file_path, 'w', encoding='utf-8') as merged_file:
        for content in merged_content:
            merged_file.write(content)

    print(f'Merged content saved to {merged_file_path}')