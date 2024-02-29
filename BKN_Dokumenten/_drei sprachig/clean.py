import os

script_dir = os.path.dirname(os.path.abspath(__file__))

def clean_file(current_dir, file):
    file_path = os.path.join(current_dir, file)
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()
    content = content.replace('d&apos;addestramento', "d'addestramento")
    content = content.replace('d&#226;��aviazione', "d’aviazione")
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(content)
    print(f'Cleaned file:', file_path)

for file in os.listdir(script_dir):
    if file == 'clean.py':
        continue
    elif os.path.isdir(file):
        current_dir = os.path.join(script_dir, file)
        for file in os.listdir(current_dir):
            clean_file(current_dir, file)