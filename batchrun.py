import os
import parse_estratto_ing
import re

files = []
scandir_path = '../'
for entry in os.scandir(scandir_path):
    if entry.is_file() and entry.name.endswith('.pdf'):
        if re.search('20\d\dQ[1-4]', entry.name):
            files.append(os.path.join(scandir_path, entry.name))

files = sorted(files, reverse=True)
for f in files:
    print('=== WORKING ON', f, '===')
    parse_estratto_ing.parse_file(f)
