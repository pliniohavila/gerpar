import os

def get_letters_name(folder):
    files = []
    for file in os.listdir(folder):
        if file.endswith(".docx") or file.endswith(".doc"):
            files.append(file)
    return files

folder = "of_aditivos"
files = get_letters_name(folder)
for file in files:
  print(file)
