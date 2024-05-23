import re
from docx import Document
from transliterate import translit
import pandas as pd
from fuzzywuzzy import fuzz
from concurrent.futures import ThreadPoolExecutor, as_completed

# reading files
doc = Document("list_files_2.docx")
list_files = [para.text for para in doc.paragraphs if '.' in para.text.split('\\')[-1]]

df = pd.read_excel("BazaIDCulture.xlsx", engine='openpyxl')
df["Project_name"] = df["Project_name"].fillna('').astype(str)
df["Author"] = df["Author"].fillna('').astype(str)

# cleaned names
def brus_cleaned(name):
    transliterated_name = translit(name, 'uk', reversed=True)

    replacements = {
        ';': '', ':': '', '"': '', "'": '', '-': '', ')': '', '(': '', ',': '', '.': '', '!': '', '?': '', '_': '',
        ' ': '', '„': '', '”': '', '’': '', 'persha': '1', 'druga': '2', 'tretja': '3', 'chetverta': '4', 'pjata': '5',
        'pyata': '5', 'shosta': '6', 'sioma': '7', 'soma': '7', 'vosma': '8', 'devyata': '9', 'desyata': '10',
        'chast': 'ch', 'chastyna': 'ch',
    }

    for key, value in replacements.items():
        transliterated_name = transliterated_name.replace(key, value)

    chastyna = re.findall(r".*(\d+chastyna).*", transliterated_name)
    if len(chastyna) > 0:
        chastyna = chastyna[0]
        ch = 'ch' + chastyna[0:-8]
        transliterated_name = transliterated_name.replace(chastyna, ch)

    chastyna2 = re.findall(r".*(\d+ch).*", transliterated_name)
    if len(chastyna2) > 0:
        chastyna2 = chastyna2[0]
        ch = 'ch' + chastyna2[0:-2]
        transliterated_name = transliterated_name.replace(chastyna2, ch)

    return transliterated_name.lower()


def mining_file_name(file_path):
    parts = file_path.split('\\')
    clean_folder = brus_cleaned(parts[2]) if len(parts) > 3 else None
    last_part = parts[-1]
    if '.' in last_part:
        name = last_part.split('.')[0]
    else:
        name = last_part
    clean_file_name = brus_cleaned(name)
    return clean_folder, clean_file_name, file_path


preprocessed_files = [mining_file_name(file) for file in list_files]


# founded match
def find_best_match(index, row):
    program_name = row["Project_name"]
    author = row["Author"]
    clean_program_name = brus_cleaned(program_name)
    clean_author = brus_cleaned(author)

    best_dist = 0
    best_name = ""
    best_match_author = 0

    for folder, file_name, full_path in preprocessed_files:
        cur_distance = fuzz.ratio(clean_program_name, file_name) / 100.0
        match_author = fuzz.ratio(clean_author, folder) / 100.0

        if cur_distance > best_dist:
            best_dist = cur_distance
            best_name = full_path
            best_match_author = match_author
        elif 0.7 <= cur_distance < 0.98:
            if match_author > 0.9 and cur_distance > 0.7:
                best_dist = cur_distance
                best_name = full_path
                best_match_author = match_author

    if best_dist >= 0.98:
        return index, best_name, best_dist
    elif 0.7 <= best_dist < 0.98 and best_match_author > 0.9:
        return index, best_name, best_dist
    return index, "dl found nothing", best_dist


# executing
results = []
count = 0
total_rows = len(df)

with ThreadPoolExecutor() as executor:
    futures = [executor.submit(find_best_match, index, row) for index, row in df.iterrows()]
    for future in as_completed(futures):
        result = future.result()
        results.append(result)

        # progress
        count += 1
        if count % 100 == 0:
            print(f"{count / total_rows:.2%} complete")

# create result file
for index, best_name, best_dist in results:
    df.at[index, 'appropriate_file_name'] = best_name
    df.at[index, 'probability'] = best_dist

df.to_excel("bazaIDCulture_updated11.xlsx", index=False)