from docx import Document
from transliterate import translit
import pandas as pd
from fuzzywuzzy import fuzz
from functools import lru_cache
import swifter


# reading files
doc = Document("list_files_2.docx")
list_files = [para.text for para in doc.paragraphs if '.' in para.text.split('\\')[-1]]

df = pd.read_excel("BazaIDCulture.xlsx", engine='openpyxl')
df["Project_name"] = df["Project_name"].fillna('').astype(str)
df["Author"] = df["Author"].fillna('').astype(str)
df["supername"] = df["supername"].fillna('').astype(str)


# cleaned names with caching
@lru_cache(maxsize=None)
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
    return transliterated_name.lower()


def mining_file_name(file_path):
    parts = file_path.split('\\')
    clean_folder = brus_cleaned(parts[2]) if len(parts) > 2 else None
    last_part = parts[-1]
    if '.' in last_part:
        name = last_part.split('.')[0]
    else:
        name = last_part
    clean_file_name = brus_cleaned(name)
    return clean_folder, clean_file_name, file_path


preprocessed_files = [mining_file_name(file) for file in list_files]


def find_best_match(row):
    program_name = row["Project_name"]
    author = row["Author"]
    super_name = row["supername"]
    clean_program_name = brus_cleaned(program_name)
    clean_author = brus_cleaned(author)
    clean_super_name = brus_cleaned(super_name)

    folder_exist = [file for file in preprocessed_files if file[0] == clean_author]
    file_in_folder_or_no = folder_exist if folder_exist else preprocessed_files

    best_dist = 0
    best_match_author = 0
    best_name = ""

    for folder, file_name, full_path in file_in_folder_or_no:
        cur_distance = fuzz.ratio(clean_program_name, file_name) / 100
        match_author = fuzz.ratio(clean_author, folder) / 100 if folder else 0

        if clean_super_name == file_name:
            return full_path, 100.0, match_author

        if match_author > 0.8:
            if cur_distance > best_dist:
                best_dist = cur_distance
                best_name = full_path
                best_match_author = 1
        else:
            if cur_distance > best_dist:
                best_dist = cur_distance
                best_name = full_path
                best_match_author = match_author

    if best_match_author == 1 and best_dist >= 0.0:
        return best_name, best_dist, best_match_author
    elif best_match_author != 1 and best_dist >= 0.98:
        return best_name, best_dist, best_match_author
    return best_name, best_dist, best_match_author


df[['appropriate_file_name', 'probability_name', 'probability_author']] = df.swifter.apply(find_best_match, axis=1, result_type='expand')

df.to_excel("bazaIDCulture_updated25.xlsx", index=False)
