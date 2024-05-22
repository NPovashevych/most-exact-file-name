from docx import Document
from transliterate import translit
import pandas as pd
import editdistance


# Read exel file & word file
doc = Document("list_files_2.docx")
list_files = [para.text for para in doc.paragraphs]

df = pd.read_excel("bazaIDCulture.xlsx")
df["Project_name"] = df["Project_name"].fillna('').astype(str)


# cleaned names files & show
def brus_cleaned(name):
    replacements = {
        ';': '', ':': '', '"': '', "'": '', '-': '', ')': '', '(': '', ',': '', '.': '', '!': '', '?': '', '_': '',
        ' ': '', '„': '', '”': '', '’': '', 'persha': '1', 'druga': '2', 'tretja': '3', 'chetverta': '4', 'pjata': '5',
        'pyata': '5', 'shosta': '6', 'sioma': '7', 'soma': '7', 'vosma': '8', 'devyata': '9', 'desyata': '10',
        'chast': 'ch', 'chastyna': 'ch'
    }
    for key, value in replacements.items():
        name = name.replace(key, value)
    return name.lower()


def cleaned_tv_show_name(show_name):
    transliterated_name = translit(show_name, 'uk', reversed=True)
    normalized_name = brus_cleaned(transliterated_name)
    return normalized_name


def cleaned_file_name(file_name):
    parts = file_name.split('\\')
    last_part = parts[-1]
    if '.' in last_part:
        name = last_part.split('.')[0]
    else:
        name = last_part
    transliterated_name = translit(name, 'uk', reversed=True)
    normalized_name = transliterated_name.replace('?', 'i').replace('_', ' ')
    clean_file_name = brus_cleaned(normalized_name)
    return clean_file_name.lower()


clean_files = {file_name: cleaned_file_name(file_name) for file_name in list_files}
count = 0

# founder match
for index, row in df.iloc[:7050].iterrows():
    program_name_ukr = row["Project_name"]
    clean_tv_show_name = cleaned_tv_show_name(program_name_ukr)

    best_dist = 0
    best_name = ""
    for file_name, clean_file_name in clean_files.items():

        cur_distance = editdistance.eval(clean_tv_show_name, clean_file_name)
        norm_distance = round(1 - (cur_distance / max(len(clean_tv_show_name), len(clean_file_name))), 3)
        if norm_distance > best_dist:
            best_dist = norm_distance
            best_name = file_name

    if best_dist >= 0.70:
        df.at[index, 'appropriate_file_name'] = best_name
#        print(f"find for {best_name}")
    else:
        df.at[index, 'appropriate_file_name'] = "dl found nothing"
    df.at[index, 'probability'] = best_dist

# progress
    count += 1
    print(f"{count/7050:.2%} complete")

# Save result to Excel
df.to_excel("bazaIDCulture_updated.xlsx", index=False)
