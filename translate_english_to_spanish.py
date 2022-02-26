
import re
import docx2txt
from googletrans import Translator


def save_to_file(doc, name):
    with open(name, mode="w+", encoding='utf-8') as myfile:
        myfile.write(''.join(doc))
        myfile.write('')


def split_english(row_id, row_text):
    if not row_text.strip():
        return False
    english_pattern = re.compile(r"([a-zA-Z]+([, \.\-\’…0-9]?)+)+")
    match_object = english_pattern.search(row_text)
    if match_object:
        extracted = match_object.group()
        position_in_row = row_text.find(extracted)
        return [row_id, position_in_row, extracted]


def pagenate_rows(rows, show_logs=False):
    pages = []
    chars_in_current_page = 0
    current_page = []
    current_page_id = 1
    for row in rows:
        if not row[-1]:
            continue

        chars_in_row = len(row[-1]) + 1
        if chars_in_current_page + chars_in_row < 5000:
            current_page.append(row)
            chars_in_current_page += chars_in_row
        else:
            pages.append(current_page)
            if show_logs:
                print(f"Current page id - {current_page_id}")
                print(f"Current page chars: {chars_in_current_page}")
                print()
            current_page_id += 1
            current_page = []
            chars_in_current_page = 0
    return pages


def translate_page(page):
    tranlsator = Translator()
    page_string = "\n".join([row[-1] for row in page])
    translation = tranlsator.translate(page_string,
                                       dest="es",
                                       src="en").text
    for index, row in enumerate(translation.split("\n")):
        page[index][-1] = row

    return page


if __name__ == "__main__":
    text = docx2txt.process("textbook.docx")
    text = text.split("\n")
    english_rows = []
    for index, data in enumerate(text):
        processed_row = split_english(index, data)
        if processed_row:
            english_rows.append(processed_row)

    spanish_rows = []
    for page in pagenate_rows(english_rows):
        spanish_rows.extend(translate_page(page))

    for row in spanish_rows:
        row_id, position_in_row, extracted = row
        text[row_id] = text[row_id][:position_in_row] + extracted

    text = "\n".join(text)
    save_to_file(text, "translated_to_spanish.doc")
