
import os
import re
import subprocess
import nltk
from nltk import word_tokenize
import spacy
import pyinputplus as pyip
from pdfminer.high_level import extract_text
from docx2python import docx2python
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt

global text
global case_sensitive


def concord(party_term: str, content: str):
    # Returns None type - Results are printed to screen
    tokens = word_tokenize(content)
    token_text = nltk.Text(tokens)
    concordance_result = token_text.concordance(party_term, width=150, lines=500)
    return concordance_result


while True:
    file = input('Enter the file path and name - only .docx, .pdf or .txt files: ')
    target_File = file.strip('\"')
    filename, file_extension = os.path.splitext(target_File)
    p = Path(target_File)
    target_File_path = p.parent
    result_Filename = p.name
    extensions = ['.docx', '.pdf', '.txt']

    if file_extension not in extensions:
        print('Invalid file type.')
        continue

    if file_extension in extensions:
        if file_extension == '.docx':
            docx_content = docx2python(target_File)
            text = docx_content.text
        if file_extension == '.pdf':
            text = extract_text(target_File)
        if file_extension == '.txt':
            target = open(target_File, 'r', encoding='utf8')
            text = target.read()
    break

text = text.replace("\n", " ")
text = re.sub(r'”', '\"', text)  # replace double smartquote open quote
text = re.sub(r'“', '\"', text)  # replace double smartquote close quote
text = re.sub(r'’', '\'', text)  # replace single smartquote close quote
text = re.sub(r'‘', '\'', text)  # replace single smartquote open quote

print('\n', 'Is this search case sensitive?')
case_sensitive = pyip.inputMenu(['Yes', 'No'], numbered=True)
if case_sensitive == 'No':
    text = text.lower()

# PDF clean-up processing
if file_extension == '.pdf':
    text = re.sub(u'[^\u0020-\uD7FF\u0009\u000A\u000D\uE000-\uFFFD\U00010000-\U0010FFFF]+', '', text)
    text = re.sub(r'(\b)(\s{2,4})(\b)', r'\g<1> ', text)

nlp = spacy.load("en_core_web_lg")
doc = nlp(text)
sentences: list[str] = [sentence.text for sentence in doc.sents]

while True:
    master_List = []
    search_phrase_list: list[str] = []

    party: str = input('Enter the term for the party to be searched (entry is case sensitive if that option selected:')
    if case_sensitive == 'No':
        party = party.lower()
    concord(party, text)
    print('\n', 'Run Excel?')
    excel_Response = pyip.inputMenu(['Yes', 'No'], numbered=True)
    if excel_Response == 'Yes':
        subprocess.Popen(r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE')

    while True:
        secTerm: str = input(r'''Enter a predicate search term or phrase - include "'s" or "'" for possessive case of
         the party being searched:''')

        if not secTerm.startswith("'"):  # add a space where the second term is not an apostrophe
            secTerm = " ".join(["", secTerm])

        search_phrase = party + secTerm
        if case_sensitive == 'No':
            search_phrase = search_phrase.lower()

        search_phrase_list.append(search_phrase)
        list_Status = ' | '.join(search_phrase_list)
        print('')
        print('Search Phrases:', list_Status)  # track Search Phrases as entered
        print('')

        response = pyip.inputMenu(['Enter another search term', 'Finished for this party'], numbered=True)

        if response == 'Finished for this party':
            for sent in sentences:
                for i in search_phrase_list:
                    if i in sent:
                        if sent not in master_List:
                            master_List.append(sent)
                continue
            break

    document = Document()
    style = document.styles['Normal']
    font = style.font
    font.name = "Times New Roman"
    font.size = Pt(12)
    sections = document.sections
    section = sections[0]
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)

    document.add_paragraph(f'Phrases searched: {list_Status}')
    for j in master_List:
        document.add_paragraph(j)
    document.save(rf'{target_File_path}\{result_Filename}_{party}_search_result.docx')
    subprocess.Popen([r'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE',
                      rf'{target_File_path}\{result_Filename}_{party}_search_result.docx'])
    response = pyip.inputMenu(['Search another party', 'Finished'], numbered=True)
    if response == 'Search another party':
        continue
    if response == 'Finished':
        print('Run Finished')
    break
