![eluciDoc_header](https://github.com/jblake1965/eluciDoc/assets/100727736/e7f94b7f-fb1b-4f55-8665-4dc11c6b93af)

[![CodeQL](https://github.com/jblake1965/eluciDoc/actions/workflows/github-code-scanning/codeql/badge.svg)](https://github.com/jblake1965/eluciDoc/actions/workflows/github-code-scanning/codeql) [![GitHub Discussions](https://img.shields.io/github/discussions/jblake1965/eluciDoc?labelColor=blue&color=orange)](https://github.com/jblake1965/eluciDoc/discussions/3) [![PYPI Version](https://img.shields.io/pypi/v/elucidoc?logoColor=blue&labelColor=green)](https://img.shields.io/pypi/v/elucidoc?logoColor=blue&labelColor=green)

# What this is:
This CLI Python project, written for the Windows™ environment, filters sentences and clauses containing specific user input
terms from a single document. This project was originally created as a tool to aid in the review of legal contracts, 
but can be used with any text. Documents can be in docx, .pdf or .txt file formats.
The general principle behind its function is subject-predicate sentence analysis. Searches are based on user-selected parties
in the document, followed by a user-selected phrase.  It is used in conjunction with Microsoft™ Office 365™
Word and Excel™ apps.
# How it works:
A .docx, .pdf or .txt file and path is entered (drag and drop work in the Windows terminal):

![file_input](https://github.com/jblake1965/eluciDoc/assets/100727736/c08d59a4-a019-4a42-b895-427a1815b474)

The file is then processed as utf8 text, with MS Word Smart Quotes being converted to straight quotes and non-ASCII and
non-breaking spaces removed. The term for the party being searched in the document is entered next:

![enter_party_name](https://github.com/jblake1965/eluciDoc/assets/100727736/bd1e9603-137c-4475-aa1d-d09ed157738a)

and then passed with the processed text to textacy's Keyword in Context (KWIC) function.  The result is saved as an Excel
file with the same name in the same location as the searched document, with "..._[name of the party]_search_result.xlsx" 
appended. The Excel file automatically opens with a subprocess call, and the results can be converted to a table for
further sorting:

![textacy_rendering](https://github.com/jblake1965/eluciDoc/assets/100727736/a9bfd1a8-8477-4401-8e96-bd83801d5488)

Note: the subprocess call below uses the default Office install location:

```python
subprocess.Popen([r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE', result_file])
```

If the user has Office installed in a different location, then the code must be changed to reflect that directory.

The document is chunked into sentences (or clauses, depending on the formatting) with the spaCy module.
The user is prompted to enter predicate search phrases culled from the Excel search file which phrases are stored in a list.  
Once finished entering the predicate search phrases, the script iterates through the list of search phrases looking for
a match in each sentence. Sentences and clauses containing a match are added to a result list. The user has the option of 
having the search phrases appear in ALL CAPS in the Word document containing the results, as shown:

![all_caps_rendering](https://github.com/user-attachments/assets/5d07f8a5-ed26-4341-9bf8-0445de315245)


The result list is then saved as a Word file that is opened automatically at the end of the run (as with Excel, 
note the location of the Word executable and adjust the path if it is not in the standard install location). 
# External Dependencies and Licenses

| Name:        | License:                                                              |
|--------------|-----------------------------------------------------------------------|
| docx2python  | [MIT](https://pypi.org/project/docx2python/)                          |
| openpyxl     | [MIT](https://pypi.org/project/openpyxl/)                             |
| pandas       | [BSD](https://pypi.org/project/pandas/)                               |
| pdfminer.six | [MIT/X](https://github.com/pdfminer/pdfminer.six/blob/master/LICENSE) |
| python-docx  | [MIT](https://github.com/atriumlts/python-docx/blob/master/LICENSE)   |
| rich         | [MIT](https://pypi.org/project/rich/)                                 |
| spacy        | [MIT](https://pypi.org/project/spacy/)                                |
| textacy      | [Apache 2.0](https://pypi.org/project/textacy/)                       |

# Installation
It is strongly recommended that this package be installed in a virtual environment.  The package is available at https://pypi.org/project/elucidoc/ 
and can be installed with ```pip install elucidoc``` .

***THE SPACY PIPELINE  `  en_core_web_lg  `  MUST ALSO BE INSTALLED INTO THE VIRTUAL
ENVIRONMENT FOR THE SCRIPT TO WORK***.

The pipeline can be installed as follows:
```
python -m spacy download en_core_web_lg
```
You must also be sure to verify the directory for the Office install is the same as noted above.  If not, the code must be 
changed to the directory where the Excel and Word apps are located.
# Running the Script
The project is run as a script.  It can be run with a .bat file calling the virtual environment and the executable
file per the below example:
```commandline
@"C:\Users\..\venv\Scripts\python.exe" "C:\Users\..\venv\lib\elucidoc\eluciDoc.py"

@pause
```
Additionally, the location of the ```elucidoc.py``` executable can be included in the Windows ```PATH``` environment variable.
# Case Sensitive Searches
General convention in legal texts is to capitalize defined terms.  For that reason, the user may want to make the search
case-sensitive to target the appropriate instances of the term.  For searches where the specific use of the subject term
is not important but broader capture is, the case-sensitive feature can be turned off.  Once a selection is made, it applies
for all subsequent searches until the script is restarted.
# Possessive Case and Other Punctuation
Textacy divides the party search term from both following words and punctuation including the possessive case, as shown below:

![textacy_rendering](https://github.com/jblake1965/eluciDoc/assets/100727736/1fd67f92-57bd-402a-b99f-95d5847f49f7)

To capture an instance of a possessive case of the party being searched, a 's or ' (for the plural possessive) must be
the first character in the predicate search phrase, as illustrated by the prompt below:

![enter_predicate_phrase](https://github.com/jblake1965/eluciDoc/assets/100727736/edc9f616-97d7-4bdc-8553-f89292e43332)

The same principal applies to the comma, colon, semicolon and closed parentheses immediately following the party name.

# Smart Quotes
Microsoft Word's default settings utilize smart quotes, which are the
curly type fonts. Those are problematic when searching
documents converted to text (rendered as slanted quotes in Utf8), and
are replaced with straight quotes via the following code:

```python
text = re.sub(r'”', '\"', text)  # replace double smartquote open quote
text = re.sub(r'“', '\"', text)  # replace double smartquote close quote
text = re.sub(r'’', '\'', text)  # replace single smartquote close quote
text = re.sub(r'‘', '\'', text)  # replace single smartquote open quote
```

# PDFs
Due to the nature of .pdf files and the sometimes-inconsistent results
that occur when converting pdf documents to text format, additional
processing is done. Some characters and extra spaces between word boundaries are removed as part of the
text processing:
```python
text = re.sub(u'[^\u0020-\uD7FF\u0009\u000A\u000D\uE000-\uFFFD\U00010000-\U0010FFFF]+', '', text)
text = re.sub(r'(\b)(\s{2,4})(\b)', r'\g<1> ', text)
```
The above solution is not a comprehensive fix for pdf issues. The accuracy of the results with searches of .pdf files
may be negatively impacted by the quality or formatting of the underlying document, particularly with
scanned documents.

# Open Files
If a consecutive search is run for the same party and the Excel file with the prior search results is still open,
the script will notify the user of such and not overwrite the existing Excel file.  With the Word files, the user will 
be prompted to save the existing file with
another name and close it before proceeding with a second search for the same party.
