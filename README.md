![alt text](https://github.com/jblake1965/eluciDoc/blob/6b49e8e144ac18308403b83fea5709755d24c04a/Pictures/eluciDoc%20header.jpg)

[![CodeQL](https://github.com/jblake1965/eluciDoc/actions/workflows/github-code-scanning/codeql/badge.svg)](https://github.com/jblake1965/eluciDoc/actions/workflows/github-code-scanning/codeql) [![GitHub Discussions](https://img.shields.io/github/discussions/jblake1965/eluciDoc?labelColor=blue&color=orange)](https://github.com/jblake1965/eluciDoc/discussions/3) 

# What this is:
This CLI Python 3.11 project, written in the Windows™ environment, filters
sentences containing specific phrases from a document. This project was
created for legal contracts, but can be used with any text. Documents
can be in docx, .pdf or .txt files. The general principle behind its
function is subject-predicate sentence analysis. Searches are based on a
user-selected party in the document, followed by a user-selected phrase.
It is designed to be used in conjunction with Microsoft™ Office 365™
Word and Excel™ apps.
# How it works:
A .docx, .pdf or .txt file is processed as utf8 text (note: MS Word Smart Quotes must be converted to straight quotes as discussed below).
The term for the party-in-interest in the contract is passed with the processed text to textacy's Keyword in Context (KWIC) function,
and the result then saved as an Excel file in the same location as the
file containing the document being searched, with the same name as the file
being searched with "..._[name of the party]_search_result.xlsx" appended.
The Excel file is opened with a subprocess call, returning results similar to the following, which can be further processed (e.g. alphabetized) in Excel:

![alt text](https://github.com/jblake1965/eluciDoc/blob/developer/Pictures/Screenshot%202023-12-10%20224559.jpg)

Note: the subprocess call below uses the default Office install location:

```python
 subprocess.Popen([r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE', result_file])
```

If the user has Office installed in a different location, then the above
code should be adjusted accordingly.

The document is chunked into sentences with the spaCy module. The script
iterates through the list of search phrases looking for a match in each
sentence. Sentences containing a match which are not duplicates of
existing sentences in the master list are added to a master list. The
master list is then saved as a Word file that is opened automatically at
the end of the run (as with Excel, note the location of the Word
executable). 
# External Dependencies and Licenses

| Name:        | Version: | License:                                                                |
|--------------|----------|-------------------------------------------------------------------------|
| docx2python  | 2.0.4    | [MIT](https://pypi.org/project/docx2python/)                            |
| textacy      | 0.13.0   | [Apache 2.0](https://pypi.org/project/textacy/)                         |
| pyinputplus  | 0.2.12   | [BSD](https://github.com/asweigart/pyinputplus/blob/master/LICENSE.txt) |
| python-docx  | 0.8.11   | [MIT](https://github.com/atriumlts/python-docx/blob/master/LICENSE)     |
| spacy        | 3.4.1    | [MIT](https://pypi.org/project/spacy/)                                  |
| pdfminer.six | 20220524 | [MIT/X](https://github.com/pdfminer/pdfminer.six/blob/master/LICENSE)   |
| python-docx  | 0.8.11   | [MIT](https://pypi.org/project/python-docx/)
# N.B.
## Installation
This project was created in a virtual environment.  Also, if installing the dependencies via the Requirements.txt file:

```
pip install -r requirements.txt
```
You need to separately install the spacy library `en_core_web_lg` into the virtual environment as follows:
```
python -m spacy download en_core_web_lg
```

## Case Sensitive Searches
General convention in legal texts is to capitalize defined terms.  For that reason, the user may want to make the search
case-sensitive to target the appropriate instances of the term.  For searches where the specific use of the subject term
is not important but broader capture is, the case-sensitive feature can be turned off.
## Possessive Case
NLTK will divide a word at " 's " and " ' " with a possessive case. See
below:

![alt text](https://github.com/jblake1965/eluciDoc/blob/developer/Pictures/Screenshot%202023-05-04%20181140.jpg)

Therefore, it is necessary to add " 's " and " ' " as the first search term in order
to capture an instance of a possessive case, as illustrated below:

![alt text](https://github.com/jblake1965/eluciDoc/blob/developer/Pictures/Screenshot%202023-05-27%20121145.jpg)

## PDFs
Due to the nature of .pdf files and the sometimes-inconsistent results
that occur when converting pdf documents to text format, additional
processing is done. Some characters and extra spaces between word boundaries are removed as part of the text processing:
```python
text = re.sub(u'[^\u0020-\uD7FF\u0009\u000A\u000D\uE000-\uFFFD\U00010000-\U0010FFFF]+', '', text)
text = re.sub(r'(\b)(\s{2,4})(\b)', r'\g<1> ', text)
```
The above solution is not a comprehensive fix for pdf issues. The
accuracy of the results with searches of .pdf files may be negatively
impacted by the quality of the underlying document, particularly with
scanned documents.
## Smart Quotes
Microsoft Word's default setting utilize smart quotes, which are the
curly type fonts. Those are problematic when searching
documents converted to text (rendered as slanted quotes in Utf8), and
are replaced with straight quotes via the following code:

```python
text = re.sub(r'”', '\"', text)  # replace double smartquote open quote
text = re.sub(r'“', '\"', text)  # replace double smartquote close quote
text = re.sub(r'’', '\'', text)  # replace single smartquote close quote
text = re.sub(r'‘', '\'', text)  # replace single smartquote open quote
```
## Saving Over Existing Files
If a search for the same party is run twice and the file with the prior search results is open, an error is returned as
the script cannot replace the existing file with the first search results.  Before running the same search again, the
prior search file must be closed or deleted.
