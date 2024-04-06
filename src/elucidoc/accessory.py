import re


def text_cleanup(text):
    text = text.replace('\n', ' ')
    text = text.replace('\xa0', ' ')
    text = re.sub(r'”', '\"', text)  # replace double smartquote open quote
    text = re.sub(r'“', '\"', text)  # replace double smartquote close quote
    text = re.sub(r'’', '\'', text)  # replace single smartquote close quote
    text = re.sub(r'‘', '\'', text)  # replace single smartquote open quote
    return str(text)


def pdf_cleanup(text):
    text = re.sub(u'[^\u0020-\uD7FF\u0009\u000A\u000D\uE000-\uFFFD\U00010000-\U0010FFFF]+', '', text)
    text = re.sub(r'(\b)(\s{2,4})(\b)', r'\g<1> ', text)
    return str(text)
