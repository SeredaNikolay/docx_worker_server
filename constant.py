from os import path, getcwd

TAGS_ONLY: str = r"(?:<[^>]+?>)*?"
LEFT_FIELD_SEARCH_REGEXP_PART: str = r"(?:(?<!~~)<w:t>{" \
                                     + TAGS_ONLY + r"{"
RIGHT_FIELD_SEARCH_REGEXP_PART: str = r"}" + TAGS_ONLY + r"}</w:t>)"
FIELD_SEARCH_REGEXP_STR: str = LEFT_FIELD_SEARCH_REGEXP_PART \
                               + r"(?:[^{}]*?)" \
                               + RIGHT_FIELD_SEARCH_REGEXP_PART
BAD_KEY_PARTS_SEARCH_REGEXP_STR: str = r"(?<={|>)(" \
                                       + "(?:\s+[a-zA-Z0-9_]*\s*)" \
                                       + "|(?:[a-zA-Z0-9_]+\s+)" \
                                       + "|(?:(?<={)(?=}))" \
                                       + ")(?=}|<)"
GOOD_KEY_PARTS_SEARCH_REGEXP_STR: str = r"(?<={|>)" \
                                        + "([a-zA-Z0-9_]+)" \
                                        + "(?=}|<)"
XML_PATH: str = r"word"
XML_FILENAME: str =r"document.xml"
DOCUMENT_PATH: str = path.join(getcwd(), "template")
DIC_FILENAME_EXTENSION: str = r"csv"
DOCUMENT_FILENAME_EXTENSION: str = r"docx"
ARCHIVE_EXTENSION = r"zip"
TEST_DIR_NAME: str = r"test"