from docx import Document as Doc
from docx.enum.style import WD_STYLE_TYPE
from re import compile,  match, IGNORECASE, MULTILINE
from typing import Any
from abc import ABC, abstractmethod
from os import path
from csv import reader
from constant import DOCUMENT_PATH, DIC_FILENAME_EXTENSION, \
                     DOCUMENT_FILENAME_EXTENSION

from replacement import Replacer
from paragraph import ParagraphStyleCreator, \
                      TextParagraphStyleCreator, \
                      PictureCaptionCreator, TableCaptionCreator, \
                      CommonParagraphRun, Paragraph, RunStyle, \
                      ParagraphStyle

class BaseController(ABC):

    _document_file_name: str
    _document_full_filename: str
    _dic_file_name: str
    _text_paragraph_ptrn: Any
    _picture_caption_ptrn: Any
    _table_caption_ptrn: Any
    _text_paragraph_style_cr: ParagraphStyleCreator
    _picture_caption_style_cr: ParagraphStyleCreator
    _table_caption_style_cr: ParagraphStyleCreator
    _common_text_paragraph_style: ParagraphStyle
    _common_picture_caption_style: ParagraphStyle
    _common_table_caption_style: ParagraphStyle
    _common_paragraph_run: RunStyle

    def __init__(self):
        self._document_full_filename = None
        self._document_file_name = None
        self._dic_file_name = None
        self.__patterns_init()

    def _style_init(self, document: Doc):

        self._text_paragraph_style_cr = \
            TextParagraphStyleCreator(document)      
        self._common_text_paragraph_style = \
            self._text_paragraph_style_cr \
                .get_paragraph_style("common_text_paragraph_style")

        self._picture_caption_style_cr = \
            PictureCaptionCreator(document)
        self._common_picture_caption_style = \
            self._picture_caption_style_cr \
                .get_paragraph_style("common_picture_caption_style")

        self._table_caption_style_cr = \
            TableCaptionCreator(document)
        self._common_table_caption_style = \
            self._table_caption_style_cr \
                .get_paragraph_style("common_table_caption_style")

        self._common_paragraph_run = \
            CommonParagraphRun(document.paragraphs[-1])
        
    def __patterns_init(self):
        # 1.|.1|. - not updated numbers in the document
        number: str = r"\d*\.?\d*(?:$|(?: –(?= \S)"
        self._text_paragraph_ptrn = \
            compile(r"^.+(?<=[…?!:;.])$", MULTILINE)
        self._picture_caption_ptrn = \
            compile(r"^Рисунок " + number + ".*(?<![…?!:;.]|\s)$))",
                    MULTILINE)
        self._table_caption_ptrn = \
            compile(r"^Таблица " + number + ".*(?<![…?!:;.]|\s)$))",
                    MULTILINE)
        
    def _set_paragraph(self,
                        paragrph_style: ParagraphStyle,
                        paragraph: Paragraph,
                        paragraph_run: RunStyle):
        if paragraph.style.type == WD_STYLE_TYPE.PARAGRAPH:
            paragrph_style \
                .convert_paragraph_instance_style(paragraph)
            for run in paragraph.runs:
                paragraph_run.convert_run_font(run)

    def get_document_full_filename(self) -> str:
        return self._document_full_filename

    @abstractmethod
    def set_dic_file_name(self, dic_file_name: str) -> bool:
        pass

    @abstractmethod
    def set_document_file_name(self, 
                               document_file_name: str) -> bool:
        pass
        
    @abstractmethod
    def get_full_filename_of_key_csv(
            self, key_list: list = []) -> str:
        pass

    @abstractmethod
    def get_replaced_document_full_filename_list(
            self,
            list_of_replacement_field_dic: list = []) -> list:
        pass

    @abstractmethod
    def set_document_paragraphs(
            self, document_full_filename_list: list):
        pass

#====================================================================
class Controller(BaseController):

    def __init__(self):
        super().__init__()

    def set_dic_file_name(self, dic_file_name: str) -> bool:
        is_valid: bool = True
        if dic_file_name == None or dic_file_name == "":
            is_valid = False
        else:
            full_file_name = path.join(DOCUMENT_PATH, dic_file_name) \
                + "." + DIC_FILENAME_EXTENSION
            id_regexp = compile(r"[0-9a-zа-яё_]+", IGNORECASE)
            try:
                with open(full_file_name, "r", encoding="utf8") \
                         as fhandler:
                    rows = reader(fhandler, delimiter=",")
                    lst: list =[]
                    first_row: list
                    for i, row in enumerate(rows):
                        if i == 0:
                            first_row = list(row)
                        elif len(first_row) != len(row):
                            is_valid = False
                        elif id_regexp.fullmatch(row[i]) == None:
                            is_valid = False
                        lst.append(row)
                    if is_valid:  
                        if first_row[0] != "@id@":
                            is_valid = False
                        else:
                            self._dic_file_name = dic_file_name      
            except:
                is_valid = False
        if not is_valid:
            self._dic_file_name = None
        return is_valid
    
    def set_document_file_name(self, 
                               document_file_name: str) -> bool:
        is_valid: bool
        if document_file_name == None or document_file_name == "":
            is_valid = False
        else:
            try:
                full_file_name = \
                    path.join(DOCUMENT_PATH, document_file_name) \
                        + "." + DOCUMENT_FILENAME_EXTENSION
                doc: Doc = Doc(full_file_name)
                self._document_file_name = document_file_name
                self._document_full_filename = full_file_name
                is_valid = True
            except:
                is_valid = False
        if not is_valid:
            self._dic_file_name = None
        return is_valid

    def get_full_filename_of_key_csv(
            self, key_list: list = []) -> str:
        replacer: Replacer = Replacer(self._document_file_name,
                                      self._dic_file_name)
        full_filename: str = \
            replacer.get_full_filename_of_key_csv(key_list)
        return full_filename
    
    def get_replaced_document_full_filename_list(
            self,
            list_of_replacement_field_dic: list = []) -> list:
        replacer: Replacer = Replacer(self._document_file_name,
                                      self._dic_file_name)
        document_full_filename_list: list = \
            replacer.get_replaced_document_full_filename_list(
                list_of_replacement_field_dic)
        return document_full_filename_list
    
    def set_document_paragraphs(
            self, fix_flag_dict: dict,
            document_full_filename_list: list):
        for document_full_filename in document_full_filename_list:
            document: Doc = Doc(document_full_filename)
            super()._style_init(document)
            for paragraph in document.paragraphs:
                if self._picture_caption_ptrn \
                       .match(paragraph.text) != None:
                    if fix_flag_dict["picture_caption"]:
                        super()._set_paragraph(
                            self._common_picture_caption_style,
                            paragraph,
                            self._common_paragraph_run)
                elif self._table_caption_ptrn \
                         .match(paragraph.text) != None:
                    if fix_flag_dict["table_caption"]:
                        super()._set_paragraph(
                            self._common_table_caption_style, 
                            paragraph,
                            self._common_paragraph_run)
                elif self._text_paragraph_ptrn \
                         .match(paragraph.text) != None:
                    if fix_flag_dict["text_paragraph"]:
                        super()._set_paragraph(
                            self._common_text_paragraph_style, 
                            paragraph,
                            self._common_paragraph_run)
            self._common_picture_caption_style.delete_tmp_paragraph()
            self._common_table_caption_style.delete_tmp_paragraph()
            self._common_text_paragraph_style.delete_tmp_paragraph()
            document.save(document_full_filename)

#r"D:\University\Mag\Course_1\TPIISRSII\KW\KW\template\t1.docx"
#r"D:\University\Mag\Course_1\TPIISRSII\KW\KW\template\Sereda_TPIISRSII_titul.docx"