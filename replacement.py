from os import path
from zipfile import ZipFile, ZIP_DEFLATED
from os import rename, remove
from typing import Any
from re import compile, sub, findall
from shutil import rmtree
from csv import writer, reader, QUOTE_ALL
from copy import deepcopy

from constant import XML_PATH, XML_FILENAME, ARCHIVE_EXTENSION, \
                     TAGS_ONLY, LEFT_FIELD_SEARCH_REGEXP_PART, \
                     RIGHT_FIELD_SEARCH_REGEXP_PART, DOCUMENT_PATH, \
                     DOCUMENT_FILENAME_EXTENSION, \
                     DIC_FILENAME_EXTENSION, \
                     FIELD_SEARCH_REGEXP_STR, \
                     GOOD_KEY_PARTS_SEARCH_REGEXP_STR, \
                     BAD_KEY_PARTS_SEARCH_REGEXP_STR

class Replacer:

    _document_file_name: str
    _document_full_file_name: str
    _document_full_filename: str
    _dic_file_name: str
    _dic_full_file_name: str
    _dic_full_filename: str
    _archive_full_file_name: str
    _unpacked_dir_full_name: str
    _archive_full_filename: str
    _xml_full_filename: str
    _field_ptrn: Any
    _good_key_parts_ptrn: Any
    _bad_key_parts_ptrn: Any
    _original_xml_line_dic: dict

    
    def __init__(
            self, document_file_name: str, 
            dic_file_name: str = None):
        self._document_file_name = document_file_name
        self._document_full_file_name = \
            path.join(DOCUMENT_PATH, self._document_file_name)
        self._document_full_filename = \
            self._document_full_file_name \
            + "." + DOCUMENT_FILENAME_EXTENSION
        self._dic_file_name = dic_file_name
        if self._dic_file_name == None:
            self._dic_full_file_name = None
            self._dic_full_filename = None
        else:
            self._dic_full_file_name = \
                path.join(DOCUMENT_PATH, self._dic_file_name)
            self._dic_full_filename = \
                self._dic_full_file_name \
                + "." + DIC_FILENAME_EXTENSION
        self._archive_full_file_name = self._document_full_file_name
        self._unpacked_dir_full_name = self._document_full_file_name
        self._archive_full_filename = self._document_full_file_name \
                                      + "." + ARCHIVE_EXTENSION
        self._xml_full_filename = path.join(DOCUMENT_PATH,
                                            self._document_file_name,
                                            XML_PATH, XML_FILENAME)
        self._field_ptrn = compile(FIELD_SEARCH_REGEXP_STR)
        self._good_key_parts_ptrn = \
            compile(GOOD_KEY_PARTS_SEARCH_REGEXP_STR)
        self._bad_key_parts_ptrn = \
            compile(BAD_KEY_PARTS_SEARCH_REGEXP_STR)
        self._original_xml_line_dic = {"info": "", "text": ""}
        self.__document_to_directory()
        self.__get_inner_xml_file_parts()

    def __del__(self):
        remove(self._archive_full_filename)
        rmtree(self._unpacked_dir_full_name)

    def __document_to_directory(self) -> list:
        err: list = [0, ""]
        try:
            rename(self._document_full_filename, 
                   self._archive_full_filename)
            try:
                with ZipFile(self._archive_full_filename, 'r') \
                     as zhandler:
                    zhandler.extractall(self._unpacked_dir_full_name)
            except OSError:
                err = [2, "This archive cannot be unpacked"]
        except OSError:
            err = [1, "This document cannot be renamed"]
        return err

    #Does not delete the directory itself
    def __directory_to_document(self, 
                                new_document_file_name: str) -> list:
        err: list = [0, ""]
        xml_arch_doc_full_fname: str = \
            path.join(XML_PATH, XML_FILENAME)
        out_arch_full_filename = \
            self._archive_full_file_name + "_." + ARCHIVE_EXTENSION
        try:
            with ZipFile(self._archive_full_filename) as in_arch, \
                 ZipFile(out_arch_full_filename, "w", 
                        compression=ZIP_DEFLATED) as out_arch:
                for arch_file in in_arch.infolist():
                    with in_arch.open(arch_file) as file:
                        # Filename gives linux separator
                        xml_name: str = xml_arch_doc_full_fname \
                                            .replace("\\", "/")
                        if arch_file.filename == xml_name:
                            with open(self._xml_full_filename, 
                                    encoding="utf8") as fhandler:
                                content = \
                                    "".join(fhandler.readlines())
                        else:
                            content = \
                                in_arch.read(arch_file.filename)
                        out_arch.writestr(arch_file.filename, 
                                          content)
            try:
                new_document_full_filename: str = \
                    path.join(DOCUMENT_PATH, new_document_file_name \
                              + "." + DOCUMENT_FILENAME_EXTENSION)
                rename(out_arch_full_filename,
                       new_document_full_filename)
            except OSError as e:
                err = [2, "This archive cannot be packed"]
        except OSError as e:
            err = [1, "This archive cannot be renamed"]

        return err

    def __get_inner_xml_file_parts(self) -> list:
        err: list = [0, ""]
        try:
            with open(self._xml_full_filename, encoding="utf8") \
                 as fhandler:
                # 1st line - info, 2nd - document xml
                try:
                    line_list = fhandler.readlines()
                    if (len(self._original_xml_line_dic) == 2 \
                           and len(line_list) == 2):
                        self._original_xml_line_dic.update(
                            {"info": line_list[0]})
                        self._original_xml_line_dic.update(
                            {"text": line_list[1]})
                    else:
                        err = [3, "Wrong lines count"]
                except Exception:
                    err = [2, "Wrong xml format"]
        except OSError:
            err = [1, "This file cannot be opened"]
        return err

    def __create_csv(self, key_list: list) -> str:
        new_dic_full_filename: str = \
            path.join(DOCUMENT_PATH, self._document_file_name) \
            + "." + DIC_FILENAME_EXTENSION
        with open(new_dic_full_filename, 'w') as fhandler:
            wr = writer(fhandler, delimiter=",", quoting=QUOTE_ALL)
            wr.writerow(key_list)
        return new_dic_full_filename
    
    def __get_key_list(self, key_list: list) -> None:
        field_list = \
            self._field_ptrn \
                .findall(self._original_xml_line_dic["text"])
        part = None
        if "@id@" not in key_list:
            key_list.append("@id@")
        for key_parts_str in field_list:
            part = self._bad_key_parts_ptrn.findall(key_parts_str)
            if len(part) == 0:
                key: str = "".join(self._good_key_parts_ptrn\
                                       .findall(key_parts_str))
                if key not in key_list:
                    key_list.append(key)

    def __replace_fields(self, replacement_field_dic: dict) -> list:
        err: list = [0, ""]
        xml_line_dic: dict = deepcopy(self._original_xml_line_dic)
        for key, value in replacement_field_dic.items():
            inner_part: str = TAGS_ONLY \
                            + TAGS_ONLY.join(list(key)) \
                            + TAGS_ONLY
            reg_exp: str = LEFT_FIELD_SEARCH_REGEXP_PART \
                        + inner_part \
                        + RIGHT_FIELD_SEARCH_REGEXP_PART
            search_ptrn = compile(reg_exp)
            replaced_str: str = sub(search_ptrn,
                                    "<w:t>" + value + "</w:t>", 
                                    xml_line_dic["text"])
            xml_line_dic.update({"text": replaced_str})
            try:
                with open(self._xml_full_filename, "w", 
                        encoding="utf8") as fhandler:
                    fhandler.write(xml_line_dic["info"])
                    fhandler.write(xml_line_dic["text"])
            except Exception:
                err = [1, "This xml document cannot be opened"]
        return err
    
    def set_dic_file_name(self, dic_file_name: str):
        self._dic_file_name = dic_file_name
        self._dic_full_file_name = \
                path.join(DOCUMENT_PATH, self._dic_file_name)
        self._dic_full_filename = \
            self._dic_full_file_name \
            + "." + DIC_FILENAME_EXTENSION
    
    def get_full_filename_of_key_csv(
            self, key_list: list = []) -> str:
        self.__get_key_list(key_list)
        csv_full_filename: str = self.__create_csv(key_list)
        self.__directory_to_document(self._document_file_name)
        return csv_full_filename

    def get_replaced_document_full_filename_list(
            self,
            list_of_replacement_field_dic: list = []) -> list:
        if self._dic_file_name == None:
            return []
        document_full_filename_list: list = []
        self._xml_line_dic = deepcopy(self._original_xml_line_dic)
        with open(self._dic_full_filename,
                  encoding="utf8") as fhandler:
            csv_reader = reader(fhandler, delimiter=",")
            key_list: list = []
            for i, row in enumerate(csv_reader):
                if i != 0:
                    list_of_replacement_field_dic.append({})
                    for j, key in enumerate(key_list):
                        list_of_replacement_field_dic[-1] \
                            .update({key: row[j]})
                else:
                    key_list = list(row)
        for i in range(len(list_of_replacement_field_dic)):
            id: str = list_of_replacement_field_dic[i]["@id@"]
            self.__replace_fields(list_of_replacement_field_dic[i])
            self.__directory_to_document(self._document_file_name \
                                         + "_" + id)
            document_full_filename: str = \
                path.join(DOCUMENT_PATH, 
                          self._document_file_name + "_" + id) \
                    + "." + DOCUMENT_FILENAME_EXTENSION
            document_full_filename_list \
                .append(document_full_filename)
        return document_full_filename_list