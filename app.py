from flask import Flask, request, send_file, Response, Request
from zipfile import ZipFile
from os import path, remove, listdir
from waitress import serve
from test import run_tests
from controller import BaseController, Controller
from constant import DOCUMENT_PATH, DIC_FILENAME_EXTENSION, \
                     DOCUMENT_FILENAME_EXTENSION, ARCHIVE_EXTENSION

def get_fix_flag_dic(request, fix_flag_dic: dict):
    text_paragraph_fix: bool = \
        request.args.get('text-paragraph-fix', default = False,
                         type = bool)
    picture_caption_fix: bool = \
        request.args.get('picture-caption-fix', default = False,
                         type = bool)
    table_caption_fix: bool = \
        request.args.get('table-caption-fix', default = False, 
                         type = bool)
    fix_flag_dic.update(
        {"picture_caption": picture_caption_fix, 
         "table_caption": table_caption_fix, 
         "text_paragraph": text_paragraph_fix})
    
def get_full_filename_for_send_file(full_filename_list: list) -> str:
    name: str = ""
    if len(full_filename_list) == 1:
        name = full_filename_list[0]
    elif len(full_filename_list) > 1:
        path_to_dir: str = path.split(full_filename_list[0])[0]
        arch_name: str = \
            path.split(full_filename_list[0])[1] \
                .replace("." + DOCUMENT_FILENAME_EXTENSION, "")
        full_arch_filename = \
            path.join(path_to_dir, 
                      arch_name + "." + ARCHIVE_EXTENSION)
        with ZipFile(full_arch_filename, 'w') as zf:
            for full_filename in full_filename_list:
                zf.write(full_filename, path.split(full_filename)[1])
        name = full_arch_filename
    return name

app = Flask(__name__)

@app.route("/document/<document_name>/dictionary-file",
           methods=["GET"])
def get_dictionary_file(document_name: str):
    controller: BaseController = Controller()
    is_valid_document: bool = \
        controller.set_document_file_name(document_name)
    if is_valid_document:
        dictionary_full_filename: str = \
            controller.get_full_filename_of_key_csv()
        send_full_filename: str = \
            get_full_filename_for_send_file(
                [dictionary_full_filename])
        return send_file(send_full_filename, 
                         as_attachment=True,
                         download_name=\
                            path.split(send_full_filename)[1])
    return Response(status=400)

@app.route("/dictionary_file/<dictionary_file_name>",
           methods=["POST"])
def load_dictionary_file(dictionary_file_name: str):
    file = request.files.get("dictionary_file", default=None)
    if file != None:
        filename = dictionary_file_name + "." \
                   + DIC_FILENAME_EXTENSION
        dst_path: str = path.join(DOCUMENT_PATH, filename)
        file.save(dst_path)
        return Response(status=200)
    return Response(status=400)

@app.route("/cleaning", methods=["DELETE"])
def delete_tmp_files():
    for filename in listdir(DOCUMENT_PATH):
        try:
            full_filename: str = path.join(DOCUMENT_PATH, filename)
            if path.isfile(full_filename):
                remove(full_filename)
        except:
            pass
    return Response(status=200)

@app.route("/document/<document_file_name>/fixed-document-file",
           methods=["GET"])
def get_fixed_document_file(document_file_name: str):
    controller: BaseController = Controller()
    is_valid_document: bool = \
        controller.set_document_file_name(document_file_name)
    if is_valid_document:
        doc_full_filename: str = \
            controller.get_document_full_filename()
        fix_flag_dic: dict = {}
        get_fix_flag_dic(request, fix_flag_dic)
        controller.set_document_paragraphs(fix_flag_dic, 
                                        [doc_full_filename])
        send_full_filename: str = \
            get_full_filename_for_send_file([doc_full_filename])
        return send_file(send_full_filename, 
                         as_attachment=True, 
                         download_name= \
                            path.split(send_full_filename)[1])
    return Response(status=400)

@app.route("/document/<document_file_name>", methods=["POST"])
def load_document_file(document_file_name: str):
    file = request.files.get("document", default=None)
    if file != None:
        filename = document_file_name + "." \
                   + DOCUMENT_FILENAME_EXTENSION
        dst_path: str = path.join(DOCUMENT_PATH, filename)
        file.save(dst_path)
        return Response(status=200)
    return Response(status=400)

@app.route(
    "/document/<document_name>" \
        + "/<dictionary_file_name>/fixed-document",
    methods=["GET"])
def get_document_with_replaced_fields(
        document_name: str, dictionary_file_name: str):
    controller: BaseController = Controller()
    is_valid_document: bool = \
        controller.set_document_file_name(document_name)
    is_valid_dic_file: bool = \
        controller.set_dic_file_name(dictionary_file_name)
    if is_valid_document and is_valid_dic_file: 
        fix_flag_dic: dict = {}
        get_fix_flag_dic(request, fix_flag_dic)
        replaced_document_full_filename_list: list = \
            controller.get_replaced_document_full_filename_list()
        controller.set_document_paragraphs(
            fix_flag_dic, replaced_document_full_filename_list)
        send_full_filename: str = \
            get_full_filename_for_send_file(
                replaced_document_full_filename_list)
        return send_file(send_full_filename, 
                         as_attachment=True,
                         download_name= \
                            path.split(send_full_filename)[1])
    return Response(status=400)

if __name__ == '__main__':
    run_tests()  
    serve(app, host="127.0.0.1", port=8080)