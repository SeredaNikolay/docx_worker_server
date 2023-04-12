from os import path
from unittest import TestCase, main

from constant import DOCUMENT_PATH, TEST_DIR_NAME
from controller import Controller


class TestControllerMethods(TestCase):
  
    def test_set_dic_file_name(self):
        print("hi1")
        controller: Controller = Controller()
        self.assertFalse(
            controller.set_dic_file_name(None), "None value")
        self.assertFalse(
            controller.set_dic_file_name(""), "Empty name")
        test_file_path = \
            path.join(TEST_DIR_NAME, "there_is_not_such_file")
        self.assertFalse(
            controller.set_dic_file_name(test_file_path), \
            "There is not such file")
        test_file_path = \
            path.join(TEST_DIR_NAME, "empty")
        self.assertFalse(
            controller.set_dic_file_name(test_file_path), 
            "Empty")
        test_file_path = \
            path.join(TEST_DIR_NAME, "not_csv")
        self.assertFalse(
            controller.set_dic_file_name(test_file_path), "Not csv")
        test_file_path = \
            path.join(TEST_DIR_NAME, "headers_without_id")
        self.assertFalse(
            controller.set_dic_file_name(test_file_path), 
            "Headers without id")
        test_file_path = \
            path.join(TEST_DIR_NAME, "miss_value")
        self.assertFalse(
            controller.set_dic_file_name(test_file_path), 
            "Miss value")
        test_file_path = \
            path.join(TEST_DIR_NAME, "wrong_id")
        self.assertFalse(
            controller.set_dic_file_name(test_file_path), 
            "Wrong id")
        
        test_file_path = \
            path.join(TEST_DIR_NAME, "headers_with_id")
        self.assertTrue(
            controller.set_dic_file_name(test_file_path), 
            "Headers with id")
        test_file_path = \
            path.join(TEST_DIR_NAME, "normal")
        self.assertTrue(
            controller.set_dic_file_name(test_file_path), 
            "Normal data")
        
    def test_set_document_file_name(self):
        print("hi2")
        controller: Controller = Controller()
        self.assertFalse(
            controller.set_document_file_name(None), "None value")
        self.assertFalse(
            controller.set_document_file_name(""), "Empty name")
        test_file_path = \
            path.join(TEST_DIR_NAME, "there_is_not_such_file")
        self.assertFalse(
            controller.set_document_file_name(test_file_path), \
            "There is not such file")
        test_file_path = \
            path.join(TEST_DIR_NAME, "not_docx")
        self.assertFalse(
            controller.set_document_file_name(test_file_path), "Not docx")
        
        test_file_path = \
            path.join(TEST_DIR_NAME, "normal")
        self.assertTrue(
            controller.set_document_file_name(test_file_path), 
            "Normal data")
        
def run_tests():
    main(module="test", exit=False)    