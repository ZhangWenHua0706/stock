import os
import sys
import yaml


class YamlUtil:
    def read_yamlfile(self, file_name):
        with open(os.getcwd()+'\\'+file_name, mode='r', encoding='utf-8') as f:
            value = yaml.load(stream=f, Loader=yaml.FullLoader)
            return value

    def read_test_yamlfile(self, file_name):
        with open(os.getcwd()+'\\testcases\\'+file_name, mode='r', encoding='utf-8') as f:
            value = yaml.load(stream=f, Loader=yaml.FullLoader)
            return value

    def write_extract_yaml(self, file_name, data):
        with open(os.getcwd()+'\\'+file_name, encoding='utf-8', mode='w') as f:
            try:
                yaml.dump(data=data, stream=f, allow_unicode=True)
            except Exception as e:
                print(e)

    def read_extract_yaml(self, file_name,):
        with open(os.getcwd()+'\\'+file_name, encoding='utf-8', mode='r') as f:
            value = yaml.load(stream=f, Loader=yaml.FullLoader)
            return value

