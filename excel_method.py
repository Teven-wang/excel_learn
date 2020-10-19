import hashlib
import os
import xlrd
import requests
from tornado.web import RequestHandler


class get_excel_data(RequestHandler):
    '''
    可以用flask或者django或者tornado获取文件
    '''

    def post(self):
        try:
            file = self.get_files('file')
            print('收取excel文件')
        except Exception as e:
            print(str(e))
            return
        excel_path = self.save_excel_data(file)
        try:
            checker = excel_deal(excel_path)
            print('初始化')
            res = checker.get_table_content()
        except Exception as e:
            print(str(e))
            return

        return

    def get_files(self, key):
        if key in self.request.files:
            file_metas = self.request.files[key][0]['body']
            return file_metas
        else:
            return None

    def save_excel_data(self, file):
        file_name = self.get_file_md5(file)
        excel_path = os.getcwd() + file_name
        open(excel_path, 'wb').write(file)
        return excel_path

    def get_file_md5(self, binary):
        '''
        文件md5转换
        :param binary:
        :return:
        '''
        md5 = hashlib.md5(binary).hexdigest()
        return md5


class excel_deal():
    def __init__(self, table_name):
        self.table_name = table_name
        self.sheet = self.table_content(table_name)

    def table_content(self, file_name):
        wb = xlrd.open_workbook(filename=file_name)  # 打开文件
        print(wb.sheet_names())  # 获取所有工作表名称
        sheet1 = wb.sheet_by_index(0)  # 通过索引获取工作表1
        return sheet1

    def get_table_content(self):
        all_data = []
        row = 1
        while 1:
            try:
                all_line = self.sheet.row_values(row)
                row += 1
                print(all_line)
                all_data.append(all_line)
            except Exception as e:
                print(e)
                break
        print(row)
        # print(all_data)
        return all_data


if __name__ == '__main__':
    excel = excel_deal('更改接点账号.xls')
    excel.get_table_content()
