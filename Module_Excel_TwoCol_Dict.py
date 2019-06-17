# -*- coding: utf-8 -*-
import xlrd
import xlwt
from datetime import date,datetime
import logging

class generate_static():
    
    def __init__(self,source_file,source_sheet):
        self.workbook=xlrd.open_workbook(source_file)
        self.sheet=self.workbook.sheet_by_name(source_sheet)
        self.name_list=[]
        self.name_dict={}
        self.avg_ack_dict={}
        self.ack_count_dict={}

    def generate_dict(self):

        ack_time_list = self.sheet.col_values(9,start_rowx=1,end_rowx=None)
        ack_name_list = self.sheet.col_values(17,start_rowx=1,end_rowx=None)
        service_name_list = self.sheet.col_values(4,start_rowx=1,end_rowx=None)
        asign_name_list = self.sheet.col_values(19,start_rowx=1,end_rowx=None)

        for ack_name in ack_name_list: 
            if ack_name != '':
                index = [x for x in range(len(ack_name_list)) if ack_name_list[x] == ack_name and service_name_list[x] == 'Direct Platform - Direct Cloud']
                if index:
                    self.name_list.append([ack_name,index])
        self.name_list = dict(self.name_list)            
        
        for name,index_name in self.name_list.items():
            ack_time = 0
            for index_time in index_name:
                if index == 'empty':
                    ack_time += 1800
                else:
                    ack_time += int(ack_time_list[index_time])
            avg_time = round( (ack_time / int(len(index_name))) / 60,2)        
            self.avg_ack_dict[name]=avg_time
            self.ack_count_dict[name]=len(index_name) 
            
   #     logging.debug(self.avg_ack_dict)
   #     logging.debug(self.ack_count_dict)
        return self.avg_ack_dict
        return self.ack_count_dict


if __name__ == '__main__':
  #  logging.basicConfig(level=logging.DEBUG,filename='C:\\Users\\fzhang2\\Desktop\\file\\test1.log', filemode='w')
    a=generate_static('C:\\Users\\fzhang2\\Desktop\\file\\incidents_new.xls','Sheet1')
    a.generate_dict()