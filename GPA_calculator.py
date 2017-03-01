#!/usr/bin/python  
# -*- coding: utf-8 -*-  
# __author__ = 'Ruiyang Ding'

import xlrd
class Gpa_Calc:

    def __init__(self,name,pku_gpa=None,zju_gpa=None,std_gpa=None,stdrev_gpa=None):
        self.name = name
        self.__pku_gpa = pku_gpa
        self.__zju_gpa = zju_gpa
        self.__std_gpa = std_gpa
        self.__stdrev_gpa = stdrev_gpa

    def grade_convert(self, originalScore):
        if originalScore=="优秀":
            return 95
        elif originalScore=="良好":
            return 85
        elif originalScore=="合格":
            return 65
        else:
            return originalScore

    def pku(self, sh):
        nrows = sh.nrows
        grade_point = []
        weight_list = []
        #print(sh.cell(3,4).value)
        for i in range(3, nrows-1):
            grade = float(self.grade_convert(sh.cell(i,4).value))
            weight = float(sh.cell(i,9).value)
            grade_pku = 4 - 3 * (100 - grade) * (100 - grade) / 1600
            grade_point.append(grade_pku*weight)
            weight_list.append(weight)
        gpa = sum(grade_point)/sum(weight_list)
        self.__pku_gpa = gpa
        
    def zju(self, sh):
        nrows = sh.nrows
        grade_point = []
        weight_list = []
        for i in range(3, nrows-1):
            grade = float(self.grade_convert(sh.cell(i,4).value))
            weight = float(sh.cell(i,9).value)
            if grade<60:
                grade_zju = 0
            elif grade>=60 and grade<=85:
                grade_zju = 4-(85-grade)/10
            elif grade>85:
                grade_zju = 4
            grade_point.append(grade_zju*weight)
            weight_list.append(weight)
        gpa = sum(grade_point) / sum(weight_list)
        self.__zju_gpa = gpa
    def std(self, sh):
        nrows = sh.nrows
        grade_point = []
        weight_list = []
        for i in range(3, nrows-1):
            grade = float(self.grade_convert(sh.cell(i,4).value))
            weight = float(sh.cell(i,9).value)
            if grade<60:
                grade_std = 0
            elif grade>=60 and grade<70:
                grade_std = 1
            elif grade>=70 and grade<80:
                grade_std = 2
            elif grade>=80 and grade<90:
                grade_std = 3
            else:
                grade_std = 4
            grade_point.append(grade_std*weight)
            weight_list.append(weight)
        gpa = sum(grade_point) / sum(weight_list)
        self.__std_gpa = gpa
    def stdrev(self, sh):
        nrows = sh.nrows
        grade_point = []
        weight_list = []
        for i in range(3, nrows-1):
            grade = float(self.grade_convert(sh.cell(i,4).value))
            weight = float(sh.cell(i,9).value)
            if grade<60:
                grade_stdrev = 0
            elif grade>=60 and grade<75:
                grade_stdrev = 2
            elif grade>=75 and grade<85:
                grade_stdrev = 3
            else:
                grade_stdrev = 4
            grade_point.append(grade_stdrev*weight)
            weight_list.append(weight)
        gpa = sum(grade_point) / sum(weight_list)
        self.__stdrev_gpa = gpa
    def get_pkugpa(self):
        return self.__pku_gpa
    def get_zjugpa(self):
        return self.__zju_gpa
    def get_stdgpa(self):
        return self.__std_gpa
    def get_stdrevgpa(self):
        return self.__stdrev_gpa
def get_sheet(name):
    fname = name
    bk = xlrd.open_workbook(fname)
    try:
        sh = bk.sheet_by_index(0)
    except:
        print("no sheet in %s named Sheet1" % fname)
    return sh
