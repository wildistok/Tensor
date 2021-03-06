                    #standard modules
import math         #module for exp (methods: calc_index_proj and lider_index)
import os           #module for get files in dir(method: get_files)
import sys          #module for get name of system for get path(method: get_symbol)

import xlrd         #lib for extract data from excel files


def open_file(filename):
    """open file for reading"""
    book = xlrd.open_workbook(filename)     
    sheet = book.sheet_by_index(0)          
    return sheet


def get_files():
    """get the list of files xls and xlsx in directory"""
    excel_files = []
    dir = ''                    
    if (len(sys.argv) > 1):
        dir = sys.argv[1]
        files = os.listdir(dir)
    else:
        files = os.listdir(".")

    for file in files:          #get only exele file from all files
        if((file[-4:] == ".xls") or (file[-5:] == ".xlsx")):
            if (dir != ''):
                excel_files.append(dir + get_symbol() + file)
            else:
                excel_files.append(file)    
    return excel_files


def get_symbol():
    """get symbol for do correct path in different os"""
    system = os.name
    if (system == 'nt'):
        return "\\"
    else:
        return "/"


def get_names(sheet, names):
    """get list of employees"""
    for colx in range(1, sheet.ncols, 2):
        if (colx == 1):                                 #add in list names of managers
            for rowx in range(1, sheet.nrows):  
                value = sheet.cell_value(rowx, colx)    
                if (value not in names) and (value != ''):
                    names.append(value)
        rowx = 0
        if (colx > 3):                                  #add all the others employees
            value = sheet.cell_value(rowx, colx).lower()
            value = value[0:value.find(" факт")]
            value = value.title();
            if (value not in names):
                names.append(value)


def find_worker_mark(sheet, rowx, colx):
    """calc efficiency mark for one programmer in project"""
    mark = 0

    plan_value = 0
    if (sheet.cell_type(rowx, colx) != 0):
        plan_value = sheet.cell_value(rowx, colx)

    actual_value = 0
    if (sheet.cell_type(rowx, colx + 1) != 0):
        actual_value = sheet.cell_value(rowx, colx + 1)

    if (plan_value == 0):
        if (actual_value > 0):
            mark = 1
        else:
            mark = 0
    else:
        mark = plan_value / actual_value
    return mark


def avg_worker_mark(sheet, name):
    """avg of efficeincy one programmer"""
    avg = 0
    sum = 0
    num = 0
    colx = find_colx(sheet, name)
    if (colx != 0):
        for i in range(1, sheet.nrows):
            sum += find_worker_mark(sheet, i, colx)
        num = find_sum_project(sheet, colx)
        if (num != 0):
            avg = sum / num
    return avg


def find_colx(sheet, name):
    """for find colx in row(0) by programmer name"""
    colx = 0
    row = sheet.row_values(0)
    for el in row:
        if not(el.find(name)):
            colx = row.index(el)
            break
    return colx


def find_sum_project(sheet, colx):
    """number of projects"""
    num = 0
    for i in range(1, sheet.nrows):
        if (((sheet.cell_type(i, colx) == 2) and 
            (sheet.cell_value(i, colx) != 0)) or 
            (sheet.cell_type(i, colx + 1) == 2)):
            num += 1
    return num


def calc_index_proj(sheet, name):
    """index of projects of quantity"""
    colx = find_colx(sheet, name)
    num = find_sum_project(sheet, colx)
    index = 1
    if (num  > 1):
        index += math.e ** (-num)
    return index


def avg_proj(sheet, row):
    """avg efficiency in project"""
    sum = 0
    num = 0
    for colx in range(4, sheet.ncols, 2):
        sum = sum + find_worker_mark(sheet, row, colx)
        if (find_worker_mark(sheet, row, colx) != 0):
            num += 1
    avg = sum / num
    return avg


def calc_lider_proj(sheet, name):
    """number for manager's projects"""
    num = 0
    col = sheet.col_values(1)
    for cell in col:
        if cell.find(name):
            num += 1
    return num


def avg_lider_mark(sheet, name):
    """avg of manager efficeincy""" 
    num = 0
    sum = 0
    avg = 0
    colx = 1
    for rowx in range(1, sheet.nrows):
        if (sheet.cell_value(rowx, colx) == name):
            plan_date = sheet.cell_value(rowx, 2)
            actual_date = sheet.cell_value(rowx, 3)
            if (plan_date > actual_date):
                num += 1
                sum += avg_proj(sheet, rowx)
    if (num != 0):
        avg = sum / num * lider_index(num)
    return avg


def lider_index(num):
    """index of projects of quantity"""
    index = 1
    if (num > 1):
        index += math.e ** (-num)
    return index


def eff_index(sheet, name):
    """calc full efficiency factor"""
    index = avg_worker_mark(sheet, name) * calc_index_proj(sheet, name) + avg_lider_mark(sheet, name)
    return index


#main code
if __name__ == "__main__":
    exel_files = get_files()
    names = [] 
    for file in exel_files:
        sheet = open_file(file)         #get first sheet of book
        get_names(sheet, names)

    #generate a dict with the names of employees and effectiveness = 0
    eff_workers = {name: 0 for name in names}
    for file in exel_files:
        sheet = open_file(file)
        for name in eff_workers:
            value = eff_index(sheet, name)
            if (value != 0):
                eff_workers[name] += value


    #sorted dict of employes by effectiveness
    sorted_data = sorted(eff_workers.items(), key=lambda x:x[1], reverse=True)
    for data in sorted_data:
        print(data[0])
