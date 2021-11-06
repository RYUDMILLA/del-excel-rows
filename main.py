#-*- coding: utf-8 -*-
import openpyxl
import csv

def save_to_file(dir_file, name):
    file = open(dir_file, mode = 'a', newline='')
    writer = csv.writer(file)
    writer.writerow([name])
    file.close()
    return
  
def initialize_csv(file):
    f = open(file, 'w+')
    writer = csv.writer(f)
    f.close()    

def find_and_delete(name, col):
    tmp = 1
    found = False
    for cell in ws[col]:        # search for column 
        if cell.value != None:     
            if name.strip() == cell.value.strip():  
                ws.delete_rows(tmp)
                tmp -= 1
                found = True
                break
            tmp += 1
    assert found != False

excel_file = "excel.xlsx"       # excel file to delete rows (not change)
find_list = "name.csv"          # elements you want to find and delete
result = "result.xlsx"          # row deleted excel file
deleted = "deleted.csv"         # found and deleted elements
error = "error.csv"             # not found elements or something errored

col = 'B'   # what column search for?

wb = openpyxl.load_workbook(filename=excel_file)
ws = wb.active
names = open(find_list, 'rt', encoding='utf-8', newline='')
initialize_csv(error)

total = 0
errors = 0
for name in names:
    try:
        total += 1
        find_and_delete(name,col)
        print(f"{name.strip()} deleted")
        save_to_file(deleted,name)
    except AssertionError:
        print(f"{name.strip()} not found")
        save_to_file(error,name)
        errors += 1
    except AttributeError:
        print(f"{name.strip()} attribute error")
        save_to_file(error.name)
        errors += 1

print(f"total : {total} / deleted : {total-errors} / error : {errors}")
wb.save(filename=result)
wb.close()

