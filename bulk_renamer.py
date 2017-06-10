import os, openpyxl


def rename_removing_0():
    for i in range(0, len(file_list)):
        if '04x' in file_list[i]:
            oldname = file_list[i]
            # dud_b1 = fname[0:13]
            # print(dud_b1)
            new_b1 = oldname[0:12]
            # print(new_b1)
            remainder = oldname[14:]
            # print(remainder)
            middle = '4'
            new_name = new_b1 + middle + remainder
            old_with_path = os.path.join(dir, oldname)
            new_with_path = os.path.join(dir, new_name)
            print('Renaming {} to {}'.format(old_with_path, new_with_path))
            os.rename(old_with_path, new_with_path)
            # basename, ext = os.path.splitext(filename)  # isolate basename and extensions


os.chdir(r"H:\Downloads\Server Downloads\Complete\TV Shows (Kids)\Peppa Pig")  # change cwd to the desired directory
dir = os.getcwd()
abspath = os.path.abspath('.')  # define abspath
file_list = [filename for filename in os.listdir(dir)]  # for each file in folder
file_list.sort()
for file in file_list:
    print(file)


def excel_export(eps_list):     #### THIS FUNCTION IS THE EXPORT TO EXCEL  #####
    wb = openpyxl.Workbook()  # create excel workbook object
    wb.save('peppa episodes.xlsx')  # save workbook as admin.xlsx
    sheet = wb.get_active_sheet()  # create sheet object as the Active sheet from the workbook object
    wb.save('peppa episodes.xlsx')  # save workbook as admin.xlsx
    # LIST-BASED POPULATION OF EXCEL SHEET
    for row in range(1, len(eps_list)):
        cell = sheet.cell(row=row, column=1)
        v = eps_list[row - 1]
        cell.value = v  # write the value (v) to the cell
    wb.save('peppa episodes.xlsx')  # save workbook as admin.xlsx

excel_export(file_list) # creates an excel file with one column which is the names of the Peppa episodes
