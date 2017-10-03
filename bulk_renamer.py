import os, openpyxl, shutil


# this function I didn't use for Paw Patrol
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


os.chdir(r"H:\Downloads\Server Downloads\Complete\TV Shows (Kids)\Paw Patrol")  # change cwd to the desired directory
dir = os.getcwd()
abspath = os.path.abspath('.')  # define abspath


file_list = [filename for filename in os.listdir(dir)]  # for each file in folder
file_list.sort()
for file in file_list:
     print(file)


def excel_export(eps_list):     #### THIS FUNCTION IS THE EXPORT TO EXCEL  #####
    wb = openpyxl.Workbook()  # create excel workbook object
    wb.save('paw patrol episodes.xlsx')  # save workbook as admin.xlsx
    sheet = wb.get_active_sheet()  # create sheet object as the Active sheet from the workbook object
    wb.save('paw patrol episodes.xlsx')  # save workbook as admin.xlsx
    # LIST-BASED POPULATION OF EXCEL SHEET
    for row in range(1, len(eps_list)):
        cell = sheet.cell(row=row, column=1)
        v = eps_list[row - 1]
        cell.value = v  # write the value (v) to the cell
    wb.save('paw patrol episodes.xlsx')  # save workbook as admin.xlsx


# Step 1 is to create the initial excel file with all the episode names in the left column, so turn this on:
# excel_export(file_list)


# Step 2 is to fill in the second column with the actual names - need to do this manually

# Step 3 - this section does the renaming

"""


# point the program to your 2-column file (commented off by default):
xls = r"H:\Downloads\Server Downloads\Complete\TV Shows (Kids)\Paw Patrol\Paw Patrol episodes all matched up.xlsx"


wb = openpyxl.load_workbook(xls)  # create excel workbook object
sheet = wb.get_sheet_by_name('Sheet')

originals = []
for i in range(1,86):
    originals.append(sheet['A{}'.format(i)].value)

new_names = []
for i in range(1,86):
    new_names.append(sheet['B{}'.format(i)].value)


for i in range(0,85):
    # print("{} | {}".format(originals[i], new_names[i]))
    original_with_path = os.path.join(r'H:\Downloads\Server Downloads\Complete\TV Shows (Kids)\Paw Patrol', originals[i])
    revised_with_path = os.path.join(r'H:\Downloads\Server Downloads\Complete\TV Shows (Kids)\Paw Patrol\new_files', new_names[i])  # new file in different directory
    revised_with_same_path = os.path.join(r'H:\Downloads\Server Downloads\Complete\TV Shows (Kids)\Paw Patrol', new_names[i])   # used when renaming on the spot
    print("{} | \n{}\n\n".format(original_with_path, revised_with_same_path))
    shutil.copy(original_with_path, revised_with_path) # this is the safer option - creating copies of files
    # shutil.move(original_with_path, revised_with_same_path) # this is just to rename on the spot

"""
