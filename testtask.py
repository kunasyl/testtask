import os, openpyxl

def init():
    path = os.getcwd()
    dirs = os.walk(path)

    wb = openpyxl.Workbook() 
    sheet = wb.active
    sheet.append(('Номер строки', 'Папка в которой лежит файл', 'название файла', 'расширение файла'))
    
    data = []
    counter = 0
    for i, d in enumerate(dirs):
        for f in d[2]:
            counter += 1
            f_name = f.split('.', maxsplit=1)
            r = f_name[1] if len(f_name)>1 else ""
            data.append([counter, d[0], f_name[0], r])

    for row in data:
        sheet.append(row)
    wb.save(f'{path}/result.xlsx')

if __name__ == '__main__':
    init()