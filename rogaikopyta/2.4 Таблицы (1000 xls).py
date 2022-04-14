import dload
import zipfile
import xlrd
import xlwt

url = "https://stepik.org/media/attachments/lesson/245299/rogaikopyta.zip"
dload.save_unzip(url, "rogaikopyta.zip")  # скачается архив

d1 = dict()
archiv = zipfile.ZipFile("rogaikopyta.zip", "r")
for filename in archiv.infolist():
    if filename.filename[-5:] == ".xlsx":
        archiv.extract(filename)  # распоковываем по очереди файлы из архива
        wb = xlrd.open_workbook(filename.filename)
        sheet = wb.sheet_by_index(0)  # выбираем активный лист
        vls = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
        d1[vls[1][1]] = int(vls[1][3])

  # цикл сортировки по алфавиту и записи в новую книгу
book = xlwt.Workbook()
sheet1 = book.add_sheet("Zarplata")
total = 0
for key, value in sorted(d1.items(), key=lambda x: x[0]):
    sheet1.write(total, 0, key)
    sheet1.write(total, 1, value)
    print(key, value)
    total += 1

  # ради забавы...подсчет суммы, мах, min
sheet1.write(total, 0, "Итого")
sheet1.write(total, 1, sum(d1.values()))
total += 1
d1 = sorted(d1.items(), key=lambda x: x[1])
sheet1.write(total, 0, "Cамый богатый")
sheet1.write(total, 1, d1[-1][0])
sheet1.write(total, 2, d1[-1][1])
total += 1
sheet1.write(total, 0, "Cамый бедный")
sheet1.write(total, 1, d1[0][0])
sheet1.write(total, 2, d1[0][1])
book.save("Зарплата всех сотрудников.xls")
