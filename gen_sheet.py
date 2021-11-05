import openpyxl as excel

# excel file gen
wb = excel.Workbook()

B4: tuple[str] = ("荻野", "黒須", "斎藤", "下笠", "田代")
M1: tuple[str] = ("荒井", "大島", "武田", "立石", "野村")
M2: tuple[str] = ("岩舘", "谷", "西川", "西田", "森川")
# add sheet
sheet_name: tuple[str] = ("B4", "M1", "M2")
map(lambda x: wb.create_sheet(title=x), sheet_name)


print("ファイル名を入力")
file_name: str = input()

if ".elsx" in file_name:
    wbname: str = file_name
else:
    wbname: str = file_name + ".xlsx"








wb.save(wbname)
