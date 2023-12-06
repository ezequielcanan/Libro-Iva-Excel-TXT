import openpyxl as xl
from datetime import datetime

excelFilePath = "./IVA VENTAS - Guzman 10-2023.xlsx"

workbookObject = xl.load_workbook(excelFilePath)
sheetObject = workbookObject.active

minRow = sheetObject.min_row
maxRow = sheetObject.max_row

txtFile = open("LIBRO_IVA_DIGITAL_VENTAS_CBTE.txt", "x", encoding="cp1252")

tiposComprobantes = {
  "TCK B": "006",
  "FAC A": "001",
  "TCK  ": "100"
}


def writeData(row):
  for col in row:
    colCoordinate = col.coordinate[0]

    if col.value is None and colCoordinate == "A":
      break
    
    if row[1].value == "TCK  ":
      break

    match colCoordinate:
      case "A":
        date = datetime.strptime(col.value, "%d/%m/%Y").date().strftime("%Y%m%d")
        txtFile.write(date)
      case "B":
        tipoComprobante = tiposComprobantes[col.value]
        txtFile.write(f"{tipoComprobante}00001")
      case "C":
        nroComprobante = col.value.replace("-","").strip().zfill(20)
        txtFile.write(f"{nroComprobante}{nroComprobante}")
      case "D":
        if col.value == "CONSUMIDOR FINAL                                  ":
          txtFile.write(f"99{'0'.zfill(50)}")
      case "E":
        if col.value is not None:
          txtFile.write(f"80{col.value.replace('-','').strip().zfill(20)}{row[3].value.ljust(30)[0:30]}")
      case "F":
        total = str(col.value).zfill(15)
        txtFile.write(f"{total}{''.zfill(105)}PES00010000001 00000000000000000000000\n")



for row in sheetObject.iter_rows(min_row = 2):
  writeData(row)

txtFile.close()