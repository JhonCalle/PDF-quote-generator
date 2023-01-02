import os
import getPDFInfo
import openpyxl

def check_num_cotiz(directory):
    # if there are more than 10 excel files, delete the 5 oldest
    excel_files = []
    for file_name in os.listdir(directory):
        if file_name.endswith(".xlsx"):
            excel_files.append(os.path.join(directory, file_name))
    if len(excel_files) > 10:
        excel_files.sort(key=os.path.getmtime)
        for i in range(5):
            os.remove(excel_files[i])

    # get only the name (no extention file) of the most recent file inside the directory
    excel_name = os.path.splitext(os.path.basename(excel_files[-1]))[0]
    num_cotiz = int(excel_name)+1
    return str(num_cotiz)

def inset_data_in_excel(ws, pdf_names, pdf_pages, pdf_size):
    items = len(pdf_names)
    for i in range(items):
        if i ==0:
            ws.cell(row=6, column=2).value = 1
            ws.cell(row=6, column=3).value = pdf_names[i]
            ws.cell(row=6, column=4).value = pdf_pages[i]
            ws.cell(row=6, column=5).value = pdf_size[i]
            ws.cell(row=6, column=6).value = "Blanco y negro"
        elif i == items-1:
            ws.cell(row=6+i, column=2).value = i+1
            ws.cell(row=6+i, column=3).value = pdf_names[i]
            ws.cell(row=6+i, column=4).value = pdf_pages[i]
            ws.cell(row=6+i, column=5).value = pdf_size[i]
            ws.cell(row=6+i, column=6).value = "Blanco y negro"
        else:
            ws.insert_rows(6+i)
            #copy format ans style from row 6
            for j in range(1, 7):
                ws.cell(row=6+i, column=j).value = ws.cell(row=6, column=j).value
                ws.cell(row=6+i, column=j)._style = ws.cell(row=6, column=j)._style

            ws.cell(row=6+i, column=2).value = i+1
            ws.cell(row=6+i, column=3).value = pdf_names[i]
            ws.cell(row=6+i, column=4).value = pdf_pages[i]
            ws.cell(row=6+i, column=5).value = pdf_size[i]
            ws.cell(row=6+i, column=6).value = "Blanco y negro"

#Open a excel file
wb = openpyxl.load_workbook(r"D:\Proyectos Programación\Cotizaciones\Plantilla.xlsx")
ws = wb["Master"]

# insert data in excel
pdf_name, pdf_num_pages, pdf_size = getPDFInfo.get_pdf_info("D:\iB\Trabajos\Por cotizar")
inset_data_in_excel(ws, pdf_name, pdf_num_pages, pdf_size)

#Save the file
name = check_num_cotiz(r"D:\Proyectos Programación\Cotizaciones\Done")
wb.save(r"D:\Proyectos Programación\Cotizaciones\Done\\" + name + ".xlsx")