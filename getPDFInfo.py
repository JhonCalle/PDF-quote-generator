import os
import pikepdf

def get_pdf_info(directory):
    pdf_files = []
    pdf_names = []

    # Get all pdf files names in directory
    for file_name in os.listdir(directory):
        if file_name.endswith(".pdf"):
            pdf_files.append(os.path.join(directory, file_name))
            pdf_names.append(file_name)

    pdf_pages = []
    pdf_size_cm = []

    # Get all pdf page numbers
    for pdf_file in pdf_files:
        with pikepdf.open(pdf_file) as pdf:

            #Get page numbers
            pdf_pages.append(len(pdf.pages))

            aux = round(len(pdf.pages) / 2)
            page = pdf.pages[aux]

            # Get the size of the page in points
            width, height = page.MediaBox[2], page.MediaBox[3]

            # Convert the size from points to centimeters
            cm_per_point = 0.0352778
            width_cm = float(width) * cm_per_point
            height_cm = float(height) * cm_per_point
            pdf_size_cm.append((width_cm, height_cm))
            pdf_size = decide_pdf_size(pdf_size_cm)

    return pdf_names, pdf_pages, pdf_size

def decide_pdf_size(pdf_size):
    pdf_size_cm = []
    for size in pdf_size:
        if size[0] < 16.5*1.1 and size[1] < 21.5*1.1:
            pdf_size_cm.append("Medio oficio")
        # if size[0] > 21.5*0.9 or size[1] > 28*0.9:
        #     pdf_size_cm.append("Carta")
        else:
            pdf_size_cm.append("Carta")
    return pdf_size_cm
