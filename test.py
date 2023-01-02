import weasyprint

def html_to_pdf(html_file, css_file, pdf_file):
  # Read the HTML and CSS files
  with open(html_file, "r") as f:
    html = f.read()
  with open(css_file, "r") as f:
    css = f.read()

  # Convert the HTML and CSS to a PDF
  pdf = weasyprint.HTML(string=html).write_pdf(stylesheets=[weasyprint.CSS(string=css)])

  # Save the PDF to the specified file
  with open(pdf_file, "wb") as f:
    f.write(pdf)

html_to_pdf("index.html", "style.css", "test.pdf")