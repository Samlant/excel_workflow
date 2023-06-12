# import subprocess
# import os
import fillpdf
from fillpdf import fillpdfs

# path = os.getcwd()
# path = os.path.join(path, "test2.py")

# data = path

# subprocess.run(["python", path], input=data, encoding="utf-8")


file_name = 'new_quoteform_EXAMPLE.pdf'
pdf_dict = fillpdfs.get_form_fields(file_name)
print(pdf_dict)