# import subprocess
import os
import fillpdf
from fillpdf import fillpdfs

# path = os.getcwd()
# path = os.path.join(path, "test2.py")

# data = path

# subprocess.run(["python", path], input=data, encoding="utf-8")

# home = os.path.expanduser( '~' )
# file_name = os.path.join(home, 'Novamar Insurance', 'Flordia Office Master - Documents', 'QUOTES New', '111111 QuoteForm.pdf')
keys_dict = {
            "fname": "fname",
            "lname": "lname",
            "year": "vessel_year",
            "vessel": "vessel_make_model",
            "referral": "referral",
        }

pdf_dict1 = fillpdfs.print_form_fields('test.pdf')
# print(pdf_dict)
# pdf_dict = {key: pdf_dict[key] for key in pdf_dict.keys() & keys_dict.values()}
# print(pdf_dict)