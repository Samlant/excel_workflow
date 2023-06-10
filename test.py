import os, time
import fillpdf
from fillpdf import fillpdfs
import string

# PATH_TO_WATCH = '.''


# class DirWatch:
#     def __init__(self):
#         self._begin_watch()

#     def _begin_watch(self) -> None:
#         before = dict([(f, None) for f in PATH_TO_WATCH])
#         while 1:
#             time.sleep(10)
#             after = dict([(f, None) for f in PATH_TO_WATCH])
#             added = [f for f in after if not f in before]
#             if added:
#                 print("added new file: ", added[0])


# app = DirWatch()


def get_PDF_values():
    keys_dict = {
        "fname": "fname",
        "lname": "lname",
        "year": "vessel_year",
        "vessel": "vessel_make_model",
        "referral": "referral",
    }
    pdf_dict = fillpdfs.get_form_fields("Lanteigne Samuel.pdf")
    pdf_dict = {key: pdf_dict[key] for key in pdf_dict.keys() & keys_dict.values()}

    excel_entry = {}
    fname = pdf_dict.get(keys_dict["fname"])
    excel_entry["fname"] = string.capwords(fname)
    lname = pdf_dict.get(keys_dict["lname"])
    excel_entry["lname"] = lname.upper()
    vessel = pdf_dict.get(keys_dict["vessel"])
    excel_entry["vessel"] = string.capwords(vessel)
    excel_entry["vessel_year"] = pdf_dict.get(keys_dict["year"])
    referral = pdf_dict.get(keys_dict["referral"])
    excel_entry["referral"] = referral.upper()


get_PDF_values()
