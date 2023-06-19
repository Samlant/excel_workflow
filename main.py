import os
import sys
import time
import subprocess
import shutil
import string
from tkinter import *
from datetime import datetime

import xlwings as xw
import fillpdf
from fillpdf import fillpdfs


HOME_DIR = os.path.expanduser("~")
shared_drive = os.path.join(
    HOME_DIR,
    "Novamar Insurance",
    "Flordia Office Master - Documents",
)
PATH_TO_WATCH = shared_drive
QUOTES_FOLDER = os.path.join(
    shared_drive,
    "QUOTES New",
)
RENEWALS_FOLDER = os.path.join(
    shared_drive,
    "QUOTES Renewal",
)
TRACKER_PATH = os.path.join(
    shared_drive,
    "Trackers",
    "1MASTER 2023 QUOTE TRACKER.xlsx",
)
ICON_NAME = "icon.ico"
if getattr(sys, "frozen", False):
    application_path = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

ICON = os.path.join(application_path, "resources", ICON_NAME)
print("icon path is ", ICON)

# Below is for testing-purposes only when the above shared drive is unavailable.
# PATH_TO_WATCH = os.path.join(os.getcwd(), "tests")
# QUOTES_FOLDER = os.path.join(PATH_TO_WATCH, "QUOTES New")
# RENEWALS_FOLDER = os.path.join(PATH_TO_WATCH, "QUOTES Renewals")
# TRACKER_PATH = os.path.join(
#     PATH_TO_WATCH,
#     "Trackers",
#     "1MASTER 2023 QUOTE TRACKER.xlsx",
# )


class DirWatch:
    def __init__(self):
        self._begin_watch()

    def _begin_watch(self) -> None:
        before = dict([(f, None) for f in os.listdir(PATH_TO_WATCH)])
        while 1:
            time.sleep(2)
            after = dict([(f, None) for f in os.listdir(PATH_TO_WATCH)])
            added = [f for f in after if not f in before]
            if added:
                new_file = added[0]
                if os.path.splitext(new_file)[1] == (".pdf" or ".docx"):
                    dialog = DialogNewFile(new_file)
                    dialog.root.mainloop()
            before = dict([(f, None) for f in os.listdir(PATH_TO_WATCH)])


class DialogNewFile:
    def __init__(self, file_name):
        self.excel_entry = {
            "vessel_year": "",
            "vessel": "",
            "markets": {},
            "status": "ALLOCATE AND SUBMIT TO MRKTS",
            "referral": "",
        }
        self.file_name = file_name
        self._initialize()

    def _initialize(self):
        self.root = Tk()
        self.root.geometry("300x400")
        self.root.title("Next Steps")
        self.root.iconbitmap(ICON)
        self.root.text_frame = Frame(self.root, bg="#CFEBDF")
        self.root.text_frame.pack(fill=BOTH, expand=True)
        self.root.btn_frame = Frame(self.root, bg="#CFEBDF")
        self.root.btn_frame.pack(fill=BOTH, expand=True, ipady=2)
        self.submitted_quotes = False
        self._save_client_name()
        if os.path.splitext(self.file_name)[1] == ".pdf":
            self.get_PDF_values()
        self._create_widgets()

    def _save_client_name(self) -> None:
        client_name = os.path.splitext(self.file_name)[0].split(" ")
        if len(client_name) >= 2:
            self.excel_entry["fname"] = string.capwords(client_name[1])
            self.excel_entry["lname"] = client_name[0].upper()

    def get_PDF_values(self):
        keys_dict = {
            "fname": "fname",
            "lname": "lname",
            "year": "vessel_year",
            "vessel": "vessel_make_model",
            "referral": "referral",
        }
        file_path_name = os.path.join(PATH_TO_WATCH, self.file_name)
        pdf_dict = fillpdfs.get_form_fields(file_path_name)
        pdf_dict = {key: pdf_dict[key] for key in pdf_dict.keys() & keys_dict.values()}

        self.excel_entry = {}
        fname = pdf_dict.get(keys_dict["fname"])
        self.excel_entry["fname"] = string.capwords(fname)
        lname = pdf_dict.get(keys_dict["lname"])
        self.excel_entry["lname"] = lname.upper()
        vessel = pdf_dict.get(keys_dict["vessel"])
        self.excel_entry["vessel"] = string.capwords(vessel)
        self.excel_entry["vessel_year"] = pdf_dict.get(keys_dict["year"])
        referral = pdf_dict.get(keys_dict["referral"])
        self.excel_entry["referral"] = referral.upper()
        self.excel_entry["status"] = "ALLOCATE AND SUBMIT TO MRKTS"
        # if any(chr.isdigit() for chr in self.excel_entry["vessel"]):
        #     self.excel_entry["length"] = pdf_dict.get(keys_dict["length"])

    def _create_widgets(self):
        client_name = " ".join([self.excel_entry["fname"], self.excel_entry["lname"]])
        vessel = self.excel_entry["vessel"]
        year = self.excel_entry["vessel_year"]
        referral = self.excel_entry["referral"]
        self.root.text_frame.grid_columnconfigure(0, weight=1)
        self.root.btn_frame.grid_columnconfigure(0, weight=1)

        Label(self.root.text_frame, text="Client name: ", bg="#CFEBDF").grid(
            column=0, row=0, pady=(3, 0)
        )
        name_entry = Entry(
            self.root.text_frame, width=30, justify="center", bg="#5F634F", fg="#FFCAB1"
        )
        name_entry.insert(0, client_name)
        name_entry.grid(column=0, row=1, pady=(0, 8))

        Label(self.root.text_frame, text="Vessel: ", bg="#CFEBDF").grid(column=0, row=2)
        vessel_entry = Entry(
            self.root.text_frame, width=30, justify="center", bg="#5F634F", fg="#FFCAB1"
        )
        vessel_entry.insert(0, vessel)
        vessel_entry.grid(column=0, row=3, pady=(0, 8))

        Label(self.root.text_frame, text="Vessel year: ", bg="#CFEBDF").grid(
            column=0, row=4
        )
        year_entry = Entry(
            self.root.text_frame, width=10, justify="center", bg="#5F634F", fg="#FFCAB1"
        )
        year_entry.insert(0, year)
        year_entry.grid(column=0, row=5, pady=(0, 8))

        Label(self.root.text_frame, text="Referral: ", bg="#CFEBDF").grid(
            column=0, row=6
        )
        referral_entry = Entry(
            self.root.text_frame,
            width=30,
            justify="center",
            bg="#5F634F",
            fg="#FFCAB1",
        )
        referral_entry.insert(0, referral)
        referral_entry.grid(column=0, row=7, pady=(0, 7))

        submit_btn = Button(
            self.root.btn_frame,
            text="Submit to Markets",
            width=36,
            height=3,
            command=lambda: self.choice("submit"),
            default=ACTIVE,
            bg="#1D3461",
            fg="#CFEBDF",
        )
        submit_btn.grid(row=0, column=0, padx=5, pady=(0, 0))

        allocate_btn = Button(
            self.root.btn_frame,
            text="Allocate Markets",
            width=36,
            height=3,
            command=lambda: self.choice("allocate"),
            default=ACTIVE,
            bg="#1D3461",
            fg="#CFEBDF",
        )
        allocate_btn.grid(row=1, column=0, padx=5, pady=(3, 3))

        create_folder_only_btn = Button(
            self.root.btn_frame,
            text="Only create folder",
            width=36,
            height=3,
            command=lambda: self.choice("only create folder"),
            default=ACTIVE,
            bg="#1D3461",
            fg="#CFEBDF",
        )
        create_folder_only_btn.grid(row=2, column=0, padx=5, pady=(0, 5))

    def choice(self, option: str) -> None:
        if option == "only create folder":
            self._create_folder()
            self.root.destroy()
            self._create_excel_entry()

        elif option == "allocate":
            self._create_folder()
            self.root.destroy()
            self.allocate_markets()

        else:
            self._create_folder()
            self.root.destroy()
            if self.run_quickdraw_app():
                self._create_excel_entry()

    def _create_folder(self):
        file_name_list = os.path.splitext(self.file_name)
        if file_name_list[1] == ".pdf":
            self.dir_name = self.excel_entry["lname"] + " " + self.excel_entry["fname"]
        else:
            self.dir_name = file_name_list[0]
        if self.excel_entry["referral"] == "RENEWAL":
            path = os.path.join(RENEWALS_FOLDER, self.dir_name)
        else:
            path = os.path.join(QUOTES_FOLDER, self.dir_name)
        self.path = os.path.join(PATH_TO_WATCH, self.file_name)
        os.makedirs(path, exist_ok=True)
        shutil.move(self.path, path)
        self.path = path

    def _create_excel_entry(self):
        excel = ExcelWorker(self.excel_entry)
        if self.submitted_quotes == True:
            excel.change_markets_to_pending()
        excel.create_row()
        excel.save_workbook()

    def allocate_markets(self) -> dict:
        dialog_allocate = DialogAllocateMarkets(self.excel_entry)

    def run_quickdraw_app(self):
        self.excel_entry["status"] = "Pending with Underwriting"
        path = os.path.join(HOME_DIR, "AppData", "work_tools", "QuickDraw.exe")
        subprocess.run([path], input=self.path, encoding="utf-8")
        self.submitted_quotes = True


class DialogAllocateMarkets:
    def __init__(self, excel_entry: dict):
        self.excel_entry = excel_entry
        self._initialize()

    def _initialize(self):
        self.root = Tk()
        self.root.geometry("260x560")
        self.root.title("Allocate Markets")
        self.root.iconbitmap(ICON)
        self.root.frame = Frame(self.root, bg="#CFEBDF")
        self.root.frame.pack(fill=BOTH, expand=False)
        self._create_widgets()

    def _create_widgets(self):
        Label(
            self.root.frame,
            text="ALLOCATE MARKETS",
            justify="center",
            bg="#CFEBDF",
            fg="#5F634F",
        ).pack(fill=X, ipady=6)
        self.ch_checkbtn = IntVar(self.root.frame)
        self._create_button("Chubb", self.ch_checkbtn)
        self.mk_checkbtn = IntVar(self.root.frame)
        self._create_button("Markel", self.mk_checkbtn)
        self.ai_checkbtn = IntVar(self.root.frame)
        self._create_button("American Integrity", self.ai_checkbtn)
        self.am_checkbtn = IntVar(self.root.frame)
        self._create_button("American Modern", self.am_checkbtn)
        self.pg_checkbtn = IntVar(self.root.frame)
        self._create_button("Progressive", self.pg_checkbtn)
        self.sw_checkbtn = IntVar(self.root.frame)
        self._create_button("Seawave", self.sw_checkbtn)
        self.km_checkbtn = IntVar(self.root.frame)
        self._create_button("Kemah Marine", self.km_checkbtn)
        self.cp_checkbtn = IntVar(self.root.frame)
        self._create_button("Concept Special Risks", self.cp_checkbtn)
        self.nh_checkbtn = IntVar(self.root.frame)
        self._create_button("New Hampshire", self.nh_checkbtn)
        self.In_checkbtn = IntVar(self.root.frame)
        self._create_button("Intact", self.In_checkbtn)
        self.tv_checkbtn = IntVar(self.root.frame)
        self._create_button("Travelers", self.tv_checkbtn)

        allocate_btn = Button(
            master=self.root.frame,
            text="ALLOCATE",
            width=30,
            height=10,
            bg="#1D3461",
            fg="#CFEBDF",
            command=lambda: self._process_market_choices(),
        )
        allocate_btn.pack(
            fill=X,
            expand=False,
            pady=5,
            padx=10,
            ipady=6,
            ipadx=10,
        )

    def _create_button(self, text: str, int_variable: IntVar):
        x = Checkbutton(
            self.root.frame,
            text=text,
            variable=int_variable,
            relief="raised",
            justify=CENTER,
            anchor=W,
            fg="#FFCAB1",
            bg="#5F634F",
            selectcolor="#000000",
        )
        x.pack(fill=X, expand=False, ipady=6, ipadx=10, pady=3, padx=10, anchor=NW)

    def _process_market_choices(self):
        self.excel_entry["markets"] = self._return_markets()
        self.excel_entry["status"] = "SUBMIT TO MRKTS"
        self.root.destroy()
        self._create_excel_entry()

    def _return_markets(self):
        dict_of_markets = {
            "ch": self.ch_checkbtn.get(),
            "mk": self.mk_checkbtn.get(),
            "ai": self.ai_checkbtn.get(),
            "am": self.am_checkbtn.get(),
            "pg": self.pg_checkbtn.get(),
            "sw": self.sw_checkbtn.get(),
            "km": self.km_checkbtn.get(),
            "cp": self.cp_checkbtn.get(),
            "nh": self.nh_checkbtn.get(),
            "In": self.In_checkbtn.get(),
            "tv": self.tv_checkbtn.get(),
        }
        return dict_of_markets

    def _create_excel_entry(self):
        excel = ExcelWorker(self.excel_entry)
        excel.create_row()
        excel.save_workbook()


class ExcelWorker:
    def __init__(
        self,
        excel_entry: dict,
    ):
        fname = excel_entry["fname"]
        lname = excel_entry["lname"]
        self.name = " ".join([lname, fname])
        self.date = self._get_current_date()
        self.vessel_year = excel_entry["vessel_year"]
        self.vessel = excel_entry["vessel"]
        self.markets = excel_entry["markets"]
        self.status = excel_entry["status"]
        self.referral = excel_entry["referral"]
        self.app = xw.App(visible=False)
        self.wb = xw.Book(TRACKER_PATH)
        month = self._get_current_month()
        self.ws = self.wb.sheets(month)
        markets_list = self._assign_markets()
        self.markets_list = self._list_to_str(markets_list)

    def _get_current_date(self) -> str:
        current_date = datetime.now()
        current_date = f"{current_date.month}-{current_date.day}"
        return current_date

    def _get_current_month(self):
        months_of_the_year = {
            1: "January",
            2: "February",
            3: "March",
            4: "April",
            5: "May",
            6: "June",
            7: "July",
            8: "August",
            9: "September",
            10: "October",
            11: "November",
            12: "December",
        }
        month = datetime.now().month
        return months_of_the_year.get(month).upper()

    def _assign_markets(self) -> list:
        list_of_markets = []
        for market, value in self.markets.items():
            if value == 1:
                mrkt = market.upper()
                list_of_markets.append(mrkt)
                self.markets[market] = ""
            else:
                self.markets[market] = ""
        return list_of_markets

    def _list_to_str(self, list_data: list) -> str:
        return ", ".join(list_data)

    def create_row(self) -> bool:
        self.ws.range("A2:Y2").insert("down")
        self.ws["D2"].value = self.name
        self.ws["E2"].value = self.date
        self.ws["G2"].value = self.vessel_year
        self.ws["H2"].value = self.vessel
        self.ws["X2"].value = self.status
        self.ws["Y2"].value = self.referral
        self._assign_markets_to_sheet()
        self.ws.range("A2:Y2").api.Borders.Weight = 1

    def _assign_markets_to_sheet(self):
        self.ws["I2"].value = self.markets_list
        self.ws["J2"].value = self.markets["ch"]
        self.ws["K2"].value = self.markets["mk"]
        self.ws["L2"].value = self.markets["ai"]
        self.ws["M2"].value = self.markets["am"]
        self.ws["N2"].value = self.markets["pg"]
        self.ws["O2"].value = self.markets["sw"]
        self.ws["P2"].value = self.markets["km"]
        self.ws["Q2"].value = self.markets["cp"]
        self.ws["R2"].value = self.markets["nh"]
        self.ws["S2"].value = self.markets["In"]
        self.ws["T2"].value = self.markets["tv"]

    def save_workbook(self):
        self.wb.save(TRACKER_PATH)
        self.app.quit()

    def change_markets_to_pending(self):
        """Not currently used;  implement once QuickDraw is working."""
        for x, y in self.markets.items():
            if y == 1:
                self.markets_list.append(x)
                y = "p"
            else:
                y = ""


app = DirWatch()
