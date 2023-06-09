import os
import time
import subprocess
import shutil
from tkinter import *
from datetime import datetime

from openpyxl import Workbook, load_workbook


class DirWatch:
    def __init__(self):
        self.begin_watch()

    def begin_watch(self) -> None:
        path_to_watch = "."
        before = dict([(f, None) for f in os.listdir(path_to_watch)])
        while 1:
            time.sleep(10)
            after = dict([(f, None) for f in os.listdir(path_to_watch)])
            added = [f for f in after if not f in before]
            if added:
                new_file = added[0]
                dialog = dialog_new_file(new_file)
                dialog.root.mainloop()
                before = dict([(f, None) for f in os.listdir(path_to_watch)])


class dialog_allocate_markets:
    def __init__(self):
        class dialog_new_file:
        self.initialize()

    def initialize(self):
        self.root = Tk()
        self.root.geometry("250x190")
        self.root.title("How to Proceed?")
        self.root.frame = Frame(self.root)
        self.root.frame.pack(fill=BOTH, expand=False)
    
    


class dialog_new_file:
    def __init__(self, file_name):
        self.initialize()
        self.file_name = file_name

    def initialize(self):
        self.root = Tk()
        self.root.geometry("250x190")
        self.root.title("How to Proceed?")
        self.root.frame = Frame(self.root)
        self.root.frame.pack(fill=BOTH, expand=False)
        self.create_buttons()

    def _create_excel_entry(self):
        self.excel = Excel_Worker(
                "/Report/Master_Tracker_2023.xlsx",
                self.dir_name,
            )
        
    def create_folder(self):
        parent_dir = os.path.dirname(__file__)
        print(parent_dir)
        # file_path = "".join("/", self.file_name)
        print(self.file_name)
        self.dir_name = os.path.splitext(self.file_name)[0]
        # dir_name = dir_name.split()
        self.path = os.path.join(parent_dir, self.dir_name)
        os.makedirs(self.path)
        self._move_quoteform_to_folder()
        self._create_excel_entry()

    def _move_quoteform_to_folder(self):
        shutil.move(self.file_name, self.path)

    def allocate_markets(self):
        dialog_allocate = Dialog_Allocate_Markets()

    def run_quickdraw_app(self):
        subprocess.run(["QuickDraw.exe"])

    def choice(self, option: str) -> None:
        if option == "create_folder":
            self.create_folder()
            self.root.destroy()
            
        elif option == "allocate":
            self.create_folder()
            self.root.destroy()
            self.allocate_markets()
        else:
            self.create_folder()
            self.root.destroy()
            self.run_quickdraw_app()

    def create_buttons(self):
        submit_btn = Button(
            self.root.frame,
            text="Submit to Markets",
            width=30,
            height=3,
            command=lambda: self.choice("submit"),
            default=ACTIVE,
            bg="green",
        )
        submit_btn.pack(side=TOP, fill=NONE, padx=5, pady=5)

        allocate_btn = Button(
            self.root.frame,
            text="Allocate Markets",
            width=30,
            height=3,
            command=lambda: self.choice("allocate"),
            default=ACTIVE,
            bg="yellow",
        )
        allocate_btn.pack(side=TOP, fill=NONE, expand=True, padx=5, pady=5)

        pass_btn = Button(
            self.root.frame,
            text="Only create folder",
            width=30,
            height=3,
            command=lambda: self.choice("create_folder"),
            default=ACTIVE,
            bg="orange",
        )
        pass_btn.pack(side=TOP, fill=NONE, expand=True, padx=5, pady=5)


class Excel_Worker:
    def __init__(
        self,
        workbook_name_or_path,
        name,
        vessel_year,
        vessel,
        markets,
        ch,
        mk,
        ai,
        am,
        pg,
        sw,
        km,
        nh,
        cp,
        yi,
        In,
        tv,
        status,
        referral,
    ):
        self.name = name
        self.date = str(datetime.today()).split()[0]
        self.vessel_year = vessel_year
        self.vessel = vessel
        self.markets = markets
        self.ch = ch
        self.mk = mk
        self.ai = ai
        self.am = am
        self.pg = pg
        self.sw = sw
        self.km = km
        self.nh = nh
        self.cp = cp
        self.yi = yi
        self.In = In
        self.tv = tv
        self.status = status
        self.referral = referral
        self.wb = load_workbook(workbook_name_or_path)
        month = self.get_current_month()
        self.wb.active = self.wb[month]
        self._create_entry()

    def get_current_month(self):
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

    def _create_entry(self):
        list_of_client_data = [
            self.name,
            self.date,
            self.vessel_year,
            self.vessel,
            self.markets,
            self.ch,
            self.mk,
            self.ai,
            self.am,
            self.pg,
            self.sw,
            self.km,
            self.nh,
            self.cp,
            self.yi,
            self.In,
            self.tv,
            self.status,
            self.referral,
        ]
        self.ws.append(list_of_client_data)
        self._save_workbook()
        
    def _save_workbook(self):
        self.wb.save("test.xlsx")

app = DirWatch()
