from tkinter import *
import subprocess
import os


root = Tk()

root.geometry("250x150")
root.title("How to Proceed?")

frame = Frame(root)
frame.pack(fill=BOTH, expand=True)


def create_folder():
    file_name = "tba"  # not implemented
    parent_dir = "tba"  # not decided
    directory_name = file_name
    path = os.path.join(parent_dir, directory_name)
    os.makedirs(path)


def allocate():
    pass


def run_quickdraw_app():
    subprocess.run(["QuickDraw.exe"])


def choice(option: str) -> None:
    if option == "pass":
        create_folder()
    elif option == "allocate":
        create_folder()
        allocate_market()
    else:
        create_folder()
        run_quickdraw_app()


submit_btn = Button(
    frame,
    text="Submit to Markets",
    width=30,
    command=lambda: choice("submit"),
    default=ACTIVE,
)
submit_btn.pack(side=LEFT, padx=5, pady=5)

allocate_btn = Button(
    frame,
    text="Allocate Markets",
    width=30,
    command=lambda: choice("allocate"),
)
allocate_btn.pack(fill=BOTH, expand=True)

pass_btn = Button(
    frame,
    text="Only create folder",
    width=30,
    command=lambda: choice("pass"),
)
pass_btn.pack(fill=BOTH, expand=True)
