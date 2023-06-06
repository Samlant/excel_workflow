from tkinter import *
from tkinter import simpledialog


class Assign_Dialog(simpledialog.Dialog):
    def body(self, master):
        Label(master, text="First:").grid(row=0, column=0)
        Label(master, text="Second:").grid(row=1, column=0)

        self.e1 = Entry(master)
        self.e2 = Entry(master)

        self.e1.grid(row=0, column=1)
        self.e2.grid(row=1, column=1)
        return self.e1  # initial focus

    def apply(self):
        first = self.e1.get()
        second = self.e2.get()
        print(first, second)


class Action_Dialog(simpledialog.Dialog):
    def body(self, master):
        Label(
            master, text="Would you like to allocate markets or submit to them?"
        ).pack(fill=BOTH, expand=TRUE)

    def buttonbox(self):
        box = Frame(self)

        w = Button(
            box,
            text="Submit to Markets",
            width=30,
            command=self.ok,
            default=ACTIVE,
        )
        w.pack(side=LEFT, padx=5, pady=5)
        w = Button(
            box,
            text="Allocate Markets",
            width=30,
            command=self.allocate,
            default=ACTIVE,
        )
        w.pack(side=LEFT, padx=5, pady=5)
        w = Button(
            box,
            text="No",
            width=10,
            command=self.cancel,
        )
        w.pack(side=LEFT, padx=5, pady=5)

        self.bind("<Return>", self.ok)
        self.bind("<Escape>", self.cancel)

        box.pack()

    def allocate(self):
        print("Running allocate script")

    # def apply(self):
    #     print(w.result)


root = Tk()
d = Action_Dialog(root)
print(d.result)
