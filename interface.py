import os
import tkinter as tk
from data import Contacts


IMG_PATH = os.path.join("", "data", "logo.png")


class Gui(tk.Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.title("Word template fill")
        self.minsize(500, 1000)
        self.resizable(0, 0)  # Windows only

        self.data = Contacts()

        self.template = {}

        self.companies = self.data.get_companies()
        self.companies_filter = None
        self.selected_company = None
        self.contacts = []
        self.contacts_filter = None
        self.selected_contact = None

        # Set up canvas for the logo.
        self.canvas = tk.Canvas(self, width=100, height=100)
        self.logo = tk.PhotoImage(file=IMG_PATH)
        self.canvas.create_image(0, 0, image=self.logo, anchor="nw")
        self.canvas.pack(pady=50)

        # Setup a frame with label, entry and listfield for the companies.
        self.company_frame = tk.Frame(self, borderwidth=2, relief="groove")
        self.company_filter_label = tk.Label(self.company_frame, text="Company filter:")
        self.company_filter_field = tk.Entry(self.company_frame, width=36)
        self.company_yscrollbar = tk.Scrollbar(self.company_frame, orient="vertical")
        self.company_list = tk.Listbox(
            self.company_frame, yscrollcommand=self.company_yscrollbar.set, width=50
        )
        self.company_list.bind("<<ListboxSelect>>", self.company_selected)

        self.company_frame.pack(pady=50)
        self.company_filter_label.grid(row=0, column=0)
        self.company_filter_field.grid(row=0, column=1, columnspan=2, sticky="we")
        self.company_yscrollbar.grid(row=1, column=2, sticky="ns")
        self.company_list.grid(row=1, column=0, columnspan=2)

        # Setup a frame with label, entry and listfield for the contacts.
        self.contact_frame = tk.Frame(self, borderwidth=2, relief="groove")
        self.contact_filter_label = tk.Label(self.contact_frame, text="Contact filter:")
        self.contact_filter_field = tk.Entry(self.contact_frame, width=36)
        self.contact_yscrollbar = tk.Scrollbar(self.contact_frame, orient="vertical")
        self.contact_list = tk.Listbox(
            self.contact_frame, yscrollcommand=self.contact_yscrollbar.set, width=50
        )
        self.contact_list.bind("<<ListboxSelect>>", self.contact_selected)

        self.contact_frame.pack(pady=50)
        self.contact_filter_label.grid(row=0, column=0)
        self.contact_filter_field.grid(row=0, column=1, columnspan=2, sticky="we")
        self.contact_yscrollbar.grid(row=1, column=2, sticky="ns")
        self.contact_list.grid(row=1, column=0, columnspan=2)

        # Setup a frame for the template values.
        self.template_frame = tk.Frame(self, borderwidth=2, relief="groove")
        self.template_frame.pack(pady=50)

        self.on_tick()

    def generate_template(self, labels: dict):
        for i, t in enumerate(labels.items()):
            tk.Label(
                self.template_frame,
                text=str(t[0]) + ": ",
                width=25,
                justify="left",
                anchor="w",
            ).grid(row=i, column=0)
            tk.Label(
                self.template_frame,
                text=str(t[1]),
                width=25,
                justify="right",
                anchor="e",
            ).grid(row=i, column=1)
            self.template.update({t[0].lower(): t[1]})

    def company_selected(self, event):
        selection = self.company_list.curselection()
        if selection:  # Check if a selection has been made
            self.selected_company = self.company_list.get(selection)
            self.contacts = self.data.get_contact_list(self.selected_company)
            self.contacts_filter = None

    def contact_selected(self, event):
        selection = self.contact_list.curselection()
        if selection:
            self.selected_contact = self.contact_list.get(selection)
            template = {
                "Company": self.selected_company,
                "Contact": self.selected_contact,
            }
            template.update(self.data.get_data()[self.selected_company][self.selected_contact])
            self.generate_template(template)

    def on_tick(self):
        if self.company_filter_field.get() != self.companies_filter:
            self.companies_filter = self.company_filter_field.get().lower()
            self.company_list.delete(0, "end")
            for company in self.companies:
                if self.companies_filter in company.lower():
                    self.company_list.insert("end", company)

        if self.contact_filter_field.get() != self.contacts_filter:
            self.contacts_filter = self.contact_filter_field.get().lower()
            self.contact_list.delete(0, "end")
            for contact in self.contacts:
                if self.contacts_filter in contact.lower():
                    self.contact_list.insert("end", contact)

        self.after(250, self.on_tick)


app = Gui()

app.mainloop()
