import os
import datetime
import tkinter as tk
from tkinter import filedialog

from data import TwoLevelDataset
from templating import Templating


IMG_PATH = os.path.join("", "data", "logo.png")


class Gui(tk.Tk):
    def __init__(self, data_source: str, template_start: str, template_end: str, *args, **kwargs):
        """Class for the GUI of a templating app.

        It creates a document based on a template input which is filled out
        with the data from the selected dataset."""
        super().__init__(*args, **kwargs)

        self.title("Word template fill")
        self.minsize(500, 1000)
        self.resizable(0, 0)  # Windows only

        self.template_start = template_start
        self.template_end = template_end
        self.template = {}

        # Set up canvas for the logo.
        self.canvas = tk.Canvas(self, width=100, height=100)
        self.logo = tk.PhotoImage(file=IMG_PATH)
        self.canvas.create_image(0, 0, image=self.logo, anchor="nw")
        self.canvas.pack(pady=50)

        self.main_frame = tk.Frame(self, width=50)
        self.main_frame.pack()

        # Set up frame for the input template select.
        self.input_frame = tk.Frame(
            self.main_frame,
            borderwidth=2,
            relief="groove",
        )
        self.input_file_button = tk.Button(
            self.input_frame, text="Select template", command=self._open_template
        )
        self.input_file_name = tk.StringVar()
        self.input_file_label = tk.Entry(
            self.input_frame, state="readonly", textvariable=self.input_file_name
        )

        self.input_frame.pack(fill="x")
        self.input_file_button.pack(anchor="w", padx=5, pady=5)
        self.input_file_label.pack(anchor="w", fill="x", padx=5, pady=5)

        # Set up frame for the output directory.
        self.output_frame = tk.Frame(
            self.main_frame,
            borderwidth=2,
            relief="groove",
        )
        self.output_folder_button = tk.Button(
            self.output_frame, text="Select output folder", command=self._open_output_folder
        )
        self.output_folder_name = tk.StringVar()
        self.output_folder_label = tk.Entry(
            self.output_frame, state="readonly", textvariable=self.output_folder_name
        )

        self.output_frame.pack(fill="x", pady=25)
        self.output_folder_button.pack(anchor="w", padx=5, pady=5)
        self.output_folder_label.pack(anchor="w", fill="x", padx=5, pady=5)

        # Setup a frame with label, entry and listfield for the companies.
        self.company_frame = tk.Frame(self.main_frame, borderwidth=2, relief="groove")
        self.company_filter_label = tk.Label(self.company_frame, text="Company filter:")
        self.company_filter_field = tk.Entry(self.company_frame, width=36)
        self.company_yscrollbar = tk.Scrollbar(self.company_frame, orient="vertical")
        self.company_list = tk.Listbox(
            self.company_frame, yscrollcommand=self.company_yscrollbar.set, width=50
        )
        self.company_list.bind("<<ListboxSelect>>", self._company_selected)

        self.company_frame.pack()
        self.company_filter_label.grid(row=0, column=0)
        self.company_filter_field.grid(row=0, column=1, columnspan=2, sticky="we")
        self.company_yscrollbar.grid(row=1, column=2, sticky="ns")
        self.company_list.grid(row=1, column=0, columnspan=2)

        # Setup a frame with label, entry and listfield for the contacts.
        self.contact_frame = tk.Frame(self.main_frame, borderwidth=2, relief="groove")
        self.contact_filter_label = tk.Label(self.contact_frame, text="Contact filter:")
        self.contact_filter_field = tk.Entry(self.contact_frame, width=36)
        self.contact_yscrollbar = tk.Scrollbar(self.contact_frame, orient="vertical")
        self.contact_list = tk.Listbox(
            self.contact_frame, yscrollcommand=self.contact_yscrollbar.set, width=50
        )
        self.contact_list.bind("<<ListboxSelect>>", self._contact_selected)

        self.contact_frame.pack(pady=25)
        self.contact_filter_label.grid(row=0, column=0)
        self.contact_filter_field.grid(row=0, column=1, columnspan=2, sticky="we")
        self.contact_yscrollbar.grid(row=1, column=2, sticky="ns")
        self.contact_list.grid(row=1, column=0, columnspan=2)

        # Setup a frame for the template values.
        self.template_frame = tk.Frame(self.main_frame, borderwidth=2, relief="groove")
        self.template_frame.pack()

        # Setup export frame.
        self.export_frame = tk.Frame(self.main_frame)
        self.generate_button = tk.Button(
            self.export_frame, text="Generate", command=self._generate_document, state="disabled"
        )
        self.notification = tk.StringVar()
        self.notification_label = tk.Label(self.export_frame, textvariable=self.notification)
        self.message = tk.StringVar()
        self.message_label = tk.Label(self.export_frame, textvariable=self.message, wraplength=400)

        self.export_frame.pack(pady=25)
        self.generate_button.pack(pady=5)
        self.notification_label.pack(pady=5)
        self.message_label.pack(pady=5)

        # Load the dataset.
        try:
            self.data = TwoLevelDataset(data_source)
        except FileNotFoundError:
            self.input_file_button.config(state="disabled")
            self.output_folder_button.config(state="disabled")
            self.notification.set("Error:")
            self.message.set(f"Data for templating not found, please check if the correct CSV exists: {data_source}")
        else:
            self.companies = self.data.get_primary_list()
            self.companies_filter = None
            self.selected_company = None
            self.contacts = []
            self.contacts_filter = None
            self.selected_contact = None
            self._on_tick()

    def _generate_template(self, labels: dict):
        """Create and display the templates loaded from the dataset
        that will be looked up in the input template."""
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

        if self.input_file_name.get() and self.output_folder_name.get():
            self.generate_button.config(state="normal")

    def _company_selected(self, event):
        """If a company is selected, fill the contact_list with the contacts
        belonging to the selected company."""
        selection = self.company_list.curselection()
        if selection:  # Check if a selection has been made
            self.selected_company = self.company_list.get(selection)
            self.contacts = self.data.get_secondary_list(self.selected_company)
            self.contacts_filter = None

    def _contact_selected(self, event):
        """If a contact is selected, create the template according the data in the dataset."""
        selection = self.contact_list.curselection()
        if selection:
            self.selected_contact = self.contact_list.get(selection)
            template = {
                "Company": self.selected_company,
                "Contact": self.selected_contact,
            }
            template.update(self.data.get_data()[self.selected_company][self.selected_contact])
            if "Date" not in template:
                template["Date"] = datetime.date.today().isoformat()
            self._generate_template(template)

    def _open_template(self):
        """Open the filebrowser to select the input template."""
        self.input_file_name.set(
            filedialog.askopenfilename(initialdir="", title="Select template file")
        )
        if self.output_folder_name.get() and self.template:
            self.generate_button.config(state="normal")

    def _open_output_folder(self):
        """Open the filebrowser to select the output folder."""
        self.output_folder_name.set(
            filedialog.askdirectory(initialdir="", title="Select output folder")
        )
        if self.input_file_name.get() and self.template:
            self.generate_button.config(state="normal")

    def _generate_document(self):
        """Generate the document with the templating."""
        try:
            document = Templating(
                self.get_input_file(),
                self.get_output_folder(),
                self.template_start,
                self.template_end,
                self.template_to_dict()
                )
        except FileNotFoundError as err:
            self.notification.set("Error:")
            self.message.set(err)
        document.sub_templates()
        self.notification.set("Templated document successfully created at:")
        self.message.set(f"{document.save()}")

    def _on_tick(self):
        """Periodically refreshes the elements of company_list and contact_list
        based on the filter fields."""
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

        self.after(250, self._on_tick)

    def get_input_file(self) -> str | None:
        """Returns the path for the input file."""
        if not os.path.exists(self.input_file_name.get()):
            raise FileNotFoundError("input file not found")
        return self.input_file_name.get()

    def get_output_folder(self) -> str | None:
        """Returns the path for the output folder."""
        if not os.path.exists(self.output_folder_name.get()):
            raise FileNotFoundError("output folder does not found")
        return self.output_folder_name.get()

    def template_to_dict(self) -> dict | None:
        """Returns a copy of the created templating dict."""
        if not self.template:
            return None
        return self.template.copy()
