import os
import re

import docx


TEMPLATE_START = "\[\["
TEMPLATE_END = "]]"


class Templating:
    def __init__(self, input_file=None, output_dir=None, templates: dict = {}) -> None:
        self.document = docx.Document(input_file)
        self.output_dir = output_dir
        self.templates = templates.copy()
        self.input_elements = self.get_input_elements()

    def get_input_elements(self):
        section_components = []
        section_elements = []
        for section in self.document.sections:
            section_components += [
                section.even_page_footer,
                section.even_page_header,
                section.first_page_footer,
                section.first_page_header,
                section.footer,
                section.header,
            ]
            for component in section_components:
                section_elements += component.paragraphs
                section_elements += component.tables

            section_elements += section.iter_inner_content()
        return section_elements

    def sub(self, template: str, value: str):
        reg = re.compile(TEMPLATE_START + template + TEMPLATE_END)

        for element in self.input_elements:
            try:
                if reg.search(element.text):
                    for i, run in enumerate(element.runs):
                        if reg.search(run.text):
                            run.text = reg.sub(value, run.text)
                        elif re.search(TEMPLATE_START + template + "$", run.text):
                            if re.search("^" + TEMPLATE_END, element.runs[i + 1].text):
                                run.text = re.sub(TEMPLATE_START + template + "$", value, run.text)
                                element.runs[i + 1].text = re.sub(
                                    "^" + TEMPLATE_END, "", element.runs[i + 1].text
                                )
                        elif re.search(TEMPLATE_START + "$", run.text):
                            if re.search("^" + template + TEMPLATE_END, element.runs[i + 1].text):
                                run.text = re.sub(TEMPLATE_START + "$", "", run.text)
                                element.runs[i + 1].text = re.sub(
                                    "^" + template + TEMPLATE_END, value, element.runs[i + 1].text
                                )
                            elif element.runs[i + 1].text == template and re.search(
                                "^" + TEMPLATE_END, element.runs[i + 2].text
                            ):
                                run.text = re.sub(TEMPLATE_START + "$", "", run.text)
                                element.runs[i + 1].text = re.sub(
                                    template, value, element.runs[i + 1].text
                                )
                                element.runs[i + 2].text = re.sub(
                                    "^" + TEMPLATE_END, "", element.runs[i + 2].text
                                )

            except AttributeError:
                for row in element.rows:
                    for cell in row.cells:
                        for cell_p in cell.paragraphs:
                            for i, cell_run in enumerate(cell_p.runs):
                                if reg.search(cell_run.text):
                                    cell_run.text = reg.sub(value, cell_run.text)
                                elif re.search(TEMPLATE_START + template + "$", cell_run.text):
                                    if re.search("^" + TEMPLATE_END, element.runs[i + 1].text):
                                        cell_run.text = re.sub(
                                            TEMPLATE_START + template + "$", value, cell_run.text
                                        )
                                        element.runs[i + 1].text = re.sub(
                                            "^" + TEMPLATE_END, "", element.runs[i + 1].text
                                        )
                                elif re.search(TEMPLATE_START + "$", cell_run.text):
                                    if re.search(
                                        "^" + template + TEMPLATE_END, element.runs[i + 1].text
                                    ):
                                        cell_run.text = re.sub(
                                            TEMPLATE_START + "$", "", cell_run.text
                                        )
                                        element.runs[i + 1].text = re.sub(
                                            "^" + template + TEMPLATE_END,
                                            value,
                                            element.runs[i + 1].text,
                                        )
                                    elif element.runs[i + 1].text == template and re.search(
                                        "^" + TEMPLATE_END, element.runs[i + 2].text
                                    ):
                                        cell_run.text = re.sub(
                                            TEMPLATE_START + "$", "", cell_run.text
                                        )
                                        element.runs[i + 1].text = re.sub(
                                            template, value, element.runs[i + 1].text
                                        )
                                        element.runs[i + 2].text = re.sub(
                                            "^" + TEMPLATE_END, "", element.runs[i + 2].text
                                        )

    def sub_templates(self):
        for template, value in self.templates.items():
            self.sub(template=template, value=value)

    def save(self):
        save_path = os.path.join("", self.output_dir, "output.docx")
        print(save_path)
        self.document.save(save_path)
        return save_path
