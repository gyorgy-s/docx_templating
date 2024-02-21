import os
import re

import docx


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
        reg = re.compile(str(template))

        for element in self.input_elements:
            try:
                for r in element.runs:
                    if reg.search(r.text):
                        r.text = re.sub(reg, str(value), r.text)
            except AttributeError:
                for row in element.rows:
                    for cell in row.cells:
                        for cell_p in cell.paragraphs:
                            for cell_run in cell_p.runs:
                                if reg.search(cell_run.text):
                                    cell_run.text = re.sub(reg, str(value), r.text)

    def sub_templates(self):
        for template, value in self.templates.items():
            self.sub(template=template, value=value)

    def save(self):
        save_path = self.output_dir + os.path.join("/", "output.docx")
        print(save_path)
        self.document.save(save_path)
