# Docx templating

Create a .docx file based on an input .docx in which the the marked template fields are filled out 
according a given dataset, which represents contact persons of companies.

 ## Usage
In the default setup templates should be represented as [[template]] in the input file.
The template designators can be edited in the main.py file using the "TEMPLATE_START" nad
"TEMPLATE_END" variables.

The templates are created from a CSV file, which is in the Data/ folder named data.csv.
The data.csv should contain a header, which represents the template tags that should be replaced.
The data will be represented as a company having multiple contact persons, and each person
has the same set of additional info. The first two columns (company, contact/contact_person) are
mandatory the naming of these are set in the app so the value set in the CSV will be overwritten.
Any other column can be freely added and customized.

For example:  
company,contact_person,prefix,email,phone,address  
company1,John Doe the 1,Mr.,johndoecompany1.co,45678652345,34567 Big City  

The default setup is for default CSV "," being the separator, which can be changed in the main.py
by editing the "DATA_SEPARATOR" variable.

## Installation

Clone the hole repo, the dataset can be found in the data folder (data.csv) this should be edited.

To set up a virtual env calle .venv for the app:  
in the repo folder run:  
on linux:
```
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python3 main.py
```

on windows:
```
python -m venv .venv
.venv\Scripts-activate
pip install -r  requirements.txt
python main.py
```
