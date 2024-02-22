import os

import interface

# Path for CSV containing the data for the templating.
DATA_SOURCE = os.path.join("", "data", "data.csv")

# Setup for the template markers in the template file. Regex pattern.
TEMPLATE_START = "\[\["
TEMPLATE_END = "]]"

if __name__ == "__main__":
    app = interface.Gui()
    app.mainloop()
