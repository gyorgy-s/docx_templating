import os
import pandas as pd

DATA_SOURCE = os.path.join("", "data", "data.csv")


class Contacts:

    def __init__(self) -> None:
        self._header = (
            pd.read_csv(DATA_SOURCE, index_col=None, header=None, nrows=1).iloc[0].to_list()
        )
        self._df = (
            pd.read_csv(DATA_SOURCE, index_col=None, header=0)
            .sort_values(by=["company"], ascending=True)
            .fillna("")
        )
        self._data = {}
        self.create_dict()

    def create_dict(self):
        for r in range(len(self._df)):
            row = self._df.iloc[r]

            company = row.iloc[0].title()
            contact_person = row.iloc[1].title()

            if company not in self._data:
                self._data[company] = {}
            if contact_person not in self._data[company]:
                self._data[company][contact_person] = {}
                for i, col in enumerate(self._header[2:]):
                    self._data[company][contact_person][col.title()] = row.iloc[i + 2]

    def get_data(self) -> dict:
        return self._data

    def get_companies(self) -> list:
        return list(self._data.keys())

    def get_contacts(self, company: str) -> dict | None:
        return self._data.get(company)

    def get_contact_list(self, company) -> list:
        return list(self.get_contacts(company).keys())
