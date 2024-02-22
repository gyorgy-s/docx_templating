import pandas as pd


class TwoLevelDataset:
    def __init__(self, data_source: str, separator: str = ",") -> None:
        """Class to represent a two-level dataset where a primary group
        contains multiple secondary elements, that have a specific set of data.
        For example a series of companies having multiple employees.

        DATA_SOURCE requires csv data that contains a header
        with data in the following order:
        primary, secondary, other_1, ..., other_n

        Resulting format:
        {
            "primary": {
                "secondary_1": {
                    "other_1": data,
                    ...
                    "other_n": data,
                },
                ...
                "secondary_n": {
                    "other_1": data,
                    ...
                    "other_n": data,
                },
            },
            ...
        }"""
        self._header = (
            pd.read_csv(data_source, index_col=None, header=None, nrows=1, sep=separator).iloc[0].to_list()
        )
        self._df = (
            pd.read_csv(data_source, index_col=None, header=0, sep=separator)
            .sort_values(by=[self._header[0]], ascending=True)
            .fillna("")
        )
        self._data = {}
        self._create_dict()

    def _create_dict(self):
        """Fill the _data attribute with the data from the dataset."""
        for r in range(len(self._df)):
            row = self._df.iloc[r]

            primary = row.iloc[0].title()
            secondary = row.iloc[1].title()

            if primary not in self._data:
                self._data[primary] = {}
            if secondary not in self._data[primary]:
                self._data[primary][secondary] = {}
                for i, col in enumerate(self._header[2:]):
                    self._data[primary][secondary][col.title()] = row.iloc[i + 2]

    def get_data(self) -> dict:
        """Returns the whole DataFrame."""
        return self._data

    def get_primary_list(self) -> list:
        """Returns the first level keys of the dataset in a list format."""
        return list(self._data.keys())

    def get_secondary(self, primary: str) -> dict | None:
        """Returns the second level of the dataset belonging to PRIMARY
        in a dict format.

        PRYMARY being the first level key of the dataset."""
        return self._data.get(primary)

    def get_secondary_list(self, primary: str) -> list:
        """Returns the second level keys of the dataset belonging to PRIMARY
        int a list format.

        PRYMARY being the first level key of the dataset."""
        return list(self.get_secondary(primary).keys())
