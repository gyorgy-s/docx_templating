import pandas as pd

from main import DATA_SOURCE


class TwoLevelDataset:
    def __init__(self, data_source: str) -> None:
        self._header = (
            pd.read_csv(data_source, index_col=None, header=None, nrows=1).iloc[0].to_list()
        )
        self._df = (
            pd.read_csv(data_source, index_col=None, header=0)
            .sort_values(by=[self._header[0]], ascending=True)
            .fillna("")
        )
        self._data = {}
        self._create_dict()

    def _create_dict(self):
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
        return self._data

    def get_primary_list(self) -> list:
        return list(self._data.keys())

    def get_secondary(self, primary: str) -> dict | None:
        return self._data.get(primary)

    def get_secondary_list(self, primary) -> list:
        return list(self.get_secondary(primary).keys())
