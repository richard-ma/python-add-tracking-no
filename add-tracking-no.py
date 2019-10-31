#!/usr/bin/env python
# encoding: utf-8

import os
import csv

class DataFile(object):

    def __init__(self, filename):
        self.delimiter = '_'
        self.t_list = ["express.xlsx", "ship.csv"]

        self.filename = filename
        self.prefix, self.filetype = self._getFileType()

    def _getFileType(self):
        prefix, suffix = self.filename.split(self.delimiter)

        # check filename
        if suffix in self.t_list:
            return (prefix, suffix.split('.')[1])
        else:
            raise Exception("Filename format ERROR!")

    def _checkFileExists(self):
        if not os.path.exists(self.filename):
            raise Exception("File not found!")

class ShipDataFile(DataFile):

    #[订单标识，商品交易号]
    def read(self):
        if self.filetype != "csv":
            raise Exception("DataFile type must be 'ship'!")

        self._checkFileExists()

        data = list()
        with open(self.filename, 'r+', newline='', encoding='GBK') as f:
            source = csv.reader(f)
            for r in source:
                data.append(r[:2])
        return dict(data[1:])

    def write(self, data):
        with open(self.filename, 'w+', encoding='GBK') as f:
            w = csv.writer(f)
            w.writerows(data)

if __name__ == "__main__":
    ship_file = ShipDataFile("test_ship.csv")
    data = ship_file.read()
    print(data)
