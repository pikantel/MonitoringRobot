# sortowanie importów https://www.python.org/dev/peps/pep-0008/#imports
import pandas as pd
import glob
import os
import pymongo
import csv
from pymongo.errors import DuplicateKeyError
import re


class UseDatabase:
    def __init__(self, line):
        self.line = line

    def open_collection(self):
        # connection url do mongo zdecydowanie powinien być w configu.
        # dodatkowo czy jest potrzeba tworzenia nowego klienta per request o collection?
        myclient = pymongo.MongoClient("mongodb://localhost:27017/")
        mydb = myclient["newdb"]
        collection = mydb["orders_ompt"]
        return collection

    def add_to_database(self):
        collection = self.open_collection()
        collection.insert_one(self.line)


# duplicate module to be updated in the near future

def remove_duplicates():
    collection = UseDatabase().open_collection()
    result = collection.find()
    new_list = []
    # unikaj nazywania x, y - są to niewiadome.
    # Zerknij też czy nie chciałbyś tu użyć setów - elementy są unikatowe w tej kolekcji.
    # https://www.w3schools.com/python/python_sets.asp
    # plus ten AttributeError - raczej to jest do wyrzucenia :D
    for x in result:
        for y in x.values():
            try:
                for i in y.values():
                    if not i[0] in new_list:
                        new_list.append(i[0])
                    else:
                        collection.remove(i)
            except AttributeError:
                continue
    return new_list


# utworzyłbym oddzielny skrypt - utils.py i wrzucił tam tą funkcję.
# takie utilsy są często robione w pythonie - są idealnym miejscem na wrzucenie generycznych funkcji do naszego projektu.
def clear_path(path):
    items = glob.glob(path + '\\*')
    for item in items:
        os.remove(item)


class SaveFilesIntoDatabase:
    def __init__(self, path_ompt, path_sap):
        self.path_ompt = path_ompt
        self.path_sap = path_sap

    def read_from_csv(self):  # this object is appending data from csv files (from ompt) to mongodb
        i = 1
        new_list = {}
        po_list = []
        for file in os.listdir(self.path_ompt):
            file_path = self.path_ompt + "\\" + file
            file = file.split('.')[0]
            with open(file_path, newline='', encoding='utf8') as csvfile:
                csvreader = csv.reader(csvfile, delimiter=',')
                new_dict = {}
                new_list[file] = new_dict
                for row in csvreader:
                    if "Transaction" in row:
                        continue
                    else:
                        try:
                            new_dict[file + str(i)] = row
                            if not 'products' in row[4]:
                                item = row[4][::-1]
                                item2 = item.split("-", 1)[-1]
                                item2 = item2[::-1]
                                po_list.append(item2)
                            i += 1
                        except DuplicateKeyError:
                            continue
                i = 1
        UseDatabase(new_list).add_to_database()
        return po_list

    def read_from_txt(self):  # this object is appending data from txt files (from SAP) to mongodb
        file = self.path_sap + 'data2.txt'
        with open(file, "r") as txt_file:
            lines = txt_file.readlines()
            result = []
            for x in lines[3:]:
                x = x.split('\t')[9]
                result.append(x)

            txt_file.close()

        # split and strip po num

        final_result = []
        for x in result:
            item = x.split(' ')[0]
            item = re.sub(r'^001', '', item)
            item = re.sub(r'\n*', '', item)
            final_result.append(item)
        return final_result


if __name__ == '__main__':
    directory = SaveFilesIntoDatabase(r'C:\Temp\output\OmptData', "C:\\Temp\\output\\SAP\\")
    table_from_sap = directory.read_from_txt()
    print(table_from_sap)
    table_from_ompt = directory.read_from_csv()
    print(table_from_ompt)
    not_in_sap = [item for item in table_from_ompt if item not in table_from_sap]
    in_sap = [item for item in table_from_ompt if item in table_from_sap]
    print(not_in_sap)
    print(len(not_in_sap))
    print(len(table_from_ompt))
    print(len(table_from_sap))
    print(in_sap)

