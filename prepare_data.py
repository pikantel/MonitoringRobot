
import pandas as pd
import glob
import os
import pymongo
import csv
from pymongo.errors import DuplicateKeyError
import re


def remove_duplicates():
    myclient = pymongo.MongoClient("mongodb://localhost:27017/")
    mydb = myclient["newdb"]
    collection = mydb["orders_ompt"]
    result = collection.find()
    new_list = []
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


def add_to_database(line):
    myclient = pymongo.MongoClient("mongodb://localhost:27017/")
    mydb = myclient["newdb"]
    orders_Ompt = mydb["orders_ompt"]
    orders_Ompt.insert_one(line)


class ClearPath:
    def __init__(self, path):
        self.path = path

    def clear_path(self):
        items = glob.glob(self.path + '\\*')
        for item in items:
            os.remove(item)


def read_from_txt():  # this object is appending data from txt files (from SAP) to mongodb
    path2 = "C:\\Temp\\output\\SAP\\"
    file = path2 + 'data2.txt'
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


class SaveFilesIntoDatabase:
    def __init__(self, path):
        self.path = path

    def read_from_csv(self):  # this object is appending data from csv files (from ompt) to mongodb
        i = 1
        new_list = {}
        po_list = []
        for file in os.listdir(self.path):
            file_path = self.path + "\\" + file
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
        add_to_database(new_list)
        return po_list


if __name__ == '__main__':
    directory = SaveFilesIntoDatabase(r'C:\Temp\output\OmptData')
    table_from_sap = read_from_txt()
    print(table_from_sap)
    table_from_ompt = directory.read_from_csv()
    print(table_from_ompt)
    not_in_sap = [item for item in table_from_ompt if item not in table_from_sap]
    in_sap = [item for item in table_from_ompt if item in table_from_sap]
    print(len(not_in_sap))
    print(len(table_from_ompt))
    print(len(table_from_sap))
    print(in_sap)
    print(not_in_sap)




