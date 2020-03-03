import os
import glob


def clear_path(path):
    items = glob.glob(path + '\\*')
    for item in items:
        os.remove(item)
