import pathlib
import pandas as pd


def getDirpath():
    dirpathFromUser = input("Path to directory containing HTML files: ")
    dirpath = pathlib.Path(dirpathFromUser)
    return dirpath

def readFiles(dirpath):
    filepaths = [filepath for filepath in dirpath.iterdir() if filepath.is_file()]
    filepaths = [filepath for filepath in filepaths \
                if filepath.suffix == '.htm' or filepath.suffix == '.html']
    dataTables = [extractDataFromFile(filepath) for filepath in filepaths]
    return dataTables, filepaths

def extractDataFromFile(filepath):
    return pd.read_html(filepath, flavor='bs4')


if __name__ == '__main__':
    dirpath = getDirpath()
    data, filepaths = readFiles(dirpath)
