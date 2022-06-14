import pathlib


def getFilepath():
    filepathFromUser = input("Path to directory containing HTML files: ")
    filepath = pathlib.Path(filepathFromUser)
    return filepath


if __name__ == '__main__':
    getFilepath()