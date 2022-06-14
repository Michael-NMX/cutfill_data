import pathlib


def getFilepath():
    dirpathFromUser = input("Path to directory containing HTML files: ")
    dirpath = pathlib.Path(dirpathFromUser)
    return dirpath


if __name__ == '__main__':
    getFilepath()