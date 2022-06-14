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

def createReport(data, dirpath):
    outputPath = dirpath / 'output.xlsx'
    for i, dataTables in enumerate(data):
        if outputPath.exists():
            with pd.ExcelWriter(outputPath, 
                mode="a",
                engine='openpyxl',
                if_sheet_exists="replace"
            ) as writer:
                dataRows = len(dataTables[0].index)
                dataTables[0].to_excel(writer, sheet_name=f'data_{i}', header=False, index=False)
        else:    
            with pd.ExcelWriter(outputPath, mode="w", engine='openpyxl') as writer:
                dataRows = len(dataTables[0].index)
                dataTables[0].to_excel(writer, sheet_name=f'data_{i}', header=False, index=False)
        with pd.ExcelWriter(
            outputPath,
            engine='openpyxl',
            mode='a',
            if_sheet_exists="overlay"
        ) as writer:
            dataRows += 2
            dataTables[1].to_excel(writer, sheet_name=f'data_{i}', startrow=dataRows, index=False)


if __name__ == '__main__':
    dirpath = getDirpath()
    data, filepaths = readFiles(dirpath)
    createReport(data, dirpath)
    print('Done!')
