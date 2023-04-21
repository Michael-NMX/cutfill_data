import pathlib
import datetime
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

def splitTablesByCutFillColumns(data):
    splitCutFillData = []
    columnsForCut = ['Station', 'Cut Area (Sq.m.)']
    columnsForFill = ['Station', 'Fill Area (Sq.m.)']
    totalRows = 0
    startRow = 0
    for dataTables in data:
        cutHeader, fillHeader = createHeaders(dataTables[0])
        cutFillData = dataTables[1].drop(0)
        cutData, fillData = getColumnData(cutFillData, columnsForCut, columnsForFill)
        totalRows = cutFillData.shape[0]
        startRow = cutHeader.shape[0] + 2
        volumeFormulas = createVolumeFormulasColumns(totalRows, startRow)
        cutData = cutData.join(volumeFormulas)
        fillData = fillData.join(volumeFormulas)
        splitCutFillData.append([cutHeader, cutData])
        splitCutFillData.append([fillHeader, fillData])
    return splitCutFillData

def createHeaders(projectData):
    road = projectData.iloc[1,0][10:]
    startStation = projectData.iloc[3,0][11:]
    endStation = projectData.iloc[4,0][9:]
    cutHeader = pd.DataFrame([
        'PROYECTO',
        f'{road}',
        'VOLÚMENES DE CORTE',
        f'DEL KM {startStation} AL KM {endStation}',
        f'{datetime.date.today()}'
    ])
    fillHeader = pd.DataFrame([
        'PROYECTO',
        f'{road}',
        'VOLÚMENES DE RELLENO',
        f'DEL KM {startStation} AL KM {endStation}',
        f'{datetime.date.today()}'
    ])
    return cutHeader, fillHeader

def getColumnData(dataTable, columnsForCut, columnsForFill):
    cutData = dataTable[columnsForCut]
    fillData = dataTable[columnsForFill]
    return cutData, fillData

def createVolumeFormulasColumns(totalRows, startRow):
    def createVolumeFormula(startRow, i):
        return f'=((B{startRow + i}+B{startRow + i - 1})/2)*(A{startRow + i}-A{startRow + i  - 1})'
    
    def createAcumVolumeFormula(startRow, i):
        return f'=C{startRow + i}+D{startRow + i - 1}'
        
    volumeHeader = 'VOLUMENES (m3)'
    acumVolumeHeader = 'VOLUMENES ACUMULADOS (m3)'
    volumeFormulas = [createVolumeFormula(startRow, i)  for i in range(totalRows)]
    volumeFormulas[0] = 0
    acumVolumeFormulas = [createAcumVolumeFormula(startRow, i) for i in range(totalRows)]
    acumVolumeFormulas[0] = 0
    return pd.DataFrame({
       volumeHeader : volumeFormulas,
       acumVolumeHeader : acumVolumeFormulas
    }, index=range(1, totalRows + 1))

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
    cutFillData = splitTablesByCutFillColumns(data)
    createReport(cutFillData, dirpath)
    print('Done!')
