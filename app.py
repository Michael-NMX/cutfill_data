import pathlib
import datetime
import json
import pandas as pd


def loadSettings():
    settings_path = pathlib.Path('.\\app_settings.json')
    with open(settings_path, "r") as file:
        settings = json.load(file)
    return settings

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
    gapRowsBeforeProjectData = SETTINGS["gapRowsBeforeProjectData"]
    gapRowsBeforeVolumeData = SETTINGS["gapRowsBeforeVolumeData"]
    gapRowsBeforeSumVolumeData = SETTINGS["gapRowsBeforeSumVolumeData"]
    totalVolumeRows = 0
    projectDataRows = 0
    startRowForVolumeData = 0
    for dataTables in data:
        cutHeader, fillHeader = createHeaders(dataTables[0])
        cutFillData = dataTables[1].drop(0)
        cutFillData.iloc[:, 0] = cutFillData.iloc[:, 0].apply(lambda station: float(station.replace('+', '')))
        cutData, fillData = getColumnData(cutFillData)
        totalVolumeRows = cutFillData.shape[0]
        projectDataRows = cutHeader.shape[0]
        headerVolumeDataRows = cutHeader.shape[1]
        startRowForVolumeData = gapRowsBeforeProjectData \
                                + projectDataRows \
                                + gapRowsBeforeVolumeData \
                                + headerVolumeDataRows
        volumeFormulas = createVolumeFormulasColumns(totalVolumeRows, startRowForVolumeData)
        sumRow = totalVolumeRows \
                + startRowForVolumeData \
                - gapRowsBeforeSumVolumeData
        sumVolumeRow = createSumRowFormula(sumRow, startRowForVolumeData)
        cutData = cutData.join(volumeFormulas)
        cutData.loc[sumRow] = sumVolumeRow
        fillData = fillData.join(volumeFormulas)
        fillData.loc[sumRow] = sumVolumeRow
        splitCutFillData.append([cutHeader, cutData])
        splitCutFillData.append([fillHeader, fillData])
    return splitCutFillData

def createHeaders(projectData):
    road = projectData.iloc[1,0][10:]
    startStation = projectData.iloc[3,0][11:]
    endStation = projectData.iloc[4,0][9:]
    projectTitle = SETTINGS['projectTitle']
    roadTitle = f'{road}'
    cutVolumeTitle = SETTINGS['cutVolumeTitle']
    fillVolumeTitle = SETTINGS['fillVolumeTitle']
    stationRange = f'DEL KM {startStation} AL KM {endStation}'
    date = f'{datetime.date.today()}'
    cutHeader = pd.DataFrame([
        projectTitle,
        roadTitle,
        cutVolumeTitle,
        stationRange,
        date
    ])
    fillHeader = pd.DataFrame([
        projectTitle,
        roadTitle,
        fillVolumeTitle,
        stationRange,
        date
    ])
    return cutHeader, fillHeader

def getColumnData(dataTable):
    columnsForCut = SETTINGS['columnsForCut']
    columnsForFill = SETTINGS['columnsForFill']
    cutData = dataTable[columnsForCut]
    fillData = dataTable[columnsForFill]
    return cutData, fillData

def createVolumeFormulasColumns(endRow, startRow):
    def distanceFormula(startRow, i):
        return f'=(A{startRow + i}-A{startRow + i - 1})/2'
    
    def createVolumeFormula(startRow, i):
        return f'=(B{startRow + i}+B{startRow + i - 1})*(C{startRow + i})'
    
    def createAcumVolumeFormula(startRow, i):
        return f'=D{startRow + i}+E{startRow + i - 1}'
        
    volumeHeader = SETTINGS['volumeHeader']
    acumVolumeHeader = SETTINGS['acumVolumeHeader']
    distanceHeader = SETTINGS['distanceHeader']
    indexStep = 0
    distanceFormulas = [distanceFormula(startRow, i) for i in range(endRow)]
    distanceFormulas[0] = 0
    volumeFormulas = [createVolumeFormula(startRow, i)  for i in range(endRow)]
    volumeFormulas[0] = 0
    acumVolumeFormulas = [createAcumVolumeFormula(startRow, i) for i in range(endRow)]
    acumVolumeFormulas[0] = 0
    indexStep = 1
    return pd.DataFrame({
       distanceHeader : distanceFormulas,
       volumeHeader : volumeFormulas,
       acumVolumeHeader : acumVolumeFormulas
    }, index=range(indexStep, endRow + indexStep))

def createSumRowFormula(endRow, startRow):
    return ['-', '-', 'SUMA', f'=SUM(D{startRow}:D{endRow})', '-']

def createReport(data, dirpath):
    outputPath = dirpath / 'output.xlsx'
    startRowForProjectData = SETTINGS['startRowForProjectData']
    startColumnForProjectData = SETTINGS['startColumnForProjectData']
    startRowForVolumeData = SETTINGS['startRowForVolumeData']
    for i, dataTables in enumerate(data):
        if outputPath.exists():
            with pd.ExcelWriter(outputPath, 
                mode="a",
                engine='openpyxl',
                if_sheet_exists="replace"
            ) as writer:
                dataTables[0].to_excel(
                    writer,
                    sheet_name=f'data_{i}', 
                    header=False,
                    index=False,
                    startrow=startRowForProjectData,
                    startcol=startColumnForProjectData
                )
        else:    
            with pd.ExcelWriter(outputPath, mode="w", engine='openpyxl') as writer:
                dataTables[0].to_excel(
                    writer,
                    sheet_name=f'data_{i}',
                    header=False,
                    index=False,
                    startrow=startRowForProjectData,
                    startcol=startColumnForProjectData
                )
        with pd.ExcelWriter(
            outputPath,
            engine='openpyxl',
            mode='a',
            if_sheet_exists="overlay"
        ) as writer:
            dataTables[1].to_excel(
                writer,
                sheet_name=f'data_{i}',
                startrow=startRowForVolumeData,
                index=False)


if __name__ == '__main__':
    SETTINGS = loadSettings()
    dirpath = getDirpath()
    data, filepaths = readFiles(dirpath)
    cutFillData = splitTablesByCutFillColumns(data)
    createReport(cutFillData, dirpath)
    print('Done!')
