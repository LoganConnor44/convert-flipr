import argparse
import openpyxl
import math

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    INFO = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

def convertFlipr(sourceFileName, sheetName, sourceLocation, outputLocation, verbose) :
    if verbose :
        print(f"{bcolors.HEADER}Beginning Script{bcolors.ENDC}")
    
    if verbose :
        print(f"{bcolors.INFO}Loading Original File And Targeted Sheet{bcolors.ENDC}")
    originalWorkBook = openpyxl.load_workbook(f'./{sourceLocation}/{sourceFileName}')
    targetSheet = originalWorkBook[sheetName]

    if verbose :
        print(f"{bcolors.INFO}Initializing Counter For Data Set{bcolors.ENDC}")
    setCounter = 1;

    if verbose :
        print(f"{bcolors.INFO}Creating Newly Converted Excel Document{bcolors.ENDC}")
    newWorkBook = openpyxl.Workbook()
    newWorkSheet = newWorkBook.active
    boldFont = openpyxl.styles.Font(bold=True)
    newWorkSheet.title = 'Summary'
    newWorkSheet['A1'] = 'Row'
    newWorkSheet['B1'] = 'Column'
    newWorkSheet['C1'] = 'Recording_ID'
    newWorkSheet['A1'].font = boldFont
    newWorkSheet['B1'].font = boldFont
    newWorkSheet['C1'].font = boldFont
    fileNameForNewWorkBook = targetSheet['A1'].value

    if verbose :
        print(f"{bcolors.INFO}Creating Variables For Columns Past Z{bcolors.ENDC}")
    newDataColumns = list(map(chr,range(ord('D'),ord('Z')+1)))
    fullAlphabet = list(map(chr,range(ord('A'),ord('Z')+1)))
    newDataColumnsPointer = 0;

    if verbose :
        print(f"{bcolors.INFO}Initializing Variables For Loop Logic{bcolors.ENDC}")
    desiredColumnNameIndex = 'A';
    dataColumns = list(map(chr,range(ord('B'),ord('Y')+1)))
    dataColumnIndex = 0
    beginRow = 9
    endRow = 24
    desiredRowNameIndex = 6
    while True :
        if verbose :
            print(f"{bcolors.OKGREEN}Beginning To Read Data Within Set {setCounter}{bcolors.ENDC}")
        
        isFirstBeginningOfLetterLoop = dataColumnIndex == 0
        doesColumnHeaderForDataExist = targetSheet[f'B{beginRow - 1}'].value != None
        if verbose :
            if isFirstBeginningOfLetterLoop and not doesColumnHeaderForDataExist :
                print(f"{bcolors.INFO}No New Sets Identified - Stopping Loop{bcolors.ENDC}")
            else :
                print(f"{bcolors.INFO}More Data Sets Identified - Continuing Loop{bcolors.ENDC}")
        if isFirstBeginningOfLetterLoop and not doesColumnHeaderForDataExist :
            break

        desiredRowColumn = f'{desiredColumnNameIndex}{desiredRowNameIndex}'
        desiredHeaderText = targetSheet[desiredRowColumn]
        desiredColumnName = desiredHeaderText.value
        if verbose :
            print(f"{bcolors.INFO}Retrieving Column Header Value For Newly Created Excel Document{bcolors.ENDC}")
            print(f"{bcolors.INFO}desiredColumnName: {desiredColumnName}{bcolors.ENDC}")
        
        dataRange = f'{dataColumns[dataColumnIndex]}{beginRow}:{dataColumns[dataColumnIndex]}{endRow}'
        columnData = targetSheet[dataRange]
        if verbose :
            print(f"{bcolors.INFO}Retrieving Column And Row Range: {dataRange}{bcolors.ENDC}")

        newRowNumberForSet = 2 + (16 * dataColumnIndex)
        for idx, row in enumerate(columnData) :
            for cell in row :
                newWorkSheetLocationColumnRow = f'{newDataColumns[0]}{(newRowNumberForSet)}'
                newWorkSheet[newWorkSheetLocationColumnRow] = cell.value
                
            newWorkSheet[f'A{(newRowNumberForSet)}'] = dataColumnIndex + 1
            newWorkSheet[f'B{(newRowNumberForSet)}'] = fullAlphabet[idx]
            newWorkSheet[f'C{(newRowNumberForSet)}'] = fileNameForNewWorkBook
            newRowNumberForSet += 1
        dataColumnIndex += 1

        # if end of letters list - restart indices for next set
        if dataColumnIndex == len(dataColumns) :
            reachedEndOfLThroughZ = setCounter > len(newDataColumns)
            if reachedEndOfLThroughZ :
                if newDataColumnsPointer > 26 :
                    newDataColumnsPointer = 0
                firstDoubleLetterIndex = math.floor(setCounter / 26)
                doubleLetterColumnName = f'{fullAlphabet[firstDoubleLetterIndex]}{fullAlphabet[newDataColumnsPointer]}'
                newDataColumns.extend([doubleLetterColumnName])
                newDataColumnsPointer += 1
            newWorkSheet[f'{newDataColumns[0]}1'] = desiredColumnName
            newWorkSheet[f'{newDataColumns.pop(0)}1'].font = boldFont
            dataColumnIndex = 0
            beginRow += 19
            endRow += 19
            desiredRowNameIndex += 19
            setCounter += 1
    
    if verbose :
        print(f"{bcolors.HEADER}Saving Newly Created Excel Document{bcolors.ENDC}")
    newWorkBook.save(f'./{outputLocation}/{sourceFileName}')

    return

def mainApplication() :
    parser=argparse.ArgumentParser()
    parser.add_argument(
        '--source-file-name',
        help = 'Specifies which file within the local directory to convert.',
        required = True
    )
    parser.add_argument(
        '--sheet-name',
        help = 'Specifies the sheet within the source file that will be converted.',
        required = True
    )
    parser.add_argument(
        '--verbose',
        help = 'If set to TRUE, extra messages will be printed. Default is FALSE.',
        default = 'False',
        choices = [
            'true',
            'false'
        ]
    )
    parser.add_argument(
        '--source-location',
        help = 'Specifies which local directory our original files are located at. Default is ./original/',
        default = 'original',
        required = False
    )
    parser.add_argument(
        '--output-location',
        help = 'Specifies which local directory our converted files will be generated at. Default is ./converted/',
        default = 'converted',
        required = False
    )
    args = parser.parse_args()

    sourceFileName = args.source_file_name
    sheetName = args.sheet_name
    sourceLocation = 'original'
    outputLocation = 'converted'
    verbose = False

    if args.source_location != None :
        sourceLocation = args.source_location.lower()
    if args.output_location != None :
        outputLocation = args.output_location.lower()
    if args.verbose != None and args.verbose.lower() == 'true' :
        verbose = True
    
    print(f"{bcolors.HEADER}Target File: {sourceFileName}")
    print(f"{bcolors.HEADER}Sheet Name Within File: {sheetName}")
    print(f"{bcolors.HEADER}Source Location: {sourceLocation}")
    print(f"{bcolors.HEADER}Output Location: {outputLocation}")
    print(f"{bcolors.HEADER}Verbose: {verbose}")

    convertFlipr(sourceFileName, sheetName, sourceLocation, outputLocation, verbose)

if __name__ == '__main__':
    mainApplication()
