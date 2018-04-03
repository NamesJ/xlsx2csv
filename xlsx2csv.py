import sys, os, xlrd, csv

class XLSX2CSV (object):

    def __init__(self):
        pass

    def convert(self, pathXLSX, pathCSV):
        wb = self.openWorkbook(pathXLSX)
        ws = self.openWorksheet(wb)
        csv_file = self.openCSVFile(pathCSV)
        csv_writer = self.getCSVWriter(csv_file)
        self.writeCSVFromWorksheet(csv_writer, ws)
        self.closeCSVFile(csv_file)

    def openWorkbook(self, wb_path):
        return xlrd.open_workbook(wb_path)

    def openWorksheet(self, wb):
        return wb.sheet_by_name(wb.sheet_names()[0])

    def openCSVFile(self, csv_path):
        return open(csv_path, "w", encoding="utf-8", newline="")

    def getCSVWriter(self, csv_file):
        return csv.writer(csv_file, quoting=csv.QUOTE_ALL)

    def writeCSVFromWorksheet(self, wr, ws):#wr: csv_writer
        for row in range(ws.nrows):
            wr.writerow(ws.row_values(row))

    def closeCSVFile(self, csv_file):
        csv_file.close()

### END: class XLSX2CSV


def usage(me=sys.argv[0]):
    print('Usage:', me, '</path/to/infile.xlsx> </path/to/outfile.csv>')

# Ignore 'me' argument
def parseArgs(c=len(sys.argv)-1, v=sys.argv[1:]):
    # /path/to/file.xlsx and /path/to/outfile.csv are required
    if c is not 2:
        usage()
        sys.exit(1)

    return (v[0], v[1])   # (xlsx_path, csv_path)


def main():
    pathXLSX, pathCSV = parseArgs()

    if os.path.exists(pathXLSX) is False:
        print('Path to xlsx file does not exists.')
        usage()
        sys.exit(1)

    if os.path.exists(pathCSV) is True:
        msg = '\"' + pathCSV + '\" already exists. Overwrite (y/n)?'
        request_to_overwrite = str(raw_input(msg))
        if request_to_overwrite.lower() is 'n':
            sys.exit(1)

    print('Converting', pathXLSX, 'to', pathCSV)
    XLSX2CSV().convert(pathXLSX, pathCSV)
    print('Converted!')


if __name__ == '__main__':
    main()
