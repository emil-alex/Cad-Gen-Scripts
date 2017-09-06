import win32com.client as win32
from pathlib import Path


class Fisa(object):

    def __init__(self, filenum):
        self.name = path + "\\Fise Imobil Sector 13\\" + str(filenum) + ".xls"
        self.deed_no_list = []
        self.deed_date_list = []
        self.def_status = []
        self.excel = win32.Dispatch('Excel.Application')
        self.wb = self.excel.Workbooks.Open(self.name)
        self.ws = self.wb.Sheets(1)

        my_range = self.ws.Range('B1', 'J1')  # Unmerges cells in the B1 - J1 range.
        my_range.MergeCells = False

    def populate_data(self):  # Poate fi modificata cu un sigur IF.
        """ Extracts data form CGXML file and populates the Deed Number, Deed Date and Defunct Status lists with data."""
        with open(cg_name, 'r') as fhand:
            for line in fhand:
                line = line.strip()
                if line.startswith('<DEEDNUMBER>'):
                    line = line.replace('<DEEDNUMBER>', '')
                    line = line.replace('</DEEDNUMBER>', '')
                    self.deed_no_list.append(line)
                if line.startswith('<DEEDDATE>'):
                    line = line.replace('<DEEDDATE>', '')
                    line = line.replace('T00:00:00+02:00</DEEDDATE>', '')
                    self.deed_date_list.append(line)
                if line.startswith('<DEFUNCT>'):
                    line = line.replace('<DEFUNCT>', '')
                    line = line.replace('</DEFUNCT>', '')
                    self.def_status.append(line)

    def aranjare_initiala(self):
        """ Initial setting. Deletes certain blank rows, sets certain column widths.
            If cell B1 is 0, the method will not be called in main program
            """
        self.ws.Cells(10, 8).Value = 'Cooperativizata/Necooperativizata'
        self.ws.Cells(11, 8).Value = '(Co/Nco)'

        row_list = [22, 12, 2, 1]
        for row in row_list:
            self.ws.Rows(row).Delete()

        my_range = self.ws.Range('B10', 'G10')
        for border_id in range(7, 13):
            my_range.Borders(border_id).LineStyle = 1
            my_range.Borders(border_id).Weight = 2

        self.ws.Columns(7).ColumnWidth = 15
        self.ws.Columns(8).ColumnWidth = 18
        self.ws.Columns(9).ColumnWidth = 11

        self.ws.Cells(1, 2).Value = '0'
        self.ws.Cells(1, 2).Font.Color = '&hFFFFFF'

    def modify_deed_cell(self):  # Poate fi modificat sa nu inceapa de la 28, ci sa caute singur pozitia
        """ Ads the date of the deed to the deed number cell.
            If cell C1 is 0, the method will not be called in main program
            """
        lung = len(self.deed_no_list)
        row = 28
        while True:
            val = self.ws.Cells(row, 8).Value
            for pos in range(0, lung-1):  # for j in (0, lung-1):  ---- initial
                if val == self.deed_no_list[pos]:
                    data = self.deed_date_list[pos].split('-')
                    self.ws.Cells(row, 8).Value = '{}/{}.{}.{}'.format(val, data[2], data[1], data[0])
            row += 1
            if val is None:
                break

        self.ws.Cells(1, 3).Value = '0'
        self.ws.Cells(1, 3).Font.Color = '&hFFFFFF'

    def defunct_status(self):  # Poate fi modificat sa nu inceapa de la 24, ci sa caute singur pozitia
        """ Ads 'Defunct' status to the observations cell.
            If cell D1 is 0, the method will not be called in main program
            """
        row = 24
        for def_value in self.def_status:
            val = self.ws.Cells(row, 9).Value
            if def_value == 'true':
                if val == 'CNP neidentificat':
                    self.ws.Cells(row, 9).Value = 'Defunct; CNP neidentificat'
                else:
                    my_range = self.ws.Range('I' + str(row), 'J' + str(row))
                    my_range.MergeCells = True
                    self.ws.Cells(row, 9).Value = 'Defunct'
            row += 1

        self.ws.Cells(1, 4).Value = '0'
        self.ws.Cells(1, 4).Font.Color = '&hFFFFFF'

    def close_file(self):
        self.wb.Close(True)
        self.excel.Quit()


path = r'E:\Drive\CG'

for i in range(41, 46):
    check_file = Path(path + "\\Fise Imobil Sector 13\\" + str(i) + ".xls")
    if check_file.is_file():
        modificat = False  # If the file is modified, it will notify the user.
        fisa = Fisa(i)
        cg_name = path + "\\Fisiere CGXML Sector 13\\" + str(i) + ".txt"

        if fisa.ws.Cells(1, 2).Value != '0':
            fisa.aranjare_initiala()
            modificat = True
        else:
            print("Metoda 'Aranjare initiala' a fost deja efectuata.",  check_file, "nemodificat")

        if fisa.ws.Cells(1, 3).Value != '0':
            fisa.populate_data()
            fisa.modify_deed_cell()
            modificat = True
        else:
            print("Metoda 'Modify Deed Cell' a fost deja efectuata.",  check_file, "nemodificat")

        if fisa.ws.Cells(1, 4).Value != '0':
            fisa.populate_data()
            fisa.defunct_status()
            modificat = True
        else:
            print("Metoda 'Defunct Status' a fost deja efectuata.",  check_file, "nemodificat")


        fisa.close_file()

        if modificat:
            print(check_file, 'modificat cu succes')
