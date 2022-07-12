import datetime
import re
import time
from pathlib import Path

import dateutil
import numpy as np
import pandas as pd
from openpyxl import Workbook  # Per creare un libro
from openpyxl.chart import AreaChart, BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabel, DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.legend import LegendEntry
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.fill import ColorChoice, PatternFillProperties
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.text import (CharacterProperties, Paragraph,
                                   ParagraphProperties)
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import (Alignment, Border, Font,  # Per cambiare lo stile
                             PatternFill, Side)
from openpyxl.styles.numbers import FORMAT_NUMBER_00  # Stili di numeri
from openpyxl.styles.numbers import FORMAT_PERCENTAGE_00
from openpyxl.utils import get_column_letter  # Per lavorare sulle colonne
from openpyxl.utils.dataframe import \
    dataframe_to_rows  # Per l'import di dataframe
from openpyxl.utils.units import cm_to_EMU, pixels_to_EMU  # Per l'ancoraggio
from openpyxl.worksheet.header_footer import HeaderFooter, HeaderFooterItem
from openpyxl.worksheet.page import PageMargins  # Opzioni di stampa
from sqlalchemy import MetaData, Table, create_engine


class Report():
    """Crea un report di un portafoglio."""
    # TODO:generalizza lo script per fare report mensili / bimensili / trimestrali...

    def __init__(self, t1, file_portafoglio='artes.xlsx'):
        """
        Initialize the class.

        Parameters:
        t1(str to datetime) = data finale
        t0_1Y(datetime) = t1 un anno fa
        t0_ytd(datetime) = t1 all fine dell'anno scorso
        t0_3Y(datetime) = t1 tre anni fa
        image(str) = percorso assoluto in cui si trova il logo
        cellh(float) = formula per lo spostamento in verticale
        cellw(float) = formula per lo spostamento in orizzontale
        mesi_dict(dict) = dizionario che associa ogni numero ordinale del mese con il suo nome
        """
        self.wb = Workbook()
        self.t1 = datetime.datetime.strptime(t1, '%d/%m/%Y')
        print(f"Data report : {self.t1}.")
        self.t0_1m = self.t1.replace(day=1) - dateutil.relativedelta.relativedelta(days=1)
        print(f"Un mese fa : {self.t0_1m}.")
        self.t0_1Y = (self.t1 - dateutil.relativedelta.relativedelta(years=+1))#.strftime("%d/%m/%Y") # data t1 un anno fa
        print(f"Un anno fa : {self.t0_1Y}.")
        self.t0_ytd = datetime.datetime(year=self.t0_1Y.year, month=12, day=31)
        print(f"L'ultimo giorno dell'anno scorso : {self.t0_ytd}.")
        self.t0_3Y = (self.t1 - dateutil.relativedelta.relativedelta(years=+3))#.strftime("%d/%m/%Y") # data t1 tre anni fa
        print(f"Tre anni fa : {self.t0_3Y}.")
        # L'intervallo del report lo vedi dalle date in agosto_2020.
        self.path = Path('C:\\Users\\Alessio\\Documents\\Sbwkrq\\Report')
        self.file_portafoglio = self.path.joinpath(file_portafoglio)
        self.image = self.path.joinpath('img', 'logo_B&S.bmp')
        self.cellh = lambda x: cm_to_EMU((x * 49.77)/99)
        self.cellw = lambda x: cm_to_EMU((x * (18.65-1.71))/10)
        self.mesi_dict = {1: 'Gennaio', 2: 'Febbraio', 3: 'Marzo', 4: 'Aprile', 5: 'Maggio', 6: 'Giugno', 7: 'Luglio', 8: 'Agosto', 9: 'Settembre', 10: 'Ottobre', 11: 'Novembre', 12: 'Dicembre'}
        # Pipeline with Python and Postrge
        # DATABASE_URL = 'postgres+psycopg2://postgres:bloomberg893@localhost:5432/artes'
        # self.engine = create_engine(DATABASE_URL)
        # self.connection = self.engine.connect()

        # Carica portafoglio
        portfolio = pd.read_excel(self.file_portafoglio, sheet_name='Portfolio', header=1)
        controvalore_t1 = portfolio['TOTALE t1'].sum()
        print(f"\nIl controvalore del portafoglio è : {controvalore_t1}.")
        
        controvalore_t0_1m = portfolio['TOTALE t0'].sum()
        print(f"Il controvalore del portafoglio nel mese precedente era : {controvalore_t0_1m}.")

    def logo(self, active_ws, col=5, colOff=0.3, row=34, rowOff=0, picture=''):
        """
        Aggiunge al foglio attivo l'immagine picture alla colonna col e alla riga row.
        Applica uno spostamento di colOff colonne e rowOff righe.

        Parameters:
        active_ws(str) = worksheet attivo in cui incollare l'immagine
        col(int) = colonna di partenza in cui incollare l'immagine
        colOff(float) = spostamento dalla colonna di partenza
        row(int) = riga di partenza in cui incollare l'immagine
        rowOff(float) = spostamento dalla riga di partenza
        picture(str) = percorso assoluto in cui si trova l'immagine
        """
        if not picture:
            picture = self.image
        logo = Image(picture)
        h, w = logo.height, logo.width
        size = XDRPositiveSize2D(pixels_to_EMU(w), pixels_to_EMU(h))
        maker = AnchorMarker(col=col, colOff=self.cellw(colOff), row=row, rowOff=self.cellh(rowOff))
        ancoraggio = OneCellAnchor(_from=maker, ext=size)
        active_ws.add_image(logo)
        logo.anchor = ancoraggio

    def text_box(self, ws, min_row, max_row, min_col, max_col, fill_type='solid', fill_color='FFFFFF', font_name='Times New Roman',
        font_size=12, font_color='31869B', border_style='medium', border_color='31869B'):
        """
        Simula una text-box
        
        Parameters:
            ws {class 'openpyxl.worksheet.worksheet.Worksheet'} = foglio excel in cui creare la text box
            min_row, max_row, min_col, max_col {int} = coordinate dove inserire la text box
            fill_type {str} = tipo di riempimento, 'solid' di default.
            fill_color {hex color} = colore del riempimento della text box, 'FFFFFF' di default.
            font_name {str} = nome del font da usare per il testo scritto nella text box, 'Times New Roman' di default
            font_size {int} = dimensione dei caratteri del testo scritto nella text box, 12 di default.
            font_color {hex color} = colore del testo scritto nella text box, '31869B' di default.
            border_style {str} = stile del bordo da applicare alla text box, 'medium' di default.
            border_color {hex color} = colore da applicare al bordo della text box, '31869B' di default.
        """
        
        for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            for _ in range(max_col-min_col+1):
                ws[row[_].coordinate].fill = PatternFill(fill_type=fill_type, fgColor=fill_color)
                ws[row[_].coordinate].font = Font(name=font_name, size=font_size, color=font_color)
            ws[row[0].coordinate].border = Border(left=Side(border_style=border_style, color=border_color))
            ws[row[max_col-min_col].coordinate].border = Border(right=Side(border_style=border_style, color=border_color))
            if row[0].row == min_row:
                for _ in range(max_col-min_col+1):
                    if row[_].column == min_col:
                        ws[row[_].coordinate].border = Border(top=Side(border_style=border_style, color=border_color), left=Side(border_style=border_style, color=border_color))
                    elif row[_].column == max_col:
                        ws[row[_].coordinate].border = Border(top=Side(border_style=border_style, color=border_color), right=Side(border_style=border_style, color=border_color))
                    else:
                        ws[row[_].coordinate].border = Border(top=Side(border_style=border_style, color=border_color))
            elif row[0].row == max_row:
                for _ in range(max_col-min_col+1):
                    if row[_].column == min_col:
                        ws[row[_].coordinate].border = Border(bottom=Side(border_style=border_style, color=border_color), left=Side(border_style=border_style, color=border_color))
                    elif row[_].column == max_col:
                        ws[row[_].coordinate].border = Border(bottom=Side(border_style=border_style, color=border_color), right=Side(border_style=border_style, color=border_color))
                    else:
                        ws[row[_].coordinate].border = Border(bottom=Side(border_style=border_style, color=border_color))

    def copertina_1(self):
        """
        Crea la prima pagina. 
        Solo formattazione.
        """
        ws = self.wb.active
        ws.title = '1.copertina'
        #self.wb.active = ws
        ws['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws.merge_cells('A1:L1')
        ws['A11'] = 'Benchmark & Style'
        ws['A11'].alignment = Alignment(horizontal='center', vertical='center')
        ws['A11'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws['A11'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws.merge_cells('A11:L14')
        ws['E17'] = 'ARTES' # Nome cliente
        ws['E17'].alignment = Alignment(horizontal='center', vertical='center')
        ws['E17'].font = Font(name='Times New Roman', size=18, bold=True, color='31869B')
        ws['G17'] = self.t1.strftime('%d/%m/%Y')
        ws['G17'].alignment = Alignment(horizontal='center', vertical='center')
        ws['G17'].font = Font(name='Times New Roman', size=18, bold=True, color='31869B')
        ws.merge_cells('E17:F19')
        ws.merge_cells('G17:H19')
        ws['A33'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws.merge_cells('A33:L33')
        # Logo
        logo = Image(self.image)
        ws.add_image(logo, 'F27') # centimeters = pixels * 2.54 / 96
        logo.height = 75.59
        logo.width = 128.88188976377952755905511811024

    def indice_2(self):
        """
        Crea la seconda pagina.
        Solo formattazione.
        """
        ws2 = self.wb.create_sheet('2.indice')
        ws2 = self.wb['2.indice']
        self.wb.active = ws2
        ws2.merge_cells('A1:L4')
        ws2['A1'] = 'Indice'
        ws2['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws2['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws2['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws2['B8'] = '1. Analisi Di Mercato'
        ws2['B8'].font = Font(name='Times New Roman', size=18, bold=True, color='31869B')
        ws2['B11'] = '2. Performance'
        ws2['B11'].font = Font(name='Times New Roman', size=18, bold=True, color='31869B')
        ws2['B14'] = '3. Valutazione Per Macroclasse'
        ws2['B14'].font = Font(name='Times New Roman', size=18, bold=True, color='31869B')
        ws2['B17'] = '4. Contatti'
        ws2['B17'].font = Font(name='Times New Roman', size=18, bold=True, color='31869B')
        # Logo
        self.logo(ws2, row=32)
    
    def analisi_di_mercato_3(self):
        """
        Crea la terza pagina.
        Solo formattazione.
        """
        # 3.Analisi mercato
        ws3 = self.wb.create_sheet('3.an_mkt')
        ws3 = self.wb['3.an_mkt']
        self.wb.active = ws3
        ws3['A11'] = '1. Analisi Di Mercato'
        ws3['A11'].alignment = Alignment(horizontal='center', vertical='center')
        ws3['A11'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws3['A11'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws3.merge_cells('A11:L14')
        # Logo
        self.logo(ws3)

    def analisi_rendimenti_4(self):
        """
        Crea la quarta pagina.
        Formattazione e una tabella.
        Aggiunge fogli Indici e fogli Indici_in_euro.
        """
        # Carica indici e tassi di cambio
        indici_tassi = pd.read_excel(self.file_portafoglio, sheet_name='Indici', header=[0, 1], parse_dates=True, index_col=0)
        indici_tassi.columns = indici_tassi.columns.droplevel(-1)
        # print(indici_tassi)
        # Carica indici in euro
        indici_in_euro = pd.read_excel(self.file_portafoglio, sheet_name='Indici_in_euro', header=[0, 1], parse_dates=True, index_col=0)
        indici_in_euro.columns = indici_in_euro.columns.droplevel(-1)
        # print(indici_in_euro)
        # Crea dizionario con i rendimenti degli indici e dei tassi
        indici_perf = {'S&P 500' : [],'NIKKEI' : [],'NASDAQ' : [],'FTSE 100' : [],'FTSE MIB' : [],'DAX' : [],'DOW JONES INDUSTRIAL AVERAGE' : [],'EURO STOXX 50' : [],'HANG SENG' : [],'MSCI WORLD' : [],
            'MSCI EMERGING MARKETS' : [],'HFRX EWSI' : [],'WTI CRUDE OIL FUTURE' : [],'LONDON GOLD MARKET FIXING LTD' : [],'COMMODITY RESEARCH BUREAU' : [],'LYXOR ETF EURO CASH' : [],'LYXOR ETF EURO CORP BOND' : [],
            'BARCLAYS EUROAGG CORP TR' : [],'JPM GBI EMU 1_10' : [], 'JPM GBI EMU 3_5' : [],'JPM GBI EMU 1_3' : [],'USDEUR' : [],'GBPEUR' : [],'CHFEUR' : [], 'AUDEUR' : [],'NOKEUR' : []}
        for column in indici_tassi.columns.values:
            monthly = (indici_tassi.loc[self.t1, column] - indici_tassi.loc[self.t0_1m, column]) / (indici_tassi.loc[self.t0_1m, column])
            ytd = (indici_tassi.loc[self.t1, column] - indici_tassi.loc[self.t0_ytd, column]) / (indici_tassi.loc[self.t0_ytd, column])
            one_year = (indici_tassi.loc[self.t1, column] - indici_tassi.loc[self.t0_1Y, column]) / (indici_tassi.loc[self.t0_1Y, column])
            three_years = (indici_tassi.loc[self.t1, column] - indici_tassi.loc[self.t0_3Y, column]) / (indici_tassi.loc[self.t0_3Y, column])
            indici_perf[column] = [monthly, ytd, one_year, three_years]
        for column in indici_in_euro.columns.values:
            ytd_eur = (indici_in_euro.loc[self.t1, column] - indici_in_euro.loc[self.t0_ytd, column]) / (indici_in_euro.loc[self.t0_ytd, column])
            indici_perf[column].append(ytd_eur)
        # print(indici_perf)

        ws4 = self.wb.create_sheet('4.an_mkt_rend')
        ws4 = self.wb['4.an_mkt_rend']
        self.wb.active = ws4
        ws4.merge_cells('A1:O4')
        ws4['A1'] = 'Analisi Di Mercato'
        ws4['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws4['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws4['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws4.row_dimensions[5].height = 3
        ws4['A6'] = 'Performance ' + self.mesi_dict[self.t1.month]
        ws4['A6'].alignment = Alignment(horizontal='center', vertical='center')
        ws4['A6'].font = Font(name='Times New Roman', size=10, bold=True, color='FFFFFF')
        ws4['A6'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws4['A6'].border = Border(top=Side(border_style='medium', color='000000'), bottom=Side(border_style='medium', color='000000'), right=Side(border_style='medium', color='000000'), left=Side(border_style='medium', color='000000'))
        ws4['A7'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws4.merge_cells('A6:I6')

        # Colonne tabella
        header_4 = ['', '', '', '', 'MENSILI', 'YTD €', 'YTD', '1y', '3y']
        for column in ws4.iter_cols(min_row=7, max_row=7, min_col=1, max_col=9):
            ws4[column[0].coordinate].value = header_4[0]
            del header_4[0]
            ws4[column[0].coordinate].font = Font(name='Times New Roman', size=9, bold=True, color='FFFFFF')
            ws4[column[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='31869B')
            ws4[column[0].coordinate].alignment = Alignment(horizontal='center', vertical='center')
            ws4[column[0].coordinate].border = Border(top=Side(border_style='medium', color='000000'), bottom=Side(border_style='medium', color='000000'))
            ws4.cell(row=column[0].row, column=1).border = Border(left=Side(border_style='medium', color='000000'))
            ws4.cell(row=column[0].row, column=9).border = Border(right=Side(border_style='medium', color='000000'))

        # Corpo tabella
        index_4 = ['AZIONARI', 'S&P 500', 'NIKKEI', 'NASDAQ', 'FTSE 100', 'FTSE MIB', 'DAX', 'DOW JONES INDUSTRIAL AVERAGE', 'EURO STOXX 50', 'HANG SENG', 'MSCI WORLD', 'MSCI EMERGING MARKETS',
            'HEDGE FUND', 'HFRX EWSI', 'COMMODITIES', 'WTI CRUDE OIL FUTURE', 'LONDON GOLD MARKET FIXING LTD', 'COMMODITY RESEARCH BUREAU', 'OBBLIGAZIONARI GOVERNATIVE', 'LYXOR ETF EURO CASH',
            'JPM GBI EMU 1_3', 'JPM GBI EMU 3_5', 'JPM GBI EMU 1_10', 'OBBLIGAZIONARI CORPORATE', 'LYXOR ETF EURO CORP BOND', 'BARCLAYS EUROAGG CORP TR', 'VALUTE', 'USDEUR', 'GBPEUR',
            'CHFEUR', 'AUDEUR', 'NOKEUR', "le valute sono espresse come quantità di euro per un'unità di valuta estera"]
        for row in ws4.iter_rows(min_row=8, max_row=40, min_col=1, max_col=9):
            ws4[row[0].coordinate].value = index_4[0]
            del index_4[0]
            ws4.row_dimensions[row[0].row].height = 13
            if ws4[row[0].coordinate].value == 'AZIONARI' or ws4[row[0].coordinate].value == 'HEDGE FUND' or ws4[row[0].coordinate].value == 'COMMODITIES'	or ws4[row[0].coordinate].value == 'OBBLIGAZIONARI GOVERNATIVE' or ws4[row[0].coordinate].value == 'OBBLIGAZIONARI CORPORATE' or ws4[row[0].coordinate].value == 'VALUTE' or ws4[row[0].coordinate].value == 'le valute sono espresse come quantità di euro per un\'unità di valuta estera':
                ws4[row[0].coordinate].font = Font(name='Times New Roman', size=8, bold=True, color='006666')
                ws4[row[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='92CDDC')
                ws4[row[0].coordinate].alignment = Alignment(horizontal='center', vertical='center')
                ws4[row[0].coordinate].border = Border(top=Side(border_style='medium', color='000000'), left=Side(border_style='medium', color='000000'), bottom=Side(border_style='medium', color='000000'), right=Side(border_style='medium', color='000000'))
                ws4.merge_cells(start_row=row[0].row, end_row=row[0].row, start_column=1, end_column=9)
            else:
                ws4[row[0].coordinate].font = Font(name='Times New Roman', size=8, bold=True, color='FFFFFF')
                ws4[row[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='31869B')
                ws4[row[0].coordinate].alignment = Alignment(vertical='center')
                ws4[row[0].coordinate].border = Border(left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                ws4.merge_cells(start_row=row[0].row, end_row=row[0].row, start_column=1, end_column=4)
            if ws4[row[0].coordinate].value in indici_perf.keys():
                if row[0].row < 34: # Riempi la tabella tranne valute (00B050 + FF0000 -)
                    ws4[row[4].coordinate].value = "{0:.2f}%".format(round(indici_perf[row[0].value][0] * 100, 2)).replace('.', ',')
                    if indici_perf[row[0].value][0] > 0:
                        ws4[row[4].coordinate].font = Font(color='00B050')
                    else:
                        ws4[row[4].coordinate].font = Font(color='FF0000')
                    ws4[row[4].coordinate].alignment = Alignment(horizontal='center', vertical='center')
                    ws4[row[4].coordinate].border = Border(bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                    ws4[row[5].coordinate].value = "{0:.2f}%".format(round(indici_perf[row[0].value][4] * 100, 2)).replace('.', ',')
                    if indici_perf[row[0].value][4] > 0:
                        ws4[row[5].coordinate].font = Font(color='00B050')
                    else:
                        ws4[row[5].coordinate].font = Font(color='FF0000')
                    ws4[row[5].coordinate].alignment = Alignment(horizontal='center', vertical='center')
                    ws4[row[5].coordinate].border = Border(bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                    ws4[row[6].coordinate].value = "{0:.2f}%".format(round(indici_perf[row[0].value][1] * 100, 2)).replace('.', ',')
                    if indici_perf[row[0].value][1] > 0:
                        ws4[row[6].coordinate].font = Font(color='00B050')
                    else:
                        ws4[row[6].coordinate].font = Font(color='FF0000')
                    ws4[row[6].coordinate].alignment = Alignment(horizontal='center', vertical='center')
                    ws4[row[6].coordinate].border = Border(bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                    ws4[row[7].coordinate].value = "{0:.2f}%".format(round(indici_perf[row[0].value][2] * 100, 2)).replace('.', ',')
                    if indici_perf[row[0].value][2] > 0:
                        ws4[row[7].coordinate].font = Font(color='00B050')
                    else:
                        ws4[row[7].coordinate].font = Font(color='FF0000')
                    ws4[row[7].coordinate].alignment = Alignment(horizontal='center', vertical='center')
                    ws4[row[7].coordinate].border = Border(bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                    ws4[row[8].coordinate].value = "{0:.2f}%".format(round(indici_perf[row[0].value][3] * 100, 2)).replace('.', ',')
                    if indici_perf[row[0].value][3] > 0:
                        ws4[row[8].coordinate].font = Font(color='00B050')
                    else:
                        ws4[row[8].coordinate].font = Font(color='FF0000')
                    ws4[row[8].coordinate].alignment = Alignment(horizontal='center', vertical='center')
                    ws4[row[8].coordinate].border = Border(bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                else: # Riempi valute
                    ws4[row[4].coordinate].value = "{0:.2f}%".format(round(indici_perf[row[0].value][0] * 100, 2)).replace('.', ',')
                    if indici_perf[row[0].value][0] > 0:
                        ws4[row[4].coordinate].font = Font(color='00B050')
                    else:
                        ws4[row[4].coordinate].font = Font(color='FF0000')
                    ws4[row[4].coordinate].alignment = Alignment(horizontal='center', vertical='center')
                    ws4[row[4].coordinate].border = Border(bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                    ws4[row[5].coordinate].value = "{0:.2f}%".format(round(indici_perf[row[0].value][1] * 100, 2)).replace('.', ',')
                    if indici_perf[row[0].value][1] > 0:
                        ws4[row[5].coordinate].font = Font(color='00B050')
                    else:
                        ws4[row[5].coordinate].font = Font(color='FF0000')
                    ws4[row[5].coordinate].alignment = Alignment(horizontal='center', vertical='center')
                    ws4[row[5].coordinate].border = Border(bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                    ws4[row[7].coordinate].value = "{0:.2f}%".format(round(indici_perf[row[0].value][2] * 100, 2)).replace('.', ',')
                    if indici_perf[row[0].value][2] > 0:
                        ws4[row[7].coordinate].font = Font(color='00B050')
                    else:
                        ws4[row[7].coordinate].font = Font(color='FF0000')
                    ws4[row[7].coordinate].alignment = Alignment(horizontal='center', vertical='center')
                    ws4[row[7].coordinate].border = Border(bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                    ws4[row[8].coordinate].value = "{0:.2f}%".format(round(indici_perf[row[0].value][3] * 100, 2)).replace('.', ',')
                    if indici_perf[row[0].value][3] > 0:
                        ws4[row[8].coordinate].font = Font(color='00B050')
                    else:
                        ws4[row[8].coordinate].font = Font(color='FF0000')
                    ws4[row[8].coordinate].alignment = Alignment(horizontal='center', vertical='center')
                    ws4[row[8].coordinate].border = Border(bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                    ws4.merge_cells(start_row=row[5].row, end_row=row[5].row, start_column=row[5].column, end_column=row[6].column)
        # 'Text-box'
        self.text_box(ws4, 6, 40, 10, 15)

        # Logo
        self.logo(ws4, col=6, colOff=0.8, row=43, rowOff=-0.2)

    def analisi_indici_5(self):
        """
        Crea la quinta pagina.
        Formattazione e grafici.
        Aggiunge fogli Indici_giornalieri.
        """
        # Carica indici giornalieri
        indici_giornalieri = pd.read_excel(self.file_portafoglio, sheet_name='Indici_giornalieri', names=['Date', 'S&P 500', 'Date.1', 'USDEUR', 'Date.2', 'VIX', 'Date.3', 'EURO STOXX 50'])
        indici_giornalieri['Date'] = pd.to_datetime(indici_giornalieri['Date'], format = '%Y-%m-%d %H:%M:%S').dt.strftime('%m-%Y')
        indici_giornalieri['Date.1'] = pd.to_datetime(indici_giornalieri['Date.1'], format = '%Y-%m-%d %H:%M:%S').dt.strftime('%m-%Y')
        indici_giornalieri['Date.2'] = pd.to_datetime(indici_giornalieri['Date.2'], format = '%Y-%m-%d %H:%M:%S').dt.strftime('%m-%Y')
        indici_giornalieri['Date.3'] = pd.to_datetime(indici_giornalieri['Date.3'], format = '%Y-%m-%d %H:%M:%S').dt.strftime('%m-%Y')
        # print(indici_giornalieri)
        
        # Aggiungi foglio dati per creare i grafici
        ws_dati_indici = self.wb.create_sheet('Dati_indici')
        ws_dati_indici = self.wb['Dati_indici']
        self.wb.active = ws_dati_indici
        # Carica gli indici giornalieri
        for r in dataframe_to_rows(indici_giornalieri, index=True, header=True):
            ws_dati_indici.append(r)
        ws_dati_indici.delete_rows(2)
        ws_dati_indici.sheet_state = 'hidden'

        ws5 = self.wb.create_sheet('5.an_mkt_perf')
        ws5 = self.wb['5.an_mkt_perf']
        self.wb.active = ws5
        ws5['A1'] = '1. Analisi Di Mercato'
        ws5['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws5['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws5['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws5.merge_cells('A1:L4')

        # Aggiunta primo grafico
        chart = AreaChart()
        chart.title = "S&P 500"
        cats = Reference(ws_dati_indici, min_col=2, min_row=2, max_row=ws_dati_indici.max_row)
        data = Reference(ws_dati_indici, min_col=3, min_row=2, max_row=ws_dati_indici.max_row)
        chart.add_data(data, titles_from_data=False)
        series = chart.series[0]
        fill = PatternFillProperties()
        fill.foreground = ColorChoice(srgbClr='31869B')
        fill.background = ColorChoice(srgbClr='31869B')
        series.graphicalProperties.pattFill = fill
        chart.set_categories(cats)
        chart.height = 6.7
        chart.width = 10
        chart.legend = None
        ws5.add_chart(chart, "A6")

        # Aggiunta secondo grafico
        chart = AreaChart()
        chart.title = "Usd/Eur"
        cats = Reference(ws_dati_indici, min_col=4, min_row=2, max_row=ws_dati_indici.max_row)
        data = Reference(ws_dati_indici, min_col=5, min_row=2, max_row=ws_dati_indici.max_row)
        chart.add_data(data, titles_from_data=False)
        series = chart.series[0]
        fill = PatternFillProperties()
        fill.foreground = ColorChoice(srgbClr='31869B')
        fill.background = ColorChoice(srgbClr='31869B')
        series.graphicalProperties.pattFill = fill
        chart.set_categories(cats)
        chart.height = 6.7
        chart.width = 10
        chart.legend = None
        ws5.add_chart(chart, "G6")

        # Aggiunta terzo grafico
        chart = AreaChart()
        chart.title = "Vix"
        cats = Reference(ws_dati_indici, min_col=6, min_row=2, max_row=ws_dati_indici.max_row)
        data = Reference(ws_dati_indici, min_col=7, min_row=2, max_row=ws_dati_indici.max_row)
        chart.add_data(data, titles_from_data=False)
        series = chart.series[0]
        fill = PatternFillProperties()
        fill.foreground = ColorChoice(srgbClr='31869B')
        fill.background = ColorChoice(srgbClr='31869B')
        series.graphicalProperties.pattFill = fill
        chart.set_categories(cats)
        chart.height = 6.7
        chart.width = 10
        chart.legend = None
        ws5.add_chart(chart, "A20")

        # Aggiunta quarto grafico
        chart = AreaChart()
        chart.title = "Eurostoxx 50"
        cats = Reference(ws_dati_indici, min_col=8, min_row=2, max_row=ws_dati_indici.max_row)
        data = Reference(ws_dati_indici, min_col=9, min_row=2, max_row=ws_dati_indici.max_row)
        chart.add_data(data, titles_from_data=False)
        series = chart.series[0]
        fill = PatternFillProperties()
        fill.foreground = ColorChoice(srgbClr='31869B')
        fill.background = ColorChoice(srgbClr='31869B')
        series.graphicalProperties.pattFill = fill
        chart.set_categories(cats)
        chart.height = 6.7
        chart.width = 10
        chart.legend = None
        ws5.add_chart(chart, "G20")
        
        # Logo
        self.logo(ws5)

    def performance_6(self):
        """
        Crea la sesta pagina.
        Solo formattazione.
        """
        # 6.Performance
        ws6 = self.wb.create_sheet('6.perf')
        ws6 =  self.wb['6.perf']
        self.wb.active = ws6
        # Corpo
        ws6['A11'] = '2. Performance'
        ws6['A11'].alignment = Alignment(horizontal='center', vertical='center')
        ws6['A11'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws6['A11'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws6.merge_cells('A11:L14')

        # Logo
        self.logo(ws6)

    def andamento_7(self):
        """
        Crea la settima pagina.
        Formattazione e grafici.
        """
        # Carica performance benchmark
        benchmark = pd.read_excel(self.file_portafoglio, sheet_name='Benchmark', index_col=0, header=0)
        perf_bk_2007 = (float(benchmark.loc[self.t1, 'benchmark_2007']) - 100) / 100
        perf_bk_ytd = (float(benchmark.loc[self.t1, 'benchmark_2007']) - float(benchmark.loc[self.t0_ytd, 'benchmark_2007'])) / float(benchmark.loc[self.t0_ytd, 'benchmark_2007'])
        perf_month = (float(benchmark.loc[self.t1, 'benchmark_2007']) - float(benchmark.loc[self.t0_1m, 'benchmark_2007'])) / float(benchmark.loc[self.t0_1m, 'benchmark_2007'])
        # Carica portafoglio
        ptf = pd.read_excel(self.file_portafoglio, sheet_name='Portafoglio', index_col=0, header=0)
        # print(ptf['31/10/2015':'31/12/2021'])
        perf_ptf_2007 = (float(ptf.loc[self.t1, 'ptf_2007']) - 100) / 100
        perf_ptf_ytd = (float(ptf.loc[self.t1, 'ptf_2007']) - ptf.loc[self.t0_ytd, 'ptf_2007']) / ptf.loc[self.t0_ytd, 'ptf_2007']
        perf_ptf_month = (float(ptf.loc[self.t1, 'ptf_2007']) - ptf.loc[self.t0_1m, 'ptf_2007']) / ptf.loc[self.t0_1m, 'ptf_2007']


        ws7 = self.wb.create_sheet('7.andamento')
        ws7 = self.wb['7.andamento']
        self.wb.active = ws7

        # Titolo
        ws7['A1'] = 'Andamento Del Portafoglio'
        ws7['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws7['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws7['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws7.merge_cells('A1:L4')

        # Aggiunta primo grafico
        chart = BarChart()
        chart.type = 'col'
        chart.title = "DA INIZIO MANDATO"
        chart.y_axis.scaling.min = min(perf_bk_2007, perf_ptf_2007, 0) - 0.04
        chart.y_axis.scaling.max = max(perf_bk_2007, perf_ptf_2007) + 0.04
        ws7['A13'] = 'Ptf'
        ws7['A14'] = perf_ptf_2007
        ws7['A14'].number_format = FORMAT_PERCENTAGE_00
        ws7['B13'] = 'Benchmark'
        ws7['B14'] = perf_bk_2007
        ws7['B14'].number_format = FORMAT_PERCENTAGE_00
        data = Reference(ws7, min_col=1, max_col=2, min_row=13, max_row=14)
        chart.add_data(data, titles_from_data=True)
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
        s = chart.series[0]
        s.graphicalProperties.solidFill = 'FF6600'
        s1 = chart.series[1]
        s1.graphicalProperties.solidFill = '177245'
        chart.height = 10.3
        chart.width = 7.2
        chart.legend.position = 'b' # sposta la legenda in basso
        chart.y_axis.delete = True # togli l'asse y
        chart.x_axis.delete = True # togli l'asse x
        ws7.add_chart(chart, 'A6')

        # Aggiunta secondo grafico
        chart = BarChart()
        chart.type = 'col'
        chart.title = "YEAR TO DATE"
        chart.y_axis.scaling.min = min(perf_bk_ytd, perf_ptf_ytd, 0) - 0.03
        chart.y_axis.scaling.max = max(perf_bk_ytd, perf_ptf_ytd, 0.04) + 0.03
        ws7['F13'] = 'Ptf'
        ws7['F14'] = perf_ptf_ytd
        ws7['F14'].number_format = FORMAT_PERCENTAGE_00
        ws7['G13'] = 'Benchmark'
        ws7['G14'] = perf_bk_ytd
        ws7['G14'].number_format = FORMAT_PERCENTAGE_00
        ws7['H13'] = 'Target'
        ws7['H14'] = 0.04
        ws7['H14'].number_format = FORMAT_PERCENTAGE_00
        data = Reference(ws7, min_col=6, max_col=8, min_row=13, max_row=14)
        chart.add_data(data, titles_from_data=True)
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
        s = chart.series[0]
        s.graphicalProperties.solidFill = 'FF6600'
        s1 = chart.series[1]
        s1.graphicalProperties.solidFill = '177245'
        s2 = chart.series[2]
        s2.graphicalProperties.solidFill = 'C00000'
        chart.legend.position = 'b' # sposta la legenda in basso
        chart.y_axis.delete = True # togli l'asse y
        chart.x_axis.delete = True # togli l'asse x
        size = XDRPositiveSize2D(pixels_to_EMU(272.12598), pixels_to_EMU(389.29))
        maker = AnchorMarker(col=4, colOff=0, row=5, rowOff=0)
        ancoraggio = OneCellAnchor(_from=maker, ext=size)
        ws7.add_chart(chart)
        chart.anchor = ancoraggio

        # Aggiunta terzo grafico
        chart = BarChart()
        chart.type = 'col'
        chart.title = "MENSILE"
        chart.y_axis.scaling.min = -0.05
        chart.y_axis.scaling.max = 0.05
        ws7['J13'] = 'Ptf'
        ws7['J14'] = perf_ptf_month
        ws7['J14'].number_format = FORMAT_PERCENTAGE_00
        ws7['K13'] = 'Benchmark'
        ws7['K14'] = perf_month
        ws7['K14'].number_format = FORMAT_PERCENTAGE_00
        data = Reference(ws7, min_col=10, max_col=11, min_row=13, max_row=14)
        chart.add_data(data, titles_from_data=True)
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
        s = chart.series[0]
        s.graphicalProperties.solidFill = 'FF6600'
        s1 = chart.series[1]
        s1.graphicalProperties.solidFill = '177245'
        chart.legend.position = 'b' # sposta la legenda in basso
        chart.y_axis.delete = True # togli l'asse y
        chart.x_axis.delete = True # togli l'asse x
        size = XDRPositiveSize2D(pixels_to_EMU(272.12598), pixels_to_EMU(389.29))
        maker = AnchorMarker(col=8, colOff=0, row=5, rowOff=0)
        ancoraggio = OneCellAnchor(_from=maker, ext=size)
        ws7.add_chart(chart)
        chart.anchor = ancoraggio

        # Corpo
        ws7['B27'] = '* il rendimento del P. in Strumenti è al netto della commissione di consulenza, degli eventuali prelievi e conferimenti'
        ws7['B27'].font = Font(name='Times New Roman', size=11, bold=False, color='31869B')

        # Logo
        self.logo(ws7)

    def caricamento_dati(self):
        """
        Crea l'ottava pagina.
        Formattazione e grafici.
        Aggiunta fogli Cono, Portafoglio e Benchmark.
        """
        # Carica dati per i coni
        coni = pd.read_excel(self.file_portafoglio, sheet_name='Cono', index_col=0, header=0)
        coni.index = pd.to_datetime(coni.index, format='%Y-%m-%d %H:%M:%S').strftime('%m-%Y')
        # Carica gli scenari per i coni
        ws_dati_cono = self.wb.create_sheet('Dati_cono')
        ws_dati_cono = self.wb['Dati_cono']
        self.wb.active = ws_dati_cono
        for r in dataframe_to_rows(coni, index=True, header=True):
            ws_dati_cono.append(r)
        ws_dati_cono.delete_rows(2)
        ws_dati_cono.sheet_state = 'hidden'
        # Carica rendimento portafoglio
        perf_pf = pd.read_excel(self.file_portafoglio, sheet_name='Portafoglio', index_col=0, header=0)
        perf_pf.index = pd.to_datetime(perf_pf.index, format='%Y-%m-%d %H:%M:%S').strftime('%m-%Y')
        # Carica perf ptf
        ws_dati_pf = self.wb.create_sheet('Dati_pf')
        ws_dati_pf = self.wb['Dati_pf']
        self.wb.active = ws_dati_pf
        for r in dataframe_to_rows(perf_pf, index=True, header=True):
            ws_dati_pf.append(r)
        ws_dati_pf.delete_rows(2)
        ws_dati_pf.sheet_state = 'hidden'
        # Carica performance benchmark
        perf_bk = pd.read_excel(self.file_portafoglio, sheet_name='Benchmark', index_col=0, header=0)
        perf_bk.index = pd.to_datetime(perf_bk.index, format='%Y-%m-%d %H:%M:%S').strftime('%m-%Y')
        # Carica perf bk
        ws_dati_bk = self.wb.create_sheet('Dati_bk')
        ws_dati_bk = self.wb['Dati_bk']
        self.wb.active = ws_dati_bk
        for r in dataframe_to_rows(perf_bk, index=True, header=True):
            ws_dati_bk.append(r)
        ws_dati_bk.delete_rows(2)
        ws_dati_bk.sheet_state = 'hidden'

        # # Titolo
        # ws8 = self.wb.create_sheet('8.cono_1')
        # ws8 = self.wb['8.cono_1']
        # self.wb.active = ws8
        # ws8['A1'] = 'Cono Delle Probabilità'
        # ws8['A1'].alignment = Alignment(horizontal='center', vertical='center')
        # ws8['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        # ws8['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        # ws8.merge_cells('A1:L4')

        # # Corpo
        # ws8['E6'] = 'Benchmark 2007'
        # ws8['E6'].alignment = Alignment(horizontal='center', vertical='center')
        # ws8['E6'].font = Font(name='Times New Roman', size=14, bold=True, color='31869B')
        # ws8.merge_cells('E6:H6')

        # # Aggiunta grafico
        # chart = LineChart()
        # # Riga corrispondente al mese di t1 nei tre nuovi fogli
        # riga_mese_t1 = lambda x : next(x[row[0].coordinate].row for row in x.iter_rows(min_col=0, max_col=0) if x[row[0].coordinate].value == self.t1.strftime('%m-%Y'))
        # ws_dati_cono_max_row = riga_mese_t1(ws_dati_cono)
        # # print(ws_dati_cono_max_row)
        # # for row in ws_dati_cono.iter_rows(min_col=0, max_col=0):
        # #     if ws_dati_cono[row[0].coordinate].value == self.t1.strftime('%m-%Y'):
        # #         # print(f"La riga del mese t1 è : {ws_dati_cono[row[0].coordinate].row}.")
        # #         ws_dati_cono_max_row = ws_dati_cono[row[0].coordinate].row
        # data = Reference(ws_dati_cono, min_col=3, max_col=5, min_row=1, max_row=ws_dati_cono_max_row)
        # chart.add_data(data, titles_from_data='False')
        # ws_dati_bk_max_row = riga_mese_t1(ws_dati_bk)
        # # print(ws_dati_bk_max_row)
        # # for row in ws_dati_bk.iter_rows(min_col=0, max_col=0):
        # #     if ws_dati_bk[row[0].coordinate].value == self.t1.strftime('%m-%Y'):
        # #         print(f"La riga del mese t1 è : {ws_dati_bk[row[0].coordinate].row}.")
        # #         ws_dati_bk_max_row = ws_dati_bk[row[0].coordinate].row
        # data = Reference(ws_dati_bk, min_col=25, min_row=3, max_row=ws_dati_bk_max_row)
        # chart.add_data(data, titles_from_data='False')
        # ws_dati_pf_max_row = riga_mese_t1(ws_dati_pf)
        # # print(ws_dati_pf_max_row)
        # # for row in ws_dati_pf.iter_rows(min_col=0, max_col=0):
        # #     if ws_dati_pf[row[0].coordinate].value == self.t1.strftime('%m-%Y'):
        # #         print(f"La riga del mese t1 è : {ws_dati_pf[row[0].coordinate].row}.")
        # #         ws_dati_pf_max_row = ws_dati_pf[row[0].coordinate].row
        # data = Reference(ws_dati_pf, min_col=4, min_row=1, max_row=ws_dati_pf_max_row)
        # chart.add_data(data, titles_from_data='False')

        # s0 = chart.series[0]
        # s0.graphicalProperties.line.solidFill = '0000FF'
        # s0.graphicalProperties.line.width = 12700
        # s0.dLbls = DataLabelList()
        # dl = DataLabel(dLblPos='t', idx=ws_dati_cono_max_row-2, numFmt='0.00', showVal=True)
        # s0.dLbls.dLbl.append(dl)
        # s1 = chart.series[1]
        # s1.graphicalProperties.line.solidFill = 'FF00FF'
        # s1.graphicalProperties.line.width = 12700
        # s1.dLbls = DataLabelList()
        # dl = DataLabel(dLblPos='t', idx=ws_dati_cono_max_row-2, numFmt='0.00', showVal=True)
        # s1.dLbls.dLbl.append(dl)
        # s2 = chart.series[2]
        # s2.graphicalProperties.line.solidFill = '000080'
        # s2.graphicalProperties.line.width = 12700
        # s2.dLbls = DataLabelList()
        # dl = DataLabel(dLblPos='t', idx=ws_dati_cono_max_row-2, numFmt='0.00', showVal=True)
        # s2.dLbls.dLbl.append(dl)
        # s3 = chart.series[3]
        # s3.graphicalProperties.line.solidFill = '177245'
        # s3.graphicalProperties.line.width = 25400
        # s3.dLbls = DataLabelList()
        # dl = DataLabel(dLblPos='b', idx=ws_dati_bk_max_row-4, numFmt='0.00', showVal=True)
        # s3.dLbls.dLbl.append(dl)
        # s4 = chart.series[4]
        # s4.graphicalProperties.line.solidFill = 'FF0000'
        # s4.graphicalProperties.line.width = 25400
        # s4.dLbls = DataLabelList()
        # dl = DataLabel(dLblPos='b', idx=ws_dati_pf_max_row-2, numFmt='0.00', showVal=True)
        # s4.dLbls.dLbl.append(dl)

        # dates = Reference(ws_dati_cono, min_col=1, max_col=1, min_row=2, max_row=ws_dati_cono_max_row)
        # chart.set_categories(dates)
        # chart.legend.layout = Layout(manualLayout=ManualLayout(h=1))
        # size = XDRPositiveSize2D(pixels_to_EMU(812.598), pixels_to_EMU(453.54))
        # cellw = lambda x: cm_to_EMU((x * (18.65-1.71))/10)
        # coloffset2 = cellw(0.1)
        # maker = AnchorMarker(col=0, colOff=coloffset2, row=6, rowOff=0)
        # ancoraggio = OneCellAnchor(_from=maker, ext=size)
        # ws8.add_chart(chart)
        # chart.anchor = ancoraggio
        # chart.y_axis.scaling.min = 80 # valore minimo asse y
        
        # # Logo
        # self.logo(ws8)

    def cono_8(self):
        """
        Crea la nona pagina.
        Formattazione e grafici.
        Riattiva fogli ws_dati_bk, ws_dati_cono e ws_dati_pf.
        """
        # Riattiva scenari coni
        ws_dati_cono = self.wb['Dati_cono']
        # Riattiva performance ptf
        ws_dati_pf = self.wb['Dati_pf']
        # Riattiva performance bk
        ws_dati_bk = self.wb['Dati_bk']

        ws9 = self.wb.create_sheet('8.cono_1')
        ws9 = self.wb['8.cono_1']
        self.wb.active = ws9

        # Titolo
        ws9['A1'] = 'Cono Delle Probabilità'
        ws9['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws9['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws9['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws9.merge_cells('A1:L4')

        # Corpo
        ws9['E6'] = 'Benchmark 2016'
        ws9['E6'].alignment = Alignment(horizontal='center', vertical='center')
        ws9['E6'].font = Font(name='Times New Roman', size=14, bold=True, color='31869B')
        ws9.merge_cells('E6:H6')

        # Aggiunta grafico
        chart = LineChart()
        riga_mese_t1 = lambda x : next(x[row[0].coordinate].row for row in x.iter_rows(min_col=0, max_col=0) if x[row[0].coordinate].value == self.t1.strftime('%m-%Y'))
        ws_dati_cono_max_row = riga_mese_t1(ws_dati_cono)
        # for row in ws_dati_cono.iter_rows(min_col=0, max_col=0):
        #     if ws_dati_cono[row[0].coordinate].value == self.t1.strftime('%m-%Y'):
        #         print(f"La riga del mese t1 è : {ws_dati_cono[row[0].coordinate].row}.")
        #         ws_dati_cono_max_row = ws_dati_cono[row[0].coordinate].row
        data = Reference(ws_dati_cono, min_col=7, max_col=9, min_row=109, max_row=ws_dati_cono_max_row) # hard coding
        chart.add_data(data, titles_from_data='False')
        # for row in ws_dati_bk.iter_rows(min_col=0, max_col=0):
        #     if ws_dati_bk[row[0].coordinate].value == self.t1.strftime('%m-%Y'):
        #         print(f"La riga del mese t1 è : {ws_dati_bk[row[0].coordinate].row}.")
        #         ws_dati_bk_max_row = ws_dati_bk[row[0].coordinate].row
        ws_dati_bk_max_row = riga_mese_t1(ws_dati_bk)
        data = Reference(ws_dati_bk, min_col=26, min_row=111, max_row=ws_dati_bk_max_row) # hard coding
        chart.add_data(data, titles_from_data='False')
        # for row in ws_dati_pf.iter_rows(min_col=0, max_col=0):
        #     if ws_dati_pf[row[0].coordinate].value == self.t1.strftime('%m-%Y'):
        #         print(f"La riga del mese t1 è : {ws_dati_pf[row[0].coordinate].row}.")
        #         ws_dati_pf_max_row = ws_dati_pf[row[0].coordinate].row
        ws_dati_pf_max_row = riga_mese_t1(ws_dati_pf)
        data = Reference(ws_dati_pf, min_col=5, min_row=109, max_row=ws_dati_pf_max_row) # hard coding
        chart.add_data(data, titles_from_data='False')

        s0 = chart.series[0]
        s0.graphicalProperties.line.solidFill = '0000FF'
        s0.graphicalProperties.line.width = 12700
        s0.dLbls = DataLabelList()
        dl = DataLabel(dLblPos='t', idx=ws_dati_cono_max_row-110, numFmt='0.00', showVal=True)
        s0.dLbls.dLbl.append(dl)
        s1 = chart.series[1]
        s1.graphicalProperties.line.solidFill = 'FF00FF'
        s1.graphicalProperties.line.width = 12700
        s1.dLbls = DataLabelList()
        dl = DataLabel(dLblPos='t', idx=ws_dati_cono_max_row-110, numFmt='0.00', showVal=True)
        s1.dLbls.dLbl.append(dl)
        s2 = chart.series[2]
        s2.graphicalProperties.line.solidFill = '000080'
        s2.graphicalProperties.line.width = 12700
        s2.dLbls = DataLabelList()
        dl = DataLabel(dLblPos='t', idx=ws_dati_cono_max_row-110, numFmt='0.00', showVal=True)
        s2.dLbls.dLbl.append(dl)
        s3 = chart.series[3]
        s3.graphicalProperties.line.solidFill = '177245'
        s3.graphicalProperties.line.width = 25400
        s3.dLbls = DataLabelList()
        dl = DataLabel(dLblPos='b', idx=ws_dati_bk_max_row-112, numFmt='0.00', showVal=True)
        s3.dLbls.dLbl.append(dl)
        s4 = chart.series[4]
        s4.graphicalProperties.line.solidFill = 'FF0000'
        s4.graphicalProperties.line.width = 25400
        s4.dLbls = DataLabelList()
        dl = DataLabel(dLblPos='t', idx=ws_dati_pf_max_row-110, numFmt='0.00', showVal=True)
        s4.dLbls.dLbl.append(dl)

        dates = Reference(ws_dati_cono, min_col=1, max_col=1, min_row=110, max_row=ws_dati_cono_max_row)
        chart.set_categories(dates)
        chart.legend.layout = Layout(manualLayout=ManualLayout(h=1))
        size = XDRPositiveSize2D(pixels_to_EMU(812.598), pixels_to_EMU(453.54))
        cellw = lambda x: cm_to_EMU((x * (18.65-1.71))/10)
        coloffset2 = cellw(0.1)
        maker = AnchorMarker(col=0, colOff=coloffset2, row=6, rowOff=0)
        ancoraggio = OneCellAnchor(_from=maker, ext=size)
        ws9.add_chart(chart)
        chart.anchor = ancoraggio
        chart.y_axis.scaling.min = 95 # valore minimo asse y
        ws9.row_dimensions[5].height = 11.25

        # Logo
        self.logo(ws9)

    def cono_9(self):
        """
        Crea la decima pagina.
        Formattazione e grafici.
        Riattiva fogli ws_dati_bk, ws_dati_cono e ws_dati_pf.
        In partenza 31/01/2022 si vedevano solo due dati, e mi è stato chiesto di sviluppare il cono dei tre casi probabilistici
        più a lungo, proiettandoli nel futuro. Definisco una costante chiamata SFASAMENTO_DATI che aggiunge n periodi alle serie
        storiche di quei tre casi. Nel futuro questa variabile è da togliere.
        """
        SFASAMENTO_DATI = 8
        # Riattiva scenari coni
        ws_dati_cono = self.wb['Dati_cono']
        # Riattiva performance ptf
        ws_dati_pf = self.wb['Dati_pf']
        # Riattiva performance bk
        ws_dati_bk = self.wb['Dati_bk']

        ws9_2 = self.wb.create_sheet('9.cono_2')
        ws9_2 = self.wb['9.cono_2']
        self.wb.active = ws9_2

        # Titolo
        ws9_2['A1'] = 'Cono Delle Probabilità'
        ws9_2['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws9_2['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws9_2['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws9_2.merge_cells('A1:L4')

        # Corpo
        ws9_2['E6'] = 'Benchmark 2022'
        ws9_2['E6'].alignment = Alignment(horizontal='center', vertical='center')
        ws9_2['E6'].font = Font(name='Times New Roman', size=14, bold=True, color='31869B')
        ws9_2.merge_cells('E6:H6')

        # Aggiunta grafico
        chart = LineChart()
        riga_mese_t1 = lambda x : next(x[row[0].coordinate].row for row in x.iter_rows(min_col=0, max_col=0) if x[row[0].coordinate].value == self.t1.strftime('%m-%Y'))
        ws_dati_cono_max_row = riga_mese_t1(ws_dati_cono)
        # for row in ws_dati_cono.iter_rows(min_col=0, max_col=0):
        #     if ws_dati_cono[row[0].coordinate].value == self.t1.strftime('%m-%Y'):
        #         print(f"La riga del mese t1 è : {ws_dati_cono[row[0].coordinate].row}.")
        #         ws_dati_cono_max_row = ws_dati_cono[row[0].coordinate].row
        data = Reference(ws_dati_cono, min_col=11, max_col=13, min_row=180, max_row=ws_dati_cono_max_row+SFASAMENTO_DATI) # hard coding
        chart.add_data(data, titles_from_data='False')
        # for row in ws_dati_bk.iter_rows(min_col=0, max_col=0):
        #     if ws_dati_bk[row[0].coordinate].value == self.t1.strftime('%m-%Y'):
        #         print(f"La riga del mese t1 è : {ws_dati_bk[row[0].coordinate].row}.")
        #         ws_dati_bk_max_row = ws_dati_bk[row[0].coordinate].row
        ws_dati_bk_max_row = riga_mese_t1(ws_dati_bk)
        data = Reference(ws_dati_bk, min_col=27, min_row=182, max_row=ws_dati_bk_max_row) # hard coding
        chart.add_data(data, titles_from_data='False')
        # for row in ws_dati_pf.iter_rows(min_col=0, max_col=0):
        #     if ws_dati_pf[row[0].coordinate].value == self.t1.strftime('%m-%Y'):
        #         print(f"La riga del mese t1 è : {ws_dati_pf[row[0].coordinate].row}.")
        #         ws_dati_pf_max_row = ws_dati_pf[row[0].coordinate].row
        ws_dati_pf_max_row = riga_mese_t1(ws_dati_pf)
        data = Reference(ws_dati_pf, min_col=6, min_row=180, max_row=ws_dati_pf_max_row) # hard coding
        chart.add_data(data, titles_from_data='False')

        s0 = chart.series[0]
        s0.graphicalProperties.line.solidFill = '0000FF'
        s0.graphicalProperties.line.width = 12700
        s0.dLbls = DataLabelList()
        dl = DataLabel(dLblPos='t', idx=ws_dati_cono_max_row+SFASAMENTO_DATI-181, numFmt='0.00', showVal=True)
        s0.dLbls.dLbl.append(dl)
        s1 = chart.series[1]
        s1.graphicalProperties.line.solidFill = 'FF00FF'
        s1.graphicalProperties.line.width = 12700
        s1.dLbls = DataLabelList()
        dl = DataLabel(dLblPos='t', idx=ws_dati_cono_max_row+SFASAMENTO_DATI-181, numFmt='0.00', showVal=True)
        s1.dLbls.dLbl.append(dl)
        s2 = chart.series[2]
        s2.graphicalProperties.line.solidFill = '000080'
        s2.graphicalProperties.line.width = 12700
        s2.dLbls = DataLabelList()
        dl = DataLabel(dLblPos='t', idx=ws_dati_cono_max_row+SFASAMENTO_DATI-181, numFmt='0.00', showVal=True)
        s2.dLbls.dLbl.append(dl)
        s3 = chart.series[3]
        s3.graphicalProperties.line.solidFill = '177245'
        s3.graphicalProperties.line.width = 25400
        s3.dLbls = DataLabelList()
        dl = DataLabel(dLblPos='b', idx=ws_dati_bk_max_row-183, numFmt='0.00', showVal=True)
        s3.dLbls.dLbl.append(dl)
        s4 = chart.series[4]
        s4.graphicalProperties.line.solidFill = 'FF0000'
        s4.graphicalProperties.line.width = 25400
        s4.dLbls = DataLabelList()
        dl = DataLabel(dLblPos='t', idx=ws_dati_pf_max_row-181, numFmt='0.00', showVal=True)
        s4.dLbls.dLbl.append(dl)

        dates = Reference(ws_dati_cono, min_col=1, max_col=1, min_row=181, max_row=ws_dati_cono_max_row+SFASAMENTO_DATI)
        chart.set_categories(dates)
        chart.legend.layout = Layout(manualLayout=ManualLayout(h=1))
        size = XDRPositiveSize2D(pixels_to_EMU(812.598), pixels_to_EMU(453.54))
        cellw = lambda x: cm_to_EMU((x * (18.65-1.71))/10)
        coloffset2 = cellw(0.1)
        maker = AnchorMarker(col=0, colOff=coloffset2, row=6, rowOff=0)
        ancoraggio = OneCellAnchor(_from=maker, ext=size)
        ws9_2.add_chart(chart)
        chart.anchor = ancoraggio
        chart.y_axis.scaling.min = 90 # valore minimo asse y
        ws9_2.row_dimensions[5].height = 11.25

        # Logo
        self.logo(ws9_2)

    def nuovo_bk_10(self):
        """
        Crea la decima pagina.
        Solo formattazione.
        """
        ws10 = self.wb.create_sheet('10.nuovo_bk')
        ws10 = self.wb['10.nuovo_bk']
        self.wb.active = ws10
        # Titolo
        ws10['A1'] = 'Nuovo Benchmark'
        ws10['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws10['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws10['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws10.merge_cells('A1:L4')
        # Corpo
        body_10_1 = ['Indice MTS BOT', 'Bloomberg Euro Government', 'Bloomberg Euro Corporate Index', 'Bloomberg Pan-European High Yield Index',
            'Bloomberg Global Aggregate Index', 'MSCI Europa', 'MSCI USA', 'MSCI Pacifico', 'MSCI Emerging Market Free', 
            'HFRX Absolute Return', 'Bloomberg Commodity Index']
        body_10_2 = ['20,00%', '13,15%', '4,66%', '8,05%', '10,09%', '7,27%', '14,55%', '4,56%', '11,04%', '2,85%', '3,78%']
        for row in ws10.iter_rows(min_row=8, max_row=8+len(body_10_1)-1, min_col=4, max_col=9):
            ws10[row[0].coordinate].value = body_10_1[0]
            del body_10_1[0]
            ws10[row[0].coordinate].font = Font(name='Calibri', size=11, bold=True, italic=True, color='000000') 
            ws10[row[0].coordinate].border = Border(bottom=Side(border_style='mediumDashDot', color='31869B'))
            ws10.merge_cells(start_row=row[0].row, end_row=row[0].row, start_column=row[0].column, end_column=row[4].column)
            ws10[row[5].coordinate].value = body_10_2[0]
            del body_10_2[0]
            ws10[row[5].coordinate].font = Font(name='Calibri', size=11, bold=True, italic=True, color='000000') 
            ws10[row[5].coordinate].alignment = Alignment(horizontal='right')
            ws10[row[5].coordinate].border = Border(bottom=Side(border_style='mediumDashDot', color='31869B'))
        ws10['C21'] = '           Benchmark costruito seguendo la composizione del portafoglio al 31/12/2021'
        ws10['C21'].font = Font(name='Calibri', size=11, bold=True, italic=True, color='31869B') 
        # Logo
        self.logo(ws10)

    def performance_11(self):
        """
        Crea l'undicesima pagina.
        Formattazione e tabella.
        """
        # Carica portafoglio
        portfolio = pd.read_excel(self.file_portafoglio, sheet_name='Portfolio', header=1)
        # Carica performance posizioni --- dipende da portfolio
        # rend_pf = pd.read_excel(self.file_portafoglio, sheet_name='Rendimento', index_col=0, header=0)
        delta = pd.read_excel(self.file_portafoglio, sheet_name='Delta', index_col=0, header=0)
        # print(delta)
        ws11 = self.wb.create_sheet('11.perf_mese')
        ws11 = self.wb['11.perf_mese']
        self.wb.active = ws11
        # Titolo
        ws11['A1'] = 'Performance Del Mese'
        ws11['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws11['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws11['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws11.merge_cells('A1:L4')
        # Colonne
        header_11 = ['', '', 'Totale ' + self.mesi_dict[self.t0_1m.month], 'Totale ' + self.mesi_dict[self.t1.month], 'Δ', 'Δ%', 'Δ% YTD']
        len_header_11 = len(header_11)
        for column in ws11.iter_cols(min_col=1, max_col=1+len_header_11-1, min_row=6, max_row=6):
            ws11[column[0].coordinate].value = header_11[0]
            del header_11[0]
            ws11[column[0].coordinate].font = Font(name='Times New Roman', size=10, color='FFFFFF')
            ws11[column[0].coordinate].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws11[column[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='31869B')
        # Indice
        intermediari = portfolio.loc[:, 'INTERMEDIARIO'].unique()
        intermediari = list(intermediari)
        intermediari.insert(len(intermediari), 'Interessi Anthares')
        intermediari.insert(len(intermediari), 'Totale Complessivo')
        intermediari.remove('Mediolanum') # Mediolanum non viene considerata.
        len_int = len(intermediari)
        # Corpo tabella
        for row in ws11.iter_rows(min_col=1, max_col=1+len_header_11-1, min_row=7, max_row=7+len_int-1):
            # if intermediari[0] == 'Mediolanum':
            #     intermediari = np.delete(intermediari, 0)
            #     continue
            ws11[row[0].coordinate].value = intermediari[0]
            del intermediari[0]
            #intermediari = np.delete(intermediari, 0)
            ws11.column_dimensions[row[2].column_letter].width = 10.5
            ws11.column_dimensions[row[3].column_letter].width = 10.5
            ws11.row_dimensions[row[0].row].height = 25.50
            ws11[row[0].coordinate].font = Font(name='Times New Roman', size=9, color='000000')
            ws11[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            ws11[row[0].coordinate].border = Border(left=Side(border_style='dashed', color='31869B'), bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'))
            ws11.merge_cells(start_column=row[0].column, end_column=row[1].column, start_row=row[0].row, end_row=row[0].row)
            # Corpo tabella
            ws11[row[2].coordinate].border = Border(right=Side(border_style='dashed', color='31869B'), bottom=Side(border_style='dashed', color='31869B'))
            ws11[row[2].coordinate].number_format = '€ #,0'
            ws11[row[2].coordinate].font = Font(name='Times New Roman', size=9)
            ws11[row[2].coordinate].alignment = Alignment(horizontal='center', vertical='center')
            ws11[row[3].coordinate].border = Border(right=Side(border_style='dashed', color='31869B'), bottom=Side(border_style='dashed', color='31869B'))
            ws11[row[3].coordinate].number_format = '€ #,0'
            ws11[row[3].coordinate].font = Font(name='Times New Roman', size=9)
            ws11[row[3].coordinate].alignment = Alignment(horizontal='center', vertical='center')
            ws11[row[4].coordinate].border = Border(right=Side(border_style='dashed', color='31869B'), bottom=Side(border_style='dashed', color='31869B'))
            ws11[row[4].coordinate].number_format = '€ #,0'
            ws11[row[4].coordinate].font = Font(name='Times New Roman', size=9)
            ws11[row[4].coordinate].alignment = Alignment(horizontal='center', vertical='center')
            ws11[row[5].coordinate].border = Border(right=Side(border_style='dashed', color='31869B'), bottom=Side(border_style='dashed', color='31869B'))
            ws11[row[5].coordinate].number_format = FORMAT_PERCENTAGE_00
            ws11[row[5].coordinate].font = Font(name='Times New Roman', size=9)
            ws11[row[5].coordinate].alignment = Alignment(horizontal='center', vertical='center')
            ws11[row[6].coordinate].border = Border(right=Side(border_style='dashed', color='31869B'), bottom=Side(border_style='dashed', color='31869B'))
            ws11[row[6].coordinate].number_format = FORMAT_PERCENTAGE_00
            ws11[row[6].coordinate].font = Font(name='Times New Roman', size=9)
            ws11[row[6].coordinate].alignment = Alignment(horizontal='center', vertical='center')
            
            if ws11[row[0].coordinate].row != 7+len_int-1:
                ws11[row[2].coordinate].value = delta.loc[ws11[row[0].coordinate].value, 'Totale mese passato']
                ws11[row[3].coordinate].value = delta.loc[ws11[row[0].coordinate].value, 'Totale mese corrente']
                ws11[row[4].coordinate].value = delta.loc[ws11[row[0].coordinate].value, 'Δ']
                ws11[row[5].coordinate].value = delta.loc[ws11[row[0].coordinate].value, 'Δ%']
                ws11[row[6].coordinate].value = delta.loc[ws11[row[0].coordinate].value, 'Δ% YTD']
            else:
                for _ in range(0, len_header_11):
                    ws11[row[_].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
                    ws11[row[_].coordinate].border = Border(top=Side(border_style='medium', color='31869B'), right=Side(border_style='medium', color='31869B'), left=Side(border_style='medium', color='31869B'), bottom=Side(border_style='medium', color='31869B'))
                    ws11[row[4].coordinate].value = 0
                    for __ in range(1, len_int): # Somma per tutti i valori nella colonna delta
                        ws11[row[4].coordinate].value = ws11[row[4].coordinate].value + ws11[row[4].coordinate].offset(row=-__).value
            
            
            # if ws11[row[0].coordinate].value == 'Banca Generali' or ws11[row[0].coordinate].value == 'Cassa Lombarda' or ws11[row[0].coordinate].value == 'Cassa Lombarda Trust' or ws11[row[0].coordinate].value == 'Corner' or ws11[row[0].coordinate].value == 'JPMorgan' or ws11[row[0].coordinate].value == 'Mediobanca' or ws11[row[0].coordinate].value == 'Fineco':
            #     ws11[row[2].coordinate].value = portfolio.loc[portfolio['INTERMEDIARIO']==ws11[row[0].coordinate].value, 'TOTALE t0'].sum()
            #     ws11[row[3].coordinate].value = portfolio.loc[portfolio['INTERMEDIARIO']==ws11[row[0].coordinate].value, 'TOTALE t1'].sum()
            #     ws11[row[4].coordinate].value = ws11[row[3].coordinate].value - ws11[row[2].coordinate].value
            #     # validazione delta
            #     if round(ws11[row[4].coordinate].value) != round(delta.loc[ws11[row[0].coordinate].value, 'Δ']):
            #         print(f'il valore di {ws11[row[0].coordinate].value} non corrisponde')
            #     ws11[row[5].coordinate].value = ws11[row[4].coordinate].value / ws11[row[2].coordinate].value
            #     #ws11[row[6].coordinate].value = (rend_pf.loc[self.t1, ws11[row[0].coordinate].value] - rend_pf.loc[self.t0_ytd, ws11[row[0].coordinate].value]) / rend_pf.loc[self.t0_ytd, ws11[row[0].coordinate].value]
            #     try:
            #         ws11[row[6].coordinate].value = delta.loc[ws11[row[0].coordinate].value, 'Δ% YTD']
            #     except KeyError:
            #         pass
            # elif ws11[row[0].coordinate].value == 'Credito Artigiano Artes' or ws11[row[0].coordinate].value == 'Credito Artigiano B.N.' or ws11[row[0].coordinate].value == 'Mediolanum':
            #     ws11[row[2].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws11[row[0].coordinate].value) & (portfolio['CATEGORIA']!='CASH'), 'TOTALE t0'].sum()
            #     ws11[row[3].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws11[row[0].coordinate].value) & (portfolio['CATEGORIA']!='CASH'), 'TOTALE t1'].sum()
            #     ws11[row[4].coordinate].value = ws11[row[3].coordinate].value - ws11[row[2].coordinate].value
            #     # validazione delta
            #     if round(ws11[row[4].coordinate].value) != round(delta.loc[ws11[row[0].coordinate].value, 'Δ']):
            #         print(f'il valore di {ws11[row[0].coordinate].value} non corrisponde')
            #     ws11[row[5].coordinate].value = ws11[row[4].coordinate].value / ws11[row[2].coordinate].value
            #     #ws11[row[6].coordinate].value = (rend_pf.loc[self.t1, ws11[row[0].coordinate].value] - rend_pf.loc[self.t0_ytd, ws11[row[0].coordinate].value]) / rend_pf.loc[self.t0_ytd, ws11[row[0].coordinate].value]
            #     try:
            #         ws11[row[6].coordinate].value = delta.loc[ws11[row[0].coordinate].value, 'Δ% YTD']
            #     except KeyError:
            #         pass
            # elif ws11[row[0].coordinate].value == 'Ubi Ita':
            #     ws11[row[2].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws11[row[0].coordinate].value) & (portfolio['PRODOTTO']!='RAGANELLA') & (portfolio['PRODOTTO']!='ANTHARES SPA'), 'TOTALE t0'].sum()
            #     ws11[row[3].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws11[row[0].coordinate].value) & (portfolio['PRODOTTO']!='RAGANELLA') & (portfolio['PRODOTTO']!='ANTHARES SPA'), 'TOTALE t1'].sum()
            #     ws11[row[4].coordinate].value = ws11[row[3].coordinate].value - ws11[row[2].coordinate].value
            #     ws11[row[5].coordinate].value = ws11[row[4].coordinate].value / ws11[row[2].coordinate].value
            #     #ws11[row[6].coordinate].value = (rend_pf.loc[self.t1, ws11[row[0].coordinate].value] - rend_pf.loc[self.t0_ytd, ws11[row[0].coordinate].value]) / rend_pf.loc[self.t0_ytd, ws11[row[0].coordinate].value]
            #     ws11[row[6].coordinate].value = delta.loc[ws11[row[0].coordinate].value, 'Δ% YTD']
            #     try:
            #         ws11[row[6].coordinate].value = delta.loc[ws11[row[0].coordinate].value, 'Δ% YTD']
            #     except KeyError:
            #         pass
            # elif ws11[row[0].coordinate].value == 'Interessi Anthares':
            #     ws11[row[4].coordinate].value = float(portfolio.loc[portfolio['PRODOTTO']=='ANTHARES SPA', 'TOTALE t1'].values[0]) * 0.005
            # else:
            #     pass
            # if ws11[row[0].coordinate].row == 7+len_int-1:
            #     for _ in range(0, len_header_11):
            #         ws11[row[_].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
            #         ws11[row[_].coordinate].border = Border(top=Side(border_style='medium', color='31869B'), right=Side(border_style='medium', color='31869B'), left=Side(border_style='medium', color='31869B'), bottom=Side(border_style='medium', color='31869B'))
            #         ws11[row[4].coordinate].value = 0
            #         for __ in range(1, len_int): # Somma per tutti i valori nella colonna delta
            #             ws11[row[4].coordinate].value = ws11[row[4].coordinate].value + ws11[row[4].coordinate].offset(row=-__).value

        # 'Text box'
        self.text_box(ws11, 6, 18, 8, 12)

        # Logo
        self.logo(ws11, colOff=0, row=26)

    def prezzi_12(self):
        """
        Crea la dodicesima pagina.
        Formattazione e tabella.
        """
        # Carica portafoglio
        portfolio = pd.read_excel(self.file_portafoglio, sheet_name='Portfolio', header=1)
        # 12.Prezzi 1
        ws12 = self.wb.create_sheet('12.prezzi_1')
        ws12 = self.wb['12.prezzi_1']
        self.wb.active = ws12
        # Titolo
        ws12['A1'] = 'Corner'
        ws12['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws12['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws12['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws12.merge_cells('A1:L4')
        # Creazione tabella
        header_12 = ['Nome', '', '', '', '', 'Valuta', 'Quantità', 'Prezzo di carico', 'Prezzo attuale', '∆ prezzo', 'Ctv', '']
        for column in ws12.iter_cols(min_col=1, max_col=12, min_row=6, max_row=6):
            ws12[column[0].coordinate].value = header_12[0]
            del header_12[0]
            ws12[column[0].coordinate].font = Font(name='Times New Roman', size=10, color='FFFFFF', bold=True)
            ws12[column[0].coordinate].alignment = Alignment(horizontal='center', vertical='center')
            ws12[column[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='31869B')

        ws12.merge_cells('A6:E7')
        ws12.merge_cells('F6:F7')
        ws12.merge_cells('G6:G7')
        ws12.merge_cells('H6:H7')
        ws12['H6'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws12.merge_cells('I6:I7')
        ws12['I6'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws12.merge_cells('J6:J7')
        ws12.merge_cells('K6:L7')
        # Corpo tabella
        corner_strumenti_liquidi = portfolio[(portfolio['INTERMEDIARIO']=='Corner') & (portfolio['CATEGORIA']!='CASH') & (portfolio['CATEGORIA']!='CASH_FOREIGN_CURR')]
        corner_proddoti = corner_strumenti_liquidi.PRODOTTO
        corner_proddoti = list(corner_proddoti)
        for row in ws12.iter_rows(min_row=8, max_row=corner_strumenti_liquidi.shape[0] + 8 -1, min_col=1, max_col=12):
            ws12[row[0].coordinate].value = corner_proddoti[0]
            del corner_proddoti[0]
            ws12[row[0].coordinate].font = Font(name='Times New Roman', size=10)
            ws12[row[5].coordinate].value = (corner_strumenti_liquidi.loc[corner_strumenti_liquidi['PRODOTTO']==ws12[row[0].coordinate].value, 'DIVISA']).values[0]
            ws12[row[5].coordinate].font = Font(name='Times New Roman', size=10)
            ws12[row[5].coordinate].alignment = Alignment(horizontal='center')
            ws12[row[6].coordinate].value = (corner_strumenti_liquidi.loc[corner_strumenti_liquidi['PRODOTTO']==ws12[row[0].coordinate].value, 'QUANTITA t1']).values[0]
            ws12[row[6].coordinate].font = Font(name='Times New Roman', size=10)
            ws12[row[6].coordinate].number_format = '#,##0.00'
            ws12[row[7].coordinate].value = (corner_strumenti_liquidi.loc[corner_strumenti_liquidi['PRODOTTO']==ws12[row[0].coordinate].value, 'Prezzo_di_carico']).values[0]
            ws12[row[7].coordinate].font = Font(name='Times New Roman', size=10)
            ws12[row[7].coordinate].number_format = FORMAT_NUMBER_00
            ws12[row[8].coordinate].value = (corner_strumenti_liquidi.loc[corner_strumenti_liquidi['PRODOTTO']==ws12[row[0].coordinate].value, 'PREZZO t1']).values[0]
            ws12[row[8].coordinate].font = Font(name='Times New Roman', size=10)
            ws12[row[8].coordinate].number_format = FORMAT_NUMBER_00
            ws12[row[9].coordinate].value = (ws12[row[8].coordinate].value / ws12[row[7].coordinate].value) - 1
            ws12[row[9].coordinate].font = Font(name='Times New Roman', size=10)
            ws12[row[9].coordinate].alignment = Alignment(horizontal='center')
            ws12[row[9].coordinate].number_format = FORMAT_PERCENTAGE_00
            ws12[row[10].coordinate].value = (corner_strumenti_liquidi.loc[corner_strumenti_liquidi['PRODOTTO']==ws12[row[0].coordinate].value, 'TOTALE t1']).values[0]
            ws12[row[10].coordinate].font = Font(name='Times New Roman', size=10)
            ws12[row[10].coordinate].alignment = Alignment(horizontal='center')
            ws12[row[10].coordinate].number_format = '€ #,##0.00'
            ws12.merge_cells(start_row=row[0].row, end_row=row[0].row, start_column=row[0].column, end_column=row[4].column)
            ws12.merge_cells(start_row=row[0].row, end_row=row[0].row, start_column=row[10].column, end_column=row[11].column)
        
        if len(corner_strumenti_liquidi) > 26:
            self.logo(ws12, row=35+(len(corner_strumenti_liquidi)-26))
        else:
            self.logo(ws12)

    def prezzi_13(self):
        """
        Crea la tredicesima pagina.
        Formattazione e tabella.
        """
        # Carica portafoglio
        portfolio = pd.read_excel(self.file_portafoglio, sheet_name='Portfolio', header=1)

        ws13 = self.wb.create_sheet('13.prezzi_2')
        ws13 = self.wb['13.prezzi_2']
        self.wb.active = ws13

        ws13['A1'] = 'Credito Artigiano'
        ws13['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws13['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws13['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws13.merge_cells('A1:L4')

        # Creazione tabella
        header_13 = ['Nome', '', '', '', '', 'Valuta', 'Quantità', 'Prezzo di carico', 'Prezzo attuale', '∆ prezzo', 'Ctv', '']
        for column in ws13.iter_cols(min_col=1, max_col=12, min_row=6, max_row=6):
            ws13[column[0].coordinate].value = header_13[0]
            del header_13[0]
            ws13[column[0].coordinate].font = Font(name='Times New Roman', size=10, color='FFFFFF', bold=True)
            ws13[column[0].coordinate].alignment = Alignment(horizontal='center', vertical='center')
            ws13[column[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='31869B')

        ws13.merge_cells('A6:E7')
        ws13.merge_cells('F6:F7')
        ws13.merge_cells('G6:G7')
        ws13.merge_cells('H6:H7')
        ws13['H6'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws13.merge_cells('I6:I7')
        ws13['I6'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws13.merge_cells('J6:J7')
        ws13.merge_cells('K6:L7')

        credito_artigiano_strumenti_liquidi = portfolio[((portfolio['INTERMEDIARIO']=='Credito Artigiano Artes') | (portfolio['INTERMEDIARIO']=='Credito Artigiano B.N.')) & (portfolio['CATEGORIA']!='CASH') & (portfolio['CATEGORIA']!='CASH_FOREIGN_CURR')]
        credito_artigiano_proddoti = credito_artigiano_strumenti_liquidi.PRODOTTO
        credito_artigiano_proddoti = list(credito_artigiano_proddoti)
        for row in ws13.iter_rows(min_row=8, max_row=credito_artigiano_strumenti_liquidi.shape[0] + 8 -1, min_col=1, max_col=12):
            ws13[row[0].coordinate].value = credito_artigiano_proddoti[0]
            del credito_artigiano_proddoti[0]
            ws13[row[0].coordinate].font = Font(name='Times New Roman', size=10)
            ws13[row[5].coordinate].value = (credito_artigiano_strumenti_liquidi.loc[credito_artigiano_strumenti_liquidi['PRODOTTO']==ws13[row[0].coordinate].value, 'DIVISA']).values[0]
            ws13[row[5].coordinate].font = Font(name='Times New Roman', size=10)
            ws13[row[5].coordinate].alignment = Alignment(horizontal='center')
            ws13[row[6].coordinate].value = (credito_artigiano_strumenti_liquidi.loc[credito_artigiano_strumenti_liquidi['PRODOTTO']==ws13[row[0].coordinate].value, 'QUANTITA t1']).values[0]
            ws13[row[6].coordinate].font = Font(name='Times New Roman', size=10)
            ws13[row[6].coordinate].number_format = '#,##0.00'
            ws13[row[7].coordinate].value = (credito_artigiano_strumenti_liquidi.loc[credito_artigiano_strumenti_liquidi['PRODOTTO']==ws13[row[0].coordinate].value, 'Prezzo_di_carico']).values[0]
            ws13[row[7].coordinate].font = Font(name='Times New Roman', size=10)
            ws13[row[7].coordinate].number_format = FORMAT_NUMBER_00
            ws13[row[8].coordinate].value = (credito_artigiano_strumenti_liquidi.loc[credito_artigiano_strumenti_liquidi['PRODOTTO']==ws13[row[0].coordinate].value, 'PREZZO t1']).values[0]
            ws13[row[8].coordinate].font = Font(name='Times New Roman', size=10)
            ws13[row[8].coordinate].number_format = FORMAT_NUMBER_00
            ws13[row[9].coordinate].value = (ws13[row[8].coordinate].value / ws13[row[7].coordinate].value) - 1
            ws13[row[9].coordinate].font = Font(name='Times New Roman', size=10)
            ws13[row[9].coordinate].alignment = Alignment(horizontal='center')
            ws13[row[9].coordinate].number_format = FORMAT_PERCENTAGE_00
            ws13[row[10].coordinate].value = (credito_artigiano_strumenti_liquidi.loc[credito_artigiano_strumenti_liquidi['PRODOTTO']==ws13[row[0].coordinate].value, 'TOTALE t1']).values[0]
            ws13[row[10].coordinate].font = Font(name='Times New Roman', size=10)
            ws13[row[10].coordinate].alignment = Alignment(horizontal='center')
            ws13[row[10].coordinate].number_format = '€ #,##0.00'
            ws13.merge_cells(start_row=row[0].row, end_row=row[0].row, start_column=row[0].column, end_column=row[4].column)
            ws13.merge_cells(start_row=row[0].row, end_row=row[0].row, start_column=row[10].column, end_column=row[11].column)

        # Creazione tabella
        #print(ws13.cell(row=jpm_strumenti_liquidi.shape[0] + 8 + 1, column=1).coordinate)
        jpm_strumenti_liquidi = portfolio[portfolio['PRODOTTO']=='HIGHBRIDGE CAP CORP']
        jpm_proddoti = jpm_strumenti_liquidi.PRODOTTO
        jpm_proddoti = list(jpm_proddoti)
        ws13.cell(row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 1, column=1,  value='JPMorgan')
        ws13.cell(row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 1, column=1).alignment = Alignment(horizontal='center', vertical='center')
        ws13.cell(row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 1, column=1).font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws13.cell(row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 1, column=1).fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws13.merge_cells(start_row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 1, end_row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 4, start_column=1, end_column=12)
        #print(jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 2 - 1)

        header_13 = ['Nome', '', '', '', '', 'Valuta', 'Quantità', 'Prezzo di carico', 'Prezzo attuale', '∆ prezzo', 'Ctv', '']
        for column in ws13.iter_cols(min_col=1, max_col=12, min_row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 6, max_row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 6):
            ws13[column[0].coordinate].value = header_13[0]
            del header_13[0]
            ws13[column[0].coordinate].font = Font(name='Times New Roman', size=10, color='FFFFFF', bold=True)
            ws13[column[0].coordinate].alignment = Alignment(horizontal='center', vertical='center')
            ws13[column[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='31869B')

        ws13.merge_cells(start_row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 6, end_row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 7, start_column=1, end_column=5)
        ws13.merge_cells(start_row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 6, end_row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 7, start_column=6, end_column=6)
        ws13.merge_cells(start_row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 6, end_row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 7, start_column=7, end_column=7)
        ws13.merge_cells(start_row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 6, end_row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 7, start_column=8, end_column=8)
        ws13.cell(row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 6, column=8).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws13.merge_cells(start_row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 6, end_row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 7, start_column=9, end_column=9)
        ws13.cell(row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 6, column=9).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws13.merge_cells(start_row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 6, end_row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 7, start_column=10, end_column=10)
        ws13.merge_cells(start_row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 6, end_row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 7, start_column=11, end_column=12)

        for row in ws13.iter_rows(min_row=credito_artigiano_strumenti_liquidi.shape[0] + 8 + 9, max_row=jpm_strumenti_liquidi.shape[0] + credito_artigiano_strumenti_liquidi.shape[0] + 8 + 9 - 1, min_col=1, max_col=12):
            ws13[row[0].coordinate].value = jpm_proddoti[0]
            del jpm_proddoti[0]
            ws13[row[0].coordinate].font = Font(name='Times New Roman', size=10)
            ws13[row[5].coordinate].value = (jpm_strumenti_liquidi.loc[jpm_strumenti_liquidi['PRODOTTO']==ws13[row[0].coordinate].value, 'DIVISA']).values[0]
            ws13[row[5].coordinate].font = Font(name='Times New Roman', size=10)
            ws13[row[5].coordinate].alignment = Alignment(horizontal='center')
            ws13[row[6].coordinate].value = (jpm_strumenti_liquidi.loc[jpm_strumenti_liquidi['PRODOTTO']==ws13[row[0].coordinate].value, 'QUANTITA t1']).values[0]
            ws13[row[6].coordinate].font = Font(name='Times New Roman', size=10)
            ws13[row[6].coordinate].number_format = '#,##0.00'
            ws13[row[7].coordinate].value = (jpm_strumenti_liquidi.loc[jpm_strumenti_liquidi['PRODOTTO']==ws13[row[0].coordinate].value, 'Prezzo_di_carico']).values[0]
            ws13[row[7].coordinate].font = Font(name='Times New Roman', size=10)
            ws13[row[7].coordinate].number_format = FORMAT_NUMBER_00
            ws13[row[8].coordinate].value = (jpm_strumenti_liquidi.loc[jpm_strumenti_liquidi['PRODOTTO']==ws13[row[0].coordinate].value, 'PREZZO t1']).values[0]
            ws13[row[8].coordinate].font = Font(name='Times New Roman', size=10)
            ws13[row[8].coordinate].number_format = FORMAT_NUMBER_00
            ws13[row[9].coordinate].value = 0.6089 # viene inserito il delta massimo, perché il prezzo dipende dal controvalore e non il contrario
            ws13[row[9].coordinate].font = Font(name='Times New Roman', size=10)
            ws13[row[9].coordinate].alignment = Alignment(horizontal='center')
            ws13[row[9].coordinate].number_format = FORMAT_PERCENTAGE_00
            ws13[row[10].coordinate].value = (jpm_strumenti_liquidi.loc[jpm_strumenti_liquidi['PRODOTTO']==ws13[row[0].coordinate].value, 'TOTALE t1']).values[0]
            ws13[row[10].coordinate].font = Font(name='Times New Roman', size=10)
            ws13[row[10].coordinate].alignment = Alignment(horizontal='center')
            ws13[row[10].coordinate].number_format = '€ #,##0.00'
            ws13.merge_cells(start_row=row[0].row, end_row=row[0].row, start_column=row[0].column, end_column=row[4].column)
            ws13.merge_cells(start_row=row[0].row, end_row=row[0].row, start_column=row[10].column, end_column=row[11].column)

        self.logo(ws13)

    def prezzi_14(self):
        """
        Crea la quattordicesima pagina.
        Formattazione e tabella.
        """
        # Carica portafoglio
        portfolio = pd.read_excel(self.file_portafoglio, sheet_name='Portfolio', header=1)

        # 14.Prezzi 3
        ws14 = self.wb.create_sheet('14.prezzi_3')
        ws14 = self.wb['14.prezzi_3']
        self.wb.active = ws14

        ws14['A1'] = 'Ubi Ita'
        ws14['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws14['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws14['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws14.merge_cells('A1:L4')

        # Creazione tabella
        header_14 = ['Nome', '', '', '', '', 'Valuta', 'Quantità', 'Prezzo di carico', 'Prezzo attuale', '∆ prezzo', 'Ctv', '']
        for column in ws14.iter_cols(min_col=1, max_col=12, min_row=6, max_row=6):
            ws14[column[0].coordinate].value = header_14[0]
            del header_14[0]
            ws14[column[0].coordinate].font = Font(name='Times New Roman', size=10, color='FFFFFF', bold=True)
            ws14[column[0].coordinate].alignment = Alignment(horizontal='center', vertical='center')
            ws14[column[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='31869B')

        ws14.merge_cells('A6:E7')
        ws14.merge_cells('F6:F7')
        ws14.merge_cells('G6:G7')
        ws14.merge_cells('H6:H7')
        ws14['H6'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws14.merge_cells('I6:I7')
        ws14['I6'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws14.merge_cells('J6:J7')
        ws14.merge_cells('K6:L7')

        ubi_strumenti_liquidi = portfolio[(portfolio['INTERMEDIARIO']=='Ubi Ita') & (portfolio['CATEGORIA']!='CASH') & (portfolio['CATEGORIA']!='CASH_FOREIGN_CURR') & ((portfolio['PRODOTTO']!='RAGANELLA') & (portfolio['PRODOTTO']!='ANTHARES SPA'))]
        ubi_proddoti = ubi_strumenti_liquidi.PRODOTTO
        ubi_proddoti = list(ubi_proddoti)
        for row in ws14.iter_rows(min_row=8, max_row=ubi_strumenti_liquidi.shape[0] + 8 -1, min_col=1, max_col=12):
            ws14[row[0].coordinate].value = ubi_proddoti[0]
            del ubi_proddoti[0]
            ws14[row[0].coordinate].font = Font(name='Times New Roman', size=10)
            ws14[row[5].coordinate].value = (ubi_strumenti_liquidi.loc[ubi_strumenti_liquidi['PRODOTTO']==ws14[row[0].coordinate].value, 'DIVISA']).values[0]
            ws14[row[5].coordinate].font = Font(name='Times New Roman', size=10)
            ws14[row[5].coordinate].alignment = Alignment(horizontal='center')
            ws14[row[6].coordinate].value = (ubi_strumenti_liquidi.loc[ubi_strumenti_liquidi['PRODOTTO']==ws14[row[0].coordinate].value, 'QUANTITA t1']).values[0]
            ws14[row[6].coordinate].font = Font(name='Times New Roman', size=10)
            ws14[row[6].coordinate].number_format = '#,##0.00'
            ws14[row[7].coordinate].value = (ubi_strumenti_liquidi.loc[ubi_strumenti_liquidi['PRODOTTO']==ws14[row[0].coordinate].value, 'Prezzo_di_carico']).values[0]
            ws14[row[7].coordinate].font = Font(name='Times New Roman', size=10)
            ws14[row[7].coordinate].number_format = FORMAT_NUMBER_00
            ws14[row[8].coordinate].value = (ubi_strumenti_liquidi.loc[ubi_strumenti_liquidi['PRODOTTO']==ws14[row[0].coordinate].value, 'PREZZO t1']).values[0]
            ws14[row[8].coordinate].font = Font(name='Times New Roman', size=10)
            ws14[row[8].coordinate].number_format = FORMAT_NUMBER_00
            ws14[row[9].coordinate].value = (ws14[row[8].coordinate].value / ws14[row[7].coordinate].value) - 1
            ws14[row[9].coordinate].font = Font(name='Times New Roman', size=10)
            ws14[row[9].coordinate].alignment = Alignment(horizontal='center')
            ws14[row[9].coordinate].number_format = FORMAT_PERCENTAGE_00
            ws14[row[10].coordinate].value = (ubi_strumenti_liquidi.loc[ubi_strumenti_liquidi['PRODOTTO']==ws14[row[0].coordinate].value, 'TOTALE t1']).values[0]
            ws14[row[10].coordinate].font = Font(name='Times New Roman', size=10)
            ws14[row[10].coordinate].alignment = Alignment(horizontal='center')
            ws14[row[10].coordinate].number_format = '€ #,##0.00'
            ws14.merge_cells(start_row=row[0].row, end_row=row[0].row, start_column=row[0].column, end_column=row[4].column)
            ws14.merge_cells(start_row=row[0].row, end_row=row[0].row, start_column=row[10].column, end_column=row[11].column)

        # Creazione tabella
        #print(ws14.cell(row=mediolanum_strumenti_liquidi.shape[0] + 8 + 1, column=1).coordinate)
        mediolanum_strumenti_liquidi = portfolio[(portfolio['INTERMEDIARIO']=='Mediolanum') & (portfolio['CATEGORIA']!='CASH') & (portfolio['CATEGORIA']!='CASH_FOREIGN_CURR')]
        jpm_proddoti = mediolanum_strumenti_liquidi.PRODOTTO
        jpm_proddoti = list(jpm_proddoti)
        ws14.cell(row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 1, column=1,  value='Mediolanum')
        ws14.cell(row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 1, column=1).alignment = Alignment(horizontal='center', vertical='center')
        ws14.cell(row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 1, column=1).font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws14.cell(row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 1, column=1).fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws14.merge_cells(start_row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 1, end_row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 4, start_column=1, end_column=12)
        #print(mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 2 - 1)

        header_14 = ['Nome', '', '', '', '', 'Valuta', 'Quantità', 'Prezzo di carico', 'Prezzo attuale', '∆ prezzo', 'Ctv', '']
        for column in ws14.iter_cols(min_col=1, max_col=12, min_row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 6, max_row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 6):
            ws14[column[0].coordinate].value = header_14[0]
            del header_14[0]
            ws14[column[0].coordinate].font = Font(name='Times New Roman', size=10, color='FFFFFF', bold=True)
            ws14[column[0].coordinate].alignment = Alignment(horizontal='center', vertical='center')
            ws14[column[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='31869B')

        ws14.merge_cells(start_row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 6, end_row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 7, start_column=1, end_column=5)
        ws14.merge_cells(start_row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 6, end_row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 7, start_column=6, end_column=6)
        ws14.merge_cells(start_row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 6, end_row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 7, start_column=7, end_column=7)
        ws14.merge_cells(start_row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 6, end_row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 7, start_column=8, end_column=8)
        ws14.cell(row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 6, column=8).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws14.merge_cells(start_row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 6, end_row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 7, start_column=9, end_column=9)
        ws14.cell(row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 6, column=9).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws14.merge_cells(start_row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 6, end_row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 7, start_column=10, end_column=10)
        ws14.merge_cells(start_row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 6, end_row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 7, start_column=11, end_column=12)
        
        for row in ws14.iter_rows(min_row=ubi_strumenti_liquidi.shape[0] + 8 + 9, max_row=mediolanum_strumenti_liquidi.shape[0] + ubi_strumenti_liquidi.shape[0] + 8 + 9 - 1, min_col=1, max_col=12):
            ws14[row[0].coordinate].value = jpm_proddoti[0]
            del jpm_proddoti[0]
            ws14[row[0].coordinate].font = Font(name='Times New Roman', size=10)
            ws14[row[5].coordinate].value = (mediolanum_strumenti_liquidi.loc[mediolanum_strumenti_liquidi['PRODOTTO']==ws14[row[0].coordinate].value, 'DIVISA']).values[0]
            ws14[row[5].coordinate].font = Font(name='Times New Roman', size=10)
            ws14[row[5].coordinate].alignment = Alignment(horizontal='center')
            ws14[row[6].coordinate].value = (mediolanum_strumenti_liquidi.loc[mediolanum_strumenti_liquidi['PRODOTTO']==ws14[row[0].coordinate].value, 'QUANTITA t1']).values[0]
            ws14[row[6].coordinate].font = Font(name='Times New Roman', size=10)
            ws14[row[6].coordinate].number_format = '#,##0.00'
            ws14[row[7].coordinate].value = (mediolanum_strumenti_liquidi.loc[mediolanum_strumenti_liquidi['PRODOTTO']==ws14[row[0].coordinate].value, 'Prezzo_di_carico']).values[0]
            ws14[row[7].coordinate].font = Font(name='Times New Roman', size=10)
            ws14[row[7].coordinate].number_format = FORMAT_NUMBER_00
            ws14[row[8].coordinate].value = (mediolanum_strumenti_liquidi.loc[mediolanum_strumenti_liquidi['PRODOTTO']==ws14[row[0].coordinate].value, 'PREZZO t1']).values[0]
            ws14[row[8].coordinate].font = Font(name='Times New Roman', size=10)
            ws14[row[8].coordinate].number_format = FORMAT_NUMBER_00
            ws14[row[9].coordinate].value = (ws14[row[8].coordinate].value / ws14[row[7].coordinate].value) - 1
            ws14[row[9].coordinate].font = Font(name='Times New Roman', size=10)
            ws14[row[9].coordinate].alignment = Alignment(horizontal='center')
            ws14[row[9].coordinate].number_format = FORMAT_PERCENTAGE_00
            ws14[row[10].coordinate].value = (mediolanum_strumenti_liquidi.loc[mediolanum_strumenti_liquidi['PRODOTTO']==ws14[row[0].coordinate].value, 'TOTALE t1']).values[0]
            ws14[row[10].coordinate].font = Font(name='Times New Roman', size=10)
            ws14[row[10].coordinate].alignment = Alignment(horizontal='center')
            ws14[row[10].coordinate].number_format = '€ #,##0.00'
            ws14.merge_cells(start_row=row[0].row, end_row=row[0].row, start_column=row[0].column, end_column=row[4].column)
            ws14.merge_cells(start_row=row[0].row, end_row=row[0].row, start_column=row[10].column, end_column=row[11].column)

        self.logo(ws14)

    def att_in_corso_15(self):
        """
        Crea la quindicesima pagina
        """
        ws15 = self.wb.create_sheet('15.att')
        ws15 = self.wb['15.att']
        self.wb.active = ws15

        ws15['A1'] = 'Attività Svolte Ed In Corso'
        ws15['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws15['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws15['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws15.merge_cells('A1:L4')

        # Text box
        self.text_box(ws15, 6, 28, 1, 12)

        self.logo(ws15)

    def valutazione_per_macroclasse_16(self):
        """
        Crea la sedicesima pagina.
        """
        ws16 = self.wb.create_sheet('16.val_per_macro')
        ws16 = self.wb['16.val_per_macro']
        self.wb.active = ws16

        ws16['A11'] = '3. Valutazione Per Macroclasse'
        ws16['A11'].alignment = Alignment(horizontal='center', vertical='center')
        ws16['A11'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws16.merge_cells('A11:L14')
        ws16['A11'].font = Font(name='Times New Roman', size=36, bold=True, color='FFFFFF')

        self.logo(ws16)

    def sintesi_17(self):
        """
        Crea la diciasettesima pagina.
        Formattazione e tabella.
        """
        # Carica portafoglio
        portfolio = pd.read_excel(self.file_portafoglio, sheet_name='Portfolio', header=1)

        ws17 = self.wb.create_sheet('17.sintesi')
        ws17 = self.wb['17.sintesi']
        self.wb.active = ws17

        # Creazione tabella
        #print(portfolio['INTERMEDIARIO'].unique())
        header_17 = list(portfolio['INTERMEDIARIO'].unique())
        header_17.insert(0, '')
        header_17.extend(('Totale '+ self.mesi_dict[self.t1.month], 'Totale '+ self.mesi_dict[self.t0_1m.month]))
        len_header_17 = len(header_17)
        #print(header_17)

        # Titolo
        ws17['A1'] = 'Sintesi'
        ws17['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws17['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws17['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        if len(list(portfolio['INTERMEDIARIO'].unique())) == 1:
            lunghezza_titolo_17 = 12
            min_col = 4
        elif len(list(portfolio['INTERMEDIARIO'].unique())) == 2:
            lunghezza_titolo_17 = 12
            min_col = 4
        elif len(list(portfolio['INTERMEDIARIO'].unique())) == 3:
            lunghezza_titolo_17 = 12
            min_col = 3
        elif len(list(portfolio['INTERMEDIARIO'].unique())) == 4:
            lunghezza_titolo_17 = 12
            min_col = 3
        elif len(list(portfolio['INTERMEDIARIO'].unique())) == 5:
            lunghezza_titolo_17 = 12
            min_col = 2
        elif len(list(portfolio['INTERMEDIARIO'].unique())) == 6:
            lunghezza_titolo_17 = 12
            min_col = 2
        elif len(list(portfolio['INTERMEDIARIO'].unique())) == 7:
            lunghezza_titolo_17 = 12
            min_col = 1
        else:
            lunghezza_titolo_17 = len_header_17
            min_col = 1
        ws17.merge_cells(start_row=1, end_row=4, start_column=1, end_column=lunghezza_titolo_17)

        for col in ws17.iter_cols(min_row=8, max_row=9, min_col=min_col, max_col=min_col + len_header_17 - 1):
            ws17[col[0].coordinate].value = header_17[0]
            del header_17[0]
            ws17[col[0].coordinate].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws17[col[0].coordinate].font = Font(name='Times New Roman', size=10, color='FFFFFF')
            ws17[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='31869B')
            ws17[col[0].coordinate].border = Border(right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'))
            ws17.merge_cells(start_row=col[0].row, end_row=col[1].row, start_column=col[0].column, end_column=col[0].column)
            ws17.row_dimensions[col[0].row].height = 20
            ws17.row_dimensions[col[1].row].height = 20
            ws17.column_dimensions[col[0].column_letter].width = 12

        tipo_strumento = list(portfolio['CATEGORIA'].unique())
        len_tipo_strumento = len(tipo_strumento)
        num_intermediari = len(portfolio['INTERMEDIARIO'].unique())
        #print(num_intermediari)
        lunghezza_colonna_17 = []
        # print(tipo_strumento)
        # print(type(tipo_strumento))
        # print(len(tipo_strumento))
        tipo_strumento_dict = {'CASH' : 'LIQUIDITÀ', 'GP' : 'GESTIONI', 'EQUITY' : 'AZIONI', 'CASH_FOREIGN_CURR' : 'LIQUIDITÀ IN VALUTA', 'CORPORATE_BOND' : 'OBBLIGAZIONI CORPORATE', 'GOVERNMENT_BOND' : 'OBBLIGAZIONI GOVERNATIVE', 'ALTERNATIVE_ASSET' : 'INVESTIMENTI ALTERNATIVI', 'HEDGE_FUND' : 'HEDGE FUND'}
        #tipo_strumento_dict = {k: v for k, v in sorted(tipo_strumento_dict.items(), key=lambda item: item[1])}

        for row in ws17.iter_rows(min_row=8, max_row=10 + len_tipo_strumento -1, min_col=min_col, max_col=min_col + len_header_17):
            if row[0].row > 9:
                #print(ws17[row[0].coordinate])
                #ws17[row[0].coordinate].value = tipo_strumento_dict[tipo_strumento[0]]
                ws17[row[0].coordinate].value = tipo_strumento[0] # carica i tipi di strumenti nell'indice
                del tipo_strumento[0]
                ws17[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
                ws17[row[0].coordinate].font = Font(name='Times New Roman', size=9, color='000000')
                ws17[row[0].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws17.row_dimensions[row[0].row].height = 19
                #print(ws17[row[1].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) # mostra sempre l'intermediario
                for _ in range(1, num_intermediari+1):
                    ws17[row[_].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws17[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) & (portfolio['CATEGORIA']==ws17[row[0].coordinate].value), 'TOTALE t1'].sum() if portfolio.loc[(portfolio['INTERMEDIARIO']==ws17[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) & (portfolio['CATEGORIA']==ws17[row[0].coordinate].value), 'TOTALE t1'].sum() != 0 else ''
                    ws17[row[_].coordinate].alignment = Alignment(horizontal='center')
                    ws17[row[_].coordinate].font = Font(name='Times New Roman', size=9)
                    ws17[row[_].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                    ws17[row[_].coordinate].number_format = '#,0'
                    # somma=[]
                    # somma.append(ws17[row[_].coordinate].value)
                    # print(somma)
                    #ws17[row[num_intermediari+1].coordinate].value = 0
                    #ws17[row[num_intermediari+1].coordinate].value = (ws17[row[num_intermediari+1].coordinate].value + float(ws17[row[_].coordinate].value)) if float(ws17[row[_].coordinate].value) != ''
                    #ws17.cell(row=row[_].row, column=num_intermediari+1, value=ws17[row[_].coordinate].value) += 
                    #print(ws17[row[_].coordinate].value)
                    # print(ws17[row[num_intermediari+1].coordinate].value)
                    # ws17[row[num_intermediari+1].coordinate].value = 0
                    # try:
                    #     ws17[row[num_intermediari+1].coordinate].value += float(ws17[row[_].coordinate].value)
                    # except ValueError:
                    #     pass
                    # except TypeError:
                    #     pass
                    # if float(ws17[row[_].coordinate].value) != ''
                #ws17[row[num_intermediari+2].coordinate].value = '=SUM('+str(ws17[row[1].coordinate])+':'+str(ws17[row[num_intermediari].coordinate])+')'
                # Somma per strumenti
                ws17[row[num_intermediari+1].coordinate].value = portfolio.loc[portfolio['CATEGORIA']==ws17[row[0].coordinate].value, 'TOTALE t1'].sum()
                ws17[row[num_intermediari+1].coordinate].alignment = Alignment(horizontal='center')
                ws17[row[num_intermediari+1].coordinate].font = Font(name='Times New Roman', size=9)
                ws17[row[num_intermediari+1].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws17[row[num_intermediari+1].coordinate].number_format = '#,0'
                ws17[row[num_intermediari+2].coordinate].value = portfolio.loc[portfolio['CATEGORIA']==ws17[row[0].coordinate].value, 'TOTALE t0'].sum()
                ws17[row[num_intermediari+2].coordinate].alignment = Alignment(horizontal='center')
                ws17[row[num_intermediari+2].coordinate].font = Font(name='Times New Roman', size=9)
                ws17[row[num_intermediari+2].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws17[row[num_intermediari+2].coordinate].number_format = '#,0'

                ws17[row[0].coordinate].value = tipo_strumento_dict[ws17[row[0].coordinate].value] # aggiorna valori dell'indice con i nomi nel dizionario
                lunghezza_colonna_17.append(len(ws17.cell(row=row[0].row, column=row[0].column).value)) # ottieni la lunghezza della colonna
                ws17.column_dimensions[row[0].column_letter].width = max(lunghezza_colonna_17) + 2.5 # modifica larghezza colonna

        # Somma per intermediari
        for row in ws17.iter_rows(min_row=10 + len_tipo_strumento, max_row=10 + len_tipo_strumento, min_col=min_col, max_col=min_col + len_header_17):
            ws17[row[0].coordinate].value = 'TOTALE'
            ws17[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
            ws17[row[0].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
            ws17[row[0].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
            ws17.row_dimensions[row[0].row].height = 19
            #print(ws17.cell(row=row[1].row, column=row[1].column).offset(row=-len_tipo_strumento))
            for _ in range(1,len_header_17-2):
                ws17[row[_].coordinate].value = portfolio.loc[portfolio['INTERMEDIARIO']==ws17.cell(row=row[_].row, column=row[_].column).offset(row=-len_tipo_strumento-2).value, 'TOTALE t1'].sum()
                # assert ws17[row[_].coordinate].value == 'SUM(B10:B17)'
            ws17[row[len_header_17-2].coordinate].value = portfolio.loc[:, 'TOTALE t1'].sum()
            ws17[row[len_header_17-1].coordinate].value = portfolio.loc[:, 'TOTALE t0'].sum()
            for _ in range(1,len_header_17):
                ws17[row[_].coordinate].alignment = Alignment(horizontal='center')
                ws17[row[_].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
                ws17[row[_].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
                ws17[row[_].coordinate].number_format = '#,0'

    def valuta_18(self):
        """
        Crea la diciottesima pagina.
        Formattazione e tabella.
        """
        # Carica portafoglio
        portfolio = pd.read_excel(self.file_portafoglio, sheet_name='Portfolio', header=1)

        ws18 = self.wb.create_sheet('18.valuta')
        ws18 = self.wb['18.valuta']
        self.wb.active = ws18

        # Creazione tabella
        header_18 = list(portfolio['INTERMEDIARIO'].unique())
        header_18.insert(0, '')
        header_18.extend(('Totale '+ self.mesi_dict[self.t1.month], 'Totale '+ self.mesi_dict[self.t0_1m.month]))
        len_header_18 = len(header_18)

        # Titolo
        ws18['A1'] = 'Valuta'
        ws18['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws18['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws18['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        if len(list(portfolio['INTERMEDIARIO'].unique())) == 1:
            lunghezza_titolo_18 = 12
            min_col = 4
        elif len(list(portfolio['INTERMEDIARIO'].unique())) == 2:
            lunghezza_titolo_18 = 12
            min_col = 4
        elif len(list(portfolio['INTERMEDIARIO'].unique())) == 3:
            lunghezza_titolo_18 = 12
            min_col = 3
        elif len(list(portfolio['INTERMEDIARIO'].unique())) == 4:
            lunghezza_titolo_18 = 12
            min_col = 3
        elif len(list(portfolio['INTERMEDIARIO'].unique())) == 5:
            lunghezza_titolo_18 = 12
            min_col = 2
        elif len(list(portfolio['INTERMEDIARIO'].unique())) == 6:
            lunghezza_titolo_18 = 12
            min_col = 2
        elif len(list(portfolio['INTERMEDIARIO'].unique())) == 7:
            lunghezza_titolo_18 = 12
            min_col = 1
        else:
            lunghezza_titolo_18 = len_header_18
            min_col = 1
        ws18.merge_cells(start_row=1, end_row=4, start_column=1, end_column=lunghezza_titolo_18)

        for col in ws18.iter_cols(min_row=8, max_row=9, min_col=min_col, max_col=min_col + len_header_18 - 1):
            ws18[col[0].coordinate].value = header_18[0]
            del header_18[0]
            ws18[col[0].coordinate].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws18[col[0].coordinate].font = Font(name='Times New Roman', size=10, color='FFFFFF')
            ws18[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='31869B')
            ws18[col[0].coordinate].border = Border(right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'))
            ws18.merge_cells(start_row=col[0].row, end_row=col[1].row, start_column=col[0].column, end_column=col[0].column)
            ws18.row_dimensions[col[0].row].height = 20
            ws18.row_dimensions[col[1].row].height = 20
            ws18.column_dimensions[col[0].column_letter].width = 12

        tipo_divisa = list(portfolio['DIVISA'].unique())
        tipo_divisa.sort()
        # tipo_divisa.insert(len(tipo_divisa), 'ALTRE VALUTE')
        len_tipo_divisa = len(tipo_divisa)
        num_intermediari = len(portfolio['INTERMEDIARIO'].unique())
        lunghezza_colonna_18 = []
        #tipo_divisa_dict = {'CASH' : 'LIQUIDITÀ', 'GP' : 'GESTIONI', 'EQUITY' : 'AZIONI', 'CASH_FOREIGN_CURR' : 'LIQUIDITÀ IN VALUTA', 'CORPORATE_BOND' : 'OBBLIGAZIONI CORPORATE', 'GOVERNMENT_BOND' : 'OBBLIGAZIONI GOVERNATIVE', 'ALTERNATIVE_ASSET' : 'INVESTIMENTI ALTERNATIVI', 'HEDGE_FUND' : 'HEDGE FUND'}
        for row in ws18.iter_rows(min_row=8, max_row=10 + len_tipo_divisa -1, min_col=min_col, max_col=min_col + len_header_18):
            if row[0].row > 9:
                ws18[row[0].coordinate].value = tipo_divisa[0]
                del tipo_divisa[0]
                ws18[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
                ws18[row[0].coordinate].font = Font(name='Times New Roman', size=9, color='000000')
                ws18[row[0].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws18.row_dimensions[row[0].row].height = 19
                for _ in range(1, num_intermediari+1):
                    ws18[row[_].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws18[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) & (portfolio['DIVISA']==ws18[row[0].coordinate].value), 'TOTALE t1'].sum() if portfolio.loc[(portfolio['INTERMEDIARIO']==ws18[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) & (portfolio['DIVISA']==ws18[row[0].coordinate].value), 'TOTALE t1'].sum() != 0 else ''
                    ws18[row[_].coordinate].alignment = Alignment(horizontal='center')
                    ws18[row[_].coordinate].font = Font(name='Times New Roman', size=9)
                    ws18[row[_].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                    ws18[row[_].coordinate].number_format = '#,0'

                ws18[row[num_intermediari+1].coordinate].value = portfolio.loc[portfolio['DIVISA']==ws18[row[0].coordinate].value, 'TOTALE t1'].sum()
                ws18[row[num_intermediari+1].coordinate].alignment = Alignment(horizontal='center')
                ws18[row[num_intermediari+1].coordinate].font = Font(name='Times New Roman', size=9)
                ws18[row[num_intermediari+1].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws18[row[num_intermediari+1].coordinate].number_format = '#,0'
                ws18[row[num_intermediari+2].coordinate].value = portfolio.loc[portfolio['DIVISA']==ws18[row[0].coordinate].value, 'TOTALE t0'].sum()
                ws18[row[num_intermediari+2].coordinate].alignment = Alignment(horizontal='center')
                ws18[row[num_intermediari+2].coordinate].font = Font(name='Times New Roman', size=9)
                ws18[row[num_intermediari+2].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws18[row[num_intermediari+2].coordinate].number_format = '#,0'

                #ws18[row[0].coordinate].value = tipo_divisa_dict[ws18[row[0].coordinate].value] # aggiorna valori dell'indice con i nomi nel dizionario
                lunghezza_colonna_18.append(len(ws18.cell(row=row[0].row, column=row[0].column).value)) # ottieni la lunghezza della colonna
                ws18.column_dimensions[row[0].column_letter].width = max(lunghezza_colonna_18) + 7.5 # modifica larghezza colonna


        # Somma per intermediari
        for row in ws18.iter_rows(min_row=10 + len_tipo_divisa, max_row=10 + len_tipo_divisa, min_col=min_col, max_col=min_col + len_header_18):
            ws18[row[0].coordinate].value = 'TOTALE'
            ws18[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
            ws18[row[0].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
            ws18[row[0].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
            ws18.row_dimensions[row[0].row].height = 19
            for _ in range(1,len_header_18-2):
                ws18[row[_].coordinate].value = portfolio.loc[portfolio['INTERMEDIARIO']==ws18.cell(row=row[_].row, column=row[_].column).offset(row=-len_tipo_divisa-2).value, 'TOTALE t1'].sum()
            ws18[row[len_header_18-2].coordinate].value = portfolio.loc[:, 'TOTALE t1'].sum()
            ws18[row[len_header_18-1].coordinate].value = portfolio.loc[:, 'TOTALE t0'].sum()
            for _ in range(1,len_header_18):
                ws18[row[_].coordinate].alignment = Alignment(horizontal='center')
                ws18[row[_].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
                ws18[row[_].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
                ws18[row[_].coordinate].number_format = '#,0'

    def azioni_19(self):
        """
        Crea la diciannovesima pagina.
        Formattazione e tabella.
        """
        # Carica portafoglio
        portfolio = pd.read_excel(self.file_portafoglio, sheet_name='Portfolio', header=1)

        ws19 = self.wb.create_sheet('19.azioni')
        ws19 = self.wb['19.azioni']
        self.wb.active = ws19

        # Creazione tabella
        header_19 = list(portfolio.loc[portfolio['CATEGORIA']=='EQUITY','INTERMEDIARIO'].unique())
        header_19.insert(0, '')
        header_19.extend(('Totale '+ self.mesi_dict[self.t1.month], 'Totale '+ self.mesi_dict[self.t0_1m.month], 'Delta'))
        len_header_19 = len(header_19)

        # Titolo
        ws19['A1'] = 'Azioni'
        ws19['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws19['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws19['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        if len(list(portfolio.loc[portfolio['CATEGORIA']=='EQUITY','INTERMEDIARIO'].unique())) == 1:
            lunghezza_titolo_19 = 12
            min_col = 4
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='EQUITY','INTERMEDIARIO'].unique())) == 2:
            lunghezza_titolo_19 = 12
            min_col = 4
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='EQUITY','INTERMEDIARIO'].unique())) == 3:
            lunghezza_titolo_19 = 12
            min_col = 3
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='EQUITY','INTERMEDIARIO'].unique())) == 4:
            lunghezza_titolo_19 = 12
            min_col = 3
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='EQUITY','INTERMEDIARIO'].unique())) == 5:
            lunghezza_titolo_19 = 12
            min_col = 2
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='EQUITY','INTERMEDIARIO'].unique())) == 6:
            lunghezza_titolo_19 = 12
            min_col = 2
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='EQUITY','INTERMEDIARIO'].unique())) == 7:
            lunghezza_titolo_19 = 12
            min_col = 1
        else:
            lunghezza_titolo_19 = len_header_19
            min_col = 1
        ws19.merge_cells(start_row=1, end_row=4, start_column=1, end_column=lunghezza_titolo_19)


        for col in ws19.iter_cols(min_row=8, max_row=9, min_col=min_col, max_col=min_col + len_header_19 - 1):
            ws19[col[0].coordinate].value = header_19[0]
            del header_19[0]
            ws19[col[0].coordinate].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws19[col[0].coordinate].font = Font(name='Times New Roman', size=10, color='FFFFFF')
            ws19[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='31869B')
            ws19[col[0].coordinate].border = Border(right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'))
            ws19.merge_cells(start_row=col[0].row, end_row=col[1].row, start_column=col[0].column, end_column=col[0].column)
            ws19.row_dimensions[col[0].row].height = 20
            ws19.row_dimensions[col[1].row].height = 20
            ws19.column_dimensions[col[0].column_letter].width = 12

        nome_azioni = list(portfolio.loc[portfolio['CATEGORIA']=='EQUITY','PRODOTTO'])
        len_nome_azioni = len(nome_azioni)
        num_intermediari = len(portfolio.loc[portfolio['CATEGORIA']=='EQUITY', 'INTERMEDIARIO'].unique())
        lunghezza_colonna_19 = []
        for row in ws19.iter_rows(min_row=8, max_row=10 + len_nome_azioni -1, min_col=min_col, max_col=min_col + len_header_19):
            if row[0].row > 9:
                ws19[row[0].coordinate].value = nome_azioni[0] 
                del nome_azioni[0]
                ws19[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
                ws19[row[0].coordinate].font = Font(name='Times New Roman', size=9, color='000000')
                ws19[row[0].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                #ws19[row[1].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']
                for _ in range(1, num_intermediari+1):
                    ws19[row[_].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws19[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) & (portfolio['PRODOTTO']==ws19[row[0].coordinate].value), 'TOTALE t1'].sum() if portfolio.loc[(portfolio['INTERMEDIARIO']==ws19[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) & (portfolio['PRODOTTO']==ws19[row[0].coordinate].value), 'TOTALE t1'].sum() != 0 else ''
                    ws19[row[_].coordinate].alignment = Alignment(horizontal='center')
                    ws19[row[_].coordinate].font = Font(name='Times New Roman', size=9)
                    ws19[row[_].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                    ws19[row[_].coordinate].number_format = '#,0'

                ws19[row[num_intermediari+1].coordinate].value = portfolio.loc[portfolio['PRODOTTO']==ws19[row[0].coordinate].value, 'TOTALE t1'].sum()
                ws19[row[num_intermediari+1].coordinate].alignment = Alignment(horizontal='center')
                ws19[row[num_intermediari+1].coordinate].font = Font(name='Times New Roman', size=9)
                ws19[row[num_intermediari+1].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws19[row[num_intermediari+1].coordinate].number_format = '#,0'
                ws19[row[num_intermediari+2].coordinate].value = portfolio.loc[portfolio['PRODOTTO']==ws19[row[0].coordinate].value, 'TOTALE t0'].sum()
                ws19[row[num_intermediari+2].coordinate].alignment = Alignment(horizontal='center')
                ws19[row[num_intermediari+2].coordinate].font = Font(name='Times New Roman', size=9)
                ws19[row[num_intermediari+2].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws19[row[num_intermediari+2].coordinate].number_format = '#,0'
                ws19[row[num_intermediari+3].coordinate].value = (ws19[row[num_intermediari+1].coordinate].value -  ws19[row[num_intermediari+2].coordinate].value) / ws19[row[num_intermediari+2].coordinate].value if ws19[row[num_intermediari+2].coordinate].value != 0 and ws19[row[num_intermediari+1].coordinate].value != 0 else '/'
                ws19[row[num_intermediari+3].coordinate].alignment = Alignment(horizontal='center')
                ws19[row[num_intermediari+3].coordinate].font = Font(name='Times New Roman', size=9)
                ws19[row[num_intermediari+3].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws19[row[num_intermediari+3].coordinate].number_format = FORMAT_PERCENTAGE_00

                lunghezza_colonna_19.append(len(ws19.cell(row=row[0].row, column=row[0].column).value)) # ottieni la lunghezza della colonna
                ws19.column_dimensions[row[0].column_letter].width = max(lunghezza_colonna_19) + 2.5 # modifica larghezza colonna

        # Somma per intermediari
        for row in ws19.iter_rows(min_row=10 + len_nome_azioni, max_row=10 + len_nome_azioni, min_col=min_col, max_col=min_col + len_header_19):
            ws19[row[0].coordinate].value = 'TOTALE'
            ws19[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
            ws19[row[0].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
            ws19[row[0].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
            for _ in range(1,len_header_19-2):
                ws19[row[_].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws19.cell(row=row[_].row, column=row[_].column).offset(row=-len_nome_azioni-2).value) & (portfolio['CATEGORIA']=='EQUITY'), 'TOTALE t1'].sum()
            ws19[row[len_header_19-3].coordinate].value = portfolio.loc[portfolio['CATEGORIA']=='EQUITY', 'TOTALE t1'].sum()
            ws19[row[len_header_19-2].coordinate].value = portfolio.loc[portfolio['CATEGORIA']=='EQUITY', 'TOTALE t0'].sum()
            ws19[row[len_header_19-1].coordinate].value = (portfolio.loc[(portfolio['CATEGORIA']=='EQUITY') & (portfolio['TOTALE t0']!=0), 'TOTALE t1'].sum() - portfolio.loc[(portfolio['CATEGORIA']=='EQUITY') & (portfolio['TOTALE t1']!=0), 'TOTALE t0'].sum()) / portfolio.loc[(portfolio['CATEGORIA']=='EQUITY') & (portfolio['TOTALE t1']!=0), 'TOTALE t0'].sum()
            for _ in range(1,len_header_19):
                ws19[row[_].coordinate].alignment = Alignment(horizontal='center')
                ws19[row[_].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
                ws19[row[_].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
                ws19[row[_].coordinate].number_format = '#,0'
            ws19[row[len_header_19-1].coordinate].number_format = FORMAT_PERCENTAGE_00

    def obb_governative_20(self):
        """
        Crea la ventesima pagina.
        Formattazione e tabella.
        """
        # Carica portafoglio
        portfolio = pd.read_excel(self.file_portafoglio, sheet_name='Portfolio', header=1)

        # 20.Obb. governative
        ws20 = self.wb.create_sheet('20.obb_gov')
        ws20 = self.wb['20.obb_gov']
        self.wb.active = ws20

        # Creazione tabella
        header_20 = list(portfolio.loc[portfolio['CATEGORIA']=='GOVERNMENT_BOND','INTERMEDIARIO'].unique())
        header_20.insert(0, '')
        header_20.extend(('Totale '+ self.mesi_dict[self.t1.month], 'Totale '+ self.mesi_dict[self.t0_1m.month], 'Delta'))
        len_header_20 = len(header_20)

        # TODO:prima crea la tabella, poi alla fine, calcola la lunghezza dell'header, aggiungi colonne davanti o dietro per centrare la tabella, e infine inserisci il titolo in alto
        ws20['A1'] = 'Obbligazioni Governative'
        ws20['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws20['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws20['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        if len(list(portfolio.loc[portfolio['CATEGORIA']=='GOVERNMENT_BOND','INTERMEDIARIO'].unique())) == 1:
            lunghezza_titolo_20 = 12
            min_col = 4
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='GOVERNMENT_BOND','INTERMEDIARIO'].unique())) == 2:
            lunghezza_titolo_20 = 12
            min_col = 4
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='GOVERNMENT_BOND','INTERMEDIARIO'].unique())) == 3:
            lunghezza_titolo_20 = 12
            min_col = 3
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='GOVERNMENT_BOND','INTERMEDIARIO'].unique())) == 4:
            lunghezza_titolo_20 = 12
            min_col = 3
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='GOVERNMENT_BOND','INTERMEDIARIO'].unique())) == 5:
            lunghezza_titolo_20 = 12
            min_col = 2
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='GOVERNMENT_BOND','INTERMEDIARIO'].unique())) == 6:
            lunghezza_titolo_20 = 12
            min_col = 2
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='GOVERNMENT_BOND','INTERMEDIARIO'].unique())) == 7:
            lunghezza_titolo_20 = 12
            min_col = 1
        else:
            lunghezza_titolo_20 = len_header_20
            min_col = 1
        # elif len(list(portfolio.loc[portfolio['CATEGORIA']=='GOVERNMENT_BOND','INTERMEDIARIO'].unique())) == 5:
        #     lunghezza_titolo_20 = 9
        # elif len(list(portfolio.loc[portfolio['CATEGORIA']=='GOVERNMENT_BOND','INTERMEDIARIO'].unique())) == 4:
        #     lunghezza_titolo_20 = 9
        #     header_20.insert(len(header_20), '')
        # elif len(list(portfolio.loc[portfolio['CATEGORIA']=='GOVERNMENT_BOND','INTERMEDIARIO'].unique())) == 3:
        #     lunghezza_titolo_20 = 9
        #     header_20.insert(len(header_20), '')
        #     header_20.insert(0, '')
        # len_header_20 = len(header_20)
        # lunghezza_titolo_20 = len(ws20.cell(row=1, column=1).value)
        # print(lunghezza_titolo_20)
        ws20.merge_cells(start_row=1, end_row=4, start_column=1, end_column=lunghezza_titolo_20)

        # Intestazione
        for col in ws20.iter_cols(min_row=8, max_row=9, min_col=min_col, max_col=min_col + len_header_20 - 1):
            ws20[col[0].coordinate].value = header_20[0]
            del header_20[0]
            ws20[col[0].coordinate].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws20[col[0].coordinate].font = Font(name='Times New Roman', size=10, color='FFFFFF')
            ws20[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='31869B')
            ws20[col[0].coordinate].border = Border(right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'))
            ws20.merge_cells(start_row=col[0].row, end_row=col[1].row, start_column=col[0].column, end_column=col[0].column)
            ws20.row_dimensions[col[0].row].height = 20
            ws20.row_dimensions[col[1].row].height = 20
            ws20.column_dimensions[col[0].column_letter].width = 12

        # Indice e riempimento tabella
        nome_obbgov = list(portfolio.loc[portfolio['CATEGORIA']=='GOVERNMENT_BOND','PRODOTTO'])
        len_nome_obbgov = len(nome_obbgov)
        num_intermediari = len(portfolio.loc[portfolio['CATEGORIA']=='GOVERNMENT_BOND', 'INTERMEDIARIO'].unique())
        lunghezza_colonna_20 = []
        for row in ws20.iter_rows(min_row=8, max_row=10 + len_nome_obbgov -1, min_col=min_col, max_col=min_col + len_header_20):
            if row[0].row > 9:
                ws20[row[0].coordinate].value = nome_obbgov[0] 
                del nome_obbgov[0]
                ws20[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
                ws20[row[0].coordinate].font = Font(name='Times New Roman', size=9, color='000000')
                ws20[row[0].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                for _ in range(1, num_intermediari+1):
                    ws20[row[_].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws20[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) & (portfolio['PRODOTTO']==ws20[row[0].coordinate].value), 'TOTALE t1'].sum() if portfolio.loc[(portfolio['INTERMEDIARIO']==ws20[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) & (portfolio['PRODOTTO']==ws20[row[0].coordinate].value), 'TOTALE t1'].sum() != 0 else ''
                    ws20[row[_].coordinate].alignment = Alignment(horizontal='center')
                    ws20[row[_].coordinate].font = Font(name='Times New Roman', size=9)
                    ws20[row[_].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                    ws20[row[_].coordinate].number_format = '#,0'

                ws20[row[num_intermediari+1].coordinate].value = portfolio.loc[portfolio['PRODOTTO']==ws20[row[0].coordinate].value, 'TOTALE t1'].sum()
                ws20[row[num_intermediari+1].coordinate].alignment = Alignment(horizontal='center')
                ws20[row[num_intermediari+1].coordinate].font = Font(name='Times New Roman', size=9)
                ws20[row[num_intermediari+1].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws20[row[num_intermediari+1].coordinate].number_format = '#,0'
                ws20[row[num_intermediari+2].coordinate].value = portfolio.loc[portfolio['PRODOTTO']==ws20[row[0].coordinate].value, 'TOTALE t0'].sum()
                ws20[row[num_intermediari+2].coordinate].alignment = Alignment(horizontal='center')
                ws20[row[num_intermediari+2].coordinate].font = Font(name='Times New Roman', size=9)
                ws20[row[num_intermediari+2].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws20[row[num_intermediari+2].coordinate].number_format = '#,0'
                ws20[row[num_intermediari+3].coordinate].value = (ws20[row[num_intermediari+1].coordinate].value -  ws20[row[num_intermediari+2].coordinate].value) / (ws20[row[num_intermediari+2].coordinate].value) if ws20[row[num_intermediari+2].coordinate].value != 0 and ws20[row[num_intermediari+1].coordinate].value != 0 else '/'
                ws20[row[num_intermediari+3].coordinate].alignment = Alignment(horizontal='center')
                ws20[row[num_intermediari+3].coordinate].font = Font(name='Times New Roman', size=9)
                ws20[row[num_intermediari+3].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws20[row[num_intermediari+3].coordinate].number_format = FORMAT_PERCENTAGE_00

                lunghezza_colonna_20.append(len(ws20.cell(row=row[0].row, column=row[0].column).value)) # ottieni la lunghezza della colonna
                ws20.column_dimensions[row[0].column_letter].width = max(lunghezza_colonna_20) + 2.5 # modifica larghezza colonna FALLO ALLA FINE

        # Somma per intermediari
        for row in ws20.iter_rows(min_row=10 + len_nome_obbgov, max_row=10 + len_nome_obbgov, min_col=min_col, max_col=min_col + len_header_20):
            ws20[row[0].coordinate].value = 'TOTALE'
            ws20[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
            ws20[row[0].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
            ws20[row[0].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
            for _ in range(1,len_header_20-2):
                ws20[row[_].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws20.cell(row=row[_].row, column=row[_].column).offset(row=-len_nome_obbgov-2).value) & (portfolio['CATEGORIA']=='GOVERNMENT_BOND'), 'TOTALE t1'].sum()
            ws20[row[len_header_20-3].coordinate].value = portfolio.loc[portfolio['CATEGORIA']=='GOVERNMENT_BOND', 'TOTALE t1'].sum()
            ws20[row[len_header_20-2].coordinate].value = portfolio.loc[portfolio['CATEGORIA']=='GOVERNMENT_BOND', 'TOTALE t0'].sum()
            ws20[row[len_header_20-1].coordinate].value = (portfolio.loc[(portfolio['CATEGORIA']=='GOVERNMENT_BOND') & (portfolio['TOTALE t0']!=0), 'TOTALE t1'].sum() - portfolio.loc[(portfolio['CATEGORIA']=='GOVERNMENT_BOND') & (portfolio['TOTALE t1']!=0), 'TOTALE t0'].sum()) / portfolio.loc[(portfolio['CATEGORIA']=='GOVERNMENT_BOND') & (portfolio['TOTALE t1']!=0), 'TOTALE t0'].sum()
            for _ in range(1,len_header_20):
                ws20[row[_].coordinate].alignment = Alignment(horizontal='center')
                ws20[row[_].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
                ws20[row[_].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
                ws20[row[_].coordinate].number_format = '#,0'
            ws20[row[len_header_20-1].coordinate].number_format = FORMAT_PERCENTAGE_00

    def obb_corporate_21(self):
        """
        Crea la ventunesima pagina.
        Formattazione e tabella.
        """
        # Carica portafoglio
        portfolio = pd.read_excel(self.file_portafoglio, sheet_name='Portfolio', header=1)

        ws21 = self.wb.create_sheet('21.obb_cor')
        ws21 = self.wb['21.obb_cor']
        self.wb.active = ws21

        # Creazione tabella
        header_21 = list(portfolio.loc[portfolio['CATEGORIA']=='CORPORATE_BOND','INTERMEDIARIO'].unique())
        header_21.insert(0, '')
        header_21.extend(('Totale '+ self.mesi_dict[self.t1.month], 'Totale '+ self.mesi_dict[self.t0_1m.month], 'Delta'))
        len_header_21 = len(header_21)

        # Titolo
        ws21['A1'] = 'Obbligazioni Corporate'
        ws21['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws21['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws21['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        if len(list(portfolio.loc[portfolio['CATEGORIA']=='CORPORATE_BOND','INTERMEDIARIO'].unique())) == 1:
            lunghezza_titolo_21 = 12
            min_col = 4
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='CORPORATE_BOND','INTERMEDIARIO'].unique())) == 2:
            lunghezza_titolo_21 = 12
            min_col = 4
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='CORPORATE_BOND','INTERMEDIARIO'].unique())) == 3:
            lunghezza_titolo_21 = 12
            min_col = 3
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='CORPORATE_BOND','INTERMEDIARIO'].unique())) == 4:
            lunghezza_titolo_21 = 12
            min_col = 3
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='CORPORATE_BOND','INTERMEDIARIO'].unique())) == 5:
            lunghezza_titolo_21 = 12
            min_col = 2
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='CORPORATE_BOND','INTERMEDIARIO'].unique())) == 6:
            lunghezza_titolo_21 = 12
            min_col = 2
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='CORPORATE_BOND','INTERMEDIARIO'].unique())) == 7:
            lunghezza_titolo_21 = 12
            min_col = 1
        else:
            lunghezza_titolo_21 = len_header_21
            min_col = 1
        ws21.merge_cells(start_row=1, end_row=4, start_column=1, end_column=lunghezza_titolo_21)

        # Intestazione
        for col in ws21.iter_cols(min_row=8, max_row=9, min_col=min_col, max_col=min_col + len_header_21 - 1):
            ws21[col[0].coordinate].value = header_21[0]
            del header_21[0]
            ws21[col[0].coordinate].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws21[col[0].coordinate].font = Font(name='Times New Roman', size=10, color='FFFFFF')
            ws21[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='31869B')
            ws21[col[0].coordinate].border = Border(right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'))
            ws21.merge_cells(start_row=col[0].row, end_row=col[1].row, start_column=col[0].column, end_column=col[0].column)
            ws21.row_dimensions[col[0].row].height = 20
            ws21.row_dimensions[col[1].row].height = 20
            ws21.column_dimensions[col[0].column_letter].width = 12

        # Indice e riempimento tabella
        nome_obbcorp = list(portfolio.loc[portfolio['CATEGORIA']=='CORPORATE_BOND','PRODOTTO'])
        len_nome_obbcorp = len(nome_obbcorp)
        num_intermediari = len(portfolio.loc[portfolio['CATEGORIA']=='CORPORATE_BOND', 'INTERMEDIARIO'].unique())
        lunghezza_colonna_21 = []
        for row in ws21.iter_rows(min_row=8, max_row=10 + len_nome_obbcorp -1, min_col=min_col, max_col=min_col + len_header_21):
            if row[0].row > 9:
                ws21[row[0].coordinate].value = nome_obbcorp[0] 
                del nome_obbcorp[0]
                ws21[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
                ws21[row[0].coordinate].font = Font(name='Times New Roman', size=9, color='000000')
                ws21[row[0].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                for _ in range(1, num_intermediari+1):
                    ws21[row[_].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws21[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) & (portfolio['PRODOTTO']==ws21[row[0].coordinate].value), 'TOTALE t1'].sum() if portfolio.loc[(portfolio['INTERMEDIARIO']==ws21[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) & (portfolio['PRODOTTO']==ws21[row[0].coordinate].value), 'TOTALE t1'].sum() != 0 else ''
                    ws21[row[_].coordinate].alignment = Alignment(horizontal='center')
                    ws21[row[_].coordinate].font = Font(name='Times New Roman', size=9)
                    ws21[row[_].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                    ws21[row[_].coordinate].number_format = '#,0'

                ws21[row[num_intermediari+1].coordinate].value = portfolio.loc[portfolio['PRODOTTO']==ws21[row[0].coordinate].value, 'TOTALE t1'].sum()
                ws21[row[num_intermediari+1].coordinate].alignment = Alignment(horizontal='center')
                ws21[row[num_intermediari+1].coordinate].font = Font(name='Times New Roman', size=9)
                ws21[row[num_intermediari+1].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws21[row[num_intermediari+1].coordinate].number_format = '#,0'
                ws21[row[num_intermediari+2].coordinate].value = portfolio.loc[portfolio['PRODOTTO']==ws21[row[0].coordinate].value, 'TOTALE t0'].sum()
                ws21[row[num_intermediari+2].coordinate].alignment = Alignment(horizontal='center')
                ws21[row[num_intermediari+2].coordinate].font = Font(name='Times New Roman', size=9)
                ws21[row[num_intermediari+2].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws21[row[num_intermediari+2].coordinate].number_format = '#,0'
                ws21[row[num_intermediari+3].coordinate].value = (ws21[row[num_intermediari+1].coordinate].value -  ws21[row[num_intermediari+2].coordinate].value) / (ws21[row[num_intermediari+2].coordinate].value) if ws21[row[num_intermediari+2].coordinate].value != 0 and ws21[row[num_intermediari+1].coordinate].value != 0 else '/'
                ws21[row[num_intermediari+3].coordinate].alignment = Alignment(horizontal='center')
                ws21[row[num_intermediari+3].coordinate].font = Font(name='Times New Roman', size=9)
                ws21[row[num_intermediari+3].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws21[row[num_intermediari+3].coordinate].number_format = FORMAT_PERCENTAGE_00

                lunghezza_colonna_21.append(len(ws21.cell(row=row[0].row, column=row[0].column).value)) # ottieni la lunghezza della colonna
                ws21.column_dimensions[row[0].column_letter].width = max(lunghezza_colonna_21) + 2.5 # modifica larghezza colonna FALLO ALLA FINE

        # Somma per intermediari
        for row in ws21.iter_rows(min_row=10 + len_nome_obbcorp, max_row=10 + len_nome_obbcorp, min_col=min_col, max_col=min_col + len_header_21):
            ws21[row[0].coordinate].value = 'TOTALE'
            ws21[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
            ws21[row[0].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
            ws21[row[0].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
            for _ in range(1,len_header_21-2):
                ws21[row[_].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws21.cell(row=row[_].row, column=row[_].column).offset(row=-len_nome_obbcorp-2).value) & (portfolio['CATEGORIA']=='CORPORATE_BOND'), 'TOTALE t1'].sum()
            ws21[row[len_header_21-3].coordinate].value = portfolio.loc[portfolio['CATEGORIA']=='CORPORATE_BOND', 'TOTALE t1'].sum()
            ws21[row[len_header_21-2].coordinate].value = portfolio.loc[portfolio['CATEGORIA']=='CORPORATE_BOND', 'TOTALE t0'].sum()
            ws21[row[len_header_21-1].coordinate].value = (portfolio.loc[(portfolio['CATEGORIA']=='CORPORATE_BOND') & (portfolio['TOTALE t0']!=0), 'TOTALE t1'].sum() - portfolio.loc[(portfolio['CATEGORIA']=='CORPORATE_BOND') & (portfolio['TOTALE t1']!=0), 'TOTALE t0'].sum()) / portfolio.loc[(portfolio['CATEGORIA']=='CORPORATE_BOND') & (portfolio['TOTALE t1']!=0), 'TOTALE t0'].sum()
            for _ in range(1,len_header_21):
                ws21[row[_].coordinate].alignment = Alignment(horizontal='center')
                ws21[row[_].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
                ws21[row[_].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
                ws21[row[_].coordinate].number_format = '#,0'
            ws21[row[len_header_21-1].coordinate].number_format = FORMAT_PERCENTAGE_00

    def obb_totale_22(self):
        """
        Crea la ventiduesima pagina.
        Formattazione e tabella.
        """
        # Carica portafoglio
        portfolio = pd.read_excel(self.file_portafoglio, sheet_name='Portfolio', header=1)

        # 22.Obb. totale
        ws22 = self.wb.create_sheet('22.obb_tot')
        ws22 = self.wb['22.obb_tot']
        self.wb.active = ws22

        # Creazione tabella
        header_22 = ['', 'OBBLIGAZIONE GOVERNATIVE', 'OBBLIGAZIONE CORPORATE', 'Totale '+ self.mesi_dict[self.t1.month], 'Totale '+ self.mesi_dict[self.t0_1m.month]]
        len_header_22 = len(header_22)

        # Titolo
        ws22['A1'] = 'Riepilogo Obbligazioni'
        ws22['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws22['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws22['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws22.merge_cells(start_row=1, end_row=4, start_column=1, end_column=12)
        min_col = 4

        # Intestazione
        for col in ws22.iter_cols(min_row=8, max_row=9, min_col=min_col, max_col=min_col + len_header_22 -1):
            ws22[col[0].coordinate].value = header_22[0]
            del header_22[0]
            ws22[col[0].coordinate].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws22[col[0].coordinate].font = Font(name='Times New Roman', size=10, color='FFFFFF')
            ws22[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='31869B')
            ws22[col[0].coordinate].border = Border(right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'))
            ws22.merge_cells(start_row=col[0].row, end_row=col[1].row, start_column=col[0].column, end_column=col[0].column)
            ws22.row_dimensions[col[0].row].height = 20
            ws22.row_dimensions[col[1].row].height = 20
            ws22.column_dimensions[col[0].column_letter].width = 15

        # Indice e riempimento tabella
        int_obb = list(portfolio.loc[(portfolio['CATEGORIA']=='GOVERNMENT_BOND') | (portfolio['CATEGORIA']=='CORPORATE_BOND'),'INTERMEDIARIO'].unique())
        len_int_obb = len(int_obb)
        #num_intermediari = len(portfolio.loc[portfolio['CATEGORIA']=='CORPORATE_BOND', 'INTERMEDIARIO'].unique())
        num_intermediari = 2
        lunghezza_colonna_22 = []
        for row in ws22.iter_rows(min_row=8, max_row=10 + len_int_obb -1, min_col=min_col, max_col=min_col + len_header_22):
            if row[0].row > 9:
                ws22[row[0].coordinate].value = int_obb[0]
                del int_obb[0]
                ws22[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
                ws22[row[0].coordinate].font = Font(name='Times New Roman', size=9, color='000000')
                ws22[row[0].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))

                ws22[row[1].coordinate].value = portfolio.loc[(portfolio['CATEGORIA']=='GOVERNMENT_BOND') & (portfolio['INTERMEDIARIO']==ws22[row[0].coordinate].value), 'TOTALE t1'].sum() if portfolio.loc[(portfolio['CATEGORIA']=='GOVERNMENT_BOND') & (portfolio['INTERMEDIARIO']==ws22[row[0].coordinate].value), 'TOTALE t1'].sum() != 0 else ''
                ws22[row[1].coordinate].alignment = Alignment(horizontal='center')
                ws22[row[1].coordinate].font = Font(name='Times New Roman', size=9)
                ws22[row[1].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws22[row[1].coordinate].number_format = '#,0'
                ws22[row[2].coordinate].value = portfolio.loc[(portfolio['CATEGORIA']=='CORPORATE_BOND') & (portfolio['INTERMEDIARIO']==ws22[row[0].coordinate].value), 'TOTALE t1'].sum() if portfolio.loc[(portfolio['CATEGORIA']=='CORPORATE_BOND') & (portfolio['INTERMEDIARIO']==ws22[row[0].coordinate].value), 'TOTALE t1'].sum() != 0 else ''
                ws22[row[2].coordinate].alignment = Alignment(horizontal='center')
                ws22[row[2].coordinate].font = Font(name='Times New Roman', size=9)
                ws22[row[2].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws22[row[2].coordinate].number_format = '#,0'

                ws22[row[num_intermediari+1].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws22[row[0].coordinate].value) & ((portfolio['CATEGORIA']=='GOVERNMENT_BOND') | (portfolio['CATEGORIA']=='CORPORATE_BOND')), 'TOTALE t1'].sum()
                ws22[row[num_intermediari+1].coordinate].alignment = Alignment(horizontal='center')
                ws22[row[num_intermediari+1].coordinate].font = Font(name='Times New Roman', size=9)
                ws22[row[num_intermediari+1].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws22[row[num_intermediari+1].coordinate].number_format = '#,0'
                ws22[row[num_intermediari+2].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws22[row[0].coordinate].value) & ((portfolio['CATEGORIA']=='GOVERNMENT_BOND') | (portfolio['CATEGORIA']=='CORPORATE_BOND')), 'TOTALE t0'].sum()
                ws22[row[num_intermediari+2].coordinate].alignment = Alignment(horizontal='center')
                ws22[row[num_intermediari+2].coordinate].font = Font(name='Times New Roman', size=9)
                ws22[row[num_intermediari+2].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws22[row[num_intermediari+2].coordinate].number_format = '#,0'

                lunghezza_colonna_22.append(len(ws22.cell(row=row[0].row, column=row[0].column).value)) # ottieni la lunghezza della colonna
                ws22.column_dimensions[row[0].column_letter].width = max(lunghezza_colonna_22) + 2.5 # modifica larghezza colonna FALLO ALLA FINE

        # Somma per strumento
        for row in ws22.iter_rows(min_row=10 + len_int_obb, max_row=10 + len_int_obb, min_col=min_col, max_col=min_col + len_header_22):
            ws22[row[0].coordinate].value = 'TOTALE'
            ws22[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
            ws22[row[0].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
            ws22[row[0].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
            
            ws22[row[1].coordinate].value = portfolio.loc[portfolio['CATEGORIA']=='GOVERNMENT_BOND', 'TOTALE t1'].sum()
            ws22[row[2].coordinate].value = portfolio.loc[portfolio['CATEGORIA']=='CORPORATE_BOND', 'TOTALE t1'].sum()
            
            ws22[row[len_header_22-2].coordinate].value = portfolio.loc[(portfolio['CATEGORIA']=='GOVERNMENT_BOND') | (portfolio['CATEGORIA']=='CORPORATE_BOND'), 'TOTALE t1'].sum()
            ws22[row[len_header_22-1].coordinate].value = portfolio.loc[(portfolio['CATEGORIA']=='GOVERNMENT_BOND') | (portfolio['CATEGORIA']=='CORPORATE_BOND'), 'TOTALE t0'].sum()

            for _ in range(1,len_header_22):
                ws22[row[_].coordinate].alignment = Alignment(horizontal='center')
                ws22[row[_].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
                ws22[row[_].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
                ws22[row[_].coordinate].number_format = '#,0'

    def liquidità_23(self):
        """
        Crea la ventitreesima pagina.
        Formattazione e tabella.
        """
        # Carica portafoglio
        portfolio = pd.read_excel(self.file_portafoglio, sheet_name='Portfolio', header=1)

        ws23 = self.wb.create_sheet('23.liq')
        ws23 = self.wb['23.liq']
        self.wb.active = ws23

        # Creazione tabella
        header_23 = list(portfolio.loc[(portfolio['CATEGORIA']=='CASH') | (portfolio['CATEGORIA']=='CASH_FOREIGN_CURR'),'INTERMEDIARIO'].unique())
        header_23.insert(0, '')
        header_23.extend(('Totale '+ self.mesi_dict[self.t1.month], 'Totale '+ self.mesi_dict[self.t0_1m.month]))
        len_header_23 = len(header_23)

        # Titolo
        ws23['A1'] = 'Liquidità'
        ws23['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws23['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws23['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        if len(list(portfolio.loc[(portfolio['CATEGORIA']=='CASH') | (portfolio['CATEGORIA']=='CASH_FOREIGN_CURR'),'INTERMEDIARIO'].unique())) == 1:
            lunghezza_titolo_23 = 12
            min_col = 4
        elif len(list(portfolio.loc[(portfolio['CATEGORIA']=='CASH') | (portfolio['CATEGORIA']=='CASH_FOREIGN_CURR'),'INTERMEDIARIO'].unique())) == 2:
            lunghezza_titolo_23 = 12
            min_col = 4
        elif len(list(portfolio.loc[(portfolio['CATEGORIA']=='CASH') | (portfolio['CATEGORIA']=='CASH_FOREIGN_CURR'),'INTERMEDIARIO'].unique())) == 3:
            lunghezza_titolo_23 = 12
            min_col = 3
        elif len(list(portfolio.loc[(portfolio['CATEGORIA']=='CASH') | (portfolio['CATEGORIA']=='CASH_FOREIGN_CURR'),'INTERMEDIARIO'].unique())) == 4:
            lunghezza_titolo_23 = 12
            min_col = 3
        elif len(list(portfolio.loc[(portfolio['CATEGORIA']=='CASH') | (portfolio['CATEGORIA']=='CASH_FOREIGN_CURR'),'INTERMEDIARIO'].unique())) == 5:
            lunghezza_titolo_23 = 12
            min_col = 2
        elif len(list(portfolio.loc[(portfolio['CATEGORIA']=='CASH') | (portfolio['CATEGORIA']=='CASH_FOREIGN_CURR'),'INTERMEDIARIO'].unique())) == 6:
            lunghezza_titolo_23 = 12
            min_col = 2
        elif len(list(portfolio.loc[(portfolio['CATEGORIA']=='CASH') | (portfolio['CATEGORIA']=='CASH_FOREIGN_CURR'),'INTERMEDIARIO'].unique())) == 7:
            lunghezza_titolo_23 = 12
            min_col = 1
        else:
            lunghezza_titolo_23 = len_header_23
            min_col = 1
        ws23.merge_cells(start_row=1, end_row=4, start_column=1, end_column=lunghezza_titolo_23)

        # Intestazione
        for col in ws23.iter_cols(min_row=8, max_row=9, min_col=min_col, max_col=min_col + len_header_23 - 1):
            ws23[col[0].coordinate].value = header_23[0]
            del header_23[0]
            ws23[col[0].coordinate].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws23[col[0].coordinate].font = Font(name='Times New Roman', size=10, color='FFFFFF')
            ws23[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='31869B')
            ws23[col[0].coordinate].border = Border(right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'))
            ws23.merge_cells(start_row=col[0].row, end_row=col[1].row, start_column=col[0].column, end_column=col[0].column)
            ws23.row_dimensions[col[0].row].height = 20
            ws23.row_dimensions[col[1].row].height = 20
            ws23.column_dimensions[col[0].column_letter].width = 12

        # Indice e riempimento tabella
        nome_liq = list(portfolio.loc[(portfolio['CATEGORIA']=='CASH') | (portfolio['CATEGORIA']=='CASH_FOREIGN_CURR'),'PRODOTTO'])
        len_nome_liq = len(nome_liq)
        num_intermediari = len(portfolio.loc[(portfolio['CATEGORIA']=='CASH') | (portfolio['CATEGORIA']=='CASH_FOREIGN_CURR'), 'INTERMEDIARIO'].unique())
        lunghezza_colonna_23 = []
        for row in ws23.iter_rows(min_row=8, max_row=10 + len_nome_liq -1, min_col=min_col, max_col=min_col + len_header_23):
            if row[0].row > 9:
                ws23[row[0].coordinate].value = nome_liq[0]
                del nome_liq[0]
                ws23[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
                ws23[row[0].coordinate].font = Font(name='Times New Roman', size=9, color='000000')
                ws23[row[0].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                for _ in range(1, num_intermediari+1):
                    ws23[row[_].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws23[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) & (portfolio['PRODOTTO']==ws23[row[0].coordinate].value), 'TOTALE t1'].sum() if portfolio.loc[(portfolio['INTERMEDIARIO']==ws23[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) & (portfolio['PRODOTTO']==ws23[row[0].coordinate].value), 'TOTALE t1'].sum() != 0 else ''
                    ws23[row[_].coordinate].alignment = Alignment(horizontal='center')
                    ws23[row[_].coordinate].font = Font(name='Times New Roman', size=9)
                    ws23[row[_].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                    ws23[row[_].coordinate].number_format = '#,0'

                ws23[row[num_intermediari+1].coordinate].value = portfolio.loc[portfolio['PRODOTTO']==ws23[row[0].coordinate].value, 'TOTALE t1'].sum()
                ws23[row[num_intermediari+1].coordinate].alignment = Alignment(horizontal='center')
                ws23[row[num_intermediari+1].coordinate].font = Font(name='Times New Roman', size=9)
                ws23[row[num_intermediari+1].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws23[row[num_intermediari+1].coordinate].number_format = '#,0'
                ws23[row[num_intermediari+2].coordinate].value = portfolio.loc[portfolio['PRODOTTO']==ws23[row[0].coordinate].value, 'TOTALE t0'].sum()
                ws23[row[num_intermediari+2].coordinate].alignment = Alignment(horizontal='center')
                ws23[row[num_intermediari+2].coordinate].font = Font(name='Times New Roman', size=9)
                ws23[row[num_intermediari+2].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws23[row[num_intermediari+2].coordinate].number_format = '#,0'

                lunghezza_colonna_23.append(len(ws23.cell(row=row[0].row, column=row[0].column).value)) # ottieni la lunghezza della colonna
                ws23.column_dimensions[row[0].column_letter].width = max(lunghezza_colonna_23) + 2.5 # modifica larghezza colonna

        # Somma per intermediari
        for row in ws23.iter_rows(min_row=10 + len_nome_liq, max_row=10 + len_nome_liq, min_col=min_col, max_col=min_col + len_header_23):
            ws23[row[0].coordinate].value = 'TOTALE'
            ws23[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
            ws23[row[0].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
            ws23[row[0].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
            for _ in range(1,len_header_23-2):
                ws23[row[_].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws23.cell(row=row[_].row, column=row[_].column).offset(row=-len_nome_liq-2).value) & ((portfolio['CATEGORIA']=='CASH') | (portfolio['CATEGORIA']=='CASH_FOREIGN_CURR')), 'TOTALE t1'].sum()
            ws23[row[len_header_23-2].coordinate].value = portfolio.loc[(portfolio['CATEGORIA']=='CASH') | (portfolio['CATEGORIA']=='CASH_FOREIGN_CURR'), 'TOTALE t1'].sum()
            ws23[row[len_header_23-1].coordinate].value = portfolio.loc[(portfolio['CATEGORIA']=='CASH') | (portfolio['CATEGORIA']=='CASH_FOREIGN_CURR'), 'TOTALE t0'].sum()
            for _ in range(1,len_header_23):
                ws23[row[_].coordinate].alignment = Alignment(horizontal='center')
                ws23[row[_].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
                ws23[row[_].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
                ws23[row[_].coordinate].number_format = '#,0'

    def liq_totale_24(self):
        """
        Crea la ventiquattresima pagina.
        Formattazione e tabella.
        """
        # Carica portafoglio
        portfolio = pd.read_excel(self.file_portafoglio, sheet_name='Portfolio', header=1)

        ws24 = self.wb.create_sheet('24.liq_tot')
        ws24 = self.wb['24.liq_tot']
        self.wb.active = ws24

        # Creazione tabella
        header_24 = ['', 'LIQUIDITÀ', 'LIQUIDITÀ IN VALUTA', 'Totale '+ self.mesi_dict[self.t1.month], 'Totale '+ self.mesi_dict[self.t0_1m.month]]
        len_header_24 = len(header_24)

        # Titolo
        ws24['A1'] = 'Riepilogo Liquidità'
        ws24['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws24['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws24['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws24.merge_cells(start_row=1, end_row=4, start_column=1, end_column=12)
        min_col = 4

        # Intestazione
        for col in ws24.iter_cols(min_row=8, max_row=9, min_col=min_col, max_col=min_col + len_header_24 -1):
            ws24[col[0].coordinate].value = header_24[0]
            del header_24[0]
            ws24[col[0].coordinate].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws24[col[0].coordinate].font = Font(name='Times New Roman', size=10, color='FFFFFF')
            ws24[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='31869B')
            ws24[col[0].coordinate].border = Border(right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'))
            ws24.merge_cells(start_row=col[0].row, end_row=col[1].row, start_column=col[0].column, end_column=col[0].column)
            ws24.row_dimensions[col[0].row].height = 20
            ws24.row_dimensions[col[1].row].height = 20
            ws24.column_dimensions[col[0].column_letter].width = 15

        # Indice e riempimento tabella
        int_liq = list(portfolio.loc[(portfolio['CATEGORIA']=='CASH') | (portfolio['CATEGORIA']=='CASH_FOREIGN_CURR'),'INTERMEDIARIO'].unique())
        len_int_liq = len(int_liq)
        num_intermediari = 2
        lunghezza_colonna_24 = []
        for row in ws24.iter_rows(min_row=8, max_row=10 + len_int_liq -1, min_col=min_col, max_col=min_col + len_header_24):
            if row[0].row > 9:
                ws24[row[0].coordinate].value = int_liq[0]
                del int_liq[0]
                ws24[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
                ws24[row[0].coordinate].font = Font(name='Times New Roman', size=9, color='000000')
                ws24[row[0].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))

                ws24[row[1].coordinate].value = portfolio.loc[(portfolio['CATEGORIA']=='CASH') & (portfolio['INTERMEDIARIO']==ws24[row[0].coordinate].value), 'TOTALE t1'].sum() if portfolio.loc[(portfolio['CATEGORIA']=='CASH') & (portfolio['INTERMEDIARIO']==ws24[row[0].coordinate].value), 'TOTALE t1'].sum() != 0 else ''
                ws24[row[1].coordinate].alignment = Alignment(horizontal='center')
                ws24[row[1].coordinate].font = Font(name='Times New Roman', size=9)
                ws24[row[1].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws24[row[1].coordinate].number_format = '#,0'
                ws24[row[2].coordinate].value = portfolio.loc[(portfolio['CATEGORIA']=='CASH_FOREIGN_CURR') & (portfolio['INTERMEDIARIO']==ws24[row[0].coordinate].value), 'TOTALE t1'].sum() if portfolio.loc[(portfolio['CATEGORIA']=='CASH_FOREIGN_CURR') & (portfolio['INTERMEDIARIO']==ws24[row[0].coordinate].value), 'TOTALE t1'].sum() != 0 else ''
                ws24[row[2].coordinate].alignment = Alignment(horizontal='center')
                ws24[row[2].coordinate].font = Font(name='Times New Roman', size=9)
                ws24[row[2].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws24[row[2].coordinate].number_format = '#,0'

                ws24[row[num_intermediari+1].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws24[row[0].coordinate].value) & ((portfolio['CATEGORIA']=='CASH') | (portfolio['CATEGORIA']=='CASH_FOREIGN_CURR')), 'TOTALE t1'].sum()
                ws24[row[num_intermediari+1].coordinate].alignment = Alignment(horizontal='center')
                ws24[row[num_intermediari+1].coordinate].font = Font(name='Times New Roman', size=9)
                ws24[row[num_intermediari+1].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws24[row[num_intermediari+1].coordinate].number_format = '#,0'
                ws24[row[num_intermediari+2].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws24[row[0].coordinate].value) & ((portfolio['CATEGORIA']=='CASH') | (portfolio['CATEGORIA']=='CASH_FOREIGN_CURR')), 'TOTALE t0'].sum()
                ws24[row[num_intermediari+2].coordinate].alignment = Alignment(horizontal='center')
                ws24[row[num_intermediari+2].coordinate].font = Font(name='Times New Roman', size=9)
                ws24[row[num_intermediari+2].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws24[row[num_intermediari+2].coordinate].number_format = '#,0'

                lunghezza_colonna_24.append(len(ws24.cell(row=row[0].row, column=row[0].column).value)) # ottieni la lunghezza della colonna
                ws24.column_dimensions[row[0].column_letter].width = max(lunghezza_colonna_24) + 2.5 # modifica larghezza colonna FALLO ALLA FINE

        # Somma per strumento
        for row in ws24.iter_rows(min_row=10 + len_int_liq, max_row=10 + len_int_liq, min_col=min_col, max_col=min_col + len_header_24):
            ws24[row[0].coordinate].value = 'TOTALE'
            ws24[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
            ws24[row[0].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
            ws24[row[0].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
            
            ws24[row[1].coordinate].value = portfolio.loc[portfolio['CATEGORIA']=='CASH', 'TOTALE t1'].sum()
            ws24[row[2].coordinate].value = portfolio.loc[portfolio['CATEGORIA']=='CASH_FOREIGN_CURR', 'TOTALE t1'].sum()
            
            ws24[row[len_header_24-2].coordinate].value = portfolio.loc[(portfolio['CATEGORIA']=='CASH') | (portfolio['CATEGORIA']=='CASH_FOREIGN_CURR'), 'TOTALE t1'].sum()
            ws24[row[len_header_24-1].coordinate].value = portfolio.loc[(portfolio['CATEGORIA']=='CASH') | (portfolio['CATEGORIA']=='CASH_FOREIGN_CURR'), 'TOTALE t0'].sum()

            for _ in range(1,len_header_24):
                ws24[row[_].coordinate].alignment = Alignment(horizontal='center')
                ws24[row[_].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
                ws24[row[_].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
                ws24[row[_].coordinate].number_format = '#,0'

    def gestioni_25(self):
        """
        Crea la venticinquesima pagina.
        Formattazione e tabella.
        """
        # Carica portafoglio
        portfolio = pd.read_excel(self.file_portafoglio, sheet_name='Portfolio', header=1)

        # 25.Gestione
        ws25 = self.wb.create_sheet('25.ges')
        ws25 = self.wb['25.ges']
        self.wb.active = ws25

        # Creazione tabella
        header_25 = list(portfolio.loc[portfolio['CATEGORIA']=='GP', 'INTERMEDIARIO'].unique())
        header_25.insert(0, '')
        header_25.extend(('Totale '+ self.mesi_dict[self.t1.month], 'Totale '+ self.mesi_dict[self.t0_1m.month], 'Delta'))
        len_header_25 = len(header_25)

        # Titolo
        ws25['A1'] = 'Gestioni'
        ws25['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws25['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws25['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        if len(list(portfolio.loc[portfolio['CATEGORIA']=='GP','INTERMEDIARIO'].unique())) == 1:
            lunghezza_titolo_25 = 12
            min_col = 4
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='GP','INTERMEDIARIO'].unique())) == 2:
            lunghezza_titolo_25 = 12
            min_col = 4
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='GP','INTERMEDIARIO'].unique())) == 3:
            lunghezza_titolo_25 = 12
            min_col = 3
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='GP','INTERMEDIARIO'].unique())) == 4:
            lunghezza_titolo_25 = 12
            min_col = 3
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='GP','INTERMEDIARIO'].unique())) == 5:
            lunghezza_titolo_25 = 12
            min_col = 2
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='GP','INTERMEDIARIO'].unique())) == 6:
            lunghezza_titolo_25 = 12
            min_col = 2
        elif len(list(portfolio.loc[portfolio['CATEGORIA']=='GP','INTERMEDIARIO'].unique())) == 7:
            lunghezza_titolo_25 = 12
            min_col = 1
        else:
            lunghezza_titolo_25 = len_header_25
            min_col = 1
        ws25.merge_cells(start_row=1, end_row=4, start_column=1, end_column=lunghezza_titolo_25)

        # Intestazione
        for col in ws25.iter_cols(min_row=8, max_row=9, min_col=min_col, max_col=min_col + len_header_25 - 1):
            ws25[col[0].coordinate].value = header_25[0]
            del header_25[0]
            ws25[col[0].coordinate].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws25[col[0].coordinate].font = Font(name='Times New Roman', size=10, color='FFFFFF')
            ws25[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='31869B')
            ws25[col[0].coordinate].border = Border(right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'))
            ws25.merge_cells(start_row=col[0].row, end_row=col[1].row, start_column=col[0].column, end_column=col[0].column)
            ws25.row_dimensions[col[0].row].height = 20
            ws25.row_dimensions[col[1].row].height = 20
            ws25.column_dimensions[col[0].column_letter].width = 12

        # Indice e riempimento tabella
        nome_ges = list(portfolio.loc[portfolio['CATEGORIA']=='GP','PRODOTTO'])
        len_nome_ges = len(nome_ges)
        num_intermediari = len(portfolio.loc[portfolio['CATEGORIA']=='GP', 'INTERMEDIARIO'].unique())
        lunghezza_colonna_25 = []
        for row in ws25.iter_rows(min_row=8, max_row=10 + len_nome_ges -1, min_col=min_col, max_col=min_col + len_header_25):
            if row[0].row > 9:
                ws25[row[0].coordinate].value = nome_ges[0] 
                del nome_ges[0]
                ws25[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
                ws25[row[0].coordinate].font = Font(name='Times New Roman', size=9, color='000000')
                ws25[row[0].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                for _ in range(1, num_intermediari+1):
                    ws25[row[_].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws25[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) & (portfolio['PRODOTTO']==ws25[row[0].coordinate].value), 'TOTALE t1'].sum() if portfolio.loc[(portfolio['INTERMEDIARIO']==ws25[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) & (portfolio['PRODOTTO']==ws25[row[0].coordinate].value), 'TOTALE t1'].sum() != 0 else ''
                    ws25[row[_].coordinate].alignment = Alignment(horizontal='center')
                    ws25[row[_].coordinate].font = Font(name='Times New Roman', size=9)
                    ws25[row[_].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                    ws25[row[_].coordinate].number_format = '#,0'

                ws25[row[num_intermediari+1].coordinate].value = portfolio.loc[portfolio['PRODOTTO']==ws25[row[0].coordinate].value, 'TOTALE t1'].sum()
                ws25[row[num_intermediari+1].coordinate].alignment = Alignment(horizontal='center')
                ws25[row[num_intermediari+1].coordinate].font = Font(name='Times New Roman', size=9)
                ws25[row[num_intermediari+1].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws25[row[num_intermediari+1].coordinate].number_format = '#,0'
                ws25[row[num_intermediari+2].coordinate].value = portfolio.loc[portfolio['PRODOTTO']==ws25[row[0].coordinate].value, 'TOTALE t0'].sum()
                ws25[row[num_intermediari+2].coordinate].alignment = Alignment(horizontal='center')
                ws25[row[num_intermediari+2].coordinate].font = Font(name='Times New Roman', size=9)
                ws25[row[num_intermediari+2].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws25[row[num_intermediari+2].coordinate].number_format = '#,0'
                ws25[row[num_intermediari+3].coordinate].value = (ws25[row[num_intermediari+1].coordinate].value -  ws25[row[num_intermediari+2].coordinate].value) / (ws25[row[num_intermediari+2].coordinate].value) if ws25[row[num_intermediari+2].coordinate].value != 0 and ws25[row[num_intermediari+1].coordinate].value != 0 else '/'
                ws25[row[num_intermediari+3].coordinate].alignment = Alignment(horizontal='center')
                ws25[row[num_intermediari+3].coordinate].font = Font(name='Times New Roman', size=9)
                ws25[row[num_intermediari+3].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws25[row[num_intermediari+3].coordinate].number_format = FORMAT_PERCENTAGE_00

                lunghezza_colonna_25.append(len(ws25.cell(row=row[0].row, column=row[0].column).value)) # ottieni la lunghezza della colonna
                ws25.column_dimensions[row[0].column_letter].width = max(lunghezza_colonna_25) + 3.5 # modifica larghezza colonna 
            
        # Somma per intermediari
        for row in ws25.iter_rows(min_row=10 + len_nome_ges, max_row=10 + len_nome_ges, min_col=min_col, max_col=min_col + len_header_25):
            ws25[row[0].coordinate].value = 'TOTALE'
            ws25[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
            ws25[row[0].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
            ws25[row[0].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
            for _ in range(1,len_header_25-2):
                ws25[row[_].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws25.cell(row=row[_].row, column=row[_].column).offset(row=-len_nome_ges-2).value) & ((portfolio['CATEGORIA']=='GP')), 'TOTALE t1'].sum()
            ws25[row[len_header_25-3].coordinate].value = portfolio.loc[portfolio['CATEGORIA']=='GP', 'TOTALE t1'].sum()
            ws25[row[len_header_25-2].coordinate].value = portfolio.loc[portfolio['CATEGORIA']=='GP', 'TOTALE t0'].sum()
            ws25[row[len_header_25-1].coordinate].value = (portfolio.loc[(portfolio['CATEGORIA']=='GP') & (portfolio['TOTALE t0']!=0), 'TOTALE t1'].sum() - portfolio.loc[(portfolio['CATEGORIA']=='GP') & (portfolio['TOTALE t1']!=0), 'TOTALE t0'].sum()) / portfolio.loc[(portfolio['CATEGORIA']=='GP') & (portfolio['TOTALE t1']!=0), 'TOTALE t0'].sum()
            for _ in range(1,len_header_25):
                ws25[row[_].coordinate].alignment = Alignment(horizontal='center')
                ws25[row[_].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
                ws25[row[_].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
                ws25[row[_].coordinate].number_format = '#,0'
            ws25[row[len_header_25-1].coordinate].number_format = FORMAT_PERCENTAGE_00

    def inv_alt_26(self):
        """
        Crea la ventiseiesima pagina.
        Formattazione e tabella.
        """
        # Carica portafoglio
        portfolio = pd.read_excel(self.file_portafoglio, sheet_name='Portfolio', header=1)


        # 26.Inv.Alt
        ws26 = self.wb.create_sheet('26.invalt')
        ws26 = self.wb['26.invalt']
        self.wb.active = ws26

        # Creazione tabella
        header_26 = list(portfolio.loc[(portfolio['CATEGORIA']=='HEDGE_FUND') | (portfolio['CATEGORIA']=='ALTERNATIVE_ASSET'), 'INTERMEDIARIO'].unique())
        header_26.insert(0, '')
        header_26.extend(('Totale '+ self.mesi_dict[self.t1.month], 'Totale '+ self.mesi_dict[self.t0_1m.month], 'Delta'))
        len_header_26 = len(header_26)

        # Titolo
        ws26['A1'] = 'Inv. Alt. e Hedge Fund'
        ws26['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws26['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws26['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        if len(list(portfolio.loc[(portfolio['CATEGORIA']=='HEDGE_FUND') | (portfolio['CATEGORIA']=='ALTERNATIVE_ASSET'),'INTERMEDIARIO'].unique())) == 1:
            lunghezza_titolo_26 = 12
            min_col = 4
        elif len(list(portfolio.loc[(portfolio['CATEGORIA']=='HEDGE_FUND') | (portfolio['CATEGORIA']=='ALTERNATIVE_ASSET'),'INTERMEDIARIO'].unique())) == 2:
            lunghezza_titolo_26 = 12
            min_col = 4
        elif len(list(portfolio.loc[(portfolio['CATEGORIA']=='HEDGE_FUND') | (portfolio['CATEGORIA']=='ALTERNATIVE_ASSET'),'INTERMEDIARIO'].unique())) == 3:
            lunghezza_titolo_26 = 12
            min_col = 3
        elif len(list(portfolio.loc[(portfolio['CATEGORIA']=='HEDGE_FUND') | (portfolio['CATEGORIA']=='ALTERNATIVE_ASSET'),'INTERMEDIARIO'].unique())) == 4:
            lunghezza_titolo_26 = 12
            min_col = 3
        elif len(list(portfolio.loc[(portfolio['CATEGORIA']=='HEDGE_FUND') | (portfolio['CATEGORIA']=='ALTERNATIVE_ASSET'),'INTERMEDIARIO'].unique())) == 5:
            lunghezza_titolo_26 = 12
            min_col = 2
        elif len(list(portfolio.loc[(portfolio['CATEGORIA']=='HEDGE_FUND') | (portfolio['CATEGORIA']=='ALTERNATIVE_ASSET'),'INTERMEDIARIO'].unique())) == 6:
            lunghezza_titolo_26 = 12
            min_col = 2
        elif len(list(portfolio.loc[(portfolio['CATEGORIA']=='HEDGE_FUND') | (portfolio['CATEGORIA']=='ALTERNATIVE_ASSET'),'INTERMEDIARIO'].unique())) == 7:
            lunghezza_titolo_26 = 12
            min_col = 1
        else:
            lunghezza_titolo_26 = len_header_26
            min_col = 1
        ws26.merge_cells(start_row=1, end_row=4, start_column=1, end_column=lunghezza_titolo_26)

        # Intestazione
        for col in ws26.iter_cols(min_row=8, max_row=9, min_col=min_col, max_col=min_col + len_header_26 - 1):
            ws26[col[0].coordinate].value = header_26[0]
            del header_26[0]
            ws26[col[0].coordinate].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws26[col[0].coordinate].font = Font(name='Times New Roman', size=10, color='FFFFFF')
            ws26[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='31869B')
            ws26[col[0].coordinate].border = Border(right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'))
            ws26.merge_cells(start_row=col[0].row, end_row=col[1].row, start_column=col[0].column, end_column=col[0].column)
            ws26.row_dimensions[col[0].row].height = 20
            ws26.row_dimensions[col[1].row].height = 20
            ws26.column_dimensions[col[0].column_letter].width = 12

        # Indice e riempimento tabella
        nome_invalt = list(portfolio.loc[(portfolio['CATEGORIA']=='HEDGE_FUND') | (portfolio['CATEGORIA']=='ALTERNATIVE_ASSET'),'PRODOTTO'])
        len_nome_invalt = len(nome_invalt)
        num_intermediari = len(portfolio.loc[(portfolio['CATEGORIA']=='HEDGE_FUND') | (portfolio['CATEGORIA']=='ALTERNATIVE_ASSET'), 'INTERMEDIARIO'].unique())
        lunghezza_colonna_26 = []
        for row in ws26.iter_rows(min_row=8, max_row=10 + len_nome_invalt -1, min_col=min_col, max_col=min_col + len_header_26):
            if row[0].row > 9:
                ws26[row[0].coordinate].value = nome_invalt[0] 
                del nome_invalt[0]
                ws26[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
                ws26[row[0].coordinate].font = Font(name='Times New Roman', size=9, color='000000')
                ws26[row[0].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                for _ in range(1, num_intermediari+1):
                    ws26[row[_].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws26[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) & (portfolio['PRODOTTO']==ws26[row[0].coordinate].value), 'TOTALE t1'].sum() if portfolio.loc[(portfolio['INTERMEDIARIO']==ws26[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) & (portfolio['PRODOTTO']==ws26[row[0].coordinate].value), 'TOTALE t1'].sum() != 0 else ''
                    ws26[row[_].coordinate].alignment = Alignment(horizontal='center')
                    ws26[row[_].coordinate].font = Font(name='Times New Roman', size=9)
                    ws26[row[_].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                    ws26[row[_].coordinate].number_format = '#,0'

                ws26[row[num_intermediari+1].coordinate].value = portfolio.loc[portfolio['PRODOTTO']==ws26[row[0].coordinate].value, 'TOTALE t1'].sum()
                ws26[row[num_intermediari+1].coordinate].alignment = Alignment(horizontal='center')
                ws26[row[num_intermediari+1].coordinate].font = Font(name='Times New Roman', size=9)
                ws26[row[num_intermediari+1].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws26[row[num_intermediari+1].coordinate].number_format = '#,0'
                ws26[row[num_intermediari+2].coordinate].value = portfolio.loc[portfolio['PRODOTTO']==ws26[row[0].coordinate].value, 'TOTALE t0'].sum()
                ws26[row[num_intermediari+2].coordinate].alignment = Alignment(horizontal='center')
                ws26[row[num_intermediari+2].coordinate].font = Font(name='Times New Roman', size=9)
                ws26[row[num_intermediari+2].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws26[row[num_intermediari+2].coordinate].number_format = '#,0'
                ws26[row[num_intermediari+3].coordinate].value = (ws26[row[num_intermediari+1].coordinate].value -  ws26[row[num_intermediari+2].coordinate].value) / (ws26[row[num_intermediari+2].coordinate].value) if ws26[row[num_intermediari+2].coordinate].value != 0 and ws26[row[num_intermediari+1].coordinate].value != 0 else '/'
                ws26[row[num_intermediari+3].coordinate].alignment = Alignment(horizontal='center')
                ws26[row[num_intermediari+3].coordinate].font = Font(name='Times New Roman', size=9)
                ws26[row[num_intermediari+3].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws26[row[num_intermediari+3].coordinate].number_format = FORMAT_PERCENTAGE_00

                lunghezza_colonna_26.append(len(ws26.cell(row=row[0].row, column=row[0].column).value)) # ottieni la lunghezza della colonna
                ws26.column_dimensions[row[0].column_letter].width = max(lunghezza_colonna_26) + 2.5 # modifica larghezza colonna

        # Somma per intermediari
        for row in ws26.iter_rows(min_row=10 + len_nome_invalt, max_row=10 + len_nome_invalt, min_col=min_col, max_col=min_col + len_header_26):
            ws26[row[0].coordinate].value = 'TOTALE'
            ws26[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
            ws26[row[0].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
            ws26[row[0].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
            for _ in range(1,len_header_26-2):
                ws26[row[_].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws26.cell(row=row[_].row, column=row[_].column).offset(row=-len_nome_invalt-2).value) & ((portfolio['CATEGORIA']=='HEDGE_FUND') | (portfolio['CATEGORIA']=='ALTERNATIVE_ASSET')), 'TOTALE t1'].sum()
            ws26[row[len_header_26-3].coordinate].value = portfolio.loc[(portfolio['CATEGORIA']=='HEDGE_FUND') | (portfolio['CATEGORIA']=='ALTERNATIVE_ASSET'), 'TOTALE t1'].sum()
            ws26[row[len_header_26-2].coordinate].value = portfolio.loc[(portfolio['CATEGORIA']=='HEDGE_FUND') | (portfolio['CATEGORIA']=='ALTERNATIVE_ASSET'), 'TOTALE t0'].sum()
            ws26[row[len_header_26-1].coordinate].value = (portfolio.loc[(portfolio['CATEGORIA']=='HEDGE_FUND') | (portfolio['CATEGORIA']=='ALTERNATIVE_ASSET') & (portfolio['TOTALE t0']!=0), 'TOTALE t1'].sum() - portfolio.loc[(portfolio['CATEGORIA']=='HEDGE_FUND') | (portfolio['CATEGORIA']=='ALTERNATIVE_ASSET') & (portfolio['TOTALE t1']!=0), 'TOTALE t0'].sum()) / portfolio.loc[(portfolio['CATEGORIA']=='HEDGE_FUND') | (portfolio['CATEGORIA']=='ALTERNATIVE_ASSET') & (portfolio['TOTALE t1']!=0), 'TOTALE t0'].sum()
            for _ in range(1,len_header_26):
                ws26[row[_].coordinate].alignment = Alignment(horizontal='center')
                ws26[row[_].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
                ws26[row[_].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
                ws26[row[_].coordinate].number_format = '#,0'
            ws26[row[len_header_26-1].coordinate].number_format = FORMAT_PERCENTAGE_00

    def asset_allocation_27(self):
        # TODO : i pezzi di torta devono avere sempre gli stessi colori
        """
        Crea la ventisettesima pagina.
        Formattazione, tabella e grafico.
        """
        # Carica portafoglio
        portfolio = pd.read_excel(self.file_portafoglio, sheet_name='Portfolio', header=1)
        # Carica asset-allocation gestioni
        ass_allocation = pd.read_excel(self.file_portafoglio, sheet_name='Gestioni', header=0)

        # 27.Sintesi
        ws27 = self.wb.create_sheet('27.ass_all')
        ws27 = self.wb['27.ass_all']
        self.wb.active = ws27

        # Creazione tabella
        header_27 = list(portfolio['INTERMEDIARIO'].unique())
        header_27.insert(0, '')
        header_27.extend(('Totale '+ self.mesi_dict[self.t1.month], 'Totale '+ self.mesi_dict[self.t0_1m.month]))
        len_header_27 = len(header_27)
        # index_banca_generali = header_27.index('Banca Generali')
        # index_cassa_lombarda = header_27.index('Banca Generali')
        # Cerca le posizioni contenenti almeno una gestione patrimoniale
        posizioni_con_gestioni = []
        for intermediario in header_27:
            if 'GP' in portfolio.loc[portfolio['INTERMEDIARIO']==intermediario, 'CATEGORIA'].unique():
                posizioni_con_gestioni.append(intermediario)
        #print(posizioni_con_gestioni)

        # Titolo
        ws27['A1'] = 'Asset Allocation'
        ws27['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws27['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws27['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        if len(list(portfolio['INTERMEDIARIO'].unique())) == 1:
            lunghezza_titolo_27 = 12
            min_col = 4
        elif len(list(portfolio['INTERMEDIARIO'].unique())) == 2:
            lunghezza_titolo_27 = 12
            min_col = 4
        elif len(list(portfolio['INTERMEDIARIO'].unique())) == 3:
            lunghezza_titolo_27 = 12
            min_col = 3
        elif len(list(portfolio['INTERMEDIARIO'].unique())) == 4:
            lunghezza_titolo_27 = 12
            min_col = 3
        elif len(list(portfolio['INTERMEDIARIO'].unique())) == 5:
            lunghezza_titolo_27 = 12
            min_col = 2
        elif len(list(portfolio['INTERMEDIARIO'].unique())) == 6:
            lunghezza_titolo_27 = 12
            min_col = 2
        elif len(list(portfolio['INTERMEDIARIO'].unique())) == 7:
            lunghezza_titolo_27 = 12
            min_col = 1
        else:
            lunghezza_titolo_27 = len_header_27
            min_col = 1
        ws27.merge_cells(start_row=1, end_row=4, start_column=1, end_column=lunghezza_titolo_27)

        for col in ws27.iter_cols(min_row=8, max_row=9, min_col=min_col, max_col=min_col + len_header_27 - 1):
            ws27[col[0].coordinate].value = header_27[0]
            del header_27[0]
            ws27[col[0].coordinate].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws27[col[0].coordinate].font = Font(name='Times New Roman', size=10, color='FFFFFF')
            ws27[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='31869B')
            ws27[col[0].coordinate].border = Border(right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'))
            ws27.merge_cells(start_row=col[0].row, end_row=col[1].row, start_column=col[0].column, end_column=col[0].column)
            ws27.row_dimensions[col[0].row].height = 20
            ws27.row_dimensions[col[1].row].height = 20
            ws27.column_dimensions[col[0].column_letter].width = 12

        tipo_strumento_nogp = list(portfolio.loc[portfolio['CATEGORIA']!='GP', 'CATEGORIA'].unique())
        len_tipo_strumento_nogp = len(tipo_strumento_nogp)
        num_intermediari = len(portfolio['INTERMEDIARIO'].unique())
        lunghezza_colonna_27 = []
        tipo_strumento_dict = {'CASH' : 'LIQUIDITÀ', 'GP' : 'GESTIONI', 'EQUITY' : 'AZIONI', 'CASH_FOREIGN_CURR' : 'LIQUIDITÀ IN VALUTA', 'CORPORATE_BOND' : 'OBBLIGAZIONI CORPORATE', 'GOVERNMENT_BOND' : 'OBBLIGAZIONI GOVERNATIVE', 'ALTERNATIVE_ASSET' : 'INVESTIMENTI ALTERNATIVI', 'HEDGE_FUND' : 'HEDGE FUND'}
        for row in ws27.iter_rows(min_row=8, max_row=10 + len_tipo_strumento_nogp -1, min_col=min_col, max_col=min_col + len_header_27):
            if row[0].row > 9:
                ws27[row[0].coordinate].value = tipo_strumento_nogp[0] # carica i tipi di strumenti nell'indice
                del tipo_strumento_nogp[0]
                ws27[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
                ws27[row[0].coordinate].font = Font(name='Times New Roman', size=9, color='000000')
                ws27[row[0].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws27.row_dimensions[row[0].row].height = 19

                for _ in range(1, num_intermediari+1):
                    ws27[row[_].coordinate].value = portfolio.loc[(portfolio['INTERMEDIARIO']==ws27[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) & (portfolio['CATEGORIA']==ws27[row[0].coordinate].value), 'TOTALE t1'].sum() if portfolio.loc[(portfolio['INTERMEDIARIO']==ws27[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) & (portfolio['CATEGORIA']==ws27[row[0].coordinate].value), 'TOTALE t1'].sum() != 0 else ''
                    ws27[row[_].coordinate].alignment = Alignment(horizontal='center')
                    ws27[row[_].coordinate].font = Font(name='Times New Roman', size=9)
                    ws27[row[_].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                    ws27[row[_].coordinate].number_format = '#,0'

                for _ in range(1, num_intermediari+1):
                    if ws27[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value in posizioni_con_gestioni:
                        ws27[row[_].coordinate].value = ass_allocation.loc[(ass_allocation['INTERMEDIARIO']==ws27[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) & (ass_allocation['CATEGORIA']==ws27[row[0].coordinate].value), 'TOTALE t1'].sum() if ass_allocation.loc[(ass_allocation['INTERMEDIARIO']==ws27[row[_].offset(column=0, row=-2-(row[0].row-10)).coordinate].value) & (ass_allocation['CATEGORIA']==ws27[row[0].coordinate].value), 'TOTALE t1'].sum() != 0 else ''
                
                ws27[row[num_intermediari+1].coordinate].value = portfolio.loc[(portfolio['CATEGORIA']==ws27[row[0].coordinate].value) & (~portfolio['INTERMEDIARIO'].isin(posizioni_con_gestioni)), 'TOTALE t1'].sum() + ass_allocation.loc[(ass_allocation['CATEGORIA']==ws27[row[0].coordinate].value) & (ass_allocation['INTERMEDIARIO'].isin(posizioni_con_gestioni)), 'TOTALE t1'].sum()
                ws27[row[num_intermediari+1].coordinate].alignment = Alignment(horizontal='center')
                ws27[row[num_intermediari+1].coordinate].font = Font(name='Times New Roman', size=9)
                ws27[row[num_intermediari+1].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws27[row[num_intermediari+1].coordinate].number_format = '#,0'
                # TODO : controlla la somma dei valori nella colonna totale mese t0 usando valori veri.
                ws27[row[num_intermediari+2].coordinate].value = portfolio.loc[(portfolio['CATEGORIA']==ws27[row[0].coordinate].value) & (~portfolio['INTERMEDIARIO'].isin(posizioni_con_gestioni)), 'TOTALE t0'].sum() + ass_allocation.loc[(ass_allocation['CATEGORIA']==ws27[row[0].coordinate].value) & (ass_allocation['INTERMEDIARIO'].isin(posizioni_con_gestioni)), 'TOTALE t0'].sum()
                ws27[row[num_intermediari+2].coordinate].alignment = Alignment(horizontal='center')
                ws27[row[num_intermediari+2].coordinate].font = Font(name='Times New Roman', size=9)
                ws27[row[num_intermediari+2].coordinate].border = Border(bottom=Side(border_style='dashed', color='31869B'), right=Side(border_style='dashed', color='31869B'), left=Side(border_style='dashed', color='31869B'))
                ws27[row[num_intermediari+2].coordinate].number_format = '#,0'

                ws27[row[0].coordinate].value = tipo_strumento_dict[ws27[row[0].coordinate].value] # aggiorna valori dell'indice con i nomi nel dizionario
                lunghezza_colonna_27.append(len(ws27.cell(row=row[0].row, column=row[0].column).value)) # ottieni la lunghezza della colonna
                ws27.column_dimensions[row[0].column_letter].width = max(lunghezza_colonna_27) + 2.5 # modifica larghezza colonna

        # Somma per intermediari
        for row in ws27.iter_rows(min_row=10 + len_tipo_strumento_nogp, max_row=10 + len_tipo_strumento_nogp, min_col=min_col, max_col=min_col + len_header_27):
            ws27[row[0].coordinate].value = 'TOTALE'
            ws27[row[0].coordinate].alignment = Alignment(horizontal='left', vertical='center')
            ws27[row[0].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
            ws27[row[0].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
            ws27.row_dimensions[row[0].row].height = 19
            for _ in range(1,len_header_27-2):
                ws27[row[_].coordinate].value = portfolio.loc[portfolio['INTERMEDIARIO']==ws27.cell(row=row[_].row, column=row[_].column).offset(row=-len_tipo_strumento_nogp-2).value, 'TOTALE t1'].sum()
            ws27[row[len_header_27-2].coordinate].value = portfolio.loc[:, 'TOTALE t1'].sum()
            ws27[row[len_header_27-1].coordinate].value = portfolio.loc[:, 'TOTALE t0'].sum()
            for _ in range(1,len_header_27):
                ws27[row[_].coordinate].alignment = Alignment(horizontal='center')
                ws27[row[_].coordinate].font = Font(name='Times New Roman', size=9, bold=True)
                ws27[row[_].coordinate].border = Border(bottom=Side(border_style='thin', color='31869B'), right=Side(border_style='thin', color='31869B'), left=Side(border_style='thin', color='31869B'), top=Side(border_style='thin', color='31869B'))
                ws27[row[_].coordinate].number_format = '#,0'

        chart = PieChart()
        labels = Reference(ws27, min_col=min_col, max_col=min_col, min_row=10, max_row=10+len_tipo_strumento_nogp-1)
        data = Reference(ws27, min_col=min_col + len_header_27 - 2, max_col=min_col + len_header_27 - 2, min_row=10, max_row=10+len_tipo_strumento_nogp-1)
        chart.add_data(data, titles_from_data=False)
        chart.set_categories(labels)
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
        chart.dataLabels.textProperties = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=1200, b=True)), endParaRPr=CharacterProperties(sz=1200, b=True))])
        chart.legend.layout = Layout(manualLayout=ManualLayout(h=1))
        ws27.add_chart(chart, 'D20')

    def contatti_28(self):
        """
        Crea la ventottesima pagina.
        Solo formattazione.
        """
        ws28 = self.wb.create_sheet('28.contatti')
        ws28 = self.wb['28.contatti']
        self.wb.active = ws28

        ws28['A1'] = '4. Contatti'
        ws28['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws28['A1'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws28['A1'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws28.merge_cells('A1:L4')

        header_28 = ['Benchmark & Style S.r.l.', 'Via San Siro, 33', '20149 Milano', '="+390258328666"', 'info@benchmarkandstyle.com']
        for row in ws28.iter_rows(min_row=6, max_row=10, min_col=1, max_col=12):
            ws28[row[0].coordinate].value = header_28[0]
            del header_28[0]
            ws28[row[0].coordinate].alignment = Alignment(horizontal='center', vertical='center')
            ws28[row[0].coordinate].font = Font(name='Times New Roman', size=11, bold=True, color='31869B')
            ws28.merge_cells(start_row=row[0].row, end_row=row[0].row, start_column=1, end_column=12)

        ws28['A13'] = 'Disclaimer'
        ws28['A13'].alignment = Alignment(horizontal='center', vertical='center')
        ws28['A13'].font = Font(name='Times New Roman', size=48, bold=True, color='FFFFFF')
        ws28['A13'].fill = PatternFill(fill_type='solid', fgColor='31869B')
        ws28.merge_cells('A13:L16')

        ws28['A18'] = 'Il presente rendiconto ha una funzione meramente informativa ed è stato redatto sulla base dei dati forniti dai singoli gestori cui è affidato il patrimonio del cliente. I dati sono stati rielaborati al fine di fornire una visione d\'insieme e progressiva dei rendimenti mensili del patrimonio e delle singole gestioni confrontati ai relativi benchmarks. Tale rielaborazione rende più semplice comprendere i contributi dei singoli gestori e delle varie classi d\'attivo alla performance del patrimonio nel periodo considerato nonchè di monitorare periodicamente la performance stessa e i rischi assunti a livello consolidato e dei singoli portafogli.'
        ws28['A18'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        ws28['A18'].font = Font(name='Times New Roman', size=11, bold=True, italic=True, color='31869B')
        ws28.merge_cells('A18:L24')

        ws28['A25'] = 'Il presente rendiconto non rappresenta in alcun caso una raccomandazione e/o sollecitazione all\'acquisto o alla vendita di titoli, fondi, strumenti finanziari derivati, valute o altro e gli eventuali contenuti in esso non potranno in nessun caso essere ritenuti responsabili delle future performance del patrimonio del cliente neppure con riferimento alle previsioni formulate circa la prevedibile evoluzione dei mercati finanziari e/o di singoli comparti di essi.'
        ws28['A25'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        ws28['A25'].font = Font(name='Times New Roman', size=11, bold=True, italic=True, color='31869B')
        ws28.merge_cells('A25:L28')

        self.logo(ws28)

    def layout(self):
        """
        Cambia il layout dei fogli non nascosti: A4, orizzontale, centrato orizzontalmente, con margini personalizzati, adatta alla pagina oriz. e vert.
        Aggiunge il numero di pagina.
        """
        # Modifica layout di pagina e aggiungi numero di pagina a tutti i fogli
        for sheet in self.wb:
            if sheet.sheet_state != 'hidden':
                self.wb.active = sheet
                sheet.set_printer_settings(paper_size=9, orientation='landscape') # 9 = 'A4'
                sheet.page_margins = PageMargins(left=0.2362204724, right=0.2362204724, top=0.7480314961, bottom=0.0480314961, header=0.3149606299, footer=0.3149606299) #margini personalizzati
                sheet.print_options.horizontalCentered = True # centrato orizzontalmente
                sheet.sheet_properties.pageSetUpPr.fitToPage = True 
                sheet.page_setup.fitToHeight = 1
                sheet.page_setup.fitToWidth = 1
                numero_pagina_regex = re.compile(r'\d*') # regex per trovare il numero del foglio
                numero_pagina_search = numero_pagina_regex.search(str(sheet.title))
                numero_pagina = numero_pagina_search.group()
                sheet.oddFooter.right.text = numero_pagina # assegna il numero del foglio al piè di pagina destro

    def salva_file(self):
        """
        Salva il file excel.
        """
        self.wb.save(self.path.joinpath('report.xlsx'))


if __name__ == "__main__":
    start = time.time()
    # TODO : riduci la costante SFASAMENTO_DATI (riga 863 di un'unità)
    _ = Report(t1='30/06/2022')
    _.copertina_1()
    _.indice_2()
    _.analisi_di_mercato_3()
    _.analisi_rendimenti_4()
    _.analisi_indici_5()
    _.performance_6()
    _.andamento_7()
    _.caricamento_dati()
    _.cono_8()
    _.cono_9()
    _.nuovo_bk_10()
    _.performance_11()
    _.prezzi_12()
    _.prezzi_13()
    _.prezzi_14()
    _.att_in_corso_15()
    _.valutazione_per_macroclasse_16()
    _.sintesi_17()
    _.valuta_18()
    _.azioni_19()
    _.obb_governative_20()
    _.obb_corporate_21()
    _.obb_totale_22()
    _.liquidità_23()
    _.liq_totale_24()
    _.gestioni_25()
    _.inv_alt_26()
    _.asset_allocation_27()
    _.contatti_28()
    _.layout()
    _.salva_file()
    end = time.time()
    print("Elapsed time : ", end - start, 'seconds')
