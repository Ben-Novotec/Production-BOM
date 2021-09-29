"""
Checks if components are ordered at Alinco by comparing the main bom list with the production bom list.
Then fills in the Status and PartSupplier column for these components.
Author: Ben Van Raemdonck
Date: 10/06/2021
"""
import pandas as pd
import openpyxl
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from my_openpyxl import find_column, write_column
from datetime import date


def production_bom(bom_production_path, bom_path, supplier='Alinco', status='B',
                   besteldatum=date.today().strftime('%d/%m')):
    """Checks if component names in the excel file bom_path are present in bom_production path.
    For these, indicates the component has been ordered and the PartSupplier is Alinco in the file bom_path."""

    # read excel files
    bom_production = pd.read_excel(bom_production_path, header=1, usecols=['Component', 'Quantity'])
    bom = pd.read_excel(bom_path, header=1, usecols=['Besteldatum', 'Component', 'Quantity', 'Status', 'PartSupplier'])

    # find components that are in the production bom list
    mask = bom.Component.isin(bom_production.Component)

    bom.loc[mask, 'Status'] = status
    bom.loc[mask, 'PartSupplier'] = supplier
    bom.loc[mask, 'Besteldatum'] = besteldatum

    # open the excel file with openpyxl
    wb = openpyxl.load_workbook(bom_path)
    ws = wb.active
    # finds the column letter in excel
    col_status = find_column(ws, 'Status')
    col_partsupplier = find_column(ws, 'PartSupplier')
    col_besteldatum = find_column(ws, 'Besteldatum')
    # write the columns
    write_column(ws, col_status, bom['Status'])
    write_column(ws, col_partsupplier, bom['PartSupplier'])
    write_column(ws, col_besteldatum, bom['Besteldatum'])

    wb.save(bom_path)


if __name__ == '__main__':
    Tk().withdraw()
    production_bom(bom_production_path=askopenfilename(title='Choose production BOM list'),
                   bom_path=askopenfilename(title='Choose main BOM list'))
