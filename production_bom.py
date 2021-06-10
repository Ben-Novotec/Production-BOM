"""
Checks if components are ordered at Alinco by comparing the main bom list with the production bom list.
Then fills in the Status and PartSupplier column for these components.
Author: Ben Van Raemdonck
Date: 10/06/2021
"""
import pandas as pd
import openpyxl
import string


def production_bom(bom_production_path, bom_path):
    """Checks if component names in the excel file bom_path are present in bom_production path.
    For these, indicates the component has been ordered and the PartSupplier is Alinco in the file bom_path."""

    bom_production = pd.read_excel(bom_production_path, header=1, usecols=['Component', 'Quantity'])
    bom = pd.read_excel(bom_path, header=1, usecols=['Component', 'Quantity', 'Status', 'PartSupplier'])

    mask = bom.Component.isin(bom_production.Component)
    bom.loc[mask, 'Status'] = 'B'
    bom.loc[mask, 'PartSupplier'] = 'Alinco'

    wb = openpyxl.load_workbook(bom_path)
    ws = wb.active

    col_status, col_partsupplier = '', ''
    for c in string.ascii_uppercase:
        if ws[f'{c}2'].value == 'Status':
            col_status = c
        if ws[f'{c}2'].value == 'PartSupplier':
            col_partsupplier = c
    if col_status == '':
        raise NameError('Status column not found')
    if col_partsupplier == '':
        raise NameError('PartSupplier column not found')

    for i, (s, ps) in enumerate(zip(bom['Status'], bom['PartSupplier'])):
        ws[f'{col_status}{i+3}'] = s
        if s == 'B':
            ws[f'{col_status}{i+3}'].style = 'Neutral'
        ws[f'{col_partsupplier}{i+3}'] = ps

    wb.save(bom_path)


if __name__ == '__main__':
    production_bom(bom_production_path='.xlsx', bom_path='.xlsx')
