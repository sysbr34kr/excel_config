import configparser
import pandas as pd
import openpyxl
import os
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Alignment, Font, numbers, PatternFill
from openpyxl.utils.cell import get_column_letter
from openpyxl.styles.borders import Border, Side



def process_quotes():
    # Locate files in the current directory
    current_directory = os.getcwd()

    # Find the first .xls file matching the pattern
    input_file = next((file for file in os.listdir(current_directory) if file.startswith("Arrematantes_Leilao_") and file.endswith(".xls")), None)
    if not input_file:
        raise FileNotFoundError("No input file matching 'Arrematantes_Leilao_#####.xls' found in the current directory.")

    input_file = os.path.join(current_directory, input_file)

    # Find the valores.ini file
    config_file = os.path.join(current_directory, "valores.ini")
    if not os.path.exists(config_file):
        raise FileNotFoundError("Configuration file 'valores.ini' not found in the current directory.")

    # Extract numeric part from input file name
    file_number = os.path.splitext(os.path.basename(input_file))[0].split('_')[-1]
    output_file = f"Cotações_Leilão_{file_number}.xlsx"

    # Read configuration values
    config = configparser.ConfigParser()
    config.read(config_file)

    im_values = {
        '0.3Kg': float(config['IM_Weights']['IM_0_3Kg']),
        '0.9Kg': float(config['IM_Weights']['IM_0_9Kg']),
        '2.0Kg': float(config['IM_Weights']['IM_2_0Kg'])
    }
    
    # Convert HTML (.xls) to DataFrame
    df = pd.read_html(input_file, header=1)[0]  # Header starts at row 3 (index 2)

    # Select required columns
    df = df.iloc[:, [1, 6]]  # Columns B and G
    df.columns = ['Nome', 'CEP']

    # Process 'Nome' column
    df['Nome'] = df['Nome'].str.upper()

    # Process 'CEP' column
    df['CEP'] = df['CEP'].astype(str)
    df['CEP'] = df['CEP'].str.replace("-", "", regex=False)
    df['CEP'] = df['CEP'].apply(lambda x: x.zfill(8))

    # Manually set column headers and initialize additional columns
    df = pd.DataFrame({
        'Nome': df['Nome'],
        'CEP': df['CEP'],
        'Modalidade': '',
        'Peso': '',
        'Alt.': '',
        'Lar.': '',
        'Com.': '',
        'Valor': ''
    })

    # Convert DataFrame to Excel with OpenPyXL
    wb = Workbook()
    ws = wb.active
    ws.title = "Quotes"

    # Write DataFrame to worksheet
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=1):
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            # Apply font and alignment
            cell.font = Font(name='Arial', size=12)
            if r_idx == 1:  # Header row
                cell.alignment = Alignment(horizontal="center")
            elif c_idx == 1 and r_idx > 1:  # Column A, non-header rows
                cell.alignment = Alignment(horizontal="left")
            else:  # Other columns
                cell.alignment = Alignment(horizontal="center")

    # Add dropdown for 'Modalidade' column (Column C)
    modalidade_options = ["RETIRA", "IM", "PAC Min.", "PAC", "2x PAC", "SEDEX", "OUTRO"]
    dv = DataValidation(type="list", formula1=f"\"{','.join(modalidade_options)}\"", showDropDown=False)
    dv.showErrorMessage = True
    ws.add_data_validation(dv)

    for row in range(2, ws.max_row):
        dv.add(ws.cell(row=row, column=3))

    # Set column widths
    column_widths = {
        'A': 71.43,
        'B': 12.14,
        'C': 16.43,
        'D': 8.57,
        'E': 10,
        'F': 10,
        'G': 10,
        'H': 12.14
    }

    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    # Apply dropdown and formulas row-by-row
    for row in range(2, ws.max_row):  # Start from row 2 to skip the header
        modalidade_cell = ws.cell(row=row, column=3)  # Column C for 'Modalidade'
        peso_cell = ws.cell(row=row, column=4)  # Column D for 'Peso'
        altura_cell = ws.cell(row=row, column=5)  # Column E for 'Alt.'
        largura_cell = ws.cell(row=row, column=6)  # Column F for 'Lar.'
        comprimento_cell = ws.cell(row=row, column=7)  # Column G for 'Com.'
        valor_cell = ws.cell(row=row, column=8)  # Column H for 'Valor'

        # Apply dropdown validation to the 'Modalidade' cell
        dv.add(modalidade_cell)

        # Apply formulas for Peso (Column D)
        peso_cell.value = (
            f'=IF({modalidade_cell.coordinate}="PAC Min.", "1.0 Kg", '
            f'IF({modalidade_cell.coordinate}="RETIRA", "-", ""))'
        )

        # Apply formulas for Alt., Lar., and Com. (Columns E, F, G)
        altura_cell.value = (
            f'=IF({modalidade_cell.coordinate}="PAC Min.", "-", '
            f'IF({modalidade_cell.coordinate}="IM", "-", '
            f'IF({modalidade_cell.coordinate}="RETIRA", "-", "")))'
        )
        largura_cell.value = (
            f'=IF({modalidade_cell.coordinate}="PAC Min.", "-", '
            f'IF({modalidade_cell.coordinate}="IM", "-", '
            f'IF({modalidade_cell.coordinate}="RETIRA", "-", "")))'
        )
        comprimento_cell.value = (
            f'=IF({modalidade_cell.coordinate}="PAC Min.", "-", '
            f'IF({modalidade_cell.coordinate}="IM", "-", '
            f'IF({modalidade_cell.coordinate}="RETIRA", "-", "")))'
        )

        # Apply formulas for Valor (Column H)
        valor_cell.value = (
            f'=IF({modalidade_cell.coordinate}="RETIRA", "-", '
            f'IF(AND({modalidade_cell.coordinate}="IM", '
            f'ISNUMBER(MATCH(LEFT({peso_cell.coordinate}, FIND(" ", {peso_cell.coordinate})-1), '
            f'{{"0.3","0.9","2.0"}}, 0))), '
            f'VLOOKUP(LEFT({peso_cell.coordinate}, FIND(" ", {peso_cell.coordinate})-1), '
            f'{{"0.3",{im_values["0.3Kg"]};"0.9",{im_values["0.9Kg"]};"2.0",{im_values["2.0Kg"]}}}, 2, FALSE), ""))'
        )

        valor_cell.number_format = '"R$" #,##0.00'  # Apply currency formatting

        # Apply custom number formatting with suffix for columns D, E, F, and G
        peso_cell.number_format = '0.0 "Kg"'
        altura_cell.number_format = '##0 "cm"'
        largura_cell.number_format = '##0 "cm"'
        comprimento_cell.number_format = '##0 "cm"'

    # Add summary row
    summary_row = ws.max_row  # Row after the last data row

    # Column A: Count the number of items (excluding the header)
    ws.cell(row=summary_row, column=1).value = f"=COUNTA(A2:A{ws.max_row - 1})"
    ws.cell(row=summary_row, column=1).font = Font(name='Arial', size=12, bold=True)
    ws.cell(row=summary_row, column=1).alignment = Alignment(horizontal="center")

    # Column H: Count the number of empty cells
    ws.cell(row=summary_row, column=8).value = f"=COUNTBLANK(H2:H{ws.max_row - 1})"
    ws.cell(row=summary_row, column=8).font = Font(name='Arial', size=12, bold=True)
    ws.cell(row=summary_row, column=8).alignment = Alignment(horizontal="center")

    # Set other columns in the summary row to empty strings
    for col in range(2, ws.max_column):
        if col != 8:  # Skip column H
            ws.cell(row=summary_row, column=col).value = ""
            ws.cell(row=summary_row, column=col).alignment = Alignment(horizontal="center")
    
    # Freeze the header row
    ws.freeze_panes = ws['A2']

    # Define styles
    header_fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
    footer_fill = PatternFill(start_color="AAAAAA", end_color="AAAAAA", fill_type="solid")
    light_gray_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    bold_font = Font(name='Arial', size=12, bold=True)

    # Style header row
    for cell in ws[1]:
        cell.font = bold_font
        cell.fill = header_fill

    # Apply alternating row colors
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row - 1), start=2):
        fill = light_gray_fill if row_idx % 2 == 0 else white_fill
        for cell in row:
            cell.fill = fill

    # Style footer row
    for cell in ws[ws.max_row]:
        cell.font = bold_font
        cell.fill = footer_fill
    
    # Define border styles
    thin_border = Side(border_style="thin", color="000000")
    header_border = Border(bottom=thin_border, left=thin_border, right=thin_border)
    footer_border = Border(top=thin_border, left=thin_border, right=thin_border)
    vertical_border = Border(left=thin_border, right=thin_border)

    # Apply borders to header row
    for cell in ws[1]:
        cell.border = header_border

    # Apply vertical borders to main document rows
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row - 1):  # Exclude footer
        for cell in row:
            cell.border = vertical_border

    # Apply borders to footer row
    for cell in ws[ws.max_row]:
        cell.border = footer_border

    # Save to file
    wb.save(output_file)

# Run the processing function
if __name__ == "__main__":
    process_quotes()
