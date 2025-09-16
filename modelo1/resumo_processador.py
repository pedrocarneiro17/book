import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def _apply_summary_styles(ws):
    """Aplica todos os estilos de formatação na aba de resumo."""
    
    # --- Estilos Gerais ---
    header_font = Font(name='Calibri', bold=True, color="FFFFFF", size=14)
    section_font_compra = Font(name='Calibri', bold=True, color="C00000", size=12)  # Vermelho
    section_font_venda = Font(name='Calibri', bold=True, color="C00000", size=12)  # Vermelho
    section_font_entrada = Font(name='Calibri', bold=True, color="000000", size=12)  # Preto
    section_font_saida = Font(name='Calibri', bold=True, color="000000", size=12)  # Preto
    
    label_font = Font(name='Calibri', bold=True, size=11)
    data_font = Font(name='Calibri', size=11)
    value_font = Font(name='Calibri', size=11)
    red_value_font = Font(name='Calibri', size=11, color="FF0000")
    diff_label_font = Font(name='Calibri', bold=True, italic=True, size=11)
    diff_value_font = Font(name='Calibri', bold=True, italic=True, size=11)
    red_diff_value_font = Font(name='Calibri', bold=True, italic=True, size=11, color="FF0000")

    blue_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    orange_fill = PatternFill(start_color="FFE46C0A", end_color="FFE46C0A", fill_type="solid")
    dark_gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

    # --- Tabela 1: Conciliação entre Books ---
    ws['B2'].fill = blue_fill
    ws['B2'].font = header_font
    
    for cell in ws['B4:E4'][0]:
        cell.fill = dark_gray_fill
    ws['B4'].font = section_font_compra
    ws['C4'].font = label_font
    ws['D4'].font = label_font
    ws['E4'].font = label_font

    for row_idx in [5, 6, 10, 11]:
        for col_idx in [2, 3]:
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = value_font
        cell = ws.cell(row=row_idx, column=3)
        cell.number_format = '#,##0.00'
        if cell.value and isinstance(cell.value, (int, float)) and cell.value < 0:
            cell.font = red_value_font
    
    for row_idx in [7, 12]:
        ws.cell(row=row_idx, column=2).font = diff_label_font
        for col_idx in [3, 4, 5]:
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = red_diff_value_font if cell.value and isinstance(cell.value, (int, float)) and cell.value < 0 else diff_value_font
            cell.number_format = '#,##0.00'
    
    for cell in ws['B9:E9'][0]:
        cell.fill = dark_gray_fill
    ws['B9'].font = section_font_venda
    ws['C9'].font = label_font
    ws['D9'].font = label_font
    ws['E9'].font = label_font

    # --- Tabela 2: Conciliação entre Livro Fiscal x Book ---
    ws['G2'].fill = orange_fill
    ws['G2'].font = header_font

    for cell in ws['G4:I4'][0]:
        cell.fill = dark_gray_fill
    ws['G4'].font = section_font_entrada
    ws['I4'].font = label_font
    
    for cell in ws['G9:I9'][0]:
        cell.fill = dark_gray_fill
    ws['G9'].font = section_font_saida
    ws['I9'].font = label_font
    
    for row_idx in [5, 6, 10, 11]:
        ws.cell(row=row_idx, column=7).font = value_font
        cell = ws.cell(row=row_idx, column=8)
        cell.font = red_value_font if cell.value and isinstance(cell.value, (int, float)) and cell.value < 0 else value_font
        cell.number_format = '#,##0.00'
    
    for row_idx in [7, 12]:
        ws.cell(row=row_idx, column=7).font = diff_label_font
        cell = ws.cell(row=row_idx, column=8)
        cell.font = red_diff_value_font if cell.value and isinstance(cell.value, (int, float)) and cell.value < 0 else diff_value_font
        cell.number_format = '#,##0.00'

    # Formata as células de data (I5 e I10)
    for row_idx in [5, 10]:
        cell = ws.cell(row=row_idx, column=9)
        cell.font = data_font
        if cell.value and isinstance(cell.value, (str, pd.Timestamp)):
            cell.number_format = 'DD/MM/YYYY'

    # Ajusta largura das colunas
    for col_letter in ['B', 'C', 'D', 'E', 'G', 'H', 'I']:
        ws.column_dimensions[col_letter].width = 25

def get_df_from_ws(workbook_obj, sheet_name):
    """Converte uma aba de um workbook openpyxl para um DataFrame pandas."""
    try:
        ws_obj = workbook_obj[sheet_name]
        data = ws_obj.values
        cols = next(data)
        data = list(data)
        return pd.DataFrame(data, columns=cols)
    except KeyError:
        print(f"Aviso: Aba '{sheet_name}' para o resumo não encontrada. Retornando DataFrame vazio.")
        return pd.DataFrame(columns=[f'col_{i}' for i in range(12)])

def criar_aba_resumo(workbook, p1_config, data_compras, data_vendas):
    """Cria e formata a aba de resumo consolidado a partir do workbook em memória."""
    print("\n--- INICIANDO CRIAÇÃO DA ABA DE RESUMO ---")
    
    ws = workbook.create_sheet("Resumo") 
    
    nomes_abas_existentes = workbook.sheetnames
    
    # --- 1. CÁLCULOS PARA A TABELA "CONCILIAÇÃO ENTRE BOOKS" ---
    
    # Compra
    nome_aba_3 = nomes_abas_existentes[2]
    nome_aba_5 = nomes_abas_existentes[4]
    df3 = get_df_from_ws(workbook, nome_aba_3)
    df5 = get_df_from_ws(workbook, nome_aba_5)
    
    soma_v_ajust_3 = pd.to_numeric(df3[p1_config['valor']], errors='coerce').sum()
    soma_v_ajust_5 = pd.to_numeric(df5[p1_config['valor']], errors='coerce').sum()
    contratos_principais_3 = set(df3[p1_config['num_contrato']].unique())
    df5['MCP'] = np.where(df5[p1_config['num_contrato']].isin(contratos_principais_3), 'N', 'S')
    soma_mcp_sim_5 = pd.to_numeric(df5[df5['MCP'] == 'S'][p1_config['valor']], errors='coerce').sum()
    
    diff_books_compra_real = soma_v_ajust_3 - soma_v_ajust_5
    diff_books_compra_display = abs(diff_books_compra_real) 
    diff_final_compra = diff_books_compra_display - soma_mcp_sim_5

    # Venda
    nome_aba_4 = nomes_abas_existentes[3]
    nome_aba_6 = nomes_abas_existentes[5]
    df4 = get_df_from_ws(workbook, nome_aba_4)
    df6 = get_df_from_ws(workbook, nome_aba_6)

    soma_v_ajust_4 = pd.to_numeric(df4[p1_config['valor']], errors='coerce').sum()
    soma_v_ajust_6 = pd.to_numeric(df6[p1_config['valor']], errors='coerce').sum()
    contratos_principais_4 = set(df4[p1_config['num_contrato']].unique())
    df6['MCP'] = np.where(df6[p1_config['num_contrato']].isin(contratos_principais_4), 'N', 'S')
    soma_mcp_sim_6 = pd.to_numeric(df6[df6['MCP'] == 'S'][p1_config['valor']], errors='coerce').sum()
    
    diff_books_venda_real = soma_v_ajust_4 - soma_v_ajust_6
    diff_books_venda_display = abs(diff_books_venda_real)
    diff_final_venda = diff_books_venda_display - soma_mcp_sim_6

    # --- 2. CÁLCULOS PARA "CONCILIAÇÃO ENTRE LIVRO FISCAL X BOOK" ---
    
    def get_totals_from_sheet(workbook_obj, sheet_name):
        """Busca a linha de 'TOTAL GERAL' e extrai as somas já calculadas."""
        try:
            ws_obj = workbook_obj[sheet_name]
            total_livro = 0
            total_book = 0
            for row in ws_obj.iter_rows(values_only=True):
                if row[0] == "TOTAL GERAL":
                    total_livro = row[4] if row[4] is not None else 0
                    total_book = row[10] if row[10] is not None else 0
                    
                    total_livro_num = pd.to_numeric(total_livro, errors='coerce')
                    total_book_num = pd.to_numeric(total_book, errors='coerce')
                    
                    return total_livro_num if not pd.isna(total_livro_num) else 0, total_book_num if not pd.isna(total_book_num) else 0

            return 0, 0
        except KeyError:
            print(f"Aviso: Aba '{sheet_name}' para o resumo não encontrada.")
            return 0, 0

    soma_livro_entrada, soma_book_entrada = get_totals_from_sheet(workbook, "Livro x Book Entrada")
    diff_entrada = soma_livro_entrada - soma_book_entrada

    soma_livro_saida, soma_book_saida = get_totals_from_sheet(workbook, "Livro x Book Saída")
    diff_saida = soma_livro_saida - soma_book_saida
    
    # --- 3. MONTAGEM DA ABA NO EXCEL ---
    ws.merge_cells('B2:E2')
    ws['B2'] = "Conciliação entre Books"
    ws.row_dimensions[2].height = 30
    
    ws['B4'] = "Compra"
    ws['C4'] = "Valor ajustado"
    ws['D4'] = "MCP (sim)"
    ws['E4'] = "Diferença"
    
    ws['B5'] = nome_aba_3
    ws['C5'] = soma_v_ajust_3
    ws['B6'] = nome_aba_5
    ws['C6'] = soma_v_ajust_5
    
    ws['B7'] = "Diferença"
    ws['C7'] = diff_books_compra_display
    ws['D7'] = soma_mcp_sim_5
    ws['E7'] = diff_final_compra
    
    ws['B9'] = "Venda"
    ws['C9'] = "Valor ajustado"
    ws['D9'] = "MCP (sim)"
    ws['E9'] = "Diferença"
    
    ws['B10'] = nome_aba_4
    ws['C10'] = soma_v_ajust_4
    ws['B11'] = nome_aba_6
    ws['C11'] = soma_v_ajust_6
    
    ws['B12'] = "Diferença"
    ws['C12'] = diff_books_venda_display
    ws['D12'] = soma_mcp_sim_6
    ws['E12'] = diff_final_venda

    ws.merge_cells('G2:I2')
    ws['G2'] = "Conciliação entre Livro Fiscal x Book"
    
    ws['G4'] = "ENTRADA"
    ws['I4'] = "Data do filtro"
    
    ws['G5'] = "Livro Fiscal"
    ws['H5'] = soma_livro_entrada
    ws['I5'] = pd.to_datetime(data_compras, errors='coerce') if data_compras else None
    
    ws['G6'] = "Book"
    ws['H6'] = soma_book_entrada
    
    ws['G7'] = "Diferença"
    ws['H7'] = diff_entrada

    ws['G9'] = "SAÍDA"
    ws['I9'] = "Data do filtro"

    ws['G10'] = "Livro Fiscal"
    ws['H10'] = soma_livro_saida
    ws['I10'] = pd.to_datetime(data_vendas, errors='coerce') if data_vendas else None
    
    ws['G11'] = "Book"
    ws['H11'] = soma_book_saida
    
    ws['G12'] = "Diferença"
    ws['H12'] = diff_saida

    _apply_summary_styles(ws)
    print("[Resumo] Aba de Resumo criada com sucesso.")
    return workbook