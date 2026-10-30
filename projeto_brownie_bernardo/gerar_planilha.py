import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

wb = openpyxl.Workbook()

font_family = "Arial"
header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
header_fill = PatternFill(start_color="3B2314", end_color="3B2314", fill_type="solid") # Tom marrom elegante
accent_fill = PatternFill(start_color="F2E6DF", end_color="F2E6DF", fill_type="solid")
border_thin = Border(
    left=Side(style='thin', color='D3D3D3'),
    right=Side(style='thin', color='D3D3D3'),
    top=Side(style='thin', color='D3D3D3'),
    bottom=Side(style='thin', color='D3D3D3')
)
border_header = Border(
    left=Side(style='thin', color='D3D3D3'),
    right=Side(style='thin', color='D3D3D3'),
    top=Side(style='medium', color='3B2314'),
    bottom=Side(style='medium', color='3B2314')
)


ws_dash = wb.active
ws_dash.title = "Resumo"
ws_dash.views.sheetView[0].showGridLines = True

ws_dash['A1'] = "CONFEITARIA BROWNIES DO BERNARDO"
ws_dash['A1'].font = Font(name=font_family, size=16, bold=True, color="3B2314")
ws_dash['A2'] = "Painel de Controle de Custos e Precificação"
ws_dash['A2'].font = Font(name=font_family, size=11, italic=True, color="666666")

dash_headers = ["Indicador", "Valor"]
for col_idx, text in enumerate(dash_headers, start=1):
    cell = ws_dash.cell(row=4, column=col_idx, value=text)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = border_header

dash_data = [
    ("Custo Total da Fornada (Receita Base)", "=Insumos!E10"),
    ("Rendimento (Unidades por Fornada)", 12),
    ("Custo Unitário por Brownie", "=B5/B6"),
    ("Margem de Lucro Desejada (%)", 1.00),
    ("Preço de Venda Sugerido (Unitário)", "=B7*(1+B8)"),
    ("Preço de Venda Praticado (Mercado)", 10.00),
    ("Lucro Estimado por Unidade", "=B10-B7")
]

for row_idx, (label, val) in enumerate(dash_data, start=5):
    c1 = ws_dash.cell(row=row_idx, column=1, value=label)
    c2 = ws_dash.cell(row=row_idx, column=2, value=val)
    
    c1.font = Font(name=font_family, size=10, bold=(row_idx in [7, 9, 11]))
    c2.font = Font(name=font_family, size=10, bold=(row_idx in [7, 9, 11]))
    c1.border = border_thin
    c2.border = border_thin
    
    if row_idx in [5, 7, 9, 10, 11]:
        c1.fill = accent_fill
        c2.fill = accent_fill
        
    if "R$" in label or row_idx in [5, 7, 9, 10, 11]:
        if isinstance(val, str) or row_idx != 6:
            c2.number_format = 'R$ #,##0.00'
        if row_idx == 8:
            c2.number_format = '0.0%'

# ----------------------------------------------------
# ABA 2: Custos e Insumos
# ----------------------------------------------------
ws_insumos = wb.create_sheet(title="Insumos")
ws_insumos.views.sheetView[0].showGridLines = True

ins_headers = ["Insumo / Ingrediente", "Unidade de Compra", "Qtd. Comprada", "Preço Pago (R$)", "Custo por Unidade de Medida (R$)"]
for col_idx, text in enumerate(ins_headers, start=1):
    cell = ws_insumos.cell(row=2, column=col_idx, value=text)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border_header

insumos_data = [
    ("Chocolate Nobre 50%", "kg", 1.0, 45.00, "=D3/C3"),
    ("Manteiga sem Sal", "kg", 0.5, 20.00, "=D4/C4"),
    ("Açúcar Cristal", "kg", 1.0, 5.00, "=D5/C5"),
    ("Ovos (Dúzia)", "unidade", 12.0, 9.00, "=D6/C6"),
    ("Farinha de Trigo", "kg", 1.0, 6.00, "=D7/C7"),
    ("Cacau em Pó 100%", "kg", 0.5, 25.00, "=D8/C8"),
    ("Embalagem Individual", "unidade", 50.0, 15.00, "=D9/C9"),
]

for i, row_data in enumerate(insumos_data, start=3):
    for j, val in enumerate(row_data, start=1):
        cell = ws_insumos.cell(row=i, column=j, value=val)
        cell.font = Font(name=font_family, size=10)
        cell.border = border_thin
        if j in [3]:
            cell.number_format = '#,##0.00'
        elif j in [4, 5]:
            cell.number_format = 'R$ #,##0.00'

# Totalizador / Custo base na linha 10
ws_insumos.cell(row=10, column=1, value="Custo Total dos Insumos da Receita Base").font = Font(name=font_family, size=10, bold=True)
ws_insumos.cell(row=10, column=5, value="=SUM(E3:E9)").font = Font(name=font_family, size=10, bold=True)
ws_insumos.cell(row=10, column=5).number_format = 'R$ #,##0.00'

# ----------------------------------------------------
# ABA 3: Controle de Estoque
# ----------------------------------------------------
ws_estoque = wb.create_sheet(title="Estoque")
ws_estoque.views.sheetView[0].showGridLines = True

est_headers = ["Item / Insumo", "Unidade", "Estoque Mínimo", "Qtd. Atual", "Status"]
for col_idx, text in enumerate(est_headers, start=1):
    cell = ws_estoque.cell(row=2, column=col_idx, value=text)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = border_header

estoque_data = [
    ("Chocolate Nobre 50%", "kg", 0.5, 2.0, '=IF(D3<=C3, "CRÍTICO - REPOR", "OK")'),
    ("Manteiga sem Sal", "kg", 0.3, 1.0, '=IF(D4<=C4, "CRÍTICO - REPOR", "OK")'),
    ("Açúcar Cristal", "kg", 1.0, 3.5, '=IF(D5<=C5, "CRÍTICO - REPOR", "OK")'),
    ("Ovos", "unidade", 12.0, 24.0, '=IF(D6<=C6, "CRÍTICO - REPOR", "OK")'),
    ("Farinha de Trigo", "kg", 1.0, 2.0, '=IF(D7<=C7, "CRÍTICO - REPOR", "OK")'),
    ("Cacau em Pó 100%", "kg", 0.2, 0.8, '=IF(D8<=C8, "CRÍTICO - REPOR", "OK")'),
    ("Embalagem Individual", "unidade", 20.0, 45.0, '=IF(D9<=C9, "CRÍTICO - REPOR", "OK")'),
]

for i, row_data in enumerate(estoque_data, start=3):
    for j, val in enumerate(row_data, start=1):
        cell = ws_estoque.cell(row=i, column=j, value=val)
        cell.font = Font(name=font_family, size=10)
        cell.border = border_thin
        if j in [3, 4]:
            cell.number_format = '#,##0.0'
        if j == 5:
            cell.alignment = Alignment(horizontal="center")

# Ajuste automático de largura das colunas em todas as abas
for sheet in wb.worksheets:
    for col in sheet.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                val_str = str(cell.value)
                if not val_str.startswith("="):
                    max_len = max(max_len, len(val_str))
        sheet.column_dimensions[col_letter].width = max(max_len + 5, 15)

# Salvar arquivo
file_name = "controle_custos_brownies_bernardo.xlsx"
wb.save(file_name)