from openpyxl import Workbook

# Criação de um workbook que contara todas as planilhas
arquivo_excel = Workbook()

# Pegando a planilha padrão e alterando o titulo
planilha1 = arquivo_excel.active
planilha1.title = "Gastos"

planilha2 = arquivo_excel.create_sheet("Ganhos")

print(arquivo_excel.sheetnames)
planilha1['A1'] = 'Categoria'
planilha1['B1'] = 'Valor'
planilha1['A2'] = "Restaurante"
planilha1['B2'] = 45.99
valores = [
    ("Categoria", "Valor"),
    ("Restaurante", 45.99),
    ("Transporte", 208.45),
    ("Viagem", 558.54)
]
for linha in valores:
    planilha1.append(linha)

arquivo_excel.save("relatorio.xlsx")