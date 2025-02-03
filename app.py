import pandas as pd
import tabula
import openpyxl
from datetime import datetime

# Lê as tabelas do PDF
def extract_from_pdf(pdf_path):
    # Lê todas as tabelas do PDF
    tables = tabula.read_pdf(pdf_path, pages='all')
    
    # Lista para armazenar todos os dados
    all_data = []
    
    # Processa cada tabela encontrada
    for table in tables:
        # Filtra apenas as colunas necessárias
        if 'Data' in table.columns and 'Descrição' in table.columns and 'Valor' in table.columns:
            filtered_data = table[['Data', 'Descrição', 'Valor']]
            all_data.append(filtered_data)
    
    # Combina todos os dados em um único DataFrame
    final_df = pd.concat(all_data, ignore_index=True)
    
    # Converte a coluna de valores para numérico, removendo 'R$' e convertendo vírgulas
    final_df['Valor'] = final_df['Valor'].str.replace('R$ ', '').str.replace('.', '').str.replace(',', '.').astype(float)
    
    # Converte a coluna de data para o formato datetime
    final_df['Data'] = pd.to_datetime(final_df['Data'], format='%d-%m-%Y')
    
    return final_df

# Salva os dados em um arquivo Excel
def save_to_excel(df, output_path):
    df.to_excel(output_path, index=False)
    
# Execução principal
pdf_path = 'extrato.pdf'  # Nome do seu arquivo PDF
output_path = 'extrato_convertido.xlsx'  # Nome do arquivo Excel de saída

try:
    # Extrai os dados do PDF
    df = extract_from_pdf(pdf_path)
    
    # Salva em Excel
    save_to_excel(df, output_path)
    print(f"Arquivo convertido com sucesso! Salvo como: {output_path}")

except Exception as e:
    print(f"Erro durante a conversão: {str(e)}")

# Created/Modified files during execution:
print("Arquivos criados/modificados:")
print("- extrato_convertido.xlsx")
