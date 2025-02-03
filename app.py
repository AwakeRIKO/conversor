import os
import pdfplumber
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from flask import Flask, request, send_file, render_template, redirect, url_for
from werkzeug.utils import secure_filename
import logging
from datetime import datetime

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Configuração do Flask
app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf'}

# Criar diretórios necessários
for folder in [UPLOAD_FOLDER, 'logs']:
    if not os.path.exists(folder):
        os.makedirs(folder)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # Limite de 16MB

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_transactions_from_pdf(pdf_path):
    """
    Extrai as transações de todas as páginas do PDF do Mercado Pago.
    """
    transactions = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            logger.info(f"PDF aberto com {len(pdf.pages)} páginas")

            for page in pdf.pages:
                logger.info(f"Processando página {page.page_number}")
                text = page.extract_text()
                lines = text.split('\n')
                logger.info(f"Encontradas {len(lines)} linhas na página {page.page_number}")

                # Padrão para extrair transações
                transaction_pattern = re.compile(
                    r'(\d{2}-\d{2}-\d{4})\s+'    # Data
                    r'(.*?)\s+'                   # Descrição
                    r'(\d{11})\s+'                # ID da operação
                    r'R\$\s*(-?[\d.,]+)'          # Valor (ajustado para capturar valores negativos)
                )

                for line in lines:
                    line = line.strip()

                    # Pular linhas de cabeçalho e rodapé
                    if any(header in line for header in [
                        'EXTRATO DE CONTA',
                        'DETALHE DOS MOVIMENTOS',
                        'Data de geração:',
                        'Saldo inicial:',
                        'Saldo final:',
                        'Data\s+Descrição',
                        'Você tem alguma dúvida?',
                        'Mercado Pago',
                        '/3'  # Indicador de página
                    ]):
                        continue

                    # Tentar encontrar uma transação
                    match = transaction_pattern.search(line)
                    if match:
                        logger.info(f"Transação encontrada: {line}")
                        date, description, operation_id, value = match.groups()

                        # Limpar e converter o valor
                        value = value.replace('.', '').replace(',', '.')
                        try:
                            value = float(value)

                            # Verificar se a transação já existe (evitar duplicatas)
                            transaction = {
                                'Data': date,
                                'Descrição': description.strip(),
                                'Valor': value
                            }

                            # Adicionar apenas se não for duplicata
                            if transaction not in transactions:
                                transactions.append(transaction)
                                logger.info(f"Nova transação adicionada: {transaction}")

                        except ValueError as e:
                            logger.error(f"Erro ao converter valor: {value} - {str(e)}")
                            continue

        logger.info(f"Total de {len(transactions)} transações extraídas do PDF")

        # Ordenar transações por data
        transactions.sort(key=lambda x: datetime.strptime(x['Data'], '%d-%m-%Y'))
        return transactions

    except Exception as e:
        logger.error(f"Erro ao processar o PDF: {str(e)}")
        raise

def format_excel(excel_path):
    """
    Aplica formatação personalizada ao arquivo Excel.
    """
    try:
        workbook = load_workbook(excel_path)
        sheet = workbook.active

        # Estilos
        header_fill = PatternFill(start_color="4CAF50", end_color="4CAF50", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)

        # Formatar cabeçalho
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')

        # Formatar células de dados
        for row in sheet.iter_rows(min_row=2):
            # Data
            row[0].alignment = Alignment(horizontal='center')
            # Descrição
            row[1].alignment = Alignment(horizontal='left')
            # Valor
            row[2].alignment = Alignment(horizontal='right')
            if row[2].value:
                valor = float(str(row[2].value).replace('.', '').replace(',', '.'))
                if valor < 0:
                    row[2].font = Font(color="FF0000")  # Vermelho para negativos
                else:
                    row[2].font = Font(color="4CAF50")  # Verde para positivos

        # Ajustar largura das colunas
        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            adjusted_width = max_length + 2
            sheet.column_dimensions[column_letter].width = adjusted_width

        workbook.save(excel_path)
        workbook.close()
        logger.info(f"Excel formatado com sucesso: {excel_path}")
    except Exception as e:
        logger.error(f"Erro ao formatar o Excel: {str(e)}")

def create_excel(transactions, excel_path):
    """
    Cria arquivo Excel com as transações.
    """
    try:
        df = pd.DataFrame(transactions)
        # Ordenar por data
        df['Data'] = pd.to_datetime(df['Data'], format='%d-%m-%Y')
        df = df.sort_values('Data')
        df['Data'] = df['Data'].dt.strftime('%d-%m-%Y')

        # Formatar valores
        df['Valor'] = df['Valor'].map(lambda x: f"{x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))

        # Salvar Excel
        df.to_excel(excel_path, sheet_name='Extrato', index=False)

        # Aplicar formatação
        format_excel(excel_path)
        logger.info(f"Arquivo Excel criado com sucesso: {excel_path}")
        return True
    except Exception as e:
        logger.error(f"Erro ao criar arquivo Excel: {str(e)}")
        return False

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/excel')
def excel():
    return render_template('excel.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files:
            return 'Nenhum arquivo enviado', 400

        file = request.files['file']
        if file.filename == '':
            return 'Nenhum arquivo selecionado', 400

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{timestamp}_{filename}")
            file.save(pdf_path)

            try:
                # Extrair transações
                transactions = extract_transactions_from_pdf(pdf_path)

                # Criar Excel
                excel_filename = f"{timestamp}_{filename.replace('.pdf', '.xlsx')}"
                excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)

                if create_excel(transactions, excel_path):
                    # Limpar arquivo PDF
                    try:
                        os.remove(pdf_path)
                    except Exception as e:
                        logger.error(f"Erro ao excluir PDF: {str(e)}")

                    return send_file(
                        excel_path,
                        as_attachment=True,
                        download_name=filename.replace('.pdf', '.xlsx'),
                        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                else:
                    return 'Erro ao criar arquivo Excel', 400
            except Exception as e:
                logger.error(f"Erro ao processar arquivo: {str(e)}")
                return f'Erro ao processar arquivo: {str(e)}', 400
    except Exception as e:
        logger.error(f"Erro no upload: {str(e)}")
        return 'Erro interno do servidor', 500

@app.errorhandler(413)
def too_large(e):
    return 'Arquivo muito grande. Tamanho máximo permitido é 16MB', 413

@app.errorhandler(500)
def internal_error(e):
    return 'Erro interno do servidor', 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
