import os
import pdfplumber
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from flask import Flask, request, send_file, render_template, jsonify
from werkzeug.utils import secure_filename
import logging
from datetime import datetime

# Configuração de logging
logging.basicConfig(level=logging.INFO)
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
    """Verifica se a extensão do arquivo é permitida."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_transactions_from_pdf(pdf_path):
    """
    Extrai as transações do PDF no formato especificado.
    Retorna uma lista de dicionários com as transações.
    """
    transactions = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                lines = text.split('\n')

                # Padrão ajustado para o formato específico
                transaction_pattern = re.compile(
                    r'^(\d{2}-\d{2}-\d{4})\s+'    # Data
                    r'(.*?)\s+'                   # Descrição (captura até o ID)
                    r'\d{11}\s+'                  # Ignora o ID da operação
                    r'R\$\s*(-?\d{1,3}(?:\.\d{3})*,\d{2})'  # Valor principal
                )

                for line in lines:
                    line = line.strip()
                    transaction_match = transaction_pattern.match(line)
                    if transaction_match:
                        date, description, value = transaction_match.groups()
                        value = float(value.replace('.', '').replace(',', '.'))
                        transactions.append({'Data': date, 'Descrição': description.strip(), 'Valor': value})
                    else:
                        if transactions and not re.match(r'\d{2}-\d{2}-\d{4}', line):
                            transactions[-1]['Descrição'] += ' ' + line.strip()
        logger.info(f"Extraídas {len(transactions)} transações do PDF")
        return transactions
    except Exception as e:
        logger.error(f"Erro ao processar o PDF {pdf_path}: {str(e)}")
        return None

def create_excel(transactions, excel_path):
    """
    Cria arquivo Excel com as transações e aplica formatação.
    """
    try:
        df = pd.DataFrame(transactions)
        df['Data'] = pd.to_datetime(df['Data'], format='%d-%m-%Y')
        df = df.sort_values('Data')
        df['Data'] = df['Data'].dt.strftime('%d-%m-%Y')
        df['Valor'] = df['Valor'].map(lambda x: f"{x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        df.to_excel(excel_path, sheet_name='Extrato', index=False)
        logger.info(f"Arquivo Excel criado com sucesso: {excel_path}")
        return True
    except Exception as e:
        logger.error(f"Erro ao criar arquivo Excel: {str(e)}")
        return False

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/excel', methods=['GET'])
def excel_page():
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
            transactions = extract_transactions_from_pdf(pdf_path)
            if not transactions:
                return 'Erro ao processar o PDF', 400
            excel_filename = f"{timestamp}_{filename.replace('.pdf', '.xlsx')}"
            excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
            if not create_excel(transactions, excel_path):
                return 'Erro ao criar arquivo Excel', 400
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
    except Exception as e:
        logger.error(f"Erro no upload: {str(e)}")
        return 'Erro interno do servidor', 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
