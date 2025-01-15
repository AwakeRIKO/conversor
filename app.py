import os
import pdfplumber
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from flask import Flask, request, send_file, render_template
from werkzeug.utils import secure_filename

# Funções do código original
def extract_transactions_from_pdf(pdf_path):
    """Extrai as transações do PDF e organiza em colunas."""
    transactions = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                lines = text.split('\n')
                for line in lines:
                    match = re.match(r'(\d{2}-\d{2}-\d{4}) (.+?) \d{11} R\$ ([\d\.,-]+)', line)
                    if match:
                        date, description, value = match.groups()
                        value = float(value.replace('.', '').replace(',', '.').strip())
                        transactions.append({'Data': date, 'Descrição': description.strip(), 'Valor': value})
        return transactions
    except Exception as e:
        print(f"Erro ao processar o PDF {pdf_path}: {str(e)}")
        return None

def format_excel(excel_path):
    """Ajusta o alinhamento e a largura das colunas no arquivo Excel."""
    try:
        workbook = load_workbook(excel_path)
        sheet = workbook.active
        for cell in sheet['C']:
            cell.alignment = Alignment(horizontal='right')
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
    except Exception as e:
        print(f"Erro ao formatar o arquivo Excel: {str(e)}")

def create_excel(transactions, excel_path):
    """Cria arquivo Excel com as transações."""
    try:
        df = pd.DataFrame(transactions)
        df['Valor'] = df['Valor'].map(lambda x: f"{x:,.2f}".replace('.', ','))
        df.to_excel(excel_path, sheet_name='Extrato', index=False)
        format_excel(excel_path)
    except Exception as e:
        print(f"Erro ao criar arquivo Excel: {str(e)}")

# Código Flask
app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf'}

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'Nenhum arquivo enviado', 400
    file = request.files['file']
    if file.filename == '':
        return 'Nenhum arquivo selecionado', 400
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(pdf_path)
        transactions = extract_transactions_from_pdf(pdf_path)
        if not transactions:
            return 'Erro ao processar o PDF', 400
        excel_filename = filename.replace('.pdf', '.xlsx')
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
        create_excel(transactions, excel_path)

        # Excluir o arquivo PDF após a conversão
        try:
            os.remove(pdf_path)
        except Exception as e:
            print(f"Erro ao excluir o arquivo PDF: {str(e)}")

        return send_file(excel_path, as_attachment=True)
    return 'Tipo de arquivo não permitido', 400

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
