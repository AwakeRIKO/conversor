def extract_transactions_from_pdf(pdf_path):
    """Extrai as transações do PDF no formato específico do extrato."""
    transactions = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                lines = text.split('\n')

                # Pular o cabeçalho
                start_index = 0
                for i, line in enumerate(lines):
                    if "DETALHE DOS MOVIMENTOS" in line:
                        start_index = i + 2  # Pular a linha de cabeçalho
                        break

                current_date = None
                for line in lines[start_index:]:
                    # Ignorar linhas de cabeçalho e rodapé
                    if "Data" in line and "Descrição" in line:
                        continue
                    if "Data de geração:" in line:
                        break

                    # Tentar extrair a data e transação
                    parts = line.strip().split()
                    if len(parts) >= 4:  # Garantir que há elementos suficientes
                        # Verificar se o primeiro elemento é uma data
                        date_pattern = r'\d{2}-\d{2}-\d{2}'
                        if re.match(date_pattern, parts[0]):
                            current_date = parts[0]

                            # Encontrar o ID da operação (número com 11 dígitos)
                            operation_id = None
                            description = []
                            value = None
                            balance = None

                            for i, part in enumerate(parts):
                                if re.match(r'\d{11}', part):
                                    operation_id = part
                                    # Descrição é tudo entre a data e o ID
                                    description = ' '.join(parts[1:i])
                                    # Valor e saldo são os dois últimos campos com R$
                                    for j in range(i+1, len(parts)):
                                        if 'R$' in parts[j]:
                                            if value is None:
                                                value = parts[j].replace('R$', '').replace('.', '').replace(',', '.').strip()
                                            else:
                                                balance = parts[j].replace('R$', '').replace('.', '').replace(',', '.').strip()
                                    break

                            if operation_id and description and value and balance:
                                try:
                                    value = float(value)
                                    balance = float(balance)
                                    transactions.append({
                                        'Data': current_date,
                                        'Descrição': description.strip(),
                                        'ID da operação': operation_id,
                                        'Valor': value,
                                        'Saldo': balance
                                    })
                                    print(f"Transação encontrada: {description.strip()} - {value} - {balance}")
                                except ValueError as e:
                                    print(f"Erro ao converter valor: {e}")
                                    continue

        return transactions
    except Exception as e:
        print(f"Erro ao processar o PDF {pdf_path}: {str(e)}")
        return None

def create_excel(transactions, excel_path):
    """Cria arquivo Excel com as transações."""
    try:
        df = pd.DataFrame(transactions)
        # Formatar valores monetários
        df['Valor'] = df['Valor'].apply(lambda x: f"R$ {x:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.'))
        df['Saldo'] = df['Saldo'].apply(lambda x: f"R$ {x:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.'))

        # Criar o arquivo Excel
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Extrato', index=False)

            # Ajustar largura das colunas
            worksheet = writer.sheets['Extrato']
            for idx, col in enumerate(df.columns):
                max_length = max(df[col].astype(str).apply(len).max(), len(col))
                worksheet.column_dimensions[chr(65 + idx)].width = max_length + 2

        format_excel(excel_path)

    except Exception as e:
        print(f"Erro ao criar arquivo Excel: {str(e)}")
