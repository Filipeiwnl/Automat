from infrastructure.excel_handler import load_excel, save_excel
from application.data_service import fetch_data_for_designator
from domain.utils import is_numeric
from datetime import datetime
from dotenv import load_dotenv
import os
load_dotenv()


def process_spreadsheet():
    file_path = "Planilhas/OTSS.xlsx"
    base_url = os.getenv('BASE_URL')
    df = load_excel(file_path)
    print("Colunas da planilha carregada:")
    print(df.columns)

    # Garante que todas as colunas necessárias estejam presentes
    required_columns = [
        'OTS', 'TX A', 'RX A', 'Span atual A', 'Span projetado A', 'Histórico A',
        'Equipamento de Envio (A)', 'TX B', 'RX B', 'Span atual B', 'Span projetado B',
        'Histórico B', 'Equipamento de Recepção (B)', 'KM: A <> B', 'Status A', 'Status B',
        'Trat Niveis A', 'Trat Niveis B', 'DB/KM A', 'DB/KM B', 'ATUALIZADO EM',
        'Consulta URL', 'Consulta Sucesso'
    ]
    for col in required_columns:
        if col not in df.columns:
            df[col] = ''

    print(f"Total de linhas na planilha: {len(df)}")

    option = input(
        "Escolha a ação:\n1. Atualizar informações específicas\n"
        "2. Atualizar toda a planilha\n"
        "3. Atualizar apenas novos dados\n"
        "Digite o número da opção desejada: "
    )

    if option == '1':
        designators = input("Digite os designadores (OTS) separados por vírgula: ").split(',')
        designators = [d.strip() for d in designators]
        df = update_spreadsheet(df, base_url, designators=designators)
    elif option == '2':
        df = update_spreadsheet(df, base_url)
    elif option == '3':
        df = update_spreadsheet(df, base_url, only_new=True)
    else:
        print("Opção inválida. Saindo do programa.")
        return

    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    df['ATUALIZADO EM'] = timestamp

    save_excel(file_path, df)
    print("Planilha salva com sucesso.")

def update_spreadsheet(df, base_url, designators=None, only_new=False):
    updated_rows = []

    for index, row in df.iterrows():
        designator = row['OTS']

        if designators and designator not in designators:
            continue

        if only_new and row.get('Equipamento de Envio (A)'):
            print(f"Linha {index + 1} já tem dados, pulando...")
            continue

        if designator:
            data = fetch_data_for_designator(base_url, designator)
            if data:
                for key, value in data.items():
                    df.at[index, key] = value
                updated_rows.append(index)
            else:
                print(f"Erro ao buscar dados para o designador {designator}.")

    return df
