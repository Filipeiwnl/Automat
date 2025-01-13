from application.spreadsheet_service import process_spreadsheet

def start_cli():
    print("Iniciando sistema de atualização de planilhas...")

    try:
        process_spreadsheet()
        print("Atualização concluída com sucesso.")
    except Exception as e:
        print(f"Erro durante o processamento: {e}")
