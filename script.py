import json
import pandas as pd

# Exemplo de payload JSON estruturado (dados recebidos de uma API)
json_data = '''
[
    {"Client Name": "Leanne Graham", "Email Address": "Sincere@april.biz", "City": "Gwenborough", "Company Name": "Romaguera-Clyne"},
    {"Client Name": "Ervin Howell", "Email Address": "Shanna@melissa.tv", "City": "Wisokyburgh", "Company Name": "Deckow-Crist"},
    {"Client Name": "Clementine Bauch", "Email Address": "Nathan@yesenia.net", "City": "McKenziehaven", "Company Name": "Romaguera-Jacobson"}
]
'''

def parse_and_export_data():
    print("[*] Starting data extraction workflow...")
    
    # Faz o parsing do payload JSON para uma lista Python
    user_data = json.loads(json_data)
    final_list = []
    
    # Processa e estrutura cada registo recebido
    for user in user_data:
        record = {
            "Client Name": user.get("Client Name"),
            "Email Address": user.get("Email Address"),
            "City": user.get("City"),
            "Company Name": user.get("Company Name")
        }
        final_list.append(record)
        
    # Converte a lista estruturada para um DataFrame do Pandas
    df = pd.DataFrame(final_list)
    
    # Exporta os dados organizados para uma planilha Excel (.xlsx)
    output_file = "clean_company_directory.xlsx"
    df.to_excel(output_file, index=False)
    
    print(f"[+] Success! The spreadsheet '{output_file}' has been generated.")
    print(f"[#] {len(final_list)} records successfully structured.")

if __name__ == "__main__":
    parse_and_export_data()
