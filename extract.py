import os
import io
import re
import openpyxl
import pdfplumber
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# Configuração das credenciais e conexão com o Google Drive
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
SERVICE_ACCOUNT_FILE = 'credentials.json'  # Substitua pelo caminho correto

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=credentials)

# Função para listar arquivos em um diretório específico do Google Drive
def list_pdfs_in_drive_folder(folder_id):
    query = f"'{folder_id}' in parents and mimeType='application/pdf'"
    results = drive_service.files().list(q=query, pageSize=100, fields="files(id, name)").execute()
    return results.get('files', [])

# Função para baixar um arquivo PDF do Google Drive
def download_pdf(file_id, file_name):
    request = drive_service.files().get_media(fileId=file_id)
    fh = io.FileIO(file_name, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    return file_name

# Função para extrair o texto desejado do PDF
def extract_pdf_data(file_path):
    with pdfplumber.open(file_path) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text()

    # Expressões regulares para extrair as informações
    nome_aluno = re.search(r'NOME DO ALUNO:\s*(.*)', text)
    carga_horaria = re.search(r'CARGA HORÁRIA CURSADA:\s*(\d+)', text)
    coeficiente_rendimento = re.search(r'COEFICIENTE DE RENDIMENTO:\s*([\d,.]+)', text)

    return {
        'nome_aluno': nome_aluno.group(1).strip() if nome_aluno else None,
        'carga_horaria': carga_horaria.group(1).strip() if carga_horaria else None,
        'coeficiente_rendimento': coeficiente_rendimento.group(1).strip() if coeficiente_rendimento else None
    }

# Função para criar a planilha e armazenar os dados
def save_to_excel(data, output_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Dados Alunos"

    # Cabeçalho
    sheet.append(["Nome do Aluno", "Carga Horária Cursada", "Coeficiente de Rendimento"])

    # Dados
    for entry in data:
        carga_horaria_numero = -1
        coeficiente_rendimento_numero = -1

        if entry['carga_horaria'] is not None:
            carga_horaria_numero = float(entry['carga_horaria'].replace('.', '').replace(',', '.'))
        
        if entry['coeficiente_rendimento'] is not None:
            coeficiente_rendimento_numero = float(entry['coeficiente_rendimento'].replace('.', '').replace(',', '.'))
        
        sheet.append([entry['nome_aluno'], carga_horaria_numero, coeficiente_rendimento_numero])

    workbook.save(output_file)

# Função principal
def main():
    # Purple
    folder_id = '<<<hash_do_diretorio>>>'

    output_file = 'saida/dados_alunos.xlsx'

    # Lista de PDFs na pasta do Google Drive
    pdf_files = list_pdfs_in_drive_folder(folder_id)
    extracted_data = []

    for pdf in pdf_files:
        print(f"Processando {pdf['name']}...")
        local_pdf = download_pdf(pdf['id'], pdf['name'])
        data = extract_pdf_data(local_pdf)
        extracted_data.append(data)
        os.remove(local_pdf)  # Remove o PDF baixado após o processamento

    # Salva os dados em uma planilha Excel
    save_to_excel(extracted_data, output_file)
    print(f"Dados salvos em {output_file}")

if __name__ == '__main__':
    main()
