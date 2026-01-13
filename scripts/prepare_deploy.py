import zipfile
import os
from pathlib import Path

def zip_project():
    # Nome do arquivo de saída
    output_filename = "rcb_deploy.zip"
    
    # Diretório atual
    source_dir = Path.cwd()
    
    # Lista de coisas para ignorar
    exclusions = [
        '.git', 
        '__pycache__', 
        '.vscode', 
        '.idea', 
        'venv', 
        'env', 
        '.zip',
        '.pyc'
    ]

    print(f"Iniciando compactação de: {source_dir}")
    print(f"Criando arquivo: {output_filename}...")

    try:
        with zipfile.ZipFile(output_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(source_dir):
                # Filtra diretórios para não entrar nos ignorados
                # Modifica a lista 'dirs' in-place
                dirs[:] = [d for d in dirs if not any(ex in d for ex in exclusions)]
                
                for file in files:
                    file_path = Path(root) / file
                    
                    # Verifica se o arquivo deve ser ignorado
                    if any(ex in file for ex in exclusions):
                        continue
                    
                    # Ignorar o próprio script de zip e logs antigos
                    if file in ['prepare_deploy.py', output_filename] or file.endswith('.log'):
                        continue

                    # Caminho relativo para o arquivo dentro do zip
                    arcname = file_path.relative_to(source_dir)
                    
                    print(f"Adicionando: {arcname}")
                    zipf.write(file_path, arcname)

        print("-" * 30)
        print(f"SUCESSO! Arquivo '{output_filename}' criado.")
        print("-" * 30)
    
    except Exception as e:
        print(f"ERRO: Algo deu errado: {e}")

if __name__ == "__main__":
    zip_project()
