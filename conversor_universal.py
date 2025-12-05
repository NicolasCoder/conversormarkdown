import os
import sys
from markitdown import MarkItDown
from colorama import init, Fore, Style

# Inicializa o colorama para que as cores funcionem no terminal do Windows
init(autoreset=True)

# --- CONFIGURAÇÕES ---
# Lista de extensões de arquivo que o programa tentará converter
# Adicione ou remova extensões conforme necessário
SUPPORTED_EXTENSIONS = (
    '.pdf', '.docx', '.pptx', '.xlsx', '.html', '.htm', 
    '.jpg', '.jpeg', '.png', '.mp3', '.wav', '.csv', 
    '.json', '.xml', '.epub'
)

# --- FUNÇÕES PRINCIPAIS ---

def print_welcome_message():
    """ Exibe a mensagem de boas-vindas com arte ASCII e instruções. """
    welcome_art = r"""
     __  __               _   _   _       _         _
    |  \/  | __ _ _ __ __| | | | | |_ __ | |__   __| |
    | |\/| |/ _` | '__/ _` | | | | | '_ \| '_ \ / _` |
    | |  | | (_| | | | (_| | | |_| | | | | | | | (_| |
    |_|  |_|\__,_|_|  \__,_|  \___/|_| |_|_| |_|\__,_|
    """
    print(Fore.CYAN + welcome_art)
    print(Style.BRIGHT + "Bem-vindo ao Conversor Universal para Markdown!")
    print("-" * 50)
    print(Fore.YELLOW + "INSTRUÇÕES:")
    print("1. Crie uma pasta em algum lugar (ex: na sua Área de Trabalho).")
    print(f"2. Coloque este programa ({os.path.basename(sys.executable)}) dentro dessa pasta.")
    print("3. Arraste um arquivo OU uma pasta para esta janela do terminal.")
    print("4. Pressione a tecla 'Enter'.")
    print("\nOs arquivos convertidos serão salvos aqui, nesta mesma pasta.")
    print("-" * 50)

def convert_single_file(file_path, output_dir, md_converter):
    """ Converte um único arquivo e exibe o status. """
    try:
        print(f"-> Convertendo '{os.path.basename(file_path)}'...", end='', flush=True)
        result = md_converter.convert(file_path)
        
        file_name_stem = Path(file_path).stem
        output_path = os.path.join(output_dir, f"{file_name_stem}.md")

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(result.text_content)
            
        print(Fore.GREEN + " Sucesso!")
        return True
    except Exception as e:
        print(Fore.RED + " Falha!")
        print(Fore.RED + f"   Erro: {e}")
        return False

def main():
    """ Função principal que executa o loop do programa. """
    print_welcome_message()
    md_converter = MarkItDown()

    while True:
        try:
            # Pede ao usuário para arrastar o arquivo/pasta
            user_input = input(Fore.MAGENTA + "\nArraste o arquivo ou pasta aqui e pressione Enter (ou feche a janela para sair): ")
            
            # Limpa o caminho (terminais às vezes adicionam aspas)
            clean_path = user_input.strip().strip('"')

            if not os.path.exists(clean_path):
                print(Fore.RED + "Erro: O arquivo ou pasta não foi encontrado. Tente arrastar novamente.")
                continue

            # --- LÓGICA PARA ARQUIVO ÚNICO ---
            if os.path.isfile(clean_path):
                output_directory = "arquivos_convertidos"
                os.makedirs(output_directory, exist_ok=True)
                print("\nIniciando conversão de arquivo único...")
                convert_single_file(clean_path, output_directory, md_converter)
                print(Fore.GREEN + f"\nConcluído! O arquivo convertido está na pasta '{output_directory}'.")

            # --- LÓGICA PARA PASTA ---
            elif os.path.isdir(clean_path):
                folder_name = os.path.basename(clean_path)
                output_directory = f"{folder_name}_convertido"
                os.makedirs(output_directory, exist_ok=True)

                print(f"\nIniciando conversão da pasta '{folder_name}'...")
                print(f"Os resultados serão salvos em '{output_directory}'.")
                
                files_to_convert = [f for f in os.listdir(clean_path) if f.lower().endswith(SUPPORTED_EXTENSIONS)]
                
                if not files_to_convert:
                    print(Fore.YELLOW + "Nenhum arquivo compatível encontrado na pasta.")
                    continue

                total = len(files_to_convert)
                success_count = 0
                for i, filename in enumerate(files_to_convert):
                    print(f"\n[{i+1}/{total}]", end=' ')
                    full_path = os.path.join(clean_path, filename)
                    if convert_single_file(full_path, output_directory, md_converter):
                        success_count += 1
                
                print(Fore.GREEN + f"\nConcluído! {success_count} de {total} arquivos foram convertidos com sucesso.")

            else:
                print(Fore.RED + "Entrada inválida. Por favor, arraste um arquivo ou uma pasta.")

        except (KeyboardInterrupt, EOFError):
            # Permite que o usuário feche o programa com Ctrl+C
            print("\nSaindo do programa. Até mais!")
            break
        except Exception as e:
            print(Fore.RED + f"\nOcorreu um erro inesperado: {e}")
            print(Fore.YELLOW + "Por favor, tente novamente ou feche o programa.")

if __name__ == "__main__":
    from pathlib import Path
    main()