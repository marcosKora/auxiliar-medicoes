import os
import sys
import subprocess
import urllib.request
import importlib

# Configurações
GITHUB_RAW_URL = "https://raw.githubusercontent.com/marcosKora/auxiliar-medicoes/refs/heads/main/auxMedWeb.py"
LOCAL_SCRIPT = "app_baixado.py"
REQUIRED_PACKAGES = ["eel", "selenium", "requests"]

def install_package(package):
    """Instala um pacote pip se não existir"""
    try:
        importlib.import_module(package.replace("-", "_"))
        print(f"✓ {package} já instalado")
    except ImportError:
        print(f"Instalando {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

def download_script():
    """Baixa o script do GitHub"""
    print(f"Baixando script de: {GITHUB_RAW_URL}")
    try:
        urllib.request.urlretrieve(GITHUB_RAW_URL, LOCAL_SCRIPT)
        print(f"✓ Script salvo como: {LOCAL_SCRIPT}")
        return True
    except Exception as e:
        print(f"✗ Erro ao baixar: {e}")
        return False

def main():
    print("=" * 50)
    print("  LAUNCHER - MEDIÇÕES v3")
    print("=" * 50)
    
    # Verifica/instala dependências
    print("\n[1/3] Verificando dependências...")
    for package in REQUIRED_PACKAGES:
        install_package(package)
    
    # Baixa o script mais recente
    print("\n[2/3] Baixando versão mais recente...")
    if not download_script():
        print("Usando script local se existir...")
        if not os.path.exists(LOCAL_SCRIPT):
            input("Nenhum script encontrado. Pressione Enter para sair...")
            return
    
    # Executa o script
    print("\n[3/3] Iniciando aplicação...")
    print("=" * 50)
    
    # Passa argumentos adicionais (ex: --headless)
    args = [sys.executable, LOCAL_SCRIPT] + sys.argv[1:]
    subprocess.run(args)

if __name__ == "__main__":
    main()
