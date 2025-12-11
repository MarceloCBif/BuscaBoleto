"""
Script para gerar o executável do BuscaBoleto usando PyInstaller.
Execute este script para criar o arquivo .exe distribuível.
"""

import subprocess
import sys
import os

def instalar_pyinstaller():
    """Instala o PyInstaller se não estiver instalado."""
    try:
        import PyInstaller
        print("✓ PyInstaller já está instalado.")
    except ImportError:
        print("Instalando PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✓ PyInstaller instalado com sucesso.")

def gerar_executavel():
    """Gera o executável usando PyInstaller."""
    
    # Diretório atual
    diretorio = os.path.dirname(os.path.abspath(__file__))
    
    # Arquivo principal
    main_file = os.path.join(diretorio, "main.py")
    
    # Arquivo de configuração (será incluído junto)
    config_file = os.path.join(diretorio, "config.ini")
    
    # Nome do executável
    nome_exe = "BuscaBoleto"
    
    # Comando PyInstaller
    comando = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",                    # Gera um único arquivo .exe
        "--windowed",                   # Não mostra console (aplicação GUI)
        "--name", nome_exe,             # Nome do executável
        f"--add-data={config_file};.",  # Inclui config.ini no executável
        "--clean",                      # Limpa cache antes de compilar
        "--noconfirm",                  # Não pede confirmação para sobrescrever
        main_file
    ]
    
    print("\n" + "="*60)
    print("Gerando executável BuscaBoleto...")
    print("="*60 + "\n")
    
    try:
        subprocess.check_call(comando)
        
        # Caminho do executável gerado
        exe_path = os.path.join(diretorio, "dist", f"{nome_exe}.exe")
        
        print("\n" + "="*60)
        print("✓ EXECUTÁVEL GERADO COM SUCESSO!")
        print("="*60)
        print(f"\nLocalização: {exe_path}")
        print("\n⚠️  IMPORTANTE:")
        print("   O arquivo 'config.ini' está embutido no executável,")
        print("   mas você pode colocar um 'config.ini' na mesma pasta")
        print("   do .exe para sobrescrever as configurações.")
        print("\n   Para distribuir, envie:")
        print(f"   1. {nome_exe}.exe (da pasta 'dist')")
        print("   2. config.ini (se quiser configuração externa)")
        print("="*60)
        
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Erro ao gerar executável: {e}")
        sys.exit(1)

if __name__ == "__main__":
    print("="*60)
    print("  GERADOR DE EXECUTÁVEL - BUSCA BOLETO")
    print("="*60)
    
    instalar_pyinstaller()
    gerar_executavel()
