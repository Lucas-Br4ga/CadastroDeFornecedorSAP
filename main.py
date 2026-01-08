import sys
import os
from pathlib import Path
from PySide6.QtWidgets import QApplication
from layout.TelaInicial import TelaInicial

# Adiciona o diretório raiz ao path para imports
ROOT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT_DIR))



def main():
    """
    Ponto de entrada principal da aplicação.
    Inicializa a interface gráfica com tela de drag & drop.
    """
    app = QApplication(sys.argv)
    
    # Define o estilo global da aplicação
    app.setStyle("Fusion")
    
    # Cria e exibe a tela inicial
    tela_inicial = TelaInicial()
    tela_inicial.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()