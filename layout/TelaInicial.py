import sys
import json
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, 
    QFileDialog, QMessageBox, QProgressBar, QHBoxLayout, QFrame
)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QPalette, QColor, QFont, QPixmap

# Adiciona o diretório raiz ao path
ROOT_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT_DIR))

# Imports das classes do projeto
from Extrator.PDFExtractor import PDFExtractor
from Extrator.LimparJson import PDFCompanyExtractor
from layout.TelaInformacoes import TelaInformacoes
from utils import get_json_paths


class ProcessadorThread(QThread):
    """Thread para processar PDF sem travar a interface"""
    
    progresso = Signal(str)
    concluido = Signal(bool, str)
    
    def __init__(self, pdf_path: Path):
        super().__init__()
        self.pdf_path = pdf_path
        
        paths = get_json_paths()
        self.json_bruto_path = paths["bruto"]
        self.json_limpo_path = paths["limpo"]
    
    def run(self):
        """Executa o processamento completo do PDF"""
        try:
            # Etapa 1: Extração
            self.progresso.emit("📄 Extraindo dados do PDF...")
            extractor = PDFExtractor(str(self.pdf_path))
            extractor.save_to_json(self.json_bruto_path)
            
            # Etapa 2: Limpeza
            self.progresso.emit("📄 Processando tabelas...")
            try:
                with open(self.json_bruto_path, "r", encoding="utf-8") as f:
                    pdf_json_bruto = json.load(f)
            except Exception as e:
                raise Exception(f"Erro ao ler JSON bruto: {str(e)}")
            
            self.progresso.emit("📋 Estruturando informações...")
            try:
                cleaner = PDFCompanyExtractor(pdf_json_bruto, self.json_limpo_path)
                resultado = cleaner.extract()
                
                if not isinstance(resultado, dict):
                    raise Exception(f"Resultado da limpeza não é um dicionário: {type(resultado)}")
                    
            except Exception as e:
                raise Exception(f"Erro na limpeza de dados: {str(e)}")
            
            # NOTA: Padronização será executada manualmente quando clicar no botão SAP
            
            self.progresso.emit("✅ Processo concluído!")
            self.concluido.emit(True, "PDF processado com sucesso!\n\nOs dados serão padronizados automaticamente ao iniciar a automação SAP.")
            
        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            print(f"Erro detalhado:\n{error_detail}")
            self.progresso.emit(f"Erro: {str(e)}")
            self.concluido.emit(False, f"Erro ao processar PDF:\n{str(e)}")


class TelaInicial(QWidget):
    """
    Tela inicial com tema claro elegante e profissional.
    Logo fixa carregada de layout/img/logo.png
    """
    
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("Sistema de Cadastro de Fornecedores")
        
        # Ajusta tamanho baseado no monitor
        self._ajustar_tamanho_janela()
        
        self.setAcceptDrops(True)
        self.apply_light_theme()
        self._build_ui()
        
        self.thread_processamento = None
        self.tela_informacoes = None
    
    def _ajustar_tamanho_janela(self):
        """Ajusta o tamanho da janela baseado nas dimensões do monitor"""
        try:
            from PySide6.QtGui import QGuiApplication
            
            # Pega a tela principal
            screen = QGuiApplication.primaryScreen()
            if screen:
                screen_geometry = screen.availableGeometry()
                screen_width = screen_geometry.width()
                screen_height = screen_geometry.height()
                
                # Define tamanho como 60% da largura e 75% da altura da tela
                window_width = int(screen_width * 0.60)
                window_height = int(screen_height * 0.75)
                
                # Limites mínimos e máximos
                window_width = max(800, min(window_width, 1200))
                window_height = max(600, min(window_height, 900))
                
                self.resize(window_width, window_height)
                
                # Centraliza a janela
                x = (screen_width - window_width) // 2
                y = (screen_height - window_height) // 2
                self.move(x, y)
                
                print(f"[INFO] Janela inicial ajustada: {window_width}x{window_height} (Tela: {screen_width}x{screen_height})")
            else:
                # Fallback para tamanho padrão
                self.resize(800, 700)
        except Exception as e:
            print(f"[AVISO] Erro ao ajustar tamanho da janela: {e}")
            # Fallback para tamanho padrão
            self.resize(800, 700)
    
    def _build_ui(self):
        """Constrói a interface principal"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # Header com logo
        header = self._criar_header()
        layout.addWidget(header)
        
        # Container principal
        container = QWidget()
        container.setStyleSheet("background-color: #f5f7fa;")
        container_layout = QVBoxLayout(container)
        container_layout.setContentsMargins(50, 40, 50, 50)
        container_layout.setSpacing(25)
        
        # Título e subtítulo
        titulo = QLabel("Extrator de Informações de Fornecedores")
        titulo.setFont(QFont("Segoe UI", 24, QFont.Bold))
        titulo.setStyleSheet("color: #2c3e50;")
        titulo.setAlignment(Qt.AlignCenter)
        container_layout.addWidget(titulo)
        
        subtitulo = QLabel("Arraste um arquivo PDF ou clique para selecionar")
        subtitulo.setFont(QFont("Segoe UI", 12))
        subtitulo.setStyleSheet("color: #7f8c8d; margin-bottom: 10px;")
        subtitulo.setAlignment(Qt.AlignCenter)
        container_layout.addWidget(subtitulo)
        
        # Área de drop com card
        drop_card = QFrame()
        drop_card.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 16px;
                border: 3px dashed #bdc3c7;
            }
        """)
        drop_layout = QVBoxLayout(drop_card)
        drop_layout.setContentsMargins(40, 60, 40, 60)
        
        self.drop_area = QLabel("📄\n\nArraste o arquivo PDF aqui\nou clique no botão abaixo")
        self.drop_area.setFont(QFont("Segoe UI", 14))
        self.drop_area.setAlignment(Qt.AlignCenter)
        self.drop_area.setStyleSheet("color: #95a5a6; border: none;")
        drop_layout.addWidget(self.drop_area)
        
        container_layout.addWidget(drop_card)
        
        # Barra de progresso
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(0)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                background-color: #ecf0f1;
                border-radius: 6px;
                text-align: center;
                color: #2c3e50;
                height: 35px;
                border: 1px solid #dfe6e9;
            }
            QProgressBar::chunk {
                background-color: #00adef;
                border-radius: 6px;
            }
        """)
        self.progress_bar.setVisible(False)
        container_layout.addWidget(self.progress_bar)
        
        # Label de status
        self.status_label = QLabel("")
        self.status_label.setFont(QFont("Segoe UI", 11))
        self.status_label.setStyleSheet("color: #7f8c8d;")
        self.status_label.setAlignment(Qt.AlignCenter)
        container_layout.addWidget(self.status_label)
        
        # Botão principal
        self.btn_selecionar = QPushButton("Selecionar Arquivo PDF")
        self.btn_selecionar.setFont(QFont("Segoe UI", 13, QFont.Medium))
        self.btn_selecionar.setMinimumHeight(55)
        self.btn_selecionar.setCursor(Qt.PointingHandCursor)
        self.btn_selecionar.setStyleSheet("""
            QPushButton {
                background-color: #00adef;
                color: white;
                border: none;
                border-radius: 12px;
                padding: 15px 30px;
            }
            QPushButton:hover {
                background-color: #0099d6;
            }
            QPushButton:pressed {
                background-color: #0088bd;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
                color: #ecf0f1;
            }
        """)
        self.btn_selecionar.clicked.connect(self.selecionar_arquivo)
        container_layout.addWidget(self.btn_selecionar)
        
        # Separador
        separador = QLabel("─────── ou ───────")
        separador.setAlignment(Qt.AlignCenter)
        separador.setStyleSheet("color: #bdc3c7; margin: 15px 0;")
        container_layout.addWidget(separador)
        
        # Botões secundários
        botoes_layout = QHBoxLayout()
        botoes_layout.setSpacing(15)
        
        btn_abrir_dados = self._criar_botao_secundario(
            "📋 Abrir Dados Existentes",
            self.abrir_dados_existentes
        )
        botoes_layout.addWidget(btn_abrir_dados)
        
        btn_automacao_sap = self._criar_botao_secundario(
            "🚀 Automação SAP",
            self.abrir_automacao_sap
        )
        botoes_layout.addWidget(btn_automacao_sap)
        
        container_layout.addLayout(botoes_layout)
        container_layout.addStretch()
        
        layout.addWidget(container)
    
    def _criar_header(self):
        """Cria header com logo fixa"""
        header = QFrame()
        header.setStyleSheet("""
            QFrame {
                background-color: white;
                border-bottom: 1px solid #e1e8ed;
            }
        """)
        header.setMinimumHeight(100)
        
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(40, 20, 40, 20)
        
        # Logo fixa
        logo_label = QLabel()
        logo_path = Path(__file__).parent / "img" / "logo.png"
        
        if logo_path.exists():
            pixmap = QPixmap(str(logo_path))
            if not pixmap.isNull():
                # Redimensiona mantendo proporção
                pixmap = pixmap.scaled(200, 60, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                logo_label.setPixmap(pixmap)
            else:
                logo_label.setText("Logo")
                logo_label.setStyleSheet("color: #00adef; font-size: 16px; font-weight: bold;")
        else:
            logo_label.setText("Logo")
            logo_label.setStyleSheet("color: #00adef; font-size: 16px; font-weight: bold;")
            print(f"[AVISO] Logo não encontrada em: {logo_path}")
        
        logo_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        header_layout.addWidget(logo_label)
        header_layout.addStretch()
        
        # Info da empresa
        info_label = QLabel("Sistema de Gestão de Fornecedores")
        info_label.setFont(QFont("Segoe UI", 11, QFont.Medium))
        info_label.setStyleSheet("color: #2c3e50;")
        header_layout.addWidget(info_label)
        
        return header
    
    def _criar_botao_secundario(self, texto, callback):
        """Cria botão secundário estilizado"""
        btn = QPushButton(texto)
        btn.setFont(QFont("Segoe UI", 11))
        btn.setMinimumHeight(50)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setStyleSheet("""
            QPushButton {
                background-color: white;
                color: #2c3e50;
                border: 2px solid #00adef;
                border-radius: 10px;
                padding: 12px 25px;
            }
            QPushButton:hover {
                background-color: #f0f9ff;
                border-color: #0099d6;
            }
            QPushButton:pressed {
                background-color: #e0f2ff;
            }
        """)
        btn.clicked.connect(callback)
        return btn
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        """Evento ao arrastar arquivo"""
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and urls[0].toLocalFile().lower().endswith('.pdf'):
                event.acceptProposedAction()
                self.drop_area.parentWidget().setStyleSheet("""
                    QFrame {
                        background-color: #f0f9ff;
                        border-radius: 16px;
                        border: 3px dashed #00adef;
                    }
                """)
                self.drop_area.setStyleSheet("color: #00adef; border: none;")
    
    def dragLeaveEvent(self, event):
        """Evento ao sair da área"""
        self.drop_area.parentWidget().setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 16px;
                border: 3px dashed #bdc3c7;
            }
        """)
        self.drop_area.setStyleSheet("color: #95a5a6; border: none;")
    
    def dropEvent(self, event: QDropEvent):
        """Evento ao soltar arquivo"""
        self.dragLeaveEvent(None)
        
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.lower().endswith('.pdf'):
                self.processar_pdf(Path(file_path))
            else:
                QMessageBox.warning(self, "Erro", "Por favor, selecione um arquivo PDF.")
    
    def selecionar_arquivo(self):
        """Abre diálogo para selecionar arquivo"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar arquivo PDF",
            "",
            "Arquivos PDF (*.pdf)"
        )
        
        if file_path:
            self.processar_pdf(Path(file_path))
    
    def processar_pdf(self, pdf_path: Path):
        """Inicia o processamento do PDF"""
        self.btn_selecionar.setEnabled(False)
        self.setAcceptDrops(False)
        
        self.progress_bar.setVisible(True)
        self.status_label.setText("Iniciando processamento...")
        
        self.thread_processamento = ProcessadorThread(pdf_path)
        self.thread_processamento.progresso.connect(self.atualizar_progresso)
        self.thread_processamento.concluido.connect(self.processamento_concluido)
        self.thread_processamento.start()
    
    def atualizar_progresso(self, mensagem: str):
        """Atualiza mensagem de progresso"""
        self.status_label.setText(mensagem)
    
    def processamento_concluido(self, sucesso: bool, mensagem: str):
        """Callback de conclusão"""
        self.btn_selecionar.setEnabled(True)
        self.setAcceptDrops(True)
        self.progress_bar.setVisible(False)
        
        if sucesso:
            self.status_label.setText("✅ Processamento concluído!")
            self.status_label.setStyleSheet("color: #27ae60;")
            QMessageBox.information(self, "Sucesso", mensagem)
            self.abrir_tela_informacoes()
        else:
            self.status_label.setText("❌ Erro no processamento")
            self.status_label.setStyleSheet("color: #e74c3c;")
            QMessageBox.critical(self, "Erro", mensagem)
    
    def abrir_dados_existentes(self):
        """Abre dados já processados"""
        paths = get_json_paths()
        json_limpo_path = paths["limpo"]
        
        if not json_limpo_path.exists():
            QMessageBox.warning(
                self,
                "Dados não encontrados",
                "Nenhum dado foi encontrado.\n\nProcesse um PDF primeiro."
            )
            return
        
        self.abrir_tela_informacoes()
    
    def abrir_automacao_sap(self):
        """Abre tela focada em automação SAP"""
        paths = get_json_paths()
        json_limpo_path = paths["limpo"]
        
        if not json_limpo_path.exists():
            resposta = QMessageBox.question(
                self,
                "Dados não encontrados",
                "A automação SAP requer dados de fornecedor.\n\nDeseja processar um PDF agora?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )
            
            if resposta == QMessageBox.Yes:
                self.selecionar_arquivo()
            return
        
        self.abrir_tela_informacoes()
        QMessageBox.information(
            self,
            "Automação SAP",
            "Verifique se todos os campos obrigatórios estão preenchidos.\n\nOs dados serão padronizados automaticamente ao clicar em 'Iniciar Automação SAP'."
        )
    
    def abrir_tela_informacoes(self):
        """Abre tela de informações"""
        paths = get_json_paths()
        json_limpo_path = paths["limpo"]
        
        self.tela_informacoes = TelaInformacoes(json_limpo_path)
        self.tela_informacoes.show()
        self.hide()
    
    def apply_light_theme(self):
        """Aplica tema claro"""
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor("#f5f7fa"))
        palette.setColor(QPalette.Base, QColor("#ffffff"))
        palette.setColor(QPalette.Text, QColor("#2c3e50"))
        palette.setColor(QPalette.Button, QColor("#ffffff"))
        palette.setColor(QPalette.ButtonText, QColor("#2c3e50"))
        palette.setColor(QPalette.Highlight, QColor("#00adef"))
        palette.setColor(QPalette.HighlightedText, QColor("#ffffff"))
        self.setPalette(palette)