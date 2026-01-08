"""
Tela para gerenciamento de anexos de arquivos do fornecedor.
Valida anexos obrigatórios e permite anexos opcionais.
"""
import sys
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFrame, QScrollArea, QFileDialog, QMessageBox, QLineEdit,
    QDialog, QDialogButtonBox
)
from PySide6.QtGui import QFont, QPalette, QColor, QPixmap
from PySide6.QtCore import Qt

# Adiciona raiz ao path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from Extrator.GerenciadorAnexos import GerenciadorAnexos, obter_caminho_anexos_json


class DialogoNomeAnexo(QDialog):
    """Diálogo para solicitar nome de anexo opcional"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Nome do Anexo")
        self.resize(400, 150)
        
        layout = QVBoxLayout(self)
        
        # Label
        label = QLabel("Digite o nome para este anexo:")
        label.setFont(QFont("Segoe UI", 11))
        layout.addWidget(label)
        
        # Campo de entrada
        self.campo_nome = QLineEdit()
        self.campo_nome.setFont(QFont("Segoe UI", 11))
        self.campo_nome.setPlaceholderText("Ex: CONTRATO SOCIAL")
        self.campo_nome.setMinimumHeight(40)
        self.campo_nome.setStyleSheet("""
            QLineEdit {
                border: 2px solid #dfe6e9;
                border-radius: 8px;
                padding: 10px;
            }
            QLineEdit:focus {
                border: 2px solid #00adef;
            }
        """)
        layout.addWidget(self.campo_nome)
        
        # Aviso
        aviso = QLabel("O nome será automaticamente convertido para MAIÚSCULAS")
        aviso.setFont(QFont("Segoe UI", 9))
        aviso.setStyleSheet("color: #7f8c8d;")
        layout.addWidget(aviso)
        
        layout.addStretch()
        
        # Botões
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
    
    def obter_nome(self) -> str:
        """Retorna nome digitado (já padronizado)"""
        return GerenciadorAnexos.padronizar_nome(self.campo_nome.text())


class TelaAnexos(QWidget):
    """
    Tela para gerenciamento de anexos de fornecedor.
    """
    
    def __init__(self, callback_continuar=None):
        super().__init__()
        
        self.callback_continuar = callback_continuar
        # MODIFICADO: limpar_ao_iniciar=True para sempre começar vazio
        self.gerenciador = GerenciadorAnexos(obter_caminho_anexos_json(), limpar_ao_iniciar=True)
        
        self.setWindowTitle("Anexos do Fornecedor")
        self._ajustar_tamanho_janela()
        
        self.apply_light_theme()
        self._build_ui()
        self._atualizar_interface()
    
    def _ajustar_tamanho_janela(self):
        """Ajusta tamanho da janela"""
        try:
            from PySide6.QtGui import QGuiApplication
            
            screen = QGuiApplication.primaryScreen()
            if screen:
                screen_geometry = screen.availableGeometry()
                screen_width = screen_geometry.width()
                screen_height = screen_geometry.height()
                
                window_width = int(screen_width * 0.70)
                window_height = int(screen_height * 0.85)
                
                window_width = max(1000, min(window_width, 1600))
                window_height = max(600, min(window_height, 1000))
                
                self.resize(window_width, window_height)
                
                x = (screen_width - window_width) // 2
                y = (screen_height - window_height) // 2
                self.move(x, y)
                
                print(f"[INFO] Janela de anexos ajustada: {window_width}x{window_height} (Tela: {screen_width}x{screen_height})")
            else:
                self.resize(1200, 800)
        except Exception as e:
            print(f"[AVISO] Erro ao ajustar tamanho da janela: {e}")
            self.resize(1200, 800)
    
    def _build_ui(self):
        """Constrói interface"""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Header
        header = self._criar_header()
        main_layout.addWidget(header)
        
        # Container principal
        container = QWidget()
        container.setStyleSheet("background-color: #f5f7fa;")
        
        scroll = QScrollArea()
        scroll.setWidget(container)
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet("QScrollArea { border: none; }")
        
        self.scroll_layout = QVBoxLayout(container)
        self.scroll_layout.setContentsMargins(40, 30, 40, 30)
        self.scroll_layout.setSpacing(30)
        
        # Título
        titulo = QLabel("Anexos do Fornecedor")
        titulo.setFont(QFont("Segoe UI", 22, QFont.Bold))
        titulo.setStyleSheet("color: #2c3e50; background-color: transparent;")
        self.scroll_layout.addWidget(titulo)
        
        subtitulo = QLabel("Anexe os documentos obrigatórios e, opcionalmente, documentos adicionais")
        subtitulo.setFont(QFont("Segoe UI", 11))
        subtitulo.setStyleSheet("color: #7f8c8d; margin-bottom: 10px; background-color: transparent;")
        self.scroll_layout.addWidget(subtitulo)
        
        # Seção de anexos obrigatórios
        self.secao_obrigatorios = self._criar_secao_obrigatorios()
        self.scroll_layout.addWidget(self.secao_obrigatorios)
        
        # Seção de anexos opcionais
        self.secao_opcionais = self._criar_secao_opcionais()
        self.scroll_layout.addWidget(self.secao_opcionais)
        
        self.scroll_layout.addStretch()
        main_layout.addWidget(scroll)
        
        # Rodapé
        self._criar_rodape(main_layout)
    
    def _criar_header(self):
        """Cria header com logo"""
        header = QFrame()
        header.setStyleSheet("""
            QFrame {
                background-color: white;
                border-bottom: 1px solid #e1e8ed;
            }
        """)
        header.setMinimumHeight(90)
        
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(40, 15, 40, 15)
        
        # Logo
        logo_label = QLabel()
        logo_path = Path(__file__).parent / "img" / "logo.png"
        
        if logo_path.exists():
            pixmap = QPixmap(str(logo_path))
            if not pixmap.isNull():
                pixmap = pixmap.scaled(200, 60, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                logo_label.setPixmap(pixmap)
            else:
                logo_label.setText("Logo")
                logo_label.setStyleSheet("color: #00adef; font-size: 16px; font-weight: bold; background-color: transparent;")
        else:
            logo_label.setText("Logo")
            logo_label.setStyleSheet("color: #00adef; font-size: 16px; font-weight: bold; background-color: transparent;")
            print(f"[AVISO] Logo não encontrada em: {logo_path}")
        
        logo_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        header_layout.addWidget(logo_label)
        header_layout.addStretch()
        
        titulo_header = QLabel("Sistema de Cadastro de Fornecedores")
        titulo_header.setFont(QFont("Segoe UI", 13, QFont.Medium))
        titulo_header.setStyleSheet("color: #2c3e50; background-color: transparent;")
        header_layout.addWidget(titulo_header)
        
        return header
    
    def _criar_secao_obrigatorios(self):
        """Cria seção de anexos obrigatórios"""
        secao = QWidget()
        secao_layout = QVBoxLayout(secao)
        secao_layout.setContentsMargins(0, 0, 0, 0)
        secao_layout.setSpacing(12)
        
        # Título
        titulo = QLabel("Anexos Obrigatórios")
        titulo.setFont(QFont("Segoe UI", 16, QFont.Bold))
        titulo.setStyleSheet("color: #2c3e50; padding-left: 5px; background-color: transparent;")
        secao_layout.addWidget(titulo)
        
        # Card
        card = QFrame()
        card.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 12px;
                border: 1px solid #e1e8ed;
            }
        """)
        
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(30, 25, 30, 25)
        card_layout.setSpacing(15)
        
        # Cria linha para cada anexo obrigatório
        self.widgets_obrigatorios = {}
        
        for nome in GerenciadorAnexos.ANEXOS_OBRIGATORIOS:
            linha = self._criar_linha_anexo_obrigatorio(nome)
            card_layout.addWidget(linha)
            self.widgets_obrigatorios[nome] = linha
        
        secao_layout.addWidget(card)
        return secao
    
    def _criar_linha_anexo_obrigatorio(self, nome: str):
        """Cria linha para um anexo obrigatório"""
        container = QFrame()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 10, 0, 10)
        
        # Label do nome
        label_nome = QLabel(f"• {nome}")
        label_nome.setFont(QFont("Segoe UI", 11, QFont.Medium))
        label_nome.setStyleSheet("color: #2c3e50; background-color: transparent;")
        label_nome.setMinimumWidth(350)
        layout.addWidget(label_nome)
        
        # Label do status/arquivo
        label_arquivo = QLabel("Nenhum arquivo anexado")
        label_arquivo.setFont(QFont("Segoe UI", 10))
        label_arquivo.setStyleSheet("color: #95a5a6; background-color: transparent;")
        label_arquivo.setWordWrap(True)
        layout.addWidget(label_arquivo, 1)
        
        # Botão anexar
        btn_anexar = QPushButton("📎 Anexar")
        btn_anexar.setFont(QFont("Segoe UI", 10))
        btn_anexar.setMinimumHeight(38)
        btn_anexar.setMinimumWidth(120)
        btn_anexar.setCursor(Qt.PointingHandCursor)
        btn_anexar.setStyleSheet("""
            QPushButton {
                background-color: #00adef;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #0099d6;
            }
        """)
        btn_anexar.clicked.connect(lambda: self._anexar_obrigatorio(nome))
        layout.addWidget(btn_anexar)
        
        # Armazena referências
        container.label_arquivo = label_arquivo
        container.btn_anexar = btn_anexar
        
        return container
    
    def _criar_secao_opcionais(self):
        """Cria seção de anexos opcionais"""
        secao = QWidget()
        secao_layout = QVBoxLayout(secao)
        secao_layout.setContentsMargins(0, 0, 0, 0)
        secao_layout.setSpacing(12)
        
        # Título
        titulo = QLabel("Anexos Opcionais")
        titulo.setFont(QFont("Segoe UI", 16, QFont.Bold))
        titulo.setStyleSheet("color: #2c3e50; padding-left: 5px; background-color: transparent;")
        secao_layout.addWidget(titulo)
        
        # Card
        card = QFrame()
        card.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 12px;
                border: 1px solid #e1e8ed;
            }
        """)
        
        self.card_layout_opcionais = QVBoxLayout(card)
        self.card_layout_opcionais.setContentsMargins(30, 25, 30, 25)
        self.card_layout_opcionais.setSpacing(15)
        
        # Área de anexos opcionais (será populada dinamicamente)
        self.container_lista_opcionais = QWidget()
        self.layout_lista_opcionais = QVBoxLayout(self.container_lista_opcionais)
        self.layout_lista_opcionais.setContentsMargins(0, 0, 0, 0)
        self.layout_lista_opcionais.setSpacing(10)
        self.card_layout_opcionais.addWidget(self.container_lista_opcionais)
        
        # Botão adicionar opcional
        btn_adicionar = QPushButton("➕ Adicionar Anexo Opcional")
        btn_adicionar.setFont(QFont("Segoe UI", 11))
        btn_adicionar.setMinimumHeight(45)
        btn_adicionar.setCursor(Qt.PointingHandCursor)
        btn_adicionar.setStyleSheet("""
            QPushButton {
                background-color: white;
                color: #00adef;
                border: 2px dashed #00adef;
                border-radius: 10px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #f0f9ff;
            }
        """)
        btn_adicionar.clicked.connect(self._adicionar_anexo_opcional)
        self.card_layout_opcionais.addWidget(btn_adicionar)
        
        secao_layout.addWidget(card)
        return secao
    
    def _criar_linha_anexo_opcional(self, nome: str, caminho: str):
        """Cria linha para um anexo opcional"""
        container = QFrame()
        container.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border-radius: 8px;
                padding: 5px;
            }
        """)
        
        layout = QHBoxLayout(container)
        layout.setContentsMargins(15, 10, 15, 10)
        
        # Nome
        label_nome = QLabel(f"📄 {nome}")
        label_nome.setFont(QFont("Segoe UI", 11, QFont.Medium))
        label_nome.setStyleSheet("color: #2c3e50; background-color: transparent;")
        label_nome.setMinimumWidth(300)
        layout.addWidget(label_nome)
        
        # Arquivo
        nome_arquivo = Path(caminho).name
        label_arquivo = QLabel(nome_arquivo)
        label_arquivo.setFont(QFont("Segoe UI", 10))
        label_arquivo.setStyleSheet("color: #7f8c8d; background-color: transparent;")
        label_arquivo.setWordWrap(True)
        layout.addWidget(label_arquivo, 1)
        
        # Botão remover
        btn_remover = QPushButton("🗑️ Remover")
        btn_remover.setFont(QFont("Segoe UI", 10))
        btn_remover.setMinimumHeight(35)
        btn_remover.setMinimumWidth(100)
        btn_remover.setCursor(Qt.PointingHandCursor)
        btn_remover.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 6px 12px;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        btn_remover.clicked.connect(lambda: self._remover_anexo_opcional(nome))
        layout.addWidget(btn_remover)
        
        return container
    
    def _criar_rodape(self, parent_layout):
        """Cria rodapé com botões"""
        rodape = QFrame()
        rodape.setStyleSheet("""
            QFrame {
                background-color: white;
                border-top: 1px solid #e1e8ed;
            }
        """)
        
        rodape_layout = QHBoxLayout(rodape)
        rodape_layout.setContentsMargins(40, 20, 40, 20)
        
        # Status
        self.status_label = QLabel("")
        self.status_label.setFont(QFont("Segoe UI", 10, QFont.Medium))
        self.status_label.setStyleSheet("background-color: transparent;")
        rodape_layout.addWidget(self.status_label)
        
        rodape_layout.addStretch()
        
        # Botão voltar
        btn_voltar = QPushButton("← Voltar")
        btn_voltar.setFont(QFont("Segoe UI", 11))
        btn_voltar.setMinimumHeight(48)
        btn_voltar.setMinimumWidth(140)
        btn_voltar.setCursor(Qt.PointingHandCursor)
        btn_voltar.setStyleSheet("""
            QPushButton {
                background-color: white;
                color: #2c3e50;
                border: 2px solid #bdc3c7;
                border-radius: 10px;
                padding: 10px 20px;
            }
            QPushButton:hover {
                background-color: #f8f9fa;
            }
        """)
        btn_voltar.clicked.connect(self.close)
        rodape_layout.addWidget(btn_voltar)
        
        # Botão continuar
        self.btn_continuar = QPushButton("🚀 Continuar para Automação")
        self.btn_continuar.setFont(QFont("Segoe UI", 11, QFont.Bold))
        self.btn_continuar.setMinimumHeight(48)
        self.btn_continuar.setMinimumWidth(240)
        self.btn_continuar.setCursor(Qt.PointingHandCursor)
        self.btn_continuar.setStyleSheet("""
            QPushButton {
                background-color: #00adef;
                color: white;
                border: none;
                border-radius: 10px;
                padding: 10px 25px;
            }
            QPushButton:hover {
                background-color: #0099d6;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
                color: #ecf0f1;
            }
        """)
        self.btn_continuar.clicked.connect(self._continuar_automacao)
        rodape_layout.addWidget(self.btn_continuar)
        
        parent_layout.addWidget(rodape)
    
    def _anexar_obrigatorio(self, nome: str):
        """Abre diálogo para anexar arquivo obrigatório"""
        arquivo, _ = QFileDialog.getOpenFileName(
            self,
            f"Selecionar arquivo: {nome}",
            "",
            "Todos os arquivos (*.*)"
        )
        
        if arquivo:
            sucesso, mensagem = self.gerenciador.adicionar_obrigatorio(nome, arquivo)
            
            if sucesso:
                self.gerenciador.salvar_dados()
                self._atualizar_interface()
            else:
                QMessageBox.warning(self, "Erro", mensagem)
    
    def _adicionar_anexo_opcional(self):
        """Adiciona novo anexo opcional"""
        # Primeiro solicita o nome
        dialogo = DialogoNomeAnexo(self)
        
        if dialogo.exec() != QDialog.Accepted:
            return
        
        nome_customizado = dialogo.obter_nome()
        
        if not nome_customizado:
            QMessageBox.warning(self, "Atenção", "Nome do anexo não pode ser vazio")
            return
        
        # Depois seleciona o arquivo
        arquivo, _ = QFileDialog.getOpenFileName(
            self,
            f"Selecionar arquivo para: {nome_customizado}",
            "",
            "Todos os arquivos (*.*)"
        )
        
        if arquivo:
            sucesso, mensagem = self.gerenciador.adicionar_opcional(nome_customizado, arquivo)
            
            if sucesso:
                self.gerenciador.salvar_dados()
                self._atualizar_interface()
            else:
                QMessageBox.warning(self, "Erro", mensagem)
    
    def _remover_anexo_opcional(self, nome: str):
        """Remove anexo opcional"""
        resposta = QMessageBox.question(
            self,
            "Confirmar Remoção",
            f"Deseja remover o anexo '{nome}'?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if resposta == QMessageBox.Yes:
            sucesso, mensagem = self.gerenciador.remover_opcional(nome)
            
            if sucesso:
                self.gerenciador.salvar_dados()
                self._atualizar_interface()
            else:
                QMessageBox.warning(self, "Erro", mensagem)
    
    def _atualizar_interface(self):
        """Atualiza interface com dados atuais"""
        # Atualiza anexos obrigatórios
        obrigatorios = self.gerenciador.obter_obrigatorios()
        
        for nome, container in self.widgets_obrigatorios.items():
            caminho = obrigatorios.get(nome, "")
            
            if caminho and Path(caminho).exists():
                nome_arquivo = Path(caminho).name
                container.label_arquivo.setText(f"✅ {nome_arquivo}")
                container.label_arquivo.setStyleSheet("color: #27ae60; background-color: transparent;")
                container.btn_anexar.setText("✏️ Alterar")
            else:
                container.label_arquivo.setText("Nenhum arquivo anexado")
                container.label_arquivo.setStyleSheet("color: #95a5a6; background-color: transparent;")
                container.btn_anexar.setText("📎 Anexar")
        
        # Atualiza anexos opcionais
        # Limpa lista atual
        while self.layout_lista_opcionais.count():
            item = self.layout_lista_opcionais.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        # Adiciona anexos opcionais
        opcionais = self.gerenciador.obter_opcionais()
        
        if opcionais:
            for nome, caminho in opcionais.items():
                if Path(caminho).exists():
                    linha = self._criar_linha_anexo_opcional(nome, caminho)
                    self.layout_lista_opcionais.addWidget(linha)
        else:
            label_vazio = QLabel("Nenhum anexo opcional adicionado")
            label_vazio.setFont(QFont("Segoe UI", 10))
            label_vazio.setStyleSheet("color: #95a5a6; padding: 20px; background-color: transparent;")
            label_vazio.setAlignment(Qt.AlignCenter)
            self.layout_lista_opcionais.addWidget(label_vazio)
        
        # Atualiza status e botão
        self._atualizar_status()
    
    def _atualizar_status(self):
        """Atualiza label de status e botão continuar"""
        valido, faltantes = self.gerenciador.validar_obrigatorios()
        obrig_ok, opcionais_count = self.gerenciador.contar_anexos()
        
        if valido:
            self.btn_continuar.setEnabled(True)
            
            if opcionais_count > 0:
                self.status_label.setText(
                    f"✅ {obrig_ok}/3 obrigatórios • {opcionais_count} opcional(is)"
                )
            else:
                self.status_label.setText(f"✅ Todos os anexos obrigatórios preenchidos")
            
            self.status_label.setStyleSheet("color: #27ae60; font-weight: 500; background-color: transparent;")
        else:
            self.btn_continuar.setEnabled(False)
            self.status_label.setText(
                f"⚠️ {len(faltantes)} anexo(s) obrigatório(s) faltando"
            )
            self.status_label.setStyleSheet("color: #e74c3c; font-weight: 500; background-color: transparent;")
    
    def _continuar_automacao(self):
        """Valida e continua para automação"""
        valido, faltantes = self.gerenciador.validar_obrigatorios()
        
        if not valido:
            mensagem = "Os seguintes anexos obrigatórios estão faltando:\n\n"
            mensagem += "\n".join([f"• {nome}" for nome in faltantes])
            QMessageBox.warning(self, "Anexos Faltando", mensagem)
            return
        
        # Salva antes de continuar
        self.gerenciador.salvar_dados()
        
        # Chama callback ou fecha
        if self.callback_continuar:
            self.callback_continuar()
        else:
            QMessageBox.information(
                self,
                "Anexos Validados",
                "Todos os anexos obrigatórios foram fornecidos!\n\nO sistema está pronto para automação SAP."
            )
            self.close()
    
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