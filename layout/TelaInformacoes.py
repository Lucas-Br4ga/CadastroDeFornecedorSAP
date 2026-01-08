# -*- coding: utf-8 -*-
import json
import re
import sys
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QLabel, QLineEdit, QVBoxLayout,
    QHBoxLayout, QScrollArea, QFrame, QSizePolicy, QPushButton, QMessageBox, QGridLayout
)
from PySide6.QtGui import QFont, QPalette, QColor, QPixmap
from PySide6.QtCore import Qt

# Adiciona o diretório raiz ao path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils import get_json_paths
from validadores import (
    aplicar_validador, aplicar_mascara_automatica,
    EmailValidator, CNPJValidator
)
from Extrator.PadronizarDados import PadronizadorDados
from layout.TelaAnexos import TelaAnexos

class TelaInformacoes(QWidget):
    """
    Tela para visualização e edição das informações extraídas do PDF.
    Design claro, elegante e profissional com a cor da empresa.
    """

    CAMPOS_OBRIGATORIOS = {
        'empresa': ['razao_social', 'nome_fantasia', 'cnpj'],
        'endereco': ['rua', 'numero', 'bairro', 'cidade', 'estado', 'cep'],
        'contato': ['celular', 'email_comercial'],
        'bancario': ['codigo_banco', 'agencia', 'conta_corrente'],
        'geral': ['prazo_pagamento', 'modalidade_frete']
    }

    TIPOS_CAMPO = {
        'cnpj': 'cnpj',
        'cep': 'cep',
        'celular': 'telefone',
        'celular_secundario': 'telefone',
        'email_comercial': 'email',
        'email_fiscal': 'email',
        'numero': 'numero',
        'conta_corrente': 'conta',
        'agencia': 'agencia_limitada',
        'codigo_banco': 'codigo_banco',
        'razao_social': 'razao_social',
        'nome_fantasia': 'nome_fantasia',
    }

    def __init__(self, json_path: Path = None, callback_automacao=None):
        super().__init__()

        if json_path is None:
            paths = get_json_paths()
            json_path = paths["limpo"]

        self.json_path = Path(json_path)
        self.callback_automacao = callback_automacao
        self.campos_widgets = {}

        self.setWindowTitle("Informações do Fornecedor")

        # Ajusta tamanho baseado no monitor
        self._ajustar_tamanho_janela()

        self.data = self.load_json()
        self.apply_light_theme()
        self._build_ui()

        # Validar limites após carregar JSON
        problemas = self._validar_limites_json()
        if problemas:
            self._mostrar_alerta_limites(problemas)

        self.atualizar_botao_automacao()

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

                # Define tamanho como 85% da largura e 90% da altura da tela
                window_width = int(screen_width * 0.85)
                window_height = int(screen_height * 0.90)

                # Limites mínimos e máximos
                window_width = max(1200, min(window_width, 1920))
                window_height = max(700, min(window_height, 1080))

                self.resize(window_width, window_height)

                # Centraliza a janela
                x = (screen_width - window_width) // 2
                y = (screen_height - window_height) // 2
                self.move(x, y)

                print(f"[INFO] Janela ajustada: {window_width}x{window_height} (Tela: {screen_width}x{screen_height})")
            else:
                # Fallback para tamanho padrão
                self.resize(1400, 900)
        except Exception as e:
            print(f"[AVISO] Erro ao ajustar tamanho da janela: {e}")
            # Fallback para tamanho padrão
            self.resize(1400, 900)

    def _build_ui(self):
        """Constrói interface principal"""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Header com logo
        header = self._criar_header()
        main_layout.addWidget(header)

        # Container com fundo
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

        # Título da seção
        titulo_secao = QLabel("Dados do Fornecedor")
        titulo_secao.setFont(QFont("Segoe UI", 22, QFont.Bold))
        titulo_secao.setStyleSheet("""
            QLabel {
                color: #2c3e50;
                margin-bottom: 5px;
                border: none;
                background-color: transparent;
            }
        """)
        self.scroll_layout.addWidget(titulo_secao)

        subtitulo_secao = QLabel("Preencha todos os campos obrigatórios marcados com *")
        subtitulo_secao.setFont(QFont("Segoe UI", 11))
        subtitulo_secao.setStyleSheet("""
            QLabel {
                color: #7f8c8d;
                margin-bottom: 10px;
                border: none;
                background-color: transparent;
            }
        """)
        self.scroll_layout.addWidget(subtitulo_secao)

        # Campos por categoria (sem cards nos títulos)
        for categoria, campos in self.data.items():
            if isinstance(campos, dict):
                self.scroll_layout.addWidget(self.criar_secao_categoria(categoria, campos))

        self.scroll_layout.addStretch()
        main_layout.addWidget(scroll)

        # Rodapé com botões
        self._criar_rodape(main_layout)

    def _criar_header(self):
        """Cria header com logo fixa"""
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

        # Logo fixa
        logo_label = QLabel()
        logo_path = Path(__file__).parent / "img" / "logo.png"

        if logo_path.exists():
            pixmap = QPixmap(str(logo_path))
            if not pixmap.isNull():
                pixmap = pixmap.scaled(200, 60, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                logo_label.setPixmap(pixmap)
            else:
                logo_label.setText("Logo")
                logo_label.setStyleSheet("""
                    QLabel {
                        color: #00adef;
                        font-size: 16px;
                        font-weight: bold;
                        border: none;
                        background-color: transparent;
                    }
                """)
        else:
            logo_label.setText("Logo")
            logo_label.setStyleSheet("""
                QLabel {
                    color: #00adef;
                    font-size: 16px;
                    font-weight: bold;
                    border: none;
                    background-color: transparent;
                }
            """)
            print(f"[AVISO] Logo não encontrada em: {logo_path}")

        logo_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        header_layout.addWidget(logo_label)
        header_layout.addStretch()

        # Título
        titulo_header = QLabel("Sistema de Cadastro de Fornecedores")
        titulo_header.setFont(QFont("Segoe UI", 13, QFont.Medium))
        titulo_header.setStyleSheet("""
            QLabel {
                color: #2c3e50;
                border: none;
                background-color: transparent;
            }
        """)
        header_layout.addWidget(titulo_header)

        return header

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
        rodape_layout.addWidget(self.status_label)

        rodape_layout.addStretch()

        # Botão validar
        btn_validar = QPushButton("Validar Dados")
        btn_validar.setFont(QFont("Segoe UI", 11))
        btn_validar.setMinimumHeight(48)
        btn_validar.setMinimumWidth(160)
        btn_validar.setCursor(Qt.PointingHandCursor)
        btn_validar.setStyleSheet("""
            QPushButton {
                background-color: white;
                color: #2c3e50;
                border: 2px solid #bdc3c7;
                border-radius: 10px;
                padding: 10px 20px;
            }
            QPushButton:hover {
                background-color: #f8f9fa;
                border-color: #95a5a6;
            }
        """)
        btn_validar.clicked.connect(self.validar_dados_completo)
        rodape_layout.addWidget(btn_validar)

        # Botão automação SAP
        self.btn_automacao = QPushButton("Iniciar Automação SAP")
        self.btn_automacao.setFont(QFont("Segoe UI", 11, QFont.Bold))
        self.btn_automacao.setMinimumHeight(48)
        self.btn_automacao.setMinimumWidth(220)
        self.btn_automacao.setCursor(Qt.PointingHandCursor)
        self.btn_automacao.setStyleSheet("""
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
            QPushButton:pressed {
                background-color: #0088bd;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
                color: #ecf0f1;
            }
        """)
        self.btn_automacao.clicked.connect(self.iniciar_automacao_sap)
        rodape_layout.addWidget(self.btn_automacao)

        parent_layout.addWidget(rodape)

    def load_json(self):
        """Carrega dados do JSON"""
        if not self.json_path.exists():
            return self._get_empty_structure()

        try:
            with open(self.json_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return self._validate_structure(data)
        except Exception as e:
            print(f"Erro ao carregar JSON: {e}")
            return self._get_empty_structure()

    def _validate_structure(self, data: dict) -> dict:
        """Valida estrutura do JSON"""
        valid_data = self._get_empty_structure()

        for categoria in valid_data.keys():
            if categoria in data and isinstance(data[categoria], dict):
                valid_data[categoria] = data[categoria]

        return valid_data

    def _get_empty_structure(self):
        """Retorna estrutura padrão"""
        return {
            "empresa": {
                "razao_social": "",
                "nome_fantasia": "",
                "cnpj": "",
                "inscricao_estadual": "",
                "inscricao_municipal": ""
            },
            "endereco": {
                "rua": "",
                "numero": "",
                "complemento": "",
                "bairro": "",
                "cidade": "",
                "estado": "",
                "cep": ""
            },
            "contato": {
                "celular": "",
                "celular_secundario": "",
                "email_comercial": "",
                "email_fiscal": ""
            },
            "bancario": {
                "banco": "",
                "codigo_banco": "",
                "agencia": "",
                "conta_corrente": ""
            },
            "geral": {
                "prazo_pagamento": "",
                "modalidade_frete": ""
            }
        }

    def _validar_limites_json(self) -> list:
        """
        Valida se dados do JSON extraído excedem limites permitidos.
        Retorna lista de problemas encontrados.
        """
        problemas = []

        # Validar Razão Social (máx 40 caracteres)
        razao = self.data.get('empresa', {}).get('razao_social', '')
        if len(razao) > 40:
            problemas.append({
                'campo': 'Razão Social',
                'valor_atual': razao,
                'tamanho_atual': len(razao),
                'limite': 40,
                'tipo': 'excede_limite'
            })

        # Validar Nome Fantasia (máx 20 caracteres)
        fantasia = self.data.get('empresa', {}).get('nome_fantasia', '')
        if len(fantasia) > 20:
            problemas.append({
                'campo': 'Nome Fantasia',
                'valor_atual': fantasia,
                'tamanho_atual': len(fantasia),
                'limite': 20,
                'tipo': 'excede_limite'
            })

        # Validar Código do Banco (máx 3 dígitos)
        codigo_banco = self.data.get('bancario', {}).get('codigo_banco', '')
        numeros_banco = re.sub(r'\D', '', str(codigo_banco))
        if len(numeros_banco) > 3:
            problemas.append({
                'campo': 'Código do Banco',
                'valor_atual': codigo_banco,
                'tamanho_atual': len(numeros_banco),
                'limite': 3,
                'tipo': 'excede_digitos'
            })

        # Validar Agência (máx 4 dígitos)
        agencia = self.data.get('bancario', {}).get('agencia', '')
        numeros_agencia = re.sub(r'\D', '', str(agencia))
        if len(numeros_agencia) > 4:
            problemas.append({
                'campo': 'Agência',
                'valor_atual': agencia,
                'tamanho_atual': len(numeros_agencia),
                'limite': 4,
                'tipo': 'excede_digitos'
            })

        return problemas

    def _mostrar_alerta_limites(self, problemas: list):
        """Mostra diálogo de alerta sobre problemas de limite encontrados"""
        if not problemas:
            return

        mensagem = "⚠️ ATENÇàO: O extrator de PDF gerou dados que excedem os limites permitidos:\n\n"
        mensagem += "Os seguintes campos precisam ser CORRIGIDOS manualmente:\n\n"

        for p in problemas:
            if p['tipo'] == 'excede_limite':
                mensagem += f"• {p['campo']}:\n"
                mensagem += f"  Atual: {p['tamanho_atual']} caracteres\n"
                mensagem += f"  Limite: {p['limite']} caracteres\n"
                mensagem += f"  Excesso: {p['tamanho_atual'] - p['limite']} caracteres\n\n"
            else:  # excede_digitos
                mensagem += f"• {p['campo']}:\n"
                mensagem += f"  Atual: {p['tamanho_atual']} dígitos\n"
                mensagem += f"  Limite: {p['limite']} dígitos\n"
                mensagem += f"  Excesso: {p['tamanho_atual'] - p['limite']} dígitos\n\n"

        mensagem += "❌ Os campos marcados em VERMELHO devem ser corrigidos antes de prosseguir.\n"
        mensagem += "Você pode editar diretamente nos campos abaixo."

        QMessageBox.warning(
            self,
            "Dados Excedem Limites Permitidos",
            mensagem
        )
    def save_json(self):
        """Salva dados no JSON"""
        try:
            self.json_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.json_path, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Erro ao salvar JSON: {e}")

    def criar_secao_categoria(self, nome_categoria, campos):
        """Cria seção com título simples e campos em card"""
        secao = QWidget()
        secao_layout = QVBoxLayout(secao)
        secao_layout.setContentsMargins(0, 0, 0, 0)
        secao_layout.setSpacing(12)

        # Título da categoria - SEM CARD, apenas texto
        titulo = QLabel(self._formatar_titulo_categoria(nome_categoria))
        titulo.setFont(QFont("Segoe UI", 16, QFont.Bold))
        titulo.setStyleSheet("""
            QLabel {
                color: #2c3e50;
                padding-left: 5px;
                border: none;
                background-color: transparent;
            }
        """)
        secao_layout.addWidget(titulo)

        # Card com os campos
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

        # Garante todos os campos da estrutura padrão
        estrutura_padrao = self._get_empty_structure()
        if nome_categoria in estrutura_padrao:
            todos_campos = estrutura_padrao[nome_categoria]
            for chave in todos_campos.keys():
                if chave not in campos:
                    campos[chave] = ""

        # Grid com 2 colunas
        grid_layout = QGridLayout()
        grid_layout.setSpacing(20)
        grid_layout.setHorizontalSpacing(30)
        grid_layout.setColumnStretch(0, 1)
        grid_layout.setColumnStretch(1, 1)

        # Organiza campos em 2 colunas
        row = 0
        col = 0
        for chave, valor in campos.items():
            campo_container = self.criar_linha_campo(nome_categoria, chave, valor)
            grid_layout.addLayout(campo_container, row, col)

            col += 1
            if col > 1:
                col = 0
                row += 1

        card_layout.addLayout(grid_layout)

        secao_layout.addWidget(card)
        return secao

    def criar_linha_campo(self, categoria, chave, valor):
        """Cria campo com label simples acima"""
        container = QVBoxLayout()
        container.setSpacing(6)

        e_obrigatorio = self.is_campo_obrigatorio(categoria, chave)

        # Label simples - SEM borda e SEM fundo
        label_texto = self.formatar_label(chave)
        if e_obrigatorio:
            label_texto += " *"

        label = QLabel(label_texto)
        label.setFont(QFont("Segoe UI", 10))
        label.setStyleSheet(f"""
            QLabel {{
                border: none;
                background-color: transparent;
                color: {'#e74c3c' if e_obrigatorio else '#555555'};
            }}
        """)
        container.addWidget(label)

        # Campo de entrada - COM borda
        campo = QLineEdit()
        campo.setText(str(valor) if valor else "")
        campo.setFont(QFont("Segoe UI", 11))
        campo.setMinimumHeight(40)

        # Aplica validadores
        tipo_campo = self.TIPOS_CAMPO.get(chave)
        if tipo_campo:
            aplicar_validador(campo, tipo_campo)
            if tipo_campo in ['cnpj', 'cep', 'telefone']:
                aplicar_mascara_automatica(campo, tipo_campo)

        self._atualizar_estilo_campo(campo, categoria, chave)

        def on_text_changed(novo_valor):
            self.atualizar_json(categoria, chave, novo_valor)
            self._atualizar_estilo_campo(campo, categoria, chave)

        campo.textChanged.connect(on_text_changed)

        campo_id = f"{categoria}.{chave}"
        self.campos_widgets[campo_id] = campo

        container.addWidget(campo)

        return container

    def _atualizar_estilo_campo(self, campo, categoria, chave):
        """Atualiza estilo do campo baseado em validação"""
        e_obrigatorio = self.is_campo_obrigatorio(categoria, chave)
        valor = campo.text().strip()

        campo_vazio = not valor
        campo_invalido = False

        if valor and e_obrigatorio:
            tipo_campo = self.TIPOS_CAMPO.get(chave)
            if tipo_campo == 'email':
                campo_invalido = not EmailValidator.is_valid(valor)
            elif tipo_campo == 'cnpj':
                import re
                numeros = re.sub(r'\D', '', valor)
                campo_invalido = len(numeros) != 14

        # Define cor da borda
        if e_obrigatorio and (campo_vazio or campo_invalido):
            border_color = "#e74c3c"
            bg_color = "#fff5f5"
        else:
            border_color = "#dfe6e9"
            bg_color = "#ffffff"

        campo.setStyleSheet(f"""
            QLineEdit {{
                background-color: {bg_color};
                border: 2px solid {border_color};
                border-radius: 8px;
                padding: 10px 14px;
                color: #2c3e50;
            }}
            QLineEdit:focus {{
                border: 2px solid #00adef;
                background-color: #f0f9ff;
            }}
        """)

    def atualizar_json(self, categoria, chave, novo_valor):
        """Atualiza JSON"""
        if categoria not in self.data:
            self.data[categoria] = {}

        self.data[categoria][chave] = novo_valor
        self.save_json()
        self.atualizar_botao_automacao()

    def is_campo_obrigatorio(self, categoria, chave):
        """Verifica se campo é obrigatório"""
        return categoria in self.CAMPOS_OBRIGATORIOS and chave in self.CAMPOS_OBRIGATORIOS[categoria]

    def validar_dados_completo(self):
        """Valida todos os campos"""
        campos_faltantes = self.get_campos_faltantes()

        if not campos_faltantes:
            QMessageBox.information(
                self,
                "“ Validação Concluída",
                "Todos os campos obrigatórios estào preenchidos corretamente!\n\nO sistema está pronto para automação SAP."
            )
        else:
            mensagem = "Os seguintes campos precisam de atençãoo:\n\n"
            mensagem += "\n".join([f" {campo}" for campo in campos_faltantes])
            QMessageBox.warning(self, "Atençãoo", mensagem)

    def get_campos_faltantes(self):
        """Retorna lista de campos com problemas"""
        campos_faltantes = []

        for categoria, campos in self.CAMPOS_OBRIGATORIOS.items():
            if categoria not in self.data:
                campos_faltantes.extend([f"{self.formatar_label(campo)} ({categoria})" for campo in campos])
                continue

            for campo in campos:
                valor = self.data[categoria].get(campo, '').strip()

                if not valor:
                    campos_faltantes.append(f"{self.formatar_label(campo)} ({self._formatar_titulo_categoria(categoria)})")
                    continue

                # Validações específicas
                tipo_campo = self.TIPOS_CAMPO.get(campo)
                if tipo_campo == 'email' and not EmailValidator.is_valid(valor):
                    campos_faltantes.append(f"{self.formatar_label(campo)} - E-mail inválido")
                elif tipo_campo == 'cnpj':
                    import re
                    numeros = re.sub(r'\D', '', valor)
                    if len(numeros) != 14:
                        campos_faltantes.append(f"{self.formatar_label(campo)} - CNPJ incompleto")

        return campos_faltantes

    def atualizar_botao_automacao(self):
        """Atualiza estado do botào"""
        campos_faltantes = self.get_campos_faltantes()

        if campos_faltantes:
            self.btn_automacao.setEnabled(False)
            self.status_label.setText(f" {len(campos_faltantes)} campo(s) precisam de atençãoo")
            self.status_label.setStyleSheet("""
                QLabel {
                    color: #e74c3c;
                    font-weight: 500;
                    border: none;
                    background-color: transparent;
                }
            """)
        else:
            self.btn_automacao.setEnabled(True)
            self.status_label.setText("Todos os campos obrigatórios preenchidos")
            self.status_label.setStyleSheet("""
                QLabel {
                    color: #27ae60;
                    font-weight: 500;
                    border: none;
                    background-color: transparent;
                }
            """)

    def iniciar_automacao_sap(self):
        """Inicia automação SAP com padronizaçãoo prévia e anexos"""
        campos_faltantes = self.get_campos_faltantes()

        if campos_faltantes:
            QMessageBox.warning(
                self,
                "Campos Obrigatórios",
                "Complete todos os campos obrigatórios antes de continuar."
            )
            return

        resposta = QMessageBox.question(
            self,
            "Confirmar Automação",
            "Deseja prosseguir para anexaçãoo de documentos?\n\nOs dados serào padronizados automaticamente.\n\nCertifique-se de que todos os dados estào corretos.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if resposta == QMessageBox.Yes:
            # Padronizar dados antes de continuar
            try:
                print("\n" + "="*60)
                print("PADRONIZANDO DADOS ANTES DA AUTOMAà‡àƒO SAP")
                print("="*60)

                padronizador = PadronizadorDados(self.json_path)

                sucesso_pad, mensagem_pad = padronizador.executar()

                if not sucesso_pad:
                    QMessageBox.critical(
                        self,
                        "Erro na Padronizaçãoo",
                        f"Erro ao padronizar dados:\n\n{mensagem_pad}\n\nA automação SAP não será iniciada."
                    )
                    return

                print(f"[INFO] {mensagem_pad}")

                # Recarrega dados padronizados
                self.data = self.load_json()
                self._atualizar_campos_na_tela()

                print("="*60)
                print("DADOS PADRONIZADOS COM SUCESSO")
                print("="*60 + "\n")

                # Mostra aviso se houver
                if padronizador.avisos:
                    avisos_texto = "\n".join([f"â€¢ {a}" for a in padronizador.avisos])
                    QMessageBox.information(
                        self,
                        "Padronizaçãoo Concluída",
                        f"Dados padronizados com sucesso!\n\nAvisos:\n{avisos_texto}\n\nProsseguindo para anexos..."
                    )

                # MODIFICADO: Fecha TelaInformacoes antes de abrir TelaAnexos
                self.tela_anexos = TelaAnexos(callback_continuar=self._executar_automacao_apos_anexos)
                self.tela_anexos.show()
                self.close()  # NOVA LINHA: Fecha a tela de informações

            except Exception as e:
                import traceback
                error_detail = traceback.format_exc()
                print(f"[ERRO] Falha na padronizaçãoo:\n{error_detail}")

                QMessageBox.critical(
                    self,
                    "Erro na Padronizaçãoo",
                    f"Erro inesperado ao padronizar dados:\n\n{str(e)}\n\nA automação SAP não será iniciada."
                )

    def _executar_automacao_apos_anexos(self):
        """Executa automação SAP após anexos validados"""
        # Fecha tela de anexos
        if hasattr(self, 'tela_anexos'):
            self.tela_anexos.close()

        # Continua com automação SAP
        if self.callback_automacao:
            self.callback_automacao()
        else:
            try:
                from SAP.AutomacaoSAP import executar_automacao

                sucesso, mensagem = executar_automacao()

                if sucesso:
                    QMessageBox.information(self, "Sucesso", mensagem)
                else:
                    QMessageBox.critical(self, "Erro", mensagem)

            except ImportError as e:
                QMessageBox.warning(
                    self,
                    "Módulo não encontrado",
                    f"O módulo de automação SAP não foi encontrado.\n\nErro: {str(e)}"
                )
            except Exception as e:
                QMessageBox.critical(
                    self,
                    "Erro",
                    f"Erro ao executar automação:\n\n{str(e)}"
                )

    def _atualizar_campos_na_tela(self):
        """Atualiza os campos na tela com os dados padronizados"""
        for campo_id, widget in self.campos_widgets.items():
            categoria, chave = campo_id.split('.')

            if categoria in self.data and chave in self.data[categoria]:
                novo_valor = self.data[categoria][chave]
                widget.setText(str(novo_valor) if novo_valor else "")

    def _formatar_titulo_categoria(self, categoria):
        """Formata título da categoria"""
        titulos = {
            'empresa': 'Informações da Empresa',
            'endereco': 'Endereço',
            'contato': 'Contatos',
            'bancario': 'Dados Bancários',
            'geral': 'Informações Gerais'
        }
        return titulos.get(categoria, categoria.capitalize())

    @staticmethod
    def formatar_label(texto):
        """Formata label do campo"""
        formatacao = {
            'razao_social': 'Razão Social',
            'nome_fantasia': 'Nome Fantasia',
            'cnpj': 'CNPJ',
            'inscricao_estadual': 'Inscrição Estadual',
            'inscricao_municipal': 'Inscrição Municipal',
            'cep': 'CEP',
            'celular': 'Celular Principal',
            'celular_secundario': 'Celular Secundário',
            'email_comercial': 'E-mail Comercial',
            'email_fiscal': 'E-mail Fiscal',
            'codigo_banco': 'Código do Banco',
            'agencia': 'Agência',
            'conta_corrente': 'Conta Corrente',
            'prazo_pagamento': 'Prazo de Pagamento',
            'modalidade_frete': 'Modalidade de Frete (CIF/FOB)'
        }
        return formatacao.get(texto, texto.replace("_", " ").capitalize())

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