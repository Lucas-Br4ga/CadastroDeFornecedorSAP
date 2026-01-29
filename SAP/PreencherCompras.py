# -*- coding: utf-8 -*-
"""
Módulo de Preenchimento de Compras - VERSÃO PORTÁVEL.

RESPONSABILIDADE:
- Adicionar papel FLVN01 (Compras)
- Preencher organização 0009 com espera robusta
- Preencher todos os campos de compras
- IMPORTANTE: NÃO salva - salvamento é responsabilidade do AutomacaoSAP.py

PORTABILIDADE:
- Esperas ativas após operações críticas
- Pausas estratégicas após ENTER
- Validação de elementos
- IDs completos SAP GUI
- Tratamento robusto de erros
"""

import time
import pythoncom
from typing import Dict


class PreencherComprasError(Exception):
    """Erro específico do módulo de compras"""
    pass


class PreencherCompras:
    """
    Classe para cadastrar papel de Compras (FLVN01) - PORTÁVEL.
    NÃO realiza salvamento - apenas preenchimento.
    """
    
    def __init__(self, session, manipulador_campos, dados_fornecedor: Dict):
        """
        Inicializa o módulo.
        
        Args:
            session: Sessão SAP ativa
            manipulador_campos: Manipulador de campos
            dados_fornecedor: Dados do fornecedor
        """
        self.session = session
        self.campos = manipulador_campos
        from .ManipuladorCampos import GerenciadorPopups
        self.popups = GerenciadorPopups(session)
        self.dados = dados_fornecedor
    
    # ========================================================================
    # ESPERAS ATIVAS (PORTÁVEIS)
    # ========================================================================
    
    def _wait_sap_ready(self, timeout: float = 10.0) -> bool:
        """
        Aguarda SAP ficar pronto (não ocupado).
        
        PORTÁVEL: Verifica session.Busy ao invés de espera fixa.
        
        Args:
            timeout: Tempo máximo de espera em segundos
            
        Returns:
            True se SAP ficou pronto, False se timeout
        """
        end_time = time.time() + timeout
        
        while time.time() < end_time:
            try:
                if not self.session.Busy:
                    return True
            except Exception:
                pass
            time.sleep(0.02)  # Polling agressivo
        
        return False
    
    def _wait_for_element(
        self, 
        element_id: str, 
        timeout: float = 10.0,
        element_name: str = ""
    ) -> bool:
        """
        Aguarda elemento existir (PORTÁVEL).
        
        Args:
            element_id: ID completo do elemento
            timeout: Tempo máximo de espera
            element_name: Nome do elemento (para logs)
            
        Returns:
            True se elemento apareceu, False se timeout
        """
        end_time = time.time() + timeout
        
        while time.time() < end_time:
            try:
                self.session.findById(element_id)
                return True
            except pythoncom.com_error:
                time.sleep(0.05)
        
        print(f"[ERRO] Elemento '{element_name}' não apareceu após {timeout}s")
        print(f"       ID: {element_id}")
        return False
    
    def _validar_campo_preenchido(self, element_id: str, valor_esperado: str) -> bool:
        """
        Valida se campo foi preenchido corretamente (PORTÁVEL).
        
        Args:
            element_id: ID do elemento
            valor_esperado: Valor que deveria estar no campo
            
        Returns:
            True se campo está com valor correto
        """
        try:
            campo = self.session.findById(element_id)
            valor_atual = str(campo.text).strip()
            return valor_esperado in valor_atual or valor_atual == valor_esperado
        except Exception:
            return False
    
    # ========================================================================
    # ADICIONAR PAPEL FLVN01 (PORTÁVEL)
    # ========================================================================
    
    def _adicionar_papel(self) -> bool:
        """
        Pressiona botão "Adicionar papel" (PORTÁVEL).
        
        Returns:
            True se pressionou com sucesso
        """
        print("[1/6] Adicionando papel...")
        
        try:
            botao_id = "wnd[0]/tbar[1]/btn[26]"
            
            # VALIDAÇÃO: Verificar se botão existe
            if not self._wait_for_element(
                botao_id, 
                timeout=5.0, 
                element_name="Botão adicionar papel"
            ):
                raise PreencherComprasError(
                    "Botão 'Adicionar papel' não encontrado"
                )
            
            # Pressionar botão
            botao = self.session.findById(botao_id)
            botao.press()
            
            print("[INFO] Botão pressionado")
            
            # ESPERA ATIVA: Aguardar processamento
            if not self._wait_sap_ready(timeout=5.0):
                print("[AVISO] SAP ainda processando após 5s")
            
            # PAUSA EXTRA: Garantir estabilização
            time.sleep(0.5)
            
            print("[OK] Papel adicionado")
            return True
            
        except PreencherComprasError:
            raise
        except Exception as e:
            raise PreencherComprasError(
                f"Erro ao adicionar papel: {e}"
            )
    
    # ========================================================================
    # SELECIONAR FLVN01 (PORTÁVEL)
    # ========================================================================
    
    def _selecionar_flvn01(self) -> bool:
        """
        Seleciona FLVN01 no combo de papel (PORTÁVEL).
        
        Returns:
            True se selecionou com sucesso
        """
        print("[2/6] Selecionando FLVN01...")
        
        try:
            combo_id = (
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/"
                "subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/"
                "subSCREEN_1100_ROLE_AND_TIME_AREA:SAPLBUPA_DIALOG_JOEL:1110/"
                "cmbBUS_JOEL_MAIN-PARTNER_ROLE"
            )
            
            # VALIDAÇÃO: Verificar se combo existe
            if not self._wait_for_element(
                combo_id, 
                timeout=5.0, 
                element_name="Combo papel"
            ):
                raise PreencherComprasError(
                    "Combo de papel não encontrado"
                )
            
            # Selecionar FLVN01
            combo = self.session.findById(combo_id)
            combo.setFocus()
            combo.key = "FLVN01"
            
            # VALIDAÇÃO: Verificar se selecionou
            if combo.key != "FLVN01":
                raise PreencherComprasError(
                    f"FLVN01 não foi selecionado. Valor atual: {combo.key}"
                )
            
            print("[OK] FLVN01 selecionado")
            
            # ESPERA ATIVA: Aguardar processamento
            if not self._wait_sap_ready(timeout=5.0):
                print("[AVISO] SAP ainda processando após 5s")
            
            # PAUSA EXTRA: Garantir estabilização
            time.sleep(0.5)
            
            return True
            
        except PreencherComprasError:
            raise
        except Exception as e:
            raise PreencherComprasError(
                f"Erro ao selecionar FLVN01: {e}"
            )
    
    # ========================================================================
    # CONFIRMAR POPUP (SE APARECER)
    # ========================================================================
    
    def _confirmar_popup_se_aparecer(self) -> None:
        """
        Confirma popup se aparecer após selecionar FLVN01 (PORTÁVEL).
        """
        print("[3/6] Verificando popup...")
        
        if self.popups.existe_popup(timeout=2):
            print("[INFO] Popup detectado, confirmando...")
            
            try:
                # Tenta botão específico primeiro
                self.session.findById("wnd[1]/usr/btnBUTTON_1").press()
                print("[OK] Popup confirmado (botão específico)")
            except Exception:
                # Fallback: botão padrão
                try:
                    self.popups.confirmar_popup()
                    print("[OK] Popup confirmado (botão padrão)")
                except Exception as e:
                    print(f"[AVISO] Erro ao confirmar popup: {e}")
            
            # ESPERA ATIVA: Aguardar processamento
            self._wait_sap_ready(timeout=3.0)
            
            # PAUSA EXTRA: Garantir fechamento
            time.sleep(0.3)
        else:
            print("[INFO] Nenhum popup detectado")
    
    # ========================================================================
    # PREENCHER ORGANIZAÇÃO 0009 (CRÍTICO - COM ESPERAS ROBUSTAS)
    # ========================================================================
    
    def _preencher_organizacao(self) -> bool:
        """
        Preenche organização 0009 com esperas ROBUSTAS (PORTÁVEL).
        
        CRÍTICO: Esta é a operação mais importante do módulo.
        O SAP precisa processar a organização antes de prosseguir.
        
        Returns:
            True se preencheu e processou com sucesso
        """
        print("[4/6] Preenchendo Organização 0009...")
        
        try:
            org_id = (
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/"
                "subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/"
                "ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/"
                "subSCREEN_1100_SUB_HEADER_AREA:SAPLCVI_FS_UI_VENDOR_PORG:0070/"
                "ctxtGV_PURCHASING_ORG"
            )
            
            # VALIDAÇÃO: Verificar se campo existe
            if not self._wait_for_element(
                org_id, 
                timeout=8.0, 
                element_name="Campo organização"
            ):
                raise PreencherComprasError(
                    "Campo de organização não encontrado após 8s"
                )
            
            # PASSO 1: Preencher campo
            campo_org = self.session.findById(org_id)
            campo_org.text = "0009"
            campo_org.setFocus()
            campo_org.caretPosition = 4
            
            print("[INFO] Campo preenchido: 0009")
            
            # PASSO 2: Pressionar ENTER
            self._wait_sap_ready(timeout=2.0)
            self.session.findById("wnd[0]").sendVKey(0)
            
            print("[INFO] ENTER pressionado, aguardando processamento...")
            
            # PASSO 3: ESPERA ATIVA - Aguardar SAP processar
            if not self._wait_sap_ready(timeout=10.0):
                print("[AVISO] SAP ainda processando após 10s, continuando...")
            
            # PASSO 4: PAUSA ESTRATÉGICA ⏱️
            # CRÍTICO: O SAP pode marcar Busy=False antes de processar completamente
            # Esta pausa garante que a organização foi realmente processada
            print("[INFO] ⏱️ Aguardando estabilização da organização...")
            time.sleep(1.2)
            
            # PASSO 5: VALIDAÇÃO - Verificar se organizou processou
            if not self._validar_campo_preenchido(org_id, "0009"):
                print("[AVISO] Organização pode não ter sido processada corretamente")
                print("[INFO] Aguardando mais tempo...")
                time.sleep(1.0)
                
                # Segunda tentativa de validação
                if not self._validar_campo_preenchido(org_id, "0009"):
                    raise PreencherComprasError(
                        "Organização não foi processada mesmo após esperas"
                    )
            
            print("[OK] ✅ Organização 0009 processada e validada")
            return True
            
        except PreencherComprasError:
            raise
        except Exception as e:
            raise PreencherComprasError(
                f"Erro ao preencher organização: {e}"
            )
    
    # ========================================================================
    # PREENCHER CAMPOS (PORTÁVEL)
    # ========================================================================
    
    def _preencher_campos(self) -> bool:
        """
        Preenche todos os campos de compras (PORTÁVEL).
        
        Returns:
            True se preencheu com sucesso
        """
        print("[5/6] Preenchendo campos...")
        
        try:
            base = (
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/"
                "subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/"
                "ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/"
                "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/"
                "ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7032"
            )
            
            # ============================================================
            # CHECKBOXES (VALIDADOS)
            # ============================================================
            print("   [5.1] Marcando checkboxes...")
            
            checkboxes = [
                {
                    'nome': 'WEBRE (Revisão fatura EM)',
                    'id': f"{base}/subA04P03:SAPLCVI_FS_UI_VENDOR_PORG:0074/chkGS_LFM1-WEBRE",
                    'obrigatorio': True
                },
                {
                    'nome': 'LEBRE (Revisão fatura serviços)',
                    'id': f"{base}/subA05P01:SAPLCVI_FS_UI_VENDOR_ENH:0048/chkGS_LFM1-LEBRE",
                    'obrigatorio': True
                },
                {
                    'nome': 'KZAUT (Geração automática OC)',
                    'id': f"{base}/subA07P03:SAPLCVI_FS_UI_VENDOR_ENH:0027/chkGS_LFM1-KZAUT",
                    'obrigatorio': False
                }
            ]
            
            for cb in checkboxes:
                try:
                    # Aguardar elemento existir
                    if not self._wait_for_element(cb['id'], timeout=3.0, element_name=cb['nome']):
                        if cb['obrigatorio']:
                            raise PreencherComprasError(
                                f"Checkbox obrigatório não encontrado: {cb['nome']}"
                            )
                        print(f"   [AVISO] {cb['nome']} não encontrado")
                        continue
                    
                    # Marcar checkbox
                    campo = self.session.findById(cb['id'])
                    campo.selected = True
                    
                    # VALIDAÇÃO
                    if not campo.selected:
                        if cb['obrigatorio']:
                            raise PreencherComprasError(
                                f"Checkbox obrigatório não marcou: {cb['nome']}"
                            )
                        print(f"   [AVISO] {cb['nome']} não marcou")
                    else:
                        print(f"   [OK] ✅ {cb['nome']}")
                    
                except PreencherComprasError:
                    raise
                except Exception as e:
                    if cb['obrigatorio']:
                        raise PreencherComprasError(
                            f"Erro ao marcar checkbox obrigatório {cb['nome']}: {e}"
                        )
                    print(f"   [AVISO] {cb['nome']}: {e}")
            
            # ============================================================
            # CAMPOS DE TEXTO (VALIDADOS)
            # ============================================================
            print("   [5.2] Preenchendo campos de texto...")
            
            prazo = self.dados['geral'].get('prazo_pagamento', 'BRFG')
            frete = self.dados['geral'].get('modalidade_frete', 'CIF')
            
            campos = [
                {
                    'nome': 'Moeda (BRL)',
                    'id': f"{base}/subA02P01:SAPLCVI_FS_UI_VENDOR_PORG:0076/ctxtGS_LFM1-WAERS",
                    'valor': 'BRL',
                    'obrigatorio': True
                },
                {
                    'nome': f'Condições pagamento ({prazo})',
                    'id': f"{base}/subA02P02:SAPLCVI_FS_UI_VENDOR_PORG:0086/ctxtGS_LFM1-ZTERM",
                    'valor': prazo,
                    'obrigatorio': True
                },
                {
                    'nome': f'Incoterms ({frete})',
                    'id': f"{base}/subA02P03:SAPLCVI_FS_UI_VENDOR_PORG:0085/ctxtGS_LFM1-INCO1",
                    'valor': frete,
                    'obrigatorio': True
                },
                {
                    'nome': 'Local incoterms (FABRICA)',
                    'id': f"{base}/subA02P03:SAPLCVI_FS_UI_VENDOR_PORG:0085/ctxtGS_LFM1-INCO2_L",
                    'valor': 'FABRICA',
                    'obrigatorio': True
                },
                {
                    'nome': 'Controle preço (1)',
                    'id': f"{base}/subA07P04:SAPLCVI_FS_UI_VENDOR_ENH:0028/ctxtGS_LFM1-MEPRF",
                    'valor': '1',
                    'obrigatorio': True
                },
                {
                    'nome': 'Controle confirmação (Z004)',
                    'id': f"{base}/subA07P07:SAPLCVI_FS_UI_VENDOR_ENH:0010/ctxtGS_LFM1-BSTAE",
                    'valor': 'Z004',
                    'obrigatorio': True
                }
            ]
            
            ultimo_campo_id = None
            
            for campo in campos:
                try:
                    # Aguardar elemento existir
                    if not self._wait_for_element(
                        campo['id'], 
                        timeout=3.0, 
                        element_name=campo['nome']
                    ):
                        if campo['obrigatorio']:
                            raise PreencherComprasError(
                                f"Campo obrigatório não encontrado: {campo['nome']}"
                            )
                        print(f"   [AVISO] {campo['nome']} não encontrado")
                        continue
                    
                    # Preencher campo
                    elem = self.session.findById(campo['id'])
                    elem.text = campo['valor']
                    
                    # Guardar último campo para dar foco depois
                    ultimo_campo_id = campo['id']
                    
                    print(f"   [OK] {campo['nome']}")
                    
                except PreencherComprasError:
                    raise
                except Exception as e:
                    if campo['obrigatorio']:
                        raise PreencherComprasError(
                            f"Erro ao preencher campo obrigatório {campo['nome']}: {e}"
                        )
                    print(f"   [AVISO] {campo['nome']}: {e}")
            
            # ============================================================
            # FINALIZAR - DAR FOCO NO ÚLTIMO CAMPO E PRESSIONAR ENTER
            # ============================================================
            print("   [5.3] Finalizando preenchimento...")
            
            if ultimo_campo_id:
                try:
                    ultimo = self.session.findById(ultimo_campo_id)
                    ultimo.setFocus()
                    ultimo.caretPosition = len(str(ultimo.text))
                    
                    # ESPERA ATIVA antes de ENTER
                    self._wait_sap_ready(timeout=2.0)
                    
                    # Pressionar ENTER
                    self.session.findById("wnd[0]").sendVKey(0)
                    
                    print("[INFO] ENTER final pressionado")
                    
                    # ESPERA ATIVA após ENTER
                    if not self._wait_sap_ready(timeout=5.0):
                        print("[AVISO] SAP ainda processando após 5s")
                    
                    # PAUSA EXTRA
                    time.sleep(0.5)
                    
                except Exception as e:
                    print(f"[AVISO] Erro ao finalizar: {e}")
            
            print("[OK] ✅ Campos preenchidos")
            return True
            
        except PreencherComprasError:
            raise
        except Exception as e:
            raise PreencherComprasError(
                f"Erro ao preencher campos: {e}"
            )
    
    # ========================================================================
    # MÉTODO PRINCIPAL
    # ========================================================================
    
    def adicionar_papel_compras(self) -> bool:
        """
        Adiciona papel FLVN01 (Compras) - PORTÁVEL.
        NÃO salva - apenas preenche.
        
        Returns:
            True se cadastrou com sucesso
        """
        print("\n" + "="*70)
        print("CADASTRANDO COMPRAS (FLVN01) - PORTÁVEL")
        print("="*70)
        
        try:
            # ETAPA 1: Adicionar papel
            if not self._adicionar_papel():
                raise PreencherComprasError("Falha ao adicionar papel")
            
            # ETAPA 2: Selecionar FLVN01
            if not self._selecionar_flvn01():
                raise PreencherComprasError("Falha ao selecionar FLVN01")
            
            # ETAPA 3: Confirmar popup (se aparecer)
            self._confirmar_popup_se_aparecer()
            
            # ETAPA 4: Preencher organização (CRÍTICO)
            if not self._preencher_organizacao():
                raise PreencherComprasError("Falha ao preencher organização")
            
            # ETAPA 5: Preencher campos
            if not self._preencher_campos():
                raise PreencherComprasError("Falha ao preencher campos")
            
            print("\n[OK] ✅✅✅ Compras cadastrado (aguardando salvamento)")
            print("="*70 + "\n")
            return True
            
        except PreencherComprasError as e:
            print(f"\n[ERRO] ❌ Falha ao cadastrar Compras:")
            print(f"       {str(e)}")
            print("="*70 + "\n")
            raise
        except Exception as e:
            print(f"\n[ERRO] ❌ Erro inesperado:")
            print(f"       {str(e)}")
            print("="*70 + "\n")
            raise PreencherComprasError(
                f"Erro inesperado durante cadastro de Compras: {e}"
            )
    
    def executar(self) -> bool:
        """
        Executa cadastro de Compras (SEM SALVAMENTO) - PORTÁVEL.
        
        Returns:
            True se cadastrou com sucesso
        """
        print("\n" + "="*70)
        print("MÓDULO: COMPRAS (PORTÁVEL)")
        print("="*70)
        
        try:
            if not self.adicionar_papel_compras():
                return False
            
            print("\n[OK] ✅✅✅ Compras COMPLETO (aguardando salvamento)")
            print("="*70 + "\n")
            return True
            
        except Exception as e:
            print(f"\n[ERRO] {e}")
            import traceback
            traceback.print_exc()
            return False