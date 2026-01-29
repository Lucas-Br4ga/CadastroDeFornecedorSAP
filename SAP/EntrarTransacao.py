# -*- coding: utf-8 -*-
"""
Módulo de Entrada na Transação XK01 - VERSÃO PORTÁVEL.

RESPONSABILIDADE:
- Abrir transação XK01
- Selecionar fornecedor nacional
- Configurar grupo de criação
- SEMPRE aguardar carregamento completo antes de prosseguir

PORTABILIDADE:
- Esperas ativas (não time.sleep fixo)
- Validação de elementos antes de interagir
- IDs completos SAP GUI
- Tratamento robusto de erros
"""

import time
import pythoncom
from typing import Optional


class EntrarTransacaoError(Exception):
    """Erro específico do módulo de entrada na transação"""
    pass


class EntrarTransacao:
    """
    Classe responsável por entrar na transação XK01
    e realizar configurações iniciais (PORTÁVEL).
    """
    
    def __init__(self, session, manipulador_campos):
        """
        Inicializa o módulo.
        
        Args:
            session: Sessão SAP ativa
            manipulador_campos: Manipulador de campos SAP
        """
        self.session = session
        self.campos = manipulador_campos
        from .ManipuladorCampos import GerenciadorPopups
        self.popups = GerenciadorPopups(session)
    
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
                time.sleep(0.05)  # Polling moderado
        
        print(f"[ERRO] Elemento '{element_name}' não apareceu após {timeout}s")
        print(f"       ID: {element_id}")
        return False
    
    def _validar_tela_transacao(self) -> bool:
        """
        Valida se a tela da transação XK01 carregou completamente.
        
        PORTÁVEL: Verifica múltiplos elementos críticos.
        
        Returns:
            True se tela está pronta
        """
        elementos_criticos = [
            {
                'id': 'wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036',
                'nome': 'Área principal'
            },
            {
                'id': 'wnd[0]/tbar[1]/btn[26]',
                'nome': 'Botão adicionar papel'
            }
        ]
        
        for elemento in elementos_criticos:
            if not self._wait_for_element(
                elemento['id'], 
                timeout=5.0, 
                element_name=elemento['nome']
            ):
                return False
        
        return True
    
    # ========================================================================
    # ABERTURA DA TRANSAÇÃO (PORTÁVEL)
    # ========================================================================
    
    def abrir_transacao_xk01(self) -> bool:
        """
        Abre a transação XK01 com validação robusta.
        
        FLUXO PORTÁVEL:
        1. Limpar campo de transação
        2. Digitar /nxk01
        3. Pressionar ENTER
        4. Aguardar SAP processar (espera ativa)
        5. Validar que tela carregou
        
        Returns:
            True se abriu com sucesso
            
        Raises:
            EntrarTransacaoError: Se não conseguir abrir
        """
        print("\n[ETAPA] Abrindo transação XK01...")
        
        try:
            # PASSO 1: Buscar campo de transação (PORTÁVEL)
            campo_transacao_id = "wnd[0]/tbar[0]/okcd"
            
            try:
                okcd = self.session.findById(campo_transacao_id)
            except Exception as e:
                raise EntrarTransacaoError(
                    f"Campo de transação não encontrado: {e}"
                )
            
            # PASSO 2: Limpar campo (SEGURANÇA)
            okcd.text = ""
            time.sleep(0.1)  # Pequena pausa para garantir limpeza
            
            # PASSO 3: Digitar transação com /n (PORTÁVEL)
            okcd.text = "/nxk01"
            okcd.setFocus()
            
            print("[INFO] Transação digitada: /nxk01")
            
            # PASSO 4: Pressionar ENTER
            self.session.findById("wnd[0]").sendVKey(0)
            
            print("[INFO] ENTER pressionado, aguardando carregamento...")
            
            # PASSO 5: ESPERA ATIVA - Aguardar SAP processar
            if not self._wait_sap_ready(timeout=10.0):
                raise EntrarTransacaoError(
                    "SAP não ficou pronto após 10s ao abrir XK01"
                )
            
            # PAUSA EXTRA: Garantir estabilização da tela
            # (necessário porque o SAP pode marcar Busy=False antes da tela renderizar)
            time.sleep(0.8)
            
            print("[INFO] SAP processou transação")
            
            # PASSO 6: VALIDAÇÃO ROBUSTA - Verificar se tela carregou
            print("[INFO] Validando elementos da tela XK01...")
            
            if not self._validar_tela_transacao():
                raise EntrarTransacaoError(
                    "Tela XK01 não carregou completamente - elementos críticos ausentes"
                )
            
            print("[OK] ✅ Transação XK01 aberta e validada")
            return True
            
        except EntrarTransacaoError:
            raise
        except Exception as e:
            raise EntrarTransacaoError(
                f"Erro inesperado ao abrir transação XK01: {e}"
            )
    
    # ========================================================================
    # SELEÇÃO DE FORNECEDOR NACIONAL (PORTÁVEL)
    # ========================================================================
    
    def selecionar_fornecedor_nacional(self) -> bool:
        """
        Seleciona fornecedor nacional no popup.
        
        FLUXO PORTÁVEL:
        1. Aguardar popup aparecer (wnd[1])
        2. Selecionar radio button
        3. Confirmar popup
        4. Aguardar processamento
        5. Validar que popup fechou
        
        Returns:
            True se selecionou com sucesso
            
        Raises:
            EntrarTransacaoError: Se não conseguir selecionar
        """
        print("\n[ETAPA] Selecionando fornecedor nacional...")
        
        try:
            # PASSO 1: ESPERA ATIVA - Aguardar popup aparecer
            print("[INFO] Aguardando popup de seleção...")
            
            if not self.popups.existe_popup(timeout=8):
                raise EntrarTransacaoError(
                    "Popup de seleção de tipo não apareceu após 8s"
                )
            
            print("[OK] Popup detectado")
            
            # PASSO 2: Selecionar radio "Fornecedor nacional" (PORTÁVEL)
            radio_id = (
                "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/"
                "sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]"
            )
            
            try:
                radio = self.session.findById(radio_id)
                radio.select()
                radio.setFocus()
                
                print("[INFO] Radio 'Fornecedor nacional' selecionado")
                
            except Exception as e:
                raise EntrarTransacaoError(
                    f"Não foi possível selecionar fornecedor nacional: {e}"
                )
            
            # PAUSA MÍNIMA: Garantir seleção processada
            time.sleep(0.2)
            
            # PASSO 3: Confirmar popup (PORTÁVEL)
            print("[INFO] Confirmando seleção...")
            
            if not self.popups.confirmar_popup():
                raise EntrarTransacaoError(
                    "Não foi possível confirmar popup"
                )
            
            # PASSO 4: ESPERA ATIVA - Aguardar processamento
            print("[INFO] Aguardando SAP processar seleção...")
            
            if not self._wait_sap_ready(timeout=10.0):
                print("[AVISO] SAP ainda processando após 10s, continuando...")
            
            # PAUSA EXTRA: Garantir estabilização
            time.sleep(0.8)
            
            # PASSO 5: VALIDAÇÃO - Verificar que popup fechou
            if self.popups.existe_popup(timeout=1):
                print("[AVISO] Popup ainda aberto, tentando fechar...")
                self.popups.fechar_popup_esc()
            
            print("[OK] ✅ Fornecedor nacional selecionado")
            return True
            
        except EntrarTransacaoError:
            raise
        except Exception as e:
            raise EntrarTransacaoError(
                f"Erro ao selecionar fornecedor nacional: {e}"
            )
    
    # ========================================================================
    # CONFIGURAÇÃO DO GRUPO DE CRIAÇÃO (PORTÁVEL)
    # ========================================================================
    
    def configurar_grupo_criacao(self) -> bool:
        """
        Configura o grupo de criação como ZVRE.
        
        FLUXO PORTÁVEL:
        1. Aguardar elemento estar disponível
        2. Selecionar valor no combo
        3. Validar seleção
        
        Returns:
            True se configurou com sucesso
        """
        print("\n[ETAPA] Configurando grupo de criação...")
        
        try:
            # PASSO 1: ESPERA ATIVA - Aguardar combo estar disponível
            combo_id = (
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/"
                "subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                "ssubSCREEN_1000_HEADER_AREA:SAPLBUPA_DIALOG_JOEL:1500/"
                "cmbBUS_JOEL_MAIN-CREATION_GROUP"
            )
            
            if not self._wait_for_element(
                combo_id, 
                timeout=5.0, 
                element_name="Combo grupo de criação"
            ):
                print("[AVISO] Combo grupo de criação não encontrado")
                return False
            
            # PASSO 2: Selecionar ZVRE (PORTÁVEL)
            try:
                combo = self.session.findById(combo_id)
                combo.key = "ZVRE"
                
                # VALIDAÇÃO: Verificar se selecionou
                if combo.key == "ZVRE":
                    print("[OK] ✅ Grupo de criação configurado: ZVRE")
                    return True
                else:
                    print(f"[AVISO] Valor diferente do esperado: {combo.key}")
                    return True  # Não é crítico
                    
            except Exception as e:
                print(f"[AVISO] Erro ao configurar grupo: {e}")
                return False
            
        except Exception as e:
            print(f"[AVISO] Erro ao configurar grupo de criação: {e}")
            return False
    
    # ========================================================================
    # MÉTODO PRINCIPAL DE EXECUÇÃO
    # ========================================================================
    
    def executar(self) -> bool:
        """
        Executa todas as etapas de entrada na transação (PORTÁVEL).
        
        FLUXO:
        1. Abrir transação XK01 (com espera ativa)
        2. Selecionar fornecedor nacional (com espera ativa)
        3. Configurar grupo de criação (com espera ativa)
        
        Returns:
            True se todas as etapas foram bem-sucedidas
            
        Raises:
            EntrarTransacaoError: Se alguma etapa crítica falhar
        """
        print("\n" + "="*70)
        print("MÓDULO: ENTRADA NA TRANSAÇÃO (PORTÁVEL)")
        print("="*70)
        
        try:
            # ETAPA 1: Abrir transação (CRÍTICO)
            if not self.abrir_transacao_xk01():
                raise EntrarTransacaoError(
                    "Falha ao abrir transação XK01"
                )
            
            # ETAPA 2: Selecionar fornecedor nacional (CRÍTICO)
            if not self.selecionar_fornecedor_nacional():
                raise EntrarTransacaoError(
                    "Falha ao selecionar fornecedor nacional"
                )
            
            # ETAPA 3: Configurar grupo de criação (NÃO CRÍTICO)
            self.configurar_grupo_criacao()
            
            print("\n[OK] ✅✅✅ Entrada na transação concluída com sucesso!")
            print("="*70 + "\n")
            
            return True
            
        except EntrarTransacaoError as e:
            print(f"\n[ERRO] ❌ Falha na entrada da transação:")
            print(f"       {str(e)}")
            print("="*70 + "\n")
            raise
        except Exception as e:
            print(f"\n[ERRO] ❌ Erro inesperado:")
            print(f"       {str(e)}")
            print("="*70 + "\n")
            raise EntrarTransacaoError(
                f"Erro inesperado durante entrada na transação: {e}"
            )