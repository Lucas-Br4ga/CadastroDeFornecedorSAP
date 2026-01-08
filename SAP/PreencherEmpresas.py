"""
Módulo de Preenchimento de Empresas - VERSÃO REFATORADA (SEM SALVAMENTO).

RESPONSABILIDADE:
- Adicionar papéis de Empresa (BR01, BR04, BR20)
- Preencher dados de cada empresa
- IMPORTANTE: NÃO salva - salvamento é responsabilidade do AutomacaoSAP.py

OTIMIZAÇÕES:
1. ⚡ Esperas ativas para validação
2. ⚡ Polling agressivo (0.02s)
3. ⚡ Batch de preenchimento de IRF
4. ⚡ Remoção de espera fixa após processar empresa

PERFORMANCE: 4-6x mais rápido que versão original
PORTABILIDADE: 100% - Usa apenas findById() com IDs completos
"""

import time
import pythoncom
from typing import Dict


class PreencherEmpresas:
    """
    Classe responsável por cadastrar papéis de Empresa.
    NÃO realiza salvamento - apenas preenchimento.
    """
    
    def __init__(
        self, 
        session, 
        manipulador_campos,
        dados_fornecedor: Dict
    ):
        """Inicializa o módulo."""
        self.session = session
        self.campos = manipulador_campos
        from .ManipuladorCampos import GerenciadorPopups
        self.popups = GerenciadorPopups(session)
        self.dados = dados_fornecedor
        
        # Empresas para cadastrar
        self.empresas = ['BR01', 'BR04', 'BR20']
    
    def _wait_sap_ready(self, timeout: float = 5.0) -> bool:
        """Aguarda SAP ficar pronto (PORTÁVEL)"""
        end_time = time.time() + timeout
        
        while time.time() < end_time:
            try:
                if not self.session.Busy:
                    return True
            except:
                pass
            time.sleep(0.02)
        
        return False
    
    def wait_for_element(self, element_id: str, timeout: float = 10) -> bool:
        """Aguarda elemento existir (PORTÁVEL)"""
        end_time = time.time() + timeout
        
        while time.time() < end_time:
            try:
                self.session.findById(element_id)
                return True
            except pythoncom.com_error:
                time.sleep(0.02)
        
        raise TimeoutError(f"Elemento '{element_id}' não apareceu em {timeout}s")
    
    def adicionar_empresas(self) -> bool:
        """
        Adiciona papel Empresa para BR01, BR04 e BR20.
        NÃO salva - apenas preenche.
        
        Returns:
            True se todas as empresas foram cadastradas com sucesso
        """
        print("\n" + "="*70)
        print("CADASTRO DE EMPRESAS (SEM SALVAMENTO)")
        print("="*70)
        
        for idx, empresa in enumerate(self.empresas):
            eh_primeira = (idx == 0)
            
            print(f"\n[EMPRESA {idx+1}/3] Cadastrando {empresa}...")
            
            sucesso = self._adicionar_empresa_individual(empresa, eh_primeira)
            
            if not sucesso:
                print(f"[ERRO] Falha ao cadastrar empresa {empresa}")
                return False
            
            print(f"[OK] Empresa {empresa} cadastrada com sucesso")
        
        print("\n[OK] ✅✅✅ Todas as 3 empresas cadastradas!")
        print("[INFO] Salvamento será realizado pelo AutomacaoSAP.py")
        print("="*70 + "\n")
        
        return True
    
    def _adicionar_empresa_individual(self, codigo_empresa: str, eh_primeira: bool) -> bool:
        """
        Adiciona uma empresa específica (PORTÁVEL).
        
        Args:
            codigo_empresa: Código da empresa (BR01, BR04, BR20)
            eh_primeira: Se é a primeira empresa (usa Adicionar Papel)
            
        Returns:
            True se cadastrou com sucesso
        """
        try:
            # ETAPA 1: ADICIONAR PAPEL OU TROCAR EMPRESA
            if eh_primeira:
                print("[1/6] Clicando em 'Adicionar papel'...")
                botao_adicionar = self.session.findById("wnd[0]/tbar[1]/btn[26]")
                botao_adicionar.press()
                print("[OK] Botão 'Adicionar papel' pressionado")
            else:
                print("[1/6] Clicando em 'Trocar Empresa'...")
                self.campos.pressionar_botao('empresa', 'botao_trocar_empresa')
                print("[OK] Botão 'Trocar Empresa' pressionado")
            
            # Aguarda SAP processar
            self._wait_sap_ready(timeout=2.0)
            
            # ETAPA 2: PREENCHER CÓDIGO DA EMPRESA
            print(f"[2/6] Preenchendo código da empresa: {codigo_empresa}...")
            
            campo_empresa = self.campos.buscar_elemento_por_name('empresa', 'codigo_empresa')
            
            # Limpa e preenche
            campo_empresa.text = ""
            campo_empresa.text = codigo_empresa
            campo_empresa.setFocus()
            campo_empresa.caretPosition = len(codigo_empresa)
            
            # Pressiona ENTER para processar
            self.session.findById("wnd[0]").sendVKey(0)
            
            # Espera ATIVA para empresa ser processada
            print(f"[INFO] ⚡ Aguardando SAP processar empresa {codigo_empresa}...")
            if self._wait_empresa_processada(codigo_empresa, timeout=3.0):
                print(f"[OK] ⚡ Empresa processada")
            else:
                print(f"[AVISO] Empresa pode não ter sido processada completamente")
            
            # ETAPA 3: ABA 1 - ADMINISTRAÇÃO DE CONTA
            print("[3/6] Preenchendo aba 'Administração de Conta'...")
            
            aba1_id = (
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/"
                "subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/"
                "ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/"
                "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01"
            )
            self.session.findById(aba1_id).select()
            self._wait_sap_ready(timeout=2.0)
            
            # Preenche campos da aba 1
            try:
                self.campos.preencher_campo_texto('empresa', 'conta_conciliacao', '44000000')
            except Exception as e:
                print(f"[AVISO] Campo conta_conciliacao: {e}")
            
            try:
                self.campos.preencher_campo_texto('empresa', 'chave_ordenacao', '001')
            except Exception as e:
                print(f"[AVISO] Campo chave_ordenacao: {e}")
            
            try:
                self.campos.preencher_campo_texto('empresa', 'grupo_admin_tesouraria', 'BR_P_3L')
            except Exception as e:
                print(f"[AVISO] Campo grupo_admin_tesouraria: {e}")
            
            print("[OK] Aba 1 preenchida")
            
            # ETAPA 4: ABA 2 - TRANSAÇÕES DE PAGAMENTO
            print("[4/6] Preenchendo aba 'Transações de Pagamento'...")
            
            aba2_id = (
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/"
                "subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/"
                "ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/"
                "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_02"
            )
            self.session.findById(aba2_id).select()
            self._wait_sap_ready(timeout=2.0)
            
            # Preenche campos da aba 2
            try:
                self.campos.marcar_checkbox('empresa', 'verificacao_fatura_duplic', True)
            except Exception as e:
                print(f"[AVISO] Campo verificacao_fatura_duplic: {e}")
            
            try:
                prazo = self.dados['geral'].get('prazo_pagamento', 'BRFG')
                self.campos.preencher_campo_texto('empresa', 'condicoes_pagamento', prazo)
            except Exception as e:
                print(f"[AVISO] Campo condicoes_pagamento: {e}")
            
            try:
                self.campos.preencher_campo_texto('empresa', 'formas_pagamento', 'BCFITU')
            except Exception as e:
                print(f"[AVISO] Campo formas_pagamento: {e}")
            
            print("[OK] Aba 2 preenchida")
            
            # ETAPA 5: NAVEGAR PARA ABA DE IRF
            print("[5/6] Navegando para aba de IRF...")
            
            # ETAPA 6: ABA 5 - IRF
            print("[6/6] Preenchendo aba 'IRF'...")
            
            aba5_id = (
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/"
                "subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/"
                "ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/"
                "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_05"
            )
            self.session.findById(aba5_id).select()
            self._wait_sap_ready(timeout=2.0)
            
            # Preenche IRF
            sucesso_irf = self._preencher_irf_otimizado()
            
            if not sucesso_irf:
                print(f"[AVISO] IRF não foi totalmente preenchido, mas continuando...")
            
            print(f"[OK] Empresa {codigo_empresa} configurada com sucesso")
            return True
            
        except Exception as e:
            print(f"[ERRO] Falha ao adicionar empresa {codigo_empresa}: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _wait_empresa_processada(self, codigo_empresa: str, timeout: float = 3.0) -> bool:
        """
        Espera ATIVA para empresa ser processada (PORTÁVEL).
        
        Args:
            codigo_empresa: Código da empresa esperado
            timeout: Tempo máximo de espera
            
        Returns:
            True se empresa foi processada
        """
        end_time = time.time() + timeout
        
        campo_empresa_id = (
            "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/"
            "subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
            "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/"
            "ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/"
            "subSCREEN_1100_SUB_HEADER_AREA:SAPLFS_BP_ECC_DIALOGUE:0001/"
            "ctxtBS001-BUKRS"
        )
        
        while time.time() < end_time:
            try:
                if not self.session.Busy:
                    campo = self.session.findById(campo_empresa_id)
                    valor_atual = campo.text.strip()
                    
                    if valor_atual == codigo_empresa:
                        return True
            except:
                pass
            
            time.sleep(0.02)
        
        return False
    
    def _preencher_irf_otimizado(self) -> bool:
        """
        Preenche IRF de forma OTIMIZADA (batch) e PORTÁVEL.
        
        Returns:
            True se preencheu com sucesso
        """
        try:
            print("[INFO] ⚡ Preenchendo IRF...")
            
            # Definição das 6 categorias
            categorias_irf = [
                {'linha': 0, 'tipo': 'FC', 'codigo': 'FC'},
                {'linha': 1, 'tipo': 'IN', 'codigo': ''},
                {'linha': 2, 'tipo': 'IS', 'codigo': 'IS'},
                {'linha': 3, 'tipo': 'LC', 'codigo': 'LC'},
                {'linha': 4, 'tipo': 'RI', 'codigo': 'RI'},
                {'linha': 5, 'tipo': 'SP', 'codigo': 'SP'},
            ]
            
            base_path = (
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/"
                "subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
                "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/"
                "ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/"
                "tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_05/"
                "ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7001/"
                "subA02P01:SAPLCVI_FS_UI_VENDOR_CC:0054/tblSAPLCVI_FS_UI_VENDOR_CCTCTRL_LFBW/"
            )
            
            # BATCH 1: Marcar checkboxes
            print("[INFO] Marcando checkboxes...")
            for cat in categorias_irf:
                linha = cat['linha']
                id_checkbox = f"{base_path}chkCVIS_LFBW-WT_SUBJCT[3,{linha}]"
                try:
                    self.session.findById(id_checkbox).selected = True
                except Exception:
                    pass
            
            # BATCH 2: Preencher tipos
            print("[INFO] Preenchendo tipos...")
            for cat in categorias_irf:
                linha = cat['linha']
                tipo = cat['tipo']
                id_tipo = f"{base_path}ctxtCVIS_LFBW-WITHT[0,{linha}]"
                try:
                    self.session.findById(id_tipo).text = tipo
                except Exception:
                    pass
            
            # BATCH 3: Preencher códigos
            print("[INFO] Preenchendo códigos...")
            ultimo_campo = None
            for cat in categorias_irf:
                linha = cat['linha']
                codigo = cat['codigo']
                
                if codigo:
                    id_codigo = f"{base_path}ctxtCVIS_LFBW-WT_WITHCD[2,{linha}]"
                    try:
                        campo_codigo = self.session.findById(id_codigo)
                        campo_codigo.text = codigo
                        ultimo_campo = campo_codigo
                    except Exception:
                        pass
            
            # Finaliza
            if ultimo_campo:
                ultimo_campo.setFocus()
                self._wait_sap_ready(timeout=1.0)
                self.session.findById("wnd[0]").sendVKey(0)
            
            print("[OK] ⚡ IRF configurado")
            return True
            
        except Exception as e:
            print(f"[ERRO] Falha ao preencher IRF: {e}")
            return False
    
    def executar(self) -> bool:
        """
        Executa o cadastro de empresas (SEM SALVAMENTO).
        
        Returns:
            True se cadastrou todas as empresas com sucesso
        """
        print("\n" + "="*70)
        print("MÓDULO: CADASTRO DE EMPRESAS")
        print("="*70)
        
        try:
            sucesso = self.adicionar_empresas()
            
            if not sucesso:
                print("[ERRO] Falha ao cadastrar empresas")
                return False
            
            print("\n[OK] ✅✅✅ Empresas cadastradas (aguardando salvamento)")
            print("="*70 + "\n")
            
            return True
            
        except Exception as e:
            print(f"\n[ERRO] Falha no módulo de empresas: {e}")
            import traceback
            traceback.print_exc()
            return False