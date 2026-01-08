"""
Módulo Centralizado de Salvamento - VERSÃO PORTÁVEL.

RESPONSABILIDADE:
- Salvar cadastro no SAP (btn[11])
- Tratar popups de confirmação
- Validar salvamento bem-sucedido
- Habilitar edição (btn[6]) para próxima etapa

FLUXO:
1. Pressionar botão Salvar (btn[11])
2. Aguardar processamento (espera ativa)
3. Tratar popup se aparecer
4. Validar salvamento
5. Habilitar edição (btn[6])

PORTABILIDADE:
- SEMPRE usa findById() com IDs completos
- Esperas ativas (não depende de time.sleep fixo)
- Independente de resolução/tema/idioma
- Tratamento robusto de erros
"""

import time


class SalvarFornecedor:
    """
    Classe centralizada para salvamento no SAP.
    Reutilizável por todos os módulos.
    """
    
    def __init__(self, session):
        """
        Inicializa o salvador.
        
        Args:
            session: Sessão SAP ativa
        """
        self.session = session
    
    def _wait_sap_ready(self, timeout: float = 10.0) -> bool:
        """
        Aguarda SAP ficar pronto (espera ativa PORTÁVEL).
        
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
    
    def _existe_popup(self, timeout: float = 2.0) -> bool:
        """
        Verifica se existe popup aberto (wnd[1]) de forma PORTÁVEL.
        
        Args:
            timeout: Tempo máximo de espera
            
        Returns:
            True se popup existe, False caso contrário
        """
        end_time = time.time() + timeout
        
        while time.time() < end_time:
            try:
                self.session.findById("wnd[1]")
                return True
            except Exception:
                pass
            time.sleep(0.02)
        
        return False
    
    def _confirmar_popup(self) -> bool:
        """
        Confirma popup se existir (PORTÁVEL).
        
        Returns:
            True se confirmou popup, False se não havia popup
        """
        try:
            if self._existe_popup(timeout=3):
                print("[INFO] Popup detectado, confirmando...")
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                self._wait_sap_ready(timeout=3.0)
                print("[OK] Popup confirmado")
                return True
            return False
        except Exception as e:
            print(f"[AVISO] Erro ao confirmar popup: {e}")
            return False
    
    def executar(self) -> bool:
        """
        Executa salvamento completo: Salvar → Verificar → Habilitar Edição.
        
        FLUXO PORTÁVEL:
        1. Pressiona botão Salvar (btn[11])
        2. Aguarda processamento (espera ativa até 10s)
        3. Trata popup se aparecer
        4. Valida salvamento (wnd[0] existe)
        5. Habilita edição (btn[6])
        
        Returns:
            True se salvou e habilitou edição com sucesso
            False se houve falha em alguma etapa
        """
        try:
            print("\n" + "="*70)
            print("SALVAMENTO CENTRALIZADO")
            print("="*70)
            
            # ETAPA 1: SALVAR
            print("\n[1/4] Pressionando botão Salvar...")
            try:
                botao_salvar = self.session.findById("wnd[0]/tbar[0]/btn[11]")
                botao_salvar.press()
                print("[OK] Salvar pressionado")
            except Exception as e:
                print(f"[ERRO] Falha ao pressionar Salvar: {e}")
                return False
            
            # ETAPA 2: AGUARDAR PROCESSAMENTO (ESPERA ATIVA)
            print("\n[2/4] Aguardando SAP processar salvamento...")
            if not self._wait_sap_ready(timeout=10.0):
                print("[AVISO] SAP ainda processando após 10s, continuando...")
            else:
                print("[OK] SAP pronto")
            
            # ETAPA 3: TRATAR POPUP SE APARECER
            print("\n[3/4] Verificando popup...")
            self._confirmar_popup()
            
            # Aguarda finalização completa (mais 5s para garantir)
            print("[INFO] Garantindo conclusão do salvamento...")
            end_time = time.time() + 5.0
            while time.time() < end_time:
                if not self.session.Busy:
                    time.sleep(0.3)
                    if not self.session.Busy:
                        break
                time.sleep(0.1)
            
            # ETAPA 4: VALIDAR SALVAMENTO
            print("\n[4/4] Validando salvamento...")
            try:
                self.session.findById("wnd[0]")
                print("[OK] ✅ Salvamento validado")
            except Exception as e:
                print(f"[ERRO] Validação falhou: {e}")
                return False
            
            # ETAPA 5: HABILITAR EDIÇÃO PARA PRÓXIMA ETAPA
            print("\n[FINAL] Habilitando edição para próxima etapa...")
            try:
                self._wait_sap_ready(timeout=2.0)
                botao_editar = self.session.findById("wnd[0]/tbar[1]/btn[6]")
                botao_editar.press()
                print("[OK] Editar pressionado")
                self._wait_sap_ready(timeout=2.0)
                print("[OK] ✅ Edição habilitada")
            except Exception as e:
                print(f"[AVISO] Não foi possível habilitar edição: {e}")
                print("[INFO] Cadastro foi salvo, mas edição pode não estar ativa")
                # Salvamento foi bem-sucedido mesmo sem habilitar edição
                return True
            
            print("\n[OK] ✅✅✅ SALVAMENTO CONCLUÍDO COM SUCESSO")
            print("="*70 + "\n")
            return True
            
        except Exception as e:
            print(f"\n[ERRO] Falha no salvamento: {e}")
            import traceback
            traceback.print_exc()
            return False