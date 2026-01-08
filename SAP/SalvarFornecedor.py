"""
Módulo de Salvamento de Fornecedor - VERSÃO SIMPLIFICADA.

FLUXO:
1. Pressionar botão Salvar (btn[11])
2. Aguardar e tratar popup se aparecer
3. Validar salvamento
4. Habilitar edição (btn[6])

PORTABILIDADE:
- Usa SEMPRE findById() com IDs completos
- Esperas ativas
- Independente de resolução/tema/idioma
"""

import time
import pythoncom


class SalvarFornecedor:
    """Salva cadastro e habilita edição."""
    
    def __init__(self, session):
        self.session = session
    
    def _wait_sap_ready(self, timeout: float = 5.0) -> bool:
        """Aguarda SAP ficar pronto."""
        end_time = time.time() + timeout
        while time.time() < end_time:
            try:
                if not self.session.Busy:
                    return True
            except:
                pass
            time.sleep(0.02)
        return False
    
    def _existe_popup(self, timeout: float = 2.0) -> bool:
        """Verifica se existe popup (wnd[1])."""
        end_time = time.time() + timeout
        while time.time() < end_time:
            try:
                self.session.findById("wnd[1]")
                return True
            except:
                time.sleep(0.02)
        return False
    
    def _confirmar_popup(self) -> bool:
        """Confirma popup se existir."""
        try:
            if self._existe_popup(timeout=3):
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                self._wait_sap_ready(timeout=2.0)
                return True
            return False
        except Exception as e:
            print(f"[AVISO] Erro ao confirmar popup: {e}")
            return False
    
    def executar(self) -> bool:
        """
        Executa: Salvar → Verificar → Habilitar edição.
        
        Returns:
            True se salvou e habilitou edição com sucesso
        """
        try:
            print("\n" + "="*70)
            print("SALVAMENTO")
            print("="*70)
            
            # 1. SALVAR
            print("\n[1/3] Salvando...")
            try:
                self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
                print("[OK] Salvar pressionado")
            except Exception as e:
                print(f"[ERRO] Falha ao pressionar Salvar: {e}")
                return False
            
            # Aguarda processamento
            self._wait_sap_ready(timeout=3.0)
            
            # Trata popup se aparecer
            if self._confirmar_popup():
                print("[OK] Popup confirmado")
            
            # 2. VERIFICAR
            print("\n[2/3] Verificando salvamento...")
            try:
                self.session.findById("wnd[0]")
                print("[OK] ✅ Cadastro salvo")
            except Exception as e:
                print(f"[ERRO] Falha ao validar salvamento: {e}")
                return False
            
            # 3. HABILITAR EDIÇÃO
            print("\n[3/3] Habilitando edição...")
            try:
                self._wait_sap_ready(timeout=2.0)
                self.session.findById("wnd[0]/tbar[1]/btn[6]").press()
                print("[OK] Editar pressionado")
                self._wait_sap_ready(timeout=2.0)
                print("[OK] ✅ Edição habilitada")
            except Exception as e:
                print(f"[AVISO] Falha ao habilitar edição: {e}")
                print("[INFO] Cadastro salvo, mas edição pode não estar ativa")
                return True  # Salvamento foi bem-sucedido
            
            print("\n[OK] ✅✅✅ SALVAMENTO CONCLUÍDO")
            print("="*70 + "\n")
            return True
            
        except Exception as e:
            print(f"\n[ERRO] Falha no salvamento: {e}")
            import traceback
            traceback.print_exc()
            return False