"""
Módulo de Entrada na Transação XK01.
Responsável por abrir a transação e realizar configurações iniciais.
"""

import time
from .ManipuladorCampos import ManipuladorCamposSAP, GerenciadorPopups


class EntrarTransacao:
    """
    Classe responsável por entrar na transação XK01
    e realizar configurações iniciais.
    """
    
    def __init__(self, session, manipulador_campos: ManipuladorCamposSAP):
        """
        Inicializa o módulo de entrada na transação.
        
        Args:
            session: Sessão SAP ativa
            manipulador_campos: Manipulador de campos SAP
        """
        self.session = session
        self.campos = manipulador_campos
        self.popups = GerenciadorPopups(session)
    
    def abrir_transacao_xk01(self) -> bool:
        """
        Abre a transação XK01.
        
        Returns:
            True se abriu com sucesso
            
        Raises:
            Exception: Se não conseguir abrir a transação
        """
        print("\n[ETAPA] Abrindo transação XK01...")
        
        try:
            # Busca campo de transação pelo name
            okcd = self.campos.buscar_elemento_por_name('transacao', 'codigo')
            
            # Limpa campo
            okcd.text = ""
            time.sleep(0.1)
            
            # Digita transação (SEMPRE com /n)
            okcd.text = "/nxk01"
            okcd.setFocus()
            
            # Pressiona ENTER
            self.session.findById("wnd[0]").sendVKey(0)
            
            print("[OK] Transação XK01 aberta")
            time.sleep(1)
            
            return True
            
        except Exception as e:
            raise Exception(f"Erro ao abrir transação XK01: {e}")
    
    def selecionar_fornecedor_nacional(self) -> bool:
        """
        Seleciona fornecedor nacional no popup que aparece após abrir XK01.
        
        Returns:
            True se selecionou com sucesso
            
        Raises:
            Exception: Se não conseguir selecionar
        """
        print("\n[ETAPA] Selecionando fornecedor nacional...")
        
        try:
            # Aguarda popup aparecer
            if not self.popups.existe_popup(timeout=5):
                raise Exception("Popup de seleção de tipo não apareceu")
            
            # Seleciona radio "Fornecedor nacional" (índice 0)
            radio_id = "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]"
            
            radio = self.session.findById(radio_id)
            radio.select()
            radio.setFocus()
            
            time.sleep(0.2)
            
            # Confirma
            self.popups.confirmar_popup()
            
            print("[OK] Fornecedor nacional selecionado")
            time.sleep(1)
            
            return True
            
        except Exception as e:
            raise Exception(f"Erro ao selecionar fornecedor nacional: {e}")
    
    def configurar_grupo_criacao(self) -> bool:
        """
        Configura o grupo de criação como ZVRE.
        
        Returns:
            True se configurou com sucesso
        """
        print("\n[ETAPA] Configurando grupo de criação...")
        
        try:
            # Seleciona combo grupo de criação
            self.campos.selecionar_combo(
                'transacao',
                'grupo_criacao',
                'ZVRE'
            )
            
            print("[OK] Grupo de criação configurado: ZVRE")
            return True
            
        except Exception as e:
            print(f"[AVISO] Erro ao configurar grupo de criação: {e}")
            return False
    
    def executar(self) -> bool:
        """
        Executa todas as etapas de entrada na transação.
        
        Returns:
            True se todas as etapas foram bem-sucedidas
            
        Raises:
            Exception: Se alguma etapa crítica falhar
        """
        print("\n" + "="*70)
        print("MÓDULO: ENTRADA NA TRANSAÇÃO")
        print("="*70)
        
        # 1. Abrir transação XK01
        self.abrir_transacao_xk01()
        
        # 2. Selecionar fornecedor nacional
        self.selecionar_fornecedor_nacional()
        
        # 3. Configurar grupo de criação
        self.configurar_grupo_criacao()
        
        print("\n[OK] Entrada na transação concluída com sucesso!")
        print("="*70 + "\n")
        
        return True