"""
Módulo de Conexão com SAP GUI.
Responsável por estabelecer e gerenciar a conexão com o SAP.
"""

import time
import win32com.client


class SAPConnectionError(Exception):
    """Erro de conexão com SAP"""
    pass


class ConexaoSAP:
    """
    Gerenciador de conexão com SAP GUI.
    Implementa padrão singleton para garantir uma única conexão.
    """
    
    _instance = None
    
    def __init__(self):
        self.sap_gui = None
        self.application = None
        self.connection = None
        self.session = None
    
    @classmethod
    def obter_instancia(cls) -> 'ConexaoSAP':
        """Retorna instância única (singleton)"""
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance
    
    def conectar(self, timeout: int = 30) -> bool:
        """
        Estabelece conexão com SAP GUI usando GetObject.
        
        Args:
            timeout: Tempo máximo de espera em segundos
            
        Returns:
            True se conectou com sucesso
            
        Raises:
            SAPConnectionError: Se não conseguir conectar
        """
        print("[INFO] Conectando ao SAP GUI...")
        
        inicio = time.time()
        
        while time.time() - inicio < timeout:
            try:
                # SEMPRE usar GetObject (NUNCA Dispatch)
                self.sap_gui = win32com.client.GetObject("SAPGUI")
                
                if not self.sap_gui:
                    raise SAPConnectionError("SAPGUI não está disponível")
                
                # Obter Scripting Engine
                self.application = self.sap_gui.GetScriptingEngine
                
                if not self.application:
                    raise SAPConnectionError("Scripting Engine não disponível")
                
                # Validar se existe conexão
                if self.application.Children.Count == 0:
                    raise SAPConnectionError("Nenhuma conexão SAP encontrada")
                
                # Obter primeira conexão
                self.connection = self.application.Children(0)
                
                if not self.connection:
                    raise SAPConnectionError("Conexão SAP não disponível")
                
                # Validar se existe sessão
                if self.connection.Children.Count == 0:
                    raise SAPConnectionError("Nenhuma sessão SAP encontrada")
                
                # Obter primeira sessão
                self.session = self.connection.Children(0)
                
                if not self.session:
                    raise SAPConnectionError("Sessão SAP não disponível")
                
                print(f"[OK] Conectado ao SAP - Sessão: {self.session.Info.SystemName}")
                return True
                
            except Exception as e:
                print(f"[AVISO] Tentando conectar... ({int(time.time() - inicio)}s)")
                time.sleep(1)
        
        raise SAPConnectionError(
            f"Não foi possível conectar ao SAP após {timeout}s. "
            "Verifique se o SAP GUI está aberto e logado."
        )
    
    def validar_sessao(self) -> bool:
        """Valida se a sessão ainda está ativa"""
        try:
            if not self.session:
                return False
            
            # Tenta acessar propriedade para validar
            _ = self.session.Info.SystemName
            return True
            
        except Exception:
            return False
    
    def maximizar_janela(self) -> None:
        """Maximiza janela principal do SAP"""
        try:
            self.session.findById("wnd[0]").maximize()
            print("[INFO] Janela SAP maximizada")
        except Exception as e:
            print(f"[AVISO] Não foi possível maximizar janela: {e}")