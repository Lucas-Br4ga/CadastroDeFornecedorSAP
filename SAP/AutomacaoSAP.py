"""
Automação SAP - Cadastro de Fornecedor (XK01) - VERSÃO FINAL

FLUXO CORRETO COM 3 SALVAMENTOS:
1-3: Preenchimento (sem save)
4: Empresas → SAVE 1/3 → EDIT
5: Compras → SAVE 2/3 → EDIT  
6: Anexos → SAVE 3/3 → EDIT
"""

import sys
import json
from pathlib import Path
from typing import Tuple, Dict

from .ConexaoSAP import ConexaoSAP, SAPConnectionError
from .ManipuladorCampos import ManipuladorCamposSAP
from .EntrarTransacao import EntrarTransacao
from .PreencherDadosGerais import PreencherDadosGerais
from .PreencherDadosBancarios import PreencherDadosBancarios
from .PreencherEmpresas import PreencherEmpresas
from .PreencherCompras import PreencherCompras
from .GerenciadorAnexos import GerenciadorAnexosSAP


class AutomacaoSAP:
    
    def __init__(self, dados_json_path: Path, campos_sap_json_path: Path):
        self.dados_json_path = dados_json_path
        self.campos_sap_json_path = campos_sap_json_path
        self.dados_fornecedor = None
        self.conexao = None
        self.session = None
        self.manipulador_campos = None
    
    def _carregar_dados_fornecedor(self) -> Dict:
        with open(self.dados_json_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def _inicializar_conexao(self):
        print("\n[ETAPA] Inicializando conexão com SAP...")
        self.conexao = ConexaoSAP.obter_instancia()
        self.conexao.conectar(timeout=30)
        self.conexao.maximizar_janela()
        self.session = self.conexao.session
        print("[OK] Conexão inicializada\n")
    
    def _inicializar_modulos(self):
        self.manipulador_campos = ManipuladorCamposSAP(self.session, self.campos_sap_json_path)
        
        self.entrar_transacao = EntrarTransacao(self.session, self.manipulador_campos)
        self.preencher_dados_gerais = PreencherDadosGerais(self.session, self.manipulador_campos, self.dados_fornecedor)
        self.preencher_dados_bancarios = PreencherDadosBancarios(self.session, self.manipulador_campos, self.dados_fornecedor)
        self.preencher_empresas = PreencherEmpresas(self.session, self.manipulador_campos, self.dados_fornecedor)
        self.preencher_compras = PreencherCompras(self.session, self.manipulador_campos, self.dados_fornecedor)
        self.gerenciador_anexos = GerenciadorAnexosSAP(self.session, self.manipulador_campos)
    
    def executar(self) -> Tuple[bool, str]:
        """
        Executa automação completa.
        
        SAVES:
        - SAVE 1/3: Empresas → btn[11] → btn[6]
        - SAVE 2/3: Compras → btn[11] → btn[6]
        - SAVE 3/3: Anexos → btn[11] → btn[6]
        """
        try:
            print("\n" + "="*70)
            print("AUTOMAÇÃO SAP - CADASTRO DE FORNECEDOR (XK01)")
            print("="*70)
            
            # Carregar dados
            self.dados_fornecedor = self._carregar_dados_fornecedor()
            print(f"[OK] Dados: {self.dados_fornecedor['empresa']['razao_social']}\n")
            
            # Conectar
            self._inicializar_conexao()
            self._inicializar_modulos()
            
            # 1. Transação
            print("\n" + "="*70)
            print("ETAPA 1/6: ENTRADA NA TRANSAÇÃO")
            print("="*70)
            if not self.entrar_transacao.executar():
                return False, "Falha ao entrar na transação XK01"
            
            # 2. Dados Gerais
            print("\n" + "="*70)
            print("ETAPA 2/6: DADOS GERAIS")
            print("="*70)
            if not self.preencher_dados_gerais.executar():
                return False, "Falha ao preencher dados gerais"
            
            # 3. Dados Bancários
            print("\n" + "="*70)
            print("ETAPA 3/6: DADOS BANCÁRIOS")
            print("="*70)
            if not self.preencher_dados_bancarios.executar():
                return False, "Falha ao preencher dados bancários"
            
            # 4. Empresas (SAVE 1/3)
            print("\n" + "="*70)
            print("ETAPA 4/6: EMPRESAS (COM SAVE 1/3)")
            print("="*70)
            if not self.preencher_empresas.executar():
                return False, "Falha no SAVE 1/3 (Empresas)"
            
            # 5. Compras (SAVE 2/3)
            print("\n" + "="*70)
            print("ETAPA 5/6: COMPRAS (COM SAVE 2/3)")
            print("="*70)
            if not self.preencher_compras.executar():
                return False, "Falha no SAVE 2/3 (Compras)"
            
            # 6. Anexos (SAVE 3/3)
            print("\n" + "="*70)
            print("ETAPA 6/6: ANEXOS (COM SAVE 3/3)")
            print("="*70)
            if not self.gerenciador_anexos.executar():
                return False, "Falha no SAVE 3/3 (Anexos)"
            
            # Sucesso
            print("\n" + "="*70)
            print("✅✅✅ AUTOMAÇÃO CONCLUÍDA COM SUCESSO ✅✅✅")
            print("="*70)
            
            mensagem = (
                "✅ AUTOMAÇÃO CONCLUÍDA!\n\n"
                "Resumo:\n"
                "✅ Dados gerais preenchidos\n"
                "✅ Dados bancários preenchidos\n"
                "✅ Empresas (BR01, BR04, BR20) → SAVE 1/3 → EDIT\n"
                "✅ Compras (FLVN01) → SAVE 2/3 → EDIT\n"
                "✅ Anexos → SAVE 3/3 → EDIT\n\n"
                "💡 Cadastro salvo e pronto para edições!"
            )
            
            return True, mensagem
            
        except SAPConnectionError as e:
            return False, f"Erro de conexão com SAP:\n\n{str(e)}"
        except Exception as e:
            import traceback
            traceback.print_exc()
            return False, f"Erro durante a automação:\n\n{str(e)}"


def executar_automacao() -> Tuple[bool, str]:
    """Função principal"""
    try:
        root_dir = Path(__file__).resolve().parents[1]
        sys.path.insert(0, str(root_dir))
        
        from utils import get_json_paths
        
        paths = get_json_paths()
        dados_json = paths["limpo"]
        campos_sap_json = root_dir / "SAP" / "campos_sap.json"
        
        if not dados_json.exists():
            return False, f"Arquivo não encontrado:\n{dados_json}"
        
        if not campos_sap_json.exists():
            return False, f"Arquivo não encontrado:\n{campos_sap_json}"
        
        automacao = AutomacaoSAP(dados_json, campos_sap_json)
        return automacao.executar()
        
    except Exception as e:
        import traceback
        erro = traceback.format_exc()
        return False, f"Erro ao inicializar:\n\n{erro}"


if __name__ == "__main__":
    sucesso, mensagem = executar_automacao()
    print("\n" + "="*70)
    print("✅ SUCESSO" if sucesso else "❌ ERRO")
    print("="*70)
    print(f"\n{mensagem}\n")