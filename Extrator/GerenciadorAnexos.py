"""
Módulo responsável por gerenciar anexos de arquivos do fornecedor.
Valida anexos obrigatórios e opcionais, padroniza nomes e persiste dados.
"""
import json
from pathlib import Path
from typing import Dict, List, Tuple, Optional


class GerenciadorAnexos:
    """
    Gerencia anexos de arquivos do fornecedor.
    
    Anexos obrigatórios:
    - FORMULARIO DE CADASTRO DE FORNECEDOR
    - CNPJ
    - COMPROVANTE BANCARIO
    
    Permite anexos opcionais com nomes customizados.
    """
    
    # Nomes padronizados dos anexos obrigatórios
    ANEXOS_OBRIGATORIOS = [
        "FORMULARIO DE CADASTRO DE FORNECEDOR",
        "CNPJ",
        "COMPROVANTE BANCARIO"
    ]
    
    def __init__(self, json_path: Path, limpar_ao_iniciar: bool = True):
        """
        Inicializa o gerenciador de anexos.
        
        Args:
            json_path: Caminho para o arquivo JSON de anexos
            limpar_ao_iniciar: Se True, limpa anexos anteriores ao iniciar
        """
        self.json_path = Path(json_path)
        
        if limpar_ao_iniciar:
            # Inicia sempre com estrutura limpa
            self.dados = {'anexos': self._estrutura_padrao()}
            print("[INFO] Anexos limpos - iniciando nova sessão")
        else:
            self.dados = self._carregar_dados()
    
    def _carregar_dados(self) -> dict:
        """
        Carrega dados de anexos do JSON ou cria estrutura padrão.
        
        Returns:
            Dicionário com estrutura de anexos
        """
        if self.json_path.exists():
            try:
                with open(self.json_path, 'r', encoding='utf-8') as f:
                    dados = json.load(f)
                
                # Valida estrutura
                if 'anexos' not in dados:
                    dados['anexos'] = self._estrutura_padrao()
                
                return dados
            except Exception as e:
                print(f"[ERRO] Falha ao carregar anexos: {e}")
                return {'anexos': self._estrutura_padrao()}
        else:
            return {'anexos': self._estrutura_padrao()}
    
    def _estrutura_padrao(self) -> dict:
        """
        Retorna estrutura padrão de anexos.
        
        Returns:
            Dicionário com estrutura vazia
        """
        return {
            'obrigatorios': {nome: "" for nome in self.ANEXOS_OBRIGATORIOS},
            'opcionais': {}
        }
    
    def salvar_dados(self) -> bool:
        """
        Salva dados de anexos no JSON.
        
        Returns:
            True se salvou com sucesso
        """
        try:
            self.json_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(self.json_path, 'w', encoding='utf-8') as f:
                json.dump(self.dados, f, ensure_ascii=False, indent=4)
            
            print(f"[INFO] Anexos salvos em: {self.json_path.name}")
            return True
        except Exception as e:
            print(f"[ERRO] Falha ao salvar anexos: {e}")
            return False
    
    @staticmethod
    def padronizar_nome(nome: str) -> str:
        """
        Padroniza nome para MAIÚSCULAS e remove espaços extras.
        
        Args:
            nome: Nome a ser padronizado
            
        Returns:
            Nome padronizado
        """
        if not nome:
            return ""
        return nome.strip().upper()
    
    def adicionar_obrigatorio(self, nome: str, caminho: str) -> Tuple[bool, str]:
        """
        Adiciona caminho de arquivo obrigatório.
        
        Args:
            nome: Nome do anexo obrigatório (será padronizado)
            caminho: Caminho físico do arquivo
            
        Returns:
            Tupla (sucesso, mensagem)
        """
        nome_padronizado = self.padronizar_nome(nome)
        
        if nome_padronizado not in self.ANEXOS_OBRIGATORIOS:
            return False, f"'{nome}' não é um anexo obrigatório válido"
        
        if not Path(caminho).exists():
            return False, f"Arquivo não encontrado: {caminho}"
        
        self.dados['anexos']['obrigatorios'][nome_padronizado] = str(caminho)
        return True, f"Anexo '{nome_padronizado}' adicionado com sucesso"
    
    def adicionar_opcional(self, nome_customizado: str, caminho: str) -> Tuple[bool, str]:
        """
        Adiciona arquivo opcional com nome customizado.
        
        Args:
            nome_customizado: Nome personalizado (será padronizado para MAIÚSCULAS)
            caminho: Caminho físico do arquivo
            
        Returns:
            Tupla (sucesso, mensagem)
        """
        if not nome_customizado or not nome_customizado.strip():
            return False, "Nome do anexo não pode ser vazio"
        
        nome_padronizado = self.padronizar_nome(nome_customizado)
        
        if not Path(caminho).exists():
            return False, f"Arquivo não encontrado: {caminho}"
        
        # Verifica se nome já existe nos opcionais
        if nome_padronizado in self.dados['anexos']['opcionais']:
            return False, f"Já existe um anexo opcional com o nome '{nome_padronizado}'"
        
        self.dados['anexos']['opcionais'][nome_padronizado] = str(caminho)
        return True, f"Anexo opcional '{nome_padronizado}' adicionado com sucesso"
    
    def remover_opcional(self, nome: str) -> Tuple[bool, str]:
        """
        Remove anexo opcional.
        
        Args:
            nome: Nome do anexo opcional
            
        Returns:
            Tupla (sucesso, mensagem)
        """
        nome_padronizado = self.padronizar_nome(nome)
        
        if nome_padronizado not in self.dados['anexos']['opcionais']:
            return False, f"Anexo '{nome}' não encontrado"
        
        del self.dados['anexos']['opcionais'][nome_padronizado]
        return True, f"Anexo '{nome_padronizado}' removido com sucesso"
    
    def validar_obrigatorios(self) -> Tuple[bool, List[str]]:
        """
        Valida se todos os anexos obrigatórios foram fornecidos.
        
        Returns:
            Tupla (todos_presentes, lista_de_faltantes)
        """
        faltantes = []
        
        for nome in self.ANEXOS_OBRIGATORIOS:
            caminho = self.dados['anexos']['obrigatorios'].get(nome, "")
            
            if not caminho or not Path(caminho).exists():
                faltantes.append(nome)
        
        return len(faltantes) == 0, faltantes
    
    def obter_todos_anexos(self) -> Dict[str, str]:
        """
        Retorna todos os anexos (obrigatórios + opcionais) em um único dicionário.
        
        Returns:
            Dicionário com nome_padronizado: caminho
        """
        todos = {}
        
        # Adiciona obrigatórios
        for nome, caminho in self.dados['anexos']['obrigatorios'].items():
            if caminho and Path(caminho).exists():
                todos[nome] = caminho
        
        # Adiciona opcionais
        for nome, caminho in self.dados['anexos']['opcionais'].items():
            if caminho and Path(caminho).exists():
                todos[nome] = caminho
        
        return todos
    
    def obter_obrigatorios(self) -> Dict[str, str]:
        """
        Retorna apenas anexos obrigatórios.
        
        Returns:
            Dicionário com anexos obrigatórios
        """
        return self.dados['anexos']['obrigatorios'].copy()
    
    def obter_opcionais(self) -> Dict[str, str]:
        """
        Retorna apenas anexos opcionais.
        
        Returns:
            Dicionário com anexos opcionais
        """
        return self.dados['anexos']['opcionais'].copy()
    
    def limpar_todos(self) -> None:
        """Limpa todos os anexos."""
        self.dados['anexos'] = self._estrutura_padrao()
    
    def contar_anexos(self) -> Tuple[int, int]:
        """
        Conta quantos anexos estão cadastrados.
        
        Returns:
            Tupla (obrigatórios_preenchidos, opcionais)
        """
        obrigatorios_ok = sum(
            1 for caminho in self.dados['anexos']['obrigatorios'].values()
            if caminho and Path(caminho).exists()
        )
        
        opcionais_count = len(self.dados['anexos']['opcionais'])
        
        return obrigatorios_ok, opcionais_count


def obter_caminho_anexos_json() -> Path:
    """
    Retorna o caminho padrão para o JSON de anexos.
    
    Returns:
        Path para fornecedor_anexos.json
    """
    import sys
    from pathlib import Path
    
    # Adiciona raiz ao path se necessário
    root_dir = Path(__file__).resolve().parents[1]
    if str(root_dir) not in sys.path:
        sys.path.insert(0, str(root_dir))
    
    from utils import get_arquivos_dir
    
    return get_arquivos_dir() / "fornecedor_anexos.json"