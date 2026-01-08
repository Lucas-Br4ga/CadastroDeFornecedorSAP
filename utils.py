"""
Funções auxiliares compartilhadas entre os módulos do projeto.
"""
from pathlib import Path


def get_project_root(start_path: Path = None) -> Path:
    """
    Encontra o diretório raiz do projeto procurando por main.py.
    
    Args:
        start_path: Caminho inicial para busca (padrão: arquivo atual)
    
    Returns:
        Path para o diretório raiz do projeto
    
    Raises:
        FileNotFoundError: Se não encontrar main.py em nenhum diretório pai
    """
    if start_path is None:
        start_path = Path(__file__).resolve().parent
    
    current = Path(start_path).resolve()
    
    # Sobe até encontrar main.py ou chegar na raiz do sistema
    max_levels = 10  # Limite de segurança
    for _ in range(max_levels):
        if (current / "main.py").exists():
            return current
        
        parent = current.parent
        if parent == current:  # Chegou na raiz do sistema
            break
        current = parent
    
    # Se não encontrou, retorna o diretório do arquivo atual
    return Path(__file__).resolve().parent


def get_arquivos_dir(create: bool = True) -> Path:
    """
    Retorna o caminho para o diretório Arquivos/.
    
    Args:
        create: Se True, cria o diretório caso não exista
    
    Returns:
        Path para o diretório Arquivos/
    """
    arquivos_dir = get_project_root() / "Arquivos"
    
    if create and not arquivos_dir.exists():
        arquivos_dir.mkdir(parents=True, exist_ok=True)
    
    return arquivos_dir


def get_json_paths() -> dict:
    """
    Retorna os caminhos padrão dos arquivos JSON.
    
    Returns:
        Dicionário com 'bruto' e 'limpo' contendo os caminhos
    """
    arquivos_dir = get_arquivos_dir()
    
    return {
        "bruto": arquivos_dir / "fornecedor_bruto.json",
        "limpo": arquivos_dir / "fornecedor_limpo.json"
    }


def get_anexos_json_path() -> Path:
    """
    Retorna o caminho para o arquivo JSON de anexos.
    
    Returns:
        Path para fornecedor_anexos.json
    """
    return get_arquivos_dir() / "fornecedor_anexos.json"