"""
Configuração centralizada de logging para automação SAP.

AUTOR: Sistema de Automação SAP
DATA: 2026-01-08
VERSÃO: 1.1 (Corrigido)

CARACTERÍSTICAS:
- Logs no console com cores
- Logs em arquivo para auditoria
- Níveis diferentes para dev/prod
- Formato padronizado
"""

import logging
import sys
from pathlib import Path
from datetime import datetime
from typing import Optional

# Tentar importar colorlog (opcional)
try:
    import colorlog
    COLORLOG_DISPONIVEL = True
except ImportError:
    COLORLOG_DISPONIVEL = False


class ConfiguradorLog:
    """Configurador centralizado de logging."""
    
    # Níveis de log
    DEBUG = logging.DEBUG
    INFO = logging.INFO
    WARNING = logging.WARNING
    ERROR = logging.ERROR
    CRITICAL = logging.CRITICAL
    
    # Diretório de logs
    DIR_LOGS = Path("logs")
    
    @classmethod
    def configurar(
        cls,
        nome_modulo: str = "AutomacaoSAP",
        nivel_console: int = None,
        nivel_arquivo: int = None,
        usar_cores: bool = True,
        salvar_em_arquivo: bool = True
    ) -> logging.Logger:
        """
        Configura logging para o módulo.
        
        Args:
            nome_modulo: Nome do módulo (ex: "AutomacaoSAP", "ConexaoSAP")
            nivel_console: Nível de log para console
            nivel_arquivo: Nível de log para arquivo
            usar_cores: Se True, usa cores no console (se colorlog disponível)
            salvar_em_arquivo: Se True, salva logs em arquivo
            
        Returns:
            logging.Logger: Logger configurado
        """
        # Valores padrão se None
        if nivel_console is None:
            nivel_console = cls.INFO
        if nivel_arquivo is None:
            nivel_arquivo = cls.DEBUG
        
        # Criar logger
        logger = logging.getLogger(nome_modulo)
        logger.setLevel(cls.DEBUG)
        logger.handlers.clear()  # Limpar handlers existentes
        
        # Handler de console
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(nivel_console)
        
        if usar_cores and COLORLOG_DISPONIVEL:
            # Formato com cores
            console_format = colorlog.ColoredFormatter(
                '%(log_color)s%(asctime)s [%(levelname)-8s]%(reset)s %(message)s',
                datefmt='%H:%M:%S',
                log_colors={
                    'DEBUG': 'cyan',
                    'INFO': 'green',
                    'WARNING': 'yellow',
                    'ERROR': 'red',
                    'CRITICAL': 'bold_red',
                }
            )
        else:
            # Formato sem cores
            console_format = logging.Formatter(
                '%(asctime)s [%(levelname)-8s] %(message)s',
                datefmt='%H:%M:%S'
            )
        
        console_handler.setFormatter(console_format)
        logger.addHandler(console_handler)
        
        # Handler de arquivo (se habilitado)
        if salvar_em_arquivo:
            arquivo_handler = cls._criar_handler_arquivo(
                nome_modulo,
                nivel_arquivo
            )
            if arquivo_handler:
                logger.addHandler(arquivo_handler)
        
        return logger
    
    @classmethod
    def _criar_handler_arquivo(
        cls,
        nome_modulo: str,
        nivel: int
    ) -> Optional[logging.FileHandler]:
        """
        Cria handler para salvar logs em arquivo.
        
        Args:
            nome_modulo: Nome do módulo
            nivel: Nível de log
            
        Returns:
            FileHandler ou None se falhar
        """
        try:
            # Criar diretório de logs
            cls.DIR_LOGS.mkdir(exist_ok=True)
            
            # Nome do arquivo com timestamp
            timestamp = datetime.now().strftime("%Y%m%d")
            arquivo_log = cls.DIR_LOGS / f"{nome_modulo}_{timestamp}.log"
            
            # Criar handler
            handler = logging.FileHandler(
                arquivo_log,
                mode='a',
                encoding='utf-8'
            )
            handler.setLevel(nivel)
            
            # Formato detalhado para arquivo
            formato = logging.Formatter(
                '%(asctime)s [%(levelname)-8s] [%(name)s] %(message)s',
                datefmt='%Y-%m-%d %H:%M:%S'
            )
            handler.setFormatter(formato)
            
            return handler
            
        except Exception as e:
            print(f"⚠️ Aviso: Não foi possível criar arquivo de log: {e}")
            return None
    
    @classmethod
    def configurar_producao(cls, nome_modulo: str = "AutomacaoSAP") -> logging.Logger:
        """
        Configuração para ambiente de produção.
        - Console: INFO
        - Arquivo: DEBUG
        - Cores: Sim (se disponível)
        
        Args:
            nome_modulo: Nome do módulo
            
        Returns:
            Logger configurado
        """
        return cls.configurar(
            nome_modulo=nome_modulo,
            nivel_console=cls.INFO,
            nivel_arquivo=cls.DEBUG,
            usar_cores=True,
            salvar_em_arquivo=True
        )
    
    @classmethod
    def configurar_desenvolvimento(cls, nome_modulo: str = "AutomacaoSAP") -> logging.Logger:
        """
        Configuração para ambiente de desenvolvimento.
        - Console: DEBUG
        - Arquivo: DEBUG
        - Cores: Sim (se disponível)
        
        Args:
            nome_modulo: Nome do módulo
            
        Returns:
            Logger configurado
        """
        return cls.configurar(
            nome_modulo=nome_modulo,
            nivel_console=cls.DEBUG,
            nivel_arquivo=cls.DEBUG,
            usar_cores=True,
            salvar_em_arquivo=True
        )
    
    @classmethod
    def configurar_minimo(cls, nome_modulo: str = "AutomacaoSAP") -> logging.Logger:
        """
        Configuração mínima (apenas erros no console).
        - Console: ERROR
        - Arquivo: INFO
        - Cores: Não
        
        Args:
            nome_modulo: Nome do módulo
            
        Returns:
            Logger configurado
        """
        return cls.configurar(
            nome_modulo=nome_modulo,
            nivel_console=cls.ERROR,
            nivel_arquivo=cls.INFO,
            usar_cores=False,
            salvar_em_arquivo=True
        )


# ===== FUNÇÕES DE CONVENIÊNCIA =====

def obter_logger(
    nome_modulo: str,
    producao: bool = True
) -> logging.Logger:
    """
    Obtém logger configurado.
    
    Args:
        nome_modulo: Nome do módulo
        producao: Se True, usa config de produção
        
    Returns:
        Logger configurado
    """
    if producao:
        return ConfiguradorLog.configurar_producao(nome_modulo)
    else:
        return ConfiguradorLog.configurar_desenvolvimento(nome_modulo)


# ===== EXEMPLO DE USO =====

if __name__ == "__main__":
    # Testar diferentes configurações
    
    print("\n" + "="*70)
    print("TESTE DE CONFIGURAÇÃO DE LOGGING")
    print("="*70 + "\n")
    
    # 1. Produção
    print("1️⃣ Configuração de PRODUÇÃO:")
    logger_prod = ConfiguradorLog.configurar_producao("TesteProducao")
    logger_prod.debug("Mensagem DEBUG (não aparece)")
    logger_prod.info("✅ Mensagem INFO")
    logger_prod.warning("⚠️ Mensagem WARNING")
    logger_prod.error("❌ Mensagem ERROR")
    
    print("\n" + "-"*70 + "\n")
    
    # 2. Desenvolvimento
    print("2️⃣ Configuração de DESENVOLVIMENTO:")
    logger_dev = ConfiguradorLog.configurar_desenvolvimento("TesteDev")
    logger_dev.debug("🔍 Mensagem DEBUG (aparece)")
    logger_dev.info("✅ Mensagem INFO")
    logger_dev.warning("⚠️ Mensagem WARNING")
    
    print("\n" + "-"*70 + "\n")
    
    # 3. Mínimo
    print("3️⃣ Configuração MÍNIMA:")
    logger_min = ConfiguradorLog.configurar_minimo("TesteMinimo")
    logger_min.info("INFO (não aparece)")
    logger_min.warning("WARNING (não aparece)")
    logger_min.error("❌ ERROR (aparece)")
    
    print("\n" + "="*70)
    print("✅ Teste concluído!")
    print(f"📁 Logs salvos em: {ConfiguradorLog.DIR_LOGS.absolute()}")
    print("="*70 + "\n")