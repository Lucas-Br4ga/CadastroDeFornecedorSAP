"""
Padronizador de dados de fornecedores para automação SAP.
Garante que todos os dados estejam no formato correto antes do envio.

VERSÃO 3.0:
- Todos os textos convertidos para MAIÚSCULAS
- Conversão de prazo de pagamento (dias → código)
- Conversão de estados (extenso → sigla)
- Preservação inteligente do código do banco
"""
import json
import re
from pathlib import Path
from typing import Dict, Tuple, List, Optional
from .PadronizadorInscricaoEstadual import PadronizadorInscricaoEstadual


class PadronizadorDados:
    """
    Padroniza dados de fornecedor para garantir compatibilidade com SAP.
    Aplica regras específicas por campo e por estado.
    """
    
    # Mapa de conversão: dias → código de prazo
    MAPA_PRAZO_PAGAMENTO = {
        0: "BRFAA",    # Vista
        5: "BRFB",
        10: "BRFC",
        15: "BRFD",
        20: "BRFE",
        25: "BRFF",
        30: "BRFG",
        35: "BRFH",
        40: "BRFI",
        45: "BRFJ",
        50: "BRFK",
        55: "BRFL",
        60: "BRFM",
        65: "BRFN",
        70: "BRFO",
        75: "BRFP",
        80: "BRFQ",
        85: "BRFR",
        90: "BRFS"
    }
    
    # Mapa de conversão: estado por extenso → sigla
    MAPA_ESTADOS = {
        "ACRE": "AC",
        "ALAGOAS": "AL",
        "AMAPA": "AP",
        "AMAZONAS": "AM",
        "BAHIA": "BA",
        "CEARA": "CE",
        "DISTRITO FEDERAL": "DF",
        "BRASILIA": "DF",
        "ESPIRITO SANTO": "ES",
        "GOIAS": "GO",
        "MARANHAO": "MA",
        "MATO GROSSO": "MT",
        "MATO GROSSO DO SUL": "MS",
        "MINAS GERAIS": "MG",
        "PARA": "PA",
        "PARAIBA": "PB",
        "PARANA": "PR",
        "PERNAMBUCO": "PE",
        "PIAUI": "PI",
        "RIO DE JANEIRO": "RJ",
        "RIO GRANDE DO NORTE": "RN",
        "RIO GRANDE DO SUL": "RS",
        "RONDONIA": "RO",
        "RORAIMA": "RR",
        "SANTA CATARINA": "SC",
        "SAO PAULO": "SP",
        "SERGIPE": "SE",
        "TOCANTINS": "TO"
    }
    
    # UFs válidas
    UFS_VALIDAS = [
        'AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA',
        'MT', 'MS', 'MG', 'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN',
        'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'
    ]
    
    def __init__(self, json_path: Path):
        """
        Inicializa o padronizador.
        
        Args:
            json_path: Caminho para o arquivo JSON limpo
        """
        self.json_path = Path(json_path)
        self.dados = None
        self.avisos = []
        self.erros = []
        
        if not self.json_path.exists():
            raise FileNotFoundError(f"JSON não encontrado: {self.json_path}")
    
    def carregar_dados(self) -> bool:
        """
        Carrega dados do JSON.
        
        Returns:
            True se sucesso, False caso contrário
        """
        try:
            with open(self.json_path, 'r', encoding='utf-8') as f:
                self.dados = json.load(f)
            
            print(f"[INFO] Dados carregados: {self.json_path.name}")
            return True
            
        except json.JSONDecodeError as e:
            erro = f"JSON inválido: {str(e)}"
            self.erros.append(erro)
            print(f"[ERRO] {erro}")
            return False
            
        except Exception as e:
            erro = f"Erro ao carregar JSON: {str(e)}"
            self.erros.append(erro)
            print(f"[ERRO] {erro}")
            return False
    
    def salvar_dados(self) -> bool:
        """
        Salva dados padronizados no JSON.
        
        Returns:
            True se sucesso, False caso contrário
        """
        try:
            with open(self.json_path, 'w', encoding='utf-8') as f:
                json.dump(self.dados, f, ensure_ascii=False, indent=4)
            
            print(f"[INFO] Dados padronizados salvos: {self.json_path.name}")
            return True
            
        except Exception as e:
            erro = f"Erro ao salvar JSON: {str(e)}"
            self.erros.append(erro)
            print(f"[ERRO] {erro}")
            return False
    
    def padronizar_cnpj(self, cnpj: str) -> str:
        """
        Padroniza CNPJ: apenas dígitos, 14 caracteres.
        
        Args:
            cnpj: CNPJ com ou sem formatação
            
        Returns:
            CNPJ padronizado (apenas dígitos)
        """
        if not cnpj:
            return ""
        
        digitos = re.sub(r'\D', '', str(cnpj))
        
        if len(digitos) != 14:
            aviso = f"CNPJ com tamanho incorreto: {len(digitos)} dígitos (esperado: 14)"
            self.avisos.append(aviso)
            print(f"[AVISO] {aviso}")
        
        return digitos
    
    def padronizar_cep(self, cep: str) -> str:
        """
        Padroniza CEP: formato XXXXX-XXX.
        
        Args:
            cep: CEP com ou sem formatação
            
        Returns:
            CEP padronizado
        """
        if not cep:
            return ""
        
        digitos = re.sub(r'\D', '', str(cep))
        
        if len(digitos) != 8:
            aviso = f"CEP com tamanho incorreto: {len(digitos)} dígitos (esperado: 8)"
            self.avisos.append(aviso)
            print(f"[AVISO] {aviso}")
            return digitos
        
        # Formato: XXXXX-XXX
        return f"{digitos[:5]}-{digitos[5:]}"
    
    def padronizar_telefone(self, telefone: str) -> str:
        """
        Padroniza telefone: formato (XX) XXXXX-XXXX ou (XX) XXXX-XXXX.
        
        Args:
            telefone: Telefone com ou sem formatação
            
        Returns:
            Telefone padronizado
        """
        if not telefone:
            return ""
        
        digitos = re.sub(r'\D', '', str(telefone))
        
        if len(digitos) < 10 or len(digitos) > 11:
            aviso = f"Telefone com tamanho incorreto: {len(digitos)} dígitos"
            self.avisos.append(aviso)
            print(f"[AVISO] {aviso}")
            return telefone
        
        # Celular com 9 dígitos: (XX) XXXXX-XXXX
        if len(digitos) == 11:
            return f"({digitos[:2]}) {digitos[2:7]}-{digitos[7:]}"
        
        # Fixo com 8 dígitos: (XX) XXXX-XXXX
        return f"({digitos[:2]}) {digitos[2:6]}-{digitos[6:]}"
    
    def padronizar_inscricao_estadual(self, inscricao: str, estado: str) -> str:
        """
        Padroniza Inscrição Estadual baseado no estado.
        
        Args:
            inscricao: Inscrição Estadual original
            estado: Sigla do estado (UF)
            
        Returns:
            Inscrição Estadual padronizada
        """
        if not estado:
            erro = "Estado não informado para padronizar IE"
            self.erros.append(erro)
            print(f"[ERRO] {erro}")
            return inscricao if inscricao else ""
        
        sucesso, ie_padronizada, erro_msg = PadronizadorInscricaoEstadual.padronizar(
            inscricao, estado, permitir_vazio=True
        )
        
        if not sucesso:
            if erro_msg:
                self.avisos.append(f"IE: {erro_msg}")
                print(f"[AVISO] IE de {estado}: {erro_msg}")
        else:
            if ie_padronizada and inscricao:
                print(f"[INFO] IE padronizada ({estado}): '{inscricao}' → '{ie_padronizada}'")
        
        return ie_padronizada
    
    def padronizar_inscricao_municipal(self, inscricao: str) -> str:
        """
        Padroniza Inscrição Municipal: apenas dígitos ou ISENTO.
        
        Args:
            inscricao: Inscrição Municipal original
            
        Returns:
            Inscrição Municipal padronizada
        """
        if not inscricao:
            return ""
        
        # Trata "ISENTO" ou similar
        if str(inscricao).upper() in ['ISENTO', 'ISENTA', 'DISPENSADO', 'DISPENSADA']:
            return "ISENTO"
        
        # Remove não-dígitos
        digitos = re.sub(r'\D', '', str(inscricao))
        
        return digitos
    
    def extrair_codigo_banco(self, banco: str) -> str:
        """
        Extrai código do banco do texto.
        
        Args:
            banco: String com formato "CODIGO - NOME" ou "CODIGO NOME"
            
        Returns:
            Código do banco (3 dígitos) ou vazio
        """
        if not banco:
            return ""
        
        banco_str = str(banco).strip()
        
        # Tenta padrão "XXX - NOME"
        match = re.match(r'^(\d{3})\s*-\s*', banco_str)
        if match:
            return match.group(1)
        
        # Tenta só dígitos iniciais "XXX"
        match = re.match(r'^(\d{3})', banco_str)
        if match:
            return match.group(1)
        
        return ""
    
    def padronizar_banco(self, banco: str) -> Tuple[str, str]:
        """
        Padroniza dados bancários: extrai código e mantém texto completo em MAIÚSCULAS.
        
        Args:
            banco: String com formato "CODIGO - NOME DO BANCO"
            
        Returns:
            Tupla (codigo, nome_banco_maiusculo)
        """
        if not banco:
            return "", ""
        
        banco_str = str(banco).strip().upper()
        codigo = self.extrair_codigo_banco(banco_str)
        
        if not codigo:
            aviso = f"Código do banco não identificado em: '{banco_str}'"
            self.avisos.append(aviso)
            print(f"[AVISO] {aviso}")
        
        return codigo, banco_str
    
    def padronizar_agencia(self, agencia: str) -> str:
        """
        Padroniza agência bancária: mantém dígitos, pontos e hífens.
        
        Args:
            agencia: Agência original
            
        Returns:
            Agência padronizada
        """
        if not agencia:
            return ""
        
        # Remove espaços mas mantém pontos e hífens
        agencia_str = str(agencia).strip()
        agencia_limpa = re.sub(r'\s+', '', agencia_str)
        
        # Valida formato
        if not re.match(r'^[\d\.\-]+$', agencia_limpa):
            aviso = f"Agência contém caracteres inválidos: '{agencia_str}'"
            self.avisos.append(aviso)
            print(f"[AVISO] {aviso}")
        
        return agencia_limpa
    
    def padronizar_conta_corrente(self, conta: str) -> str:
        """
        Padroniza conta corrente: mantém dígitos e hífen.
        
        Args:
            conta: Conta corrente original
            
        Returns:
            Conta corrente padronizada
        """
        if not conta:
            return ""
        
        # Remove espaços mas mantém hífen
        conta_str = str(conta).strip()
        conta_limpa = re.sub(r'\s+', '', conta_str)
        
        # Valida formato
        if not re.match(r'^[\d\-]+$', conta_limpa):
            aviso = f"Conta corrente contém caracteres inválidos: '{conta_str}'"
            self.avisos.append(aviso)
            print(f"[AVISO] {aviso}")
        
        return conta_limpa
    
    def padronizar_texto_maiuscula(self, texto: str, max_length: int = None) -> str:
        """
        Padroniza texto: converte para MAIÚSCULAS e remove espaços extras.
        
        Args:
            texto: Texto original
            max_length: Tamanho máximo permitido (None = sem limite)
            
        Returns:
            Texto padronizado em MAIÚSCULAS
        """
        if not texto:
            return ""
        
        # Converte para maiúsculas e remove espaços duplos
        texto_limpo = re.sub(r'\s+', ' ', str(texto).strip().upper())
        
        # Verifica tamanho máximo
        if max_length and len(texto_limpo) > max_length:
            aviso = f"Texto excede tamanho máximo ({len(texto_limpo)}/{max_length}): '{texto_limpo[:50]}...'"
            self.avisos.append(aviso)
            print(f"[AVISO] {aviso}")
        
        return texto_limpo
    
    def converter_prazo_pagamento(self, prazo: str) -> str:
        """
        Converte prazo de pagamento para código SAP.
        
        Exemplos:
            "60 dias" → "BRFM"
            "30" → "BRFG"
            "vista" → "BRFAA"
            "BRFM" → "BRFM" (já é código)
        
        Args:
            prazo: Prazo em texto ou número
            
        Returns:
            Código do prazo (BRFXX) em MAIÚSCULAS
        """
        if not prazo:
            return ""
        
        prazo_str = str(prazo).strip().upper()
        
        # Se já está no formato BRFXX, retorna
        if re.match(r'^BRF[A-Z]{1,2}$', prazo_str):
            print(f"[INFO] Prazo já em formato código: {prazo_str}")
            return prazo_str
        
        # Trata "VISTA" ou "À VISTA"
        if "VISTA" in prazo_str:
            print(f"[INFO] Prazo convertido: '{prazo}' → 'BRFAA' (VISTA)")
            return "BRFAA"
        
        # Extrai número de dias
        match = re.search(r'(\d+)', prazo_str)
        if match:
            dias = int(match.group(1))
            
            # Busca código exato
            if dias in self.MAPA_PRAZO_PAGAMENTO:
                codigo = self.MAPA_PRAZO_PAGAMENTO[dias]
                print(f"[INFO] Prazo convertido: '{prazo}' → '{codigo}' ({dias} dias)")
                return codigo
            else:
                aviso = f"Prazo de {dias} dias não tem código mapeado. Valores válidos: {sorted(self.MAPA_PRAZO_PAGAMENTO.keys())}"
                self.avisos.append(aviso)
                print(f"[AVISO] {aviso}")
                return prazo_str
        
        # Não conseguiu identificar
        aviso = f"Formato de prazo não reconhecido: '{prazo}'"
        self.avisos.append(aviso)
        print(f"[AVISO] {aviso}")
        return prazo_str
    
    def converter_estado(self, estado: str) -> str:
        """
        Converte estado por extenso para sigla.
        
        Exemplos:
            "Minas Gerais" → "MG"
            "São Paulo" → "SP"
            "MG" → "MG" (já é sigla)
        
        Args:
            estado: Estado por extenso ou sigla
            
        Returns:
            Sigla do estado (UF) em MAIÚSCULAS
        """
        if not estado:
            return ""
        
        estado_str = str(estado).strip().upper()
        
        # Remove acentos para comparação
        estado_normalizado = estado_str
        for acento, sem_acento in [
            ('Á', 'A'), ('À', 'A'), ('Ã', 'A'), ('Â', 'A'),
            ('É', 'E'), ('Ê', 'E'),
            ('Í', 'I'),
            ('Ó', 'O'), ('Ô', 'O'), ('Õ', 'O'),
            ('Ú', 'U'), ('Ü', 'U'),
            ('Ç', 'C')
        ]:
            estado_normalizado = estado_normalizado.replace(acento, sem_acento)
        
        # Se já é uma UF válida, retorna
        if estado_normalizado in self.UFS_VALIDAS:
            print(f"[INFO] Estado já em formato UF: {estado_normalizado}")
            return estado_normalizado
        
        # Busca no mapa
        if estado_normalizado in self.MAPA_ESTADOS:
            uf = self.MAPA_ESTADOS[estado_normalizado]
            print(f"[INFO] Estado convertido: '{estado}' → '{uf}'")
            return uf
        
        # Não encontrou
        aviso = f"Estado não reconhecido: '{estado}'. Esperado sigla (UF) ou nome por extenso"
        self.avisos.append(aviso)
        print(f"[AVISO] {aviso}")
        return estado_str
    
    def executar(self) -> Tuple[bool, str]:
        """
        Executa padronização completa dos dados.
        
        Returns:
            Tupla (sucesso, mensagem)
        """
        print("\n" + "="*70)
        print("INICIANDO PADRONIZAÇÃO DE DADOS (v3.0)")
        print("="*70)
        
        # Limpa listas de avisos/erros
        self.avisos.clear()
        self.erros.clear()
        
        # Carrega dados
        if not self.carregar_dados():
            return False, "Erro ao carregar dados"
        
        # Valida estrutura
        categorias_obrigatorias = ['empresa', 'endereco', 'contato', 'bancario', 'geral']
        for categoria in categorias_obrigatorias:
            if categoria not in self.dados:
                self.dados[categoria] = {}
        
        # === EMPRESA ===
        print("\n[PADRONIZAÇÃO] Empresa...")
        
        if 'razao_social' in self.dados['empresa']:
            self.dados['empresa']['razao_social'] = self.padronizar_texto_maiuscula(
                self.dados['empresa']['razao_social'], max_length=40
            )
        
        if 'nome_fantasia' in self.dados['empresa']:
            self.dados['empresa']['nome_fantasia'] = self.padronizar_texto_maiuscula(
                self.dados['empresa']['nome_fantasia'], max_length=20
            )
        
        if 'cnpj' in self.dados['empresa']:
            self.dados['empresa']['cnpj'] = self.padronizar_cnpj(
                self.dados['empresa']['cnpj']
            )
        
        if 'inscricao_municipal' in self.dados['empresa']:
            self.dados['empresa']['inscricao_municipal'] = self.padronizar_inscricao_municipal(
                self.dados['empresa']['inscricao_municipal']
            )
        
        # === ENDEREÇO ===
        print("[PADRONIZAÇÃO] Endereço...")
        
        # PRIMEIRO converte estado (por extenso → sigla)
        if 'estado' in self.dados['endereco']:
            estado_original = self.dados['endereco']['estado']
            self.dados['endereco']['estado'] = self.converter_estado(estado_original)
        
        # DEPOIS padroniza IE usando o estado já convertido
        if 'inscricao_estadual' in self.dados['empresa']:
            estado = self.dados['endereco'].get('estado', '')
            self.dados['empresa']['inscricao_estadual'] = self.padronizar_inscricao_estadual(
                self.dados['empresa']['inscricao_estadual'],
                estado
            )
        
        if 'rua' in self.dados['endereco']:
            self.dados['endereco']['rua'] = self.padronizar_texto_maiuscula(
                self.dados['endereco']['rua']
            )
        
        if 'complemento' in self.dados['endereco']:
            self.dados['endereco']['complemento'] = self.padronizar_texto_maiuscula(
                self.dados['endereco']['complemento']
            )
        
        if 'bairro' in self.dados['endereco']:
            self.dados['endereco']['bairro'] = self.padronizar_texto_maiuscula(
                self.dados['endereco']['bairro']
            )
        
        if 'cidade' in self.dados['endereco']:
            self.dados['endereco']['cidade'] = self.padronizar_texto_maiuscula(
                self.dados['endereco']['cidade']
            )
        
        if 'cep' in self.dados['endereco']:
            self.dados['endereco']['cep'] = self.padronizar_cep(
                self.dados['endereco']['cep']
            )
        
        # === CONTATO ===
        print("[PADRONIZAÇÃO] Contato...")
        
        if 'celular' in self.dados['contato']:
            self.dados['contato']['celular'] = self.padronizar_telefone(
                self.dados['contato']['celular']
            )
        
        if 'celular_secundario' in self.dados['contato']:
            self.dados['contato']['celular_secundario'] = self.padronizar_telefone(
                self.dados['contato']['celular_secundario']
            )
        
        if 'email_comercial' in self.dados['contato']:
            self.dados['contato']['email_comercial'] = str(
                self.dados['contato']['email_comercial']
            ).strip().lower()
        
        if 'email_fiscal' in self.dados['contato']:
            self.dados['contato']['email_fiscal'] = str(
                self.dados['contato']['email_fiscal']
            ).strip().lower()
        
        # === BANCÁRIO ===
        print("[PADRONIZAÇÃO] Dados bancários...")
        
        # PROTEÇÃO: Preserva codigo_banco se já existir
        codigo_banco_existente = self.dados['bancario'].get('codigo_banco', '').strip()
        
        if 'banco' in self.dados['bancario']:
            codigo_extraido, nome_completo = self.padronizar_banco(
                self.dados['bancario']['banco']
            )
            
            # Prioriza: codigo_banco existente > codigo_extraido > vazio
            if codigo_banco_existente:
                print(f"[INFO] Preservando código do banco existente: {codigo_banco_existente}")
                self.dados['bancario']['codigo_banco'] = codigo_banco_existente
            elif codigo_extraido:
                print(f"[INFO] Código do banco extraído: {codigo_extraido}")
                self.dados['bancario']['codigo_banco'] = codigo_extraido
            
            self.dados['bancario']['banco'] = nome_completo
        elif codigo_banco_existente:
            # Só tinha codigo_banco, sem campo banco
            print(f"[INFO] Mantendo código do banco: {codigo_banco_existente}")
            self.dados['bancario']['codigo_banco'] = codigo_banco_existente
        
        if 'agencia' in self.dados['bancario']:
            self.dados['bancario']['agencia'] = self.padronizar_agencia(
                self.dados['bancario']['agencia']
            )
        
        if 'conta_corrente' in self.dados['bancario']:
            self.dados['bancario']['conta_corrente'] = self.padronizar_conta_corrente(
                self.dados['bancario']['conta_corrente']
            )
        
        # === GERAL ===
        print("[PADRONIZAÇÃO] Informações gerais...")
        
        # Converte prazo de pagamento (dias → código)
        if 'prazo_pagamento' in self.dados['geral']:
            prazo_original = self.dados['geral']['prazo_pagamento']
            self.dados['geral']['prazo_pagamento'] = self.converter_prazo_pagamento(prazo_original)
        
        # Modalidade de frete em MAIÚSCULAS
        if 'modalidade_frete' in self.dados['geral']:
            self.dados['geral']['modalidade_frete'] = self.padronizar_texto_maiuscula(
                self.dados['geral']['modalidade_frete']
            )
        
        # Salva dados
        if not self.salvar_dados():
            return False, "Erro ao salvar dados padronizados"
        
        # Monta mensagem final
        print("\n" + "="*70)
        print("PADRONIZAÇÃO CONCLUÍDA")
        print("="*70)
        
        if self.avisos:
            print(f"\n⚠️  {len(self.avisos)} aviso(s):")
            for aviso in self.avisos:
                print(f"  • {aviso}")
        
        if self.erros:
            print(f"\n❌ {len(self.erros)} erro(s):")
            for erro in self.erros:
                print(f"  • {erro}")
            return False, f"Padronização com {len(self.erros)} erro(s)"
        
        mensagem = "Dados padronizados com sucesso"
        if self.avisos:
            mensagem += f" ({len(self.avisos)} aviso(s))"
        
        return True, mensagem


# Função de conveniência
def padronizar_dados_fornecedor(json_path: Path) -> Tuple[bool, str]:
    """
    Função de conveniência para padronizar dados de fornecedor.
    
    Args:
        json_path: Caminho para o JSON limpo
        
    Returns:
        Tupla (sucesso, mensagem)
    """
    padronizador = PadronizadorDados(json_path)
    return padronizador.executar()


# Teste
if __name__ == "__main__":
    from pathlib import Path
    
    # Testa com JSON de exemplo
    json_teste = Path(__file__).parent.parent / "Arquivos" / "fornecedor_limpo.json"
    
    if json_teste.exists():
        print(f"Testando padronização em: {json_teste}\n")
        sucesso, mensagem = padronizar_dados_fornecedor(json_teste)
        print(f"\nResultado: {mensagem}")
    else:
        print(f"Arquivo de teste não encontrado: {json_teste}")