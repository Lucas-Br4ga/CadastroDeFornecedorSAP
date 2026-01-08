# -*- coding: utf-8 -*-
"""
Módulo de limpeza e estruturação de dados extraídos de PDF.
Garante encoding UTF-8 correto em todas as operações.
"""
import re
import json
from pathlib import Path


class PDFCompanyExtractor:
    """
    Extrai informações estruturadas (limpas) a partir do JSON bruto gerado pelo PDFExtractor.
    Atualiza automaticamente o JSON limpo na pasta /Arquivos.
    
    CORREÇÃO v2.0: Encoding UTF-8 garantido em todas as operações.
    """

    ALIASES = {
        "razao_social": [
            "Razão social", "Razao social", "Razão social/Jurídica",
            "Razão social/Jurídica ou Física", "Razao juridica"
        ],
        "nome_fantasia": ["Fantasia", "Nome fantasia", "Nome Fantasia"],
        "cnpj": ["CNPJ", "CNPJ/CPF", "cnpj"],
        "inscricao_estadual": ["Inscrição Estadual", "IE", "Insc. Estadual"],
        "inscricao_municipal": ["Inscrição Municipal", "IM", "Insc. Municipal"],
        "rua": ["Rua", "Endereço", "Rua/nº", "Logradouro"],
        "bairro": ["Bairro"],
        "cidade": ["Cidade"],
        "estado": ["Estado", "UF"],
        "cep": ["CEP", "Código Postal"],
        "celular": ["Telefone", "Celular", "Telefone/Celular", "Telefone Principal"],
        "celular_secundario": ["Telefone Secundário", "Celular Secundário", "Telefone 2", "Celular 2"],
        "banco": ["Banco", "Banco (Nome e nº)", "Instituição"],
        "codigo_banco": ["Código Banco", "Cód. Banco", "Código do Banco", "Cod Banco"],
        "agencia": ["Agência", "Agência (Nome e nº)", "Agencia"],
        "conta": ["Conta", "Conta Corrente", "Nº de Conta Corrente", "Numero Conta"],
        "prazo": ["Prazo", "Prazo de pagamento"],
        "modalidade_frete": ["Modalidade de Frete", "Frete", "Tipo de Frete", "Modalidade Frete"],
        "email_comercial": ["E-mail (Comercial)", "Email Comercial"],
        "email_fiscal": ["E-mail (Fiscal)", "Email Fiscal"]
    }

    @staticmethod
    def normalize(text):
        """Normaliza texto removendo quebras de linha e espaços múltiplos"""
        if not text:
            return ""
        text = str(text).replace("\n", " ")
        text = re.sub(r"\s{2,}", " ", text)
        return text.strip()

    @staticmethod
    def match_key(key, aliases):
        """Verifica se a chave corresponde a algum alias"""
        key_norm = PDFCompanyExtractor.normalize(key).lower()
        return any(alias.lower() in key_norm for alias in aliases)

    @staticmethod
    def extract_emails(text):
        """Extrai emails do texto usando regex"""
        return re.findall(r"[\w\.-]+@[\w\.-]+\.\w+", text)
    
    @staticmethod
    def extract_codigo_banco(banco_text):
        """
        Extrai código do banco do texto.
        Exemplos: "082 - Villela", "001 Banco do Brasil"
        """
        if not banco_text:
            return ""
        
        # Procura por padrão: 3 dígitos no início
        match = re.match(r'^(\d{3})', str(banco_text).strip())
        if match:
            return match.group(1)
        
        return ""

    @staticmethod
    def deep_update(original: dict, new: dict):
        """Faz merge profundo entre dicionários"""
        for k, v in new.items():
            if isinstance(v, dict) and k in original and isinstance(original[k], dict):
                original[k] = PDFCompanyExtractor.deep_update(original[k], v)
            else:
                original[k] = v
        return original

    def __init__(self, pdf_json_bruto: dict, json_limpo_path: Path):
        self.pdf_json = pdf_json_bruto
        self.json_limpo_path = json_limpo_path

        self.full_text = ""
        self.table_dict = {}

    def _build_full_text(self):
        """Constrói texto completo de todas as páginas"""
        for _, txt in self.pdf_json.get("text", {}).items():
            self.full_text += "\n" + str(txt)

    def _build_table_dict(self):
        """Constrói dicionário chave-valor a partir das tabelas"""
        try:
            table = self.pdf_json["tables"]["camelot"][0]["rows"]
        except (KeyError, IndexError, TypeError):
            table = []

        for row in table:
            if len(row) < 2:
                continue

            key = self.normalize(row[0])
            val = self.normalize(row[1])
            
            if key:  # Ignora linhas vazias
                self.table_dict[key] = val

    def _find_value(self, alias_list):
        """Busca valor na tabela usando lista de aliases"""
        for k, v in self.table_dict.items():
            if self.match_key(k, alias_list):
                if v:
                    return v

                # Fallback: email escondido na chave
                emails = self.extract_emails(k)
                if emails:
                    return emails[0]
        return ""

    def extract(self):
        """
        Método principal que extrai e estrutura todos os dados.
        Retorna dicionário limpo e salva em arquivo JSON com encoding UTF-8.
        """
        self._build_full_text()
        self._build_table_dict()

        # Debug: mostra as chaves encontradas na tabela
        print(f"\n[DEBUG] Chaves encontradas na tabela: {list(self.table_dict.keys())[:10]}")

        # Função auxiliar para buscar valores
        get = lambda name: self._find_value(self.ALIASES[name])

        # Extração de dados da empresa
        razao = get("razao_social")
        fantasia = get("nome_fantasia")
        cnpj = get("cnpj")
        ie = get("inscricao_estadual")
        im = get("inscricao_municipal")

        # Extração de endereço
        rua_completa = get("rua")
        bairro = get("bairro")
        cidade = get("cidade")
        estado = get("estado")
        cep = get("cep")

        # Extração de contatos
        celular = get("celular")
        celular_secundario = get("celular_secundario")

        # Extração de dados bancários
        banco = get("banco")
        codigo_banco_texto = get("codigo_banco")
        agencia = get("agencia")
        conta = get("conta")
        
        # Tenta extrair código do banco se não encontrou diretamente
        if not codigo_banco_texto and banco:
            codigo_banco_texto = self.extract_codigo_banco(banco)

        # Extração de informações gerais
        prazo = get("prazo")
        modalidade_frete = get("modalidade_frete")

        # Extração de emails
        email_comercial = get("email_comercial")
        email_fiscal = get("email_fiscal")

        # Parse de endereço (rua, número, complemento)
        rua = numero = complemento = ""
        m = re.match(r"(.+?)[,\s]+(\d+)(.*)", rua_completa)
        if m:
            rua = self.normalize(m.group(1))
            numero = self.normalize(m.group(2))
            complemento = self.normalize(m.group(3))
        else:
            rua = rua_completa

        # Fallback para emails no texto completo
        emails_full = self.extract_emails(self.full_text)

        if not email_comercial and emails_full:
            email_comercial = emails_full[0]

        if not email_fiscal and len(emails_full) > 1:
            email_fiscal = emails_full[1]

        # Estrutura JSON limpa
        cleaned_json = {
            "empresa": {
                "razao_social": razao,
                "nome_fantasia": fantasia,
                "cnpj": cnpj,
                "inscricao_estadual": ie,
                "inscricao_municipal": im
            },
            "endereco": {
                "rua": rua,
                "numero": numero,
                "complemento": complemento,
                "bairro": bairro,
                "cidade": cidade,
                "estado": estado,
                "cep": cep
            },
            "contato": {
                "celular": celular,
                "celular_secundario": celular_secundario,
                "email_comercial": email_comercial,
                "email_fiscal": email_fiscal
            },
            "bancario": {
                "banco": banco,
                "codigo_banco": codigo_banco_texto,
                "agencia": agencia,
                "conta_corrente": conta
            },
            "geral": {
                "prazo_pagamento": prazo,
                "modalidade_frete": modalidade_frete
            }
        }

        print(f"[DEBUG] Estrutura cleaned_json criada: {list(cleaned_json.keys())}")

        # Merge com dados existentes (se houver)
        if self.json_limpo_path.exists():
            try:
                with open(self.json_limpo_path, "r", encoding="utf-8") as f:
                    old_data = json.load(f)
                print(f"[DEBUG] JSON existente carregado. Chaves: {list(old_data.keys())}")
            except Exception as e:
                print(f"[DEBUG] Erro ao carregar JSON existente: {e}")
                old_data = {}

            merged = self.deep_update(old_data, cleaned_json)
        else:
            merged = cleaned_json

        print(f"[DEBUG] Merged final. Tipo: {type(merged)}, Chaves: {list(merged.keys())}")

        # Garante que o diretório existe
        self.json_limpo_path.parent.mkdir(parents=True, exist_ok=True)

        # Salva JSON limpo com encoding UTF-8
        with open(self.json_limpo_path, "w", encoding="utf-8") as f:
            json.dump(merged, f, ensure_ascii=False, indent=4)

        print(f"[INFO] JSON salvo com UTF-8 em: {self.json_limpo_path}")

        return merged