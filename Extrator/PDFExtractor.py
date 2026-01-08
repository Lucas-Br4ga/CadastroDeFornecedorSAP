# -*- coding: utf-8 -*-
"""
Módulo de extração de dados de arquivos PDF.
Garante encoding UTF-8 correto em todas as operações.
"""
import json
import sys
import PyPDF2
import camelot
import tabula
from pathlib import Path

# Adiciona o diretório raiz ao path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils import get_json_paths


class PDFExtractor:
    """
    Classe responsável por extrair texto, tabelas e metadados de arquivos PDF.
    Utiliza múltiplas bibliotecas para garantir extração robusta.
    
    CORREÇÃO v2.0: Encoding UTF-8 garantido em todas as operações.
    """
    
    def __init__(self, pdf_path: str):
        self.pdf_path = Path(pdf_path)
        self.num_pages = 0
        
        if not self.pdf_path.exists():
            raise FileNotFoundError(f"Arquivo PDF não encontrado: {self.pdf_path}")

    def extract_text(self):
        """Extrai texto de todas as páginas usando PyPDF2 com encoding UTF-8"""
        text_by_page = {}

        with open(self.pdf_path, "rb") as file:
            pdf = PyPDF2.PdfReader(file)
            self.num_pages = len(pdf.pages)

            for i, page in enumerate(pdf.pages):
                try:
                    text = page.extract_text() or ""
                    # Garante que o texto é string UTF-8
                    if isinstance(text, bytes):
                        text = text.decode('utf-8', errors='replace')
                except Exception as e:
                    text = f"Erro ao extrair página {i+1}: {str(e)}"

                text_by_page[f"page_{i+1}"] = text

        return text_by_page

    def extract_tables_camelot(self):
        """Extrai tabelas usando Camelot (requer Ghostscript instalado)"""
        try:
            tables = camelot.read_pdf(str(self.pdf_path), pages="all", flavor="lattice")

            result = []
            for idx, table in enumerate(tables):
                df = table.df
                
                # Converte DataFrame para valores que garantem UTF-8
                rows_data = []
                for row in df.values.tolist():
                    row_cleaned = []
                    for cell in row:
                        if isinstance(cell, bytes):
                            cell = cell.decode('utf-8', errors='replace')
                        row_cleaned.append(str(cell))
                    rows_data.append(row_cleaned)
                
                result.append({
                    "table_index": idx,
                    "rows": rows_data,
                    "columns": [str(col) for col in df.columns.tolist()]
                })

            return result if result else [{"rows": [], "columns": []}]

        except Exception as e:
            return [{"error": str(e), "rows": [], "columns": []}]

    def extract_tables_tabula(self):
        """Extrai tabelas usando Tabula (requer Java instalado)"""
        try:
            dfs = tabula.read_pdf(str(self.pdf_path), pages="all", multiple_tables=True)

            result = []
            for idx, df in enumerate(dfs):
                df = df.astype(str)
                
                # Garante UTF-8 em todas as células
                rows_data = []
                for row in df.values.tolist():
                    row_cleaned = []
                    for cell in row:
                        if isinstance(cell, bytes):
                            cell = cell.decode('utf-8', errors='replace')
                        row_cleaned.append(str(cell))
                    rows_data.append(row_cleaned)
                
                result.append({
                    "table_index": idx,
                    "rows": rows_data,
                    "columns": [str(col) for col in list(df.columns)]
                })

            return result if result else [{"rows": [], "columns": []}]

        except Exception as e:
            return [{"error": str(e), "rows": [], "columns": []}]

    def extract_metadata(self):
        """Extrai metadados do PDF"""
        try:
            with open(self.pdf_path, "rb") as file:
                reader = PyPDF2.PdfReader(file)
                metadata = reader.metadata

                if not metadata:
                    return {}

                # Converte metadados para UTF-8
                metadata_clean = {}
                for k, v in metadata.items():
                    key = str(k).lstrip("/")
                    value = str(v) if v is not None else ""
                    
                    # Garante UTF-8
                    if isinstance(value, bytes):
                        value = value.decode('utf-8', errors='replace')
                    
                    metadata_clean[key] = value
                
                return metadata_clean
                
        except Exception as e:
            return {"error": str(e)}

    def extract_all(self):
        """Executa todas as extrações e retorna um dicionário completo"""
        return {
            "file": self.pdf_path.name,
            "metadata": self.extract_metadata(),
            "text": self.extract_text(),
            "tables": {
                "camelot": self.extract_tables_camelot(),
                "tabula": self.extract_tables_tabula()
            },
            "pages_total": self.num_pages
        }

    def save_to_json(self, output_path: Path = None):
        """
        Salva os dados extraídos em JSON com encoding UTF-8.
        Se output_path não for fornecido, salva em Arquivos/fornecedor_bruto.json
        """
        if output_path is None:
            # Busca o diretório raiz do projeto (onde está main.py)
            current = Path(__file__).resolve().parent
            while current.parent != current:
                if (current / "main.py").exists():
                    break
                current = current.parent
            output_path = current / "Arquivos" / "fornecedor_bruto.json"
        
        # Garante que o diretório existe
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        dados = self.extract_all()

        # Salva com encoding UTF-8 e ensure_ascii=False
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(dados, f, indent=4, ensure_ascii=False)

        print(f"[INFO] Dados extraídos salvos com UTF-8 em: {output_path}")

        return output_path