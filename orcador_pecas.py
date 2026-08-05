import os
import sys
import pandas as pd
import numpy as np
from dataclasses import dataclass
from typing import List, Dict, Optional


@dataclass
class ItemOrcamento:
    codigo: str
    descricao: str
    ncm: str
    quantidade: int
    preco_base: float
    ipi_percent: float
    canal: str  # 'REVENDA' ou 'BALCAO'


@dataclass
class ResultadoCalculoItem:
    item: ItemOrcamento
    valor_base_total: float
    valor_ipi: float
    frete_rateado: float
    base_icms_proprio: float
    valor_icms_proprio: float
    base_icms_st: float
    valor_icms_st: float
    preco_final_unitario: float
    preco_final_total: float


class TaxEngine:
    """Motor tributário para cálculo do ICMS-ST, IPI e Frete por Estado (UF)."""

    def __init__(self, df_matriz_fiscal: pd.DataFrame):
        # Mapeia a tabela da Sheet2
        self.matriz_fiscal = df_matriz_fiscal.set_index("UF Destino").to_dict(
            orient="index"
        )

    def calcular_item(
        self,
        item: ItemOrcamento,
        uf_destino: str,
        frete_unitario: float = 0.0,
    ) -> ResultadoCalculoItem:
        uf_destino = uf_destino.upper().strip()
        regra_uf = self.matriz_fiscal.get(
            uf_destino,
            {"Alíquota Interna": 18.0, "Alíquota Interestadual": 12.0, "MVA": 71.78},
        )

        aliq_interna = float(regra_uf.get("Alíquota Interna", 18.0)) / 100.0
        aliq_interestadual = (
            float(regra_uf.get("Alíquota Interestadual", 12.0)) / 100.0
        )
        mva = float(regra_uf.get("MVA", 71.78)) / 100.0

        # 1. Frete por item (calculado primeiro para compor a base)
        v_frete_total_item = frete_unitario * item.quantidade

        # 2. Valores Base e IPI (o frete agora integra a base de cálculo do IPI)
        v_base_unit = item.preco_base
        v_base_total = v_base_unit * item.quantidade
        
        base_ipi = v_base_total + v_frete_total_item
        v_ipi_unit = (base_ipi / item.quantidade) * (item.ipi_percent / 100.0)
        v_ipi_total = base_ipi * (item.ipi_percent / 100.0)

        print(f"PEÇA: {item.codigo} | Base: {v_base_total} | Frete: {v_frete_total_item} | Base IPI: {base_ipi} | IPI Gerado: {v_ipi_total}")

        # 3. ICMS Próprio (Operação Própria)
        base_icms_proprio = v_base_total + v_frete_total_item
        v_icms_proprio = base_icms_proprio * aliq_interestadual

        # 4. ICMS Substituição Tributária (ICMS-ST)
        base_icms_st = (
            v_base_total + v_ipi_total + v_frete_total_item
        ) * (1.0 + mva)
        v_icms_st_bruto = base_icms_st * aliq_interna
        v_icms_st_liquido = max(0.0, v_icms_st_bruto - v_icms_proprio)

        # 5. Preço Final
        preco_final_total = (
            v_base_total + v_ipi_total + v_icms_st_liquido + v_frete_total_item
        )
        preco_final_unitario = preco_final_total / item.quantidade

        return ResultadoCalculoItem(
            item=item,
            valor_base_total=round(v_base_total, 2),
            valor_ipi=round(v_ipi_total, 2),
            frete_rateado=round(v_frete_total_item, 2),
            base_icms_proprio=round(base_icms_proprio, 2),
            valor_icms_proprio=round(v_icms_proprio, 2),
            base_icms_st=round(base_icms_st, 2),
            valor_icms_st=round(v_icms_st_liquido, 2),
            preco_final_unitario=round(preco_final_unitario, 2),
            preco_final_total=round(preco_final_total, 2),
        )


class CatalogManager:
    """Carrega e gerencia a busca rápida nas 10.901 peças do catálogo."""

    def __init__(self, excel_path: str = "LISTA DE PREÇO PEÇAS_050526.xlsx"):
        # Garante a resolução do caminho absoluto seguro antes de carregar
        self.excel_path = self._resolver_caminho(excel_path)
        self.df_pecas = pd.DataFrame()
        self.df_fiscal = pd.DataFrame()
        self._carregar_dados()

    def _resolver_caminho(self, caminho_fornecido: str) -> str:
        """Localiza o arquivo Excel em caminhos absolutos no Windows ou na pasta _internal."""
        if os.path.isabs(caminho_fornecido) and os.path.exists(caminho_fornecido):
            return caminho_fornecido

        nome_arquivo = os.path.basename(caminho_fornecido)
        base_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.getcwd()

        # 1. Busca na pasta de instalação (ex: C:\Sistema Effa\LISTA...)
        caminho_raiz = os.path.join(base_dir, nome_arquivo)
        if os.path.exists(caminho_raiz):
            return os.path.abspath(caminho_raiz)

        # 2. Busca na pasta interna do PyInstaller (_internal)
        pasta_internal = os.path.dirname(os.path.abspath(__file__))
        caminho_internal = os.path.join(pasta_internal, nome_arquivo)
        if os.path.exists(caminho_internal):
            return os.path.abspath(caminho_internal)

        # 3. Busca no diretório de trabalho atual
        caminho_cwd = os.path.join(os.getcwd(), nome_arquivo)
        if os.path.exists(caminho_cwd):
            return os.path.abspath(caminho_cwd)

        return caminho_raiz

    def _carregar_dados(self):
        # Carrega a tabela de preços principal (linha 6 contém os cabeçalhos)
        self.df_pecas = pd.read_excel(
            self.excel_path, sheet_name="TABELA DE PREÇOS", skiprows=5
        )

        # Mapeamento dinâmico e seguro por posição de coluna caso o cabeçalho esteja vazio
        cols = list(self.df_pecas.columns)

        col_codigo = cols[2] if len(cols) > 2 else "Código"
        col_nome = cols[4] if len(cols) > 4 else "Nome"
        col_ncm = cols[75] if len(cols) > 75 else "NCM"
        col_ipi = cols[86] if len(cols) > 86 else "IPI %"
        col_revenda = cols[87] if len(cols) > 87 else "Preço Peças Concessionária 2015"
        col_balcao = cols[89] if len(cols) > 89 else "Preço Peças Balcão"

        col_map = {
            col_codigo: "codigo",
            col_nome: "nome",
            col_ncm: "ncm",
            col_ipi: "ipi_percent",
            col_revenda: "preco_revenda",
            col_balcao: "preco_balcao",
        }

        self.df_pecas.rename(columns=col_map, inplace=True)

        # Trata valores nulos e converte tipos
        self.df_pecas["codigo"] = (
            self.df_pecas["codigo"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True)
        )
        self.df_pecas["nome"] = self.df_pecas["nome"].fillna("").astype(str)
        self.df_pecas["ncm"] = self.df_pecas["ncm"].fillna("").astype(str)
        self.df_pecas["ipi_percent"] = (
            pd.to_numeric(self.df_pecas["ipi_percent"], errors="coerce")
            .fillna(0.0)
        )
        self.df_pecas["preco_revenda"] = (
            pd.to_numeric(self.df_pecas["preco_revenda"], errors="coerce")
            .fillna(0.0)
        )
        self.df_pecas["preco_balcao"] = (
            pd.to_numeric(self.df_pecas["preco_balcao"], errors="coerce")
            .fillna(0.0)
        )

        # Carrega a Sheet2 (Matriz Fiscal por UF)
        df_sheet2 = pd.read_excel(
            self.excel_path, sheet_name="Sheet2", skiprows=1
        )
        df_sheet2.columns = [str(c).strip() for c in df_sheet2.columns]
        self.df_fiscal = df_sheet2.dropna(subset=["UF Destino"]).copy()

    def buscar_pecas(self, termo: str, limite: int = 20) -> pd.DataFrame:
        if not termo or len(termo.strip()) < 2:
            return self.df_pecas.head(limite)

        termo_lower = termo.lower().strip()
        mask = (
            self.df_pecas["codigo"].str.lower().str.contains(termo_lower)
            | self.df_pecas["nome"].str.lower().str.contains(termo_lower)
            | self.df_pecas["ncm"].str.lower().str.contains(termo_lower)
        )
        return self.df_pecas[mask].head(limite)