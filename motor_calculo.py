# -*- coding: utf-8 -*-
"""
Motor de Cálculo - Demonstrativo de Faturamento (DSV)
=======================================================
Reimplementação em Python da lógica original da planilha
'Demonstrativo_-_Nouryon.xlsm', servindo como camada de validação
antes da migração para a aplicação web (HTML/JS) hospedada no GitHub Pages.

Fontes mapeadas na planilha original:
- Aba "Capa Faturamento": fórmulas XLOOKUP/VLOOKUP + FILTER (Microsoft 365)
- Aba "Relatório FollowNet": dados do processo/embarque
- Aba "Tabela Clientes": cadastro de clientes (chave: CNPJ, coluna F)
- Aba "CW1": exportação bruta do sistema operacional (tabela "Table1")
- Aba "Tabela de Despesas": de-para de categorias de despesa
- Aba "Custeio Mult": conferência cruzada (validação de totais)
- VBA real do arquivo (Module1-4): Mult(), LimparEFormatarCNPJ_FollowNet(),
  Exportar_PDF_Demonstrativo(), SalvarCapaFaturamentoDownloads()

Ajustes solicitados pelo cliente (implementados abaixo):
1. Migração de dados só ocorre com 'Ref. Cliente' preenchida.
2. Campo "Gerenciamento PO" tornou-se editável (valor BRUTO,
   padrão R$ 327,60, pois o cliente definiu esse valor como bruto).
3. "Compensation Fee" = 1,5% sobre o total de despesas pagas pela
   Comissária (DSV), calculado automaticamente (mas pode ser sobrescrito).
4. Frete internacional faturado pela DSV: o valor lançado deve ser o
   LÍQUIDO (sem impostos) — sinalizado via aviso na categoria
   "Frete Intl. Pago".
5. Geração dos valores do processo a partir da CW1: replica a lógica
   FILTER(Table1[...], Debtor=Vendor, Job Num.=Processo, ...).
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Optional


# ---------------------------------------------------------------------------
# Constantes fiscais / regras de negócio (idênticas às fórmulas da planilha)
# ---------------------------------------------------------------------------
FATOR_GROSS_UP = 0.8575          # Bruto = Líquido / 0.8575  (Líquido = Bruto * 0.8575)
ALIQ_PIS = 0.0065                # 0,65%
ALIQ_COFINS = 0.03               # 3%
ALIQ_CSLL = 0.01                 # 1%
ALIQ_IRRF = 0.015                # 1,5%
ALIQ_COMPENSATION_FEE = 0.015    # 1,5% sobre despesas pagas pela Comissária
VALOR_PO_MANAGEMENT_PADRAO = 327.60  # valor BRUTO padrão (editável)


# ---------------------------------------------------------------------------
# Estruturas de dados
# ---------------------------------------------------------------------------
@dataclass
class LinhaDespesa:
    descricao: str
    categoria: str = ""          # via Tabela de Despesas (XLOOKUP)
    pago_cliente: float = 0.0    # coluna J/K "Pagas pelo Cliente (1)"
    pago_comissaria: float = 0.0  # coluna M/N "Pagas pela Comissária (2)"
    origem: str = "manual"       # 'manual' | 'follownet' | 'cw1'


@dataclass
class ServicoTributavel:
    descricao: str
    valor_liquido: float = 0.0
    editavel: bool = True

    @property
    def valor_bruto(self) -> float:
        return round(self.valor_liquido / FATOR_GROSS_UP, 2) if self.valor_liquido else 0.0


@dataclass
class Demonstrativo:
    ref_cliente: str = ""
    dados_processo: dict = field(default_factory=dict)
    dados_cliente: dict = field(default_factory=dict)
    despesas: list[LinhaDespesa] = field(default_factory=list)
    servicos: list[ServicoTributavel] = field(default_factory=list)
    adiantamentos: list[float] = field(default_factory=list)


# ---------------------------------------------------------------------------
# 1) Migração de dados para a capa (regra do cliente #1)
# ---------------------------------------------------------------------------
def buscar_dados_processo(ref_cliente: str, follownet_rows: list[dict]) -> Optional[dict]:
    """Equivalente a VLOOKUP($F$22, 'Relatório FollowNet'!$A:$AF, MATCH(campo,...))

    Sem 'Ref. Cliente' preenchida, não migra nada — regra explícita do cliente.
    """
    if not ref_cliente or not ref_cliente.strip():
        return None
    for row in follownet_rows:
        if str(row.get("Ref. Cliente", "")).strip() == ref_cliente.strip():
            return row
    return None


def buscar_dados_cliente(cnpj: str, clientes_rows: list[dict]) -> Optional[dict]:
    """Equivalente a XLOOKUP(F16, 'Tabela Clientes'!F:F, 'Tabela Clientes'!B:B / C:C / ...).

    A chave de busca é o CNPJ (coluna F da Tabela Clientes), obtido do
    Relatório FollowNet (coluna 'CNPJ').
    """
    if not cnpj:
        return None
    cnpj_limpo = "".join(ch for ch in str(cnpj) if ch.isdigit())
    for row in clientes_rows:
        cnpj_tabela = "".join(ch for ch in str(row.get("CNPJ", "")) if ch.isdigit())
        if cnpj_tabela and cnpj_tabela == cnpj_limpo:
            return row
    return None


# ---------------------------------------------------------------------------
# 2) Categorização de despesas (Tabela de Despesas -> XLOOKUP)
# ---------------------------------------------------------------------------
def categorizar_despesa(descricao: str, tabela_despesas: list[dict]) -> str:
    """Equivalente a XLOOKUP(E, 'Tabela de Despesas'!B:B, 'Tabela de Despesas'!A:A, "Verificar")."""
    if not descricao:
        return ""
    desc_norm = descricao.strip().lower()
    for row in tabela_despesas:
        if str(row.get("Descrição da despesa no CW1/FN", "")).strip().lower() == desc_norm:
            return row.get("Tipo de Despesa", "Verificar")
    return "Verificar"


# ---------------------------------------------------------------------------
# 3) Geração dos valores do processo a partir da CW1 (ajuste solicitado #5)
#    Replica as fórmulas FILTER() reais da planilha:
#      E48 = FILTER(Table1[Description],
#              (Debtor=Vendor)*(Posted-Revenue="Y")*(Job Num.=Processo)*(Inv.Type<>"DES"))
#      N48 = FILTER(Table1[OS Sell Amt], <mesmos critérios>)
#      N75 = FILTER(Table1[Estimated Revenue],
#              (Debtor=Vendor)*(Estimated Revenue<0)*(Job Num.=Processo)*(Inv.Type="NON"))
# ---------------------------------------------------------------------------
def importar_despesas_cw1(vendor: str, num_processo: str, cw1_rows: list[dict]) -> list[LinhaDespesa]:
    """Despesas pagas pela Comissária (DSV), geradas automaticamente a partir da CW1."""
    resultado = []
    for row in cw1_rows:
        if (str(row.get("Debtor", "")).strip() == str(vendor).strip()
                and str(row.get("Job Num.", "")).strip() == str(num_processo).strip()
                and str(row.get("Posted - Revenue", "")).strip().upper() == "Y"
                and str(row.get("Inv. Type", "")).strip().upper() != "DES"):
            resultado.append(LinhaDespesa(
                descricao=str(row.get("Description", "")).strip(),
                pago_comissaria=float(row.get("OS Sell Amt", 0) or 0),
                origem="cw1",
            ))
    return resultado


def importar_adiantamentos_cw1(vendor: str, num_processo: str, cw1_rows: list[dict]) -> list[float]:
    """Adiantamentos (Estimated Revenue negativo, Inv. Type = NON)."""
    valores = []
    for row in cw1_rows:
        est_rev = row.get("Estimated Revenue", None)
        try:
            est_rev = float(est_rev)
        except (TypeError, ValueError):
            continue
        if (str(row.get("Debtor", "")).strip() == str(vendor).strip()
                and str(row.get("Job Num.", "")).strip() == str(num_processo).strip()
                and est_rev < 0
                and str(row.get("Inv. Type", "")).strip().upper() == "NON"):
            valores.append(est_rev)
    return valores


# ---------------------------------------------------------------------------
# 4) Serviços tributáveis (gross-up 0,8575) + ajustes #2 e #3
# ---------------------------------------------------------------------------
def montar_servicos_tributaveis(
    valor_despachante_liquido: float,
    po_management_bruto: float,
    total_despesas_pagas_comissaria: float,
    compensation_fee_manual: Optional[float] = None,
) -> list[ServicoTributavel]:
    """
    - Despachante: valor líquido informado manualmente (mantém comportamento original).
    - Gerenciamento PO (ajuste #2): agora o usuário informa o valor BRUTO
      (padrão R$ 327,60); o líquido é derivado por *0,8575.
    - Compensation Fee (ajuste #3): 1,5% sobre o total de despesas pagas
      pela Comissária, calculado automaticamente. Pode ser sobrescrito.
    """
    servicos = []

    if valor_despachante_liquido:
        servicos.append(ServicoTributavel(
            "Serviços Profissionais de Despachante", valor_despachante_liquido))

    po_liquido = round(po_management_bruto * FATOR_GROSS_UP, 2) if po_management_bruto else 0.0
    servicos.append(ServicoTributavel("Gerenciamento PO", po_liquido))

    if compensation_fee_manual is not None:
        fee_liquido = compensation_fee_manual
    else:
        fee_bruto = round(total_despesas_pagas_comissaria * ALIQ_COMPENSATION_FEE, 2)
        fee_liquido = round(fee_bruto * FATOR_GROSS_UP, 2)
    servicos.append(ServicoTributavel("Compensation Fee", fee_liquido))

    return servicos


# ---------------------------------------------------------------------------
# 5) Totais e tributos (idêntico às fórmulas N73:N82 da planilha)
# ---------------------------------------------------------------------------
def calcular_totais(demo: Demonstrativo) -> dict:
    total_pago_cliente = sum(d.pago_cliente for d in demo.despesas)
    total_pago_comissaria = sum(d.pago_comissaria for d in demo.despesas)
    total_despesas_pagas = total_pago_cliente + total_pago_comissaria  # N61

    soma_brutos_servicos = sum(s.valor_bruto for s in demo.servicos)   # SUM(N64:N71)
    # PIS/COFINS/CSLL incidem apenas sobre Despachante + Gerenciamento PO
    # (replica SUM($N$64:$N$65) da planilha original)
    base_pis_cofins_csll = sum(
        s.valor_bruto for s in demo.servicos[:2]
    ) if len(demo.servicos) >= 2 else soma_brutos_servicos

    pis = round(base_pis_cofins_csll * ALIQ_PIS, 2)
    cofins = round(base_pis_cofins_csll * ALIQ_COFINS, 2)
    csll = round(base_pis_cofins_csll * ALIQ_CSLL, 2)
    irrf = round(soma_brutos_servicos * ALIQ_IRRF, 2)
    total_tributos = pis + cofins + csll + irrf

    total_servicos_liquido_tributos = round(soma_brutos_servicos - total_tributos, 2)  # N73
    total_a_pagar = round(total_servicos_liquido_tributos + total_despesas_pagas, 2)    # N76
    total_adiantamentos = round(sum(demo.adiantamentos), 2)                            # N75
    saldo_adiantamento = round(total_a_pagar + total_adiantamentos, 2)                 # N77

    return {
        "total_pago_cliente": total_pago_cliente,
        "total_pago_comissaria": total_pago_comissaria,
        "total_despesas_pagas": total_despesas_pagas,
        "soma_brutos_servicos": soma_brutos_servicos,
        "pis": pis,
        "cofins": cofins,
        "csll": csll,
        "irrf": irrf,
        "total_tributos": total_tributos,
        "total_servicos_liquido_tributos": total_servicos_liquido_tributos,
        "total_adiantamentos": total_adiantamentos,
        "total_a_pagar": total_a_pagar,
        "saldo_adiantamento": saldo_adiantamento,
    }


# ---------------------------------------------------------------------------
# Validação com os dados reais extraídos do arquivo original
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    follownet_rows = [{
        "Ref. Cliente": "5103596062_4515055273",
        "Nr. Processo": "BCPQ0154213",
        "House": "271667321",
        "Importador": "Agro Bayer S.R.L.",
        "Via de Transporte": "MARITIMA",
        "Incoterm": "FCA",
        "Navio": "MAERSK GANGES 624S",
        "Agente": "HOYER GLOBAL (USA) INC",
        "Nr. DUE": "26BR001069863-7",
        "CNPJ": "43.818.418/0004-66",
        "CIF R$": 5113503.26,
        "FOB USD": 989286.48,
    }]

    clientes_rows = [{
        "Customer": 6402117463,
        "Name": "TESTE CLIENTE LTDA",
        "Adress": "RUA TESTE, 123",
        "District": "CENTRO",
        "City": "SAO PAULO - SP",
        "CNPJ": "43.818.418/0004-66",
    }]

    tabela_despesas = [
        {"Tipo de Despesa": "Frete Intl. Pago", "Descrição da despesa no CW1/FN": "Frete Internacional"},
        {"Tipo de Despesa": "Desconsolidação", "Descrição da despesa no CW1/FN": "Desconsolidação"},
    ]

    cw1_rows = [
        {"Debtor": "6402117463", "Job Num.": "BCPQ0154213", "Posted - Revenue": "Y",
         "Inv. Type": "REV", "Description": "Desconsolidação", "OS Sell Amt": 850.00},
        {"Debtor": "6402117463", "Job Num.": "BCPQ0154213", "Posted - Revenue": "N",
         "Inv. Type": "NON", "Description": "Adiantamento Numerário",
         "Estimated Revenue": -1200.00},
    ]

    # --- Teste 1: sem Ref. Cliente não migra nada ---
    assert buscar_dados_processo("", follownet_rows) is None
    print("[OK] Sem Ref. Cliente -> não migra dados (regra #1)")

    # --- Teste 2: com Ref. Cliente, migra corretamente ---
    processo = buscar_dados_processo("5103596062_4515055273", follownet_rows)
    assert processo is not None
    assert processo["Nr. Processo"] == "BCPQ0154213"
    print("[OK] Dados do processo migrados:", processo["Nr. Processo"], processo["Importador"])

    cliente = buscar_dados_cliente(processo["CNPJ"], clientes_rows)
    assert cliente is not None
    print("[OK] Cliente localizado via CNPJ:", cliente["Name"])

    # --- Teste 3: categorização de despesa ---
    cat = categorizar_despesa("Frete Internacional", tabela_despesas)
    assert cat == "Frete Intl. Pago"
    print("[OK] Categorização de despesa:", cat, "(lembrar: lançar valor líquido)")

    # --- Teste 4: importação CW1 (despesas + adiantamentos) ---
    despesas_cw1 = importar_despesas_cw1("6402117463", "BCPQ0154213", cw1_rows)
    assert len(despesas_cw1) == 1 and despesas_cw1[0].pago_comissaria == 850.00
    print("[OK] Despesas importadas da CW1:", despesas_cw1[0].descricao, despesas_cw1[0].pago_comissaria)

    adiant = importar_adiantamentos_cw1("6402117463", "BCPQ0154213", cw1_rows)
    assert adiant == [-1200.00]
    print("[OK] Adiantamentos importados da CW1:", adiant)

    # --- Teste 5: serviços tributáveis (PO Management editável + Compensation Fee auto) ---
    demo = Demonstrativo(ref_cliente="5103596062_4515055273")
    demo.despesas = despesas_cw1
    total_comissaria = sum(d.pago_comissaria for d in demo.despesas)
    demo.servicos = montar_servicos_tributaveis(
        valor_despachante_liquido=577.20,
        po_management_bruto=327.60,   # valor BRUTO padrão, editável pelo usuário
        total_despesas_pagas_comissaria=total_comissaria,
    )
    for s in demo.servicos:
        print(f"[OK] Serviço: {s.descricao} | líquido={s.valor_liquido:.2f} | bruto={s.valor_bruto:.2f}")

    demo.adiantamentos = adiant
    totais = calcular_totais(demo)
    print("\n=== TOTAIS ===")
    for k, v in totais.items():
        print(f"  {k}: {v}")

    print("\nTodos os testes de validação da lógica passaram com sucesso.")
