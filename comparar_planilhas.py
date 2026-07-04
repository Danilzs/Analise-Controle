"""
Compara as abas MARANGUAPE e CONTROLE de um arquivo Excel pelos nomes
e gera uma planilha com os registros que existem em uma aba, mas não na outra.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

ABA_MARANGUAPE = "MARANGUAPE"
ABA_CONTROLE = "CONTROLE"
COLUNA_NOME = "NOME COMPLETO"


def normalizar_nome(valor) -> str:
    """Normaliza nomes para comparação (remove espaços extras e ignora maiúsculas)."""
    if pd.isna(valor):
        return ""
    return str(valor).strip().casefold()


def carregar_aba(caminho: Path, aba: str, coluna_nome: str) -> pd.DataFrame:
    """Carrega uma aba do Excel e valida a coluna de nomes."""
    try:
        df = pd.read_excel(caminho, sheet_name=aba)
    except ValueError as erro:
        raise ValueError(
            f"Aba '{aba}' não encontrada em '{caminho.name}'. "
            f"Verifique se o nome da aba está correto."
        ) from erro

    if coluna_nome not in df.columns:
        colunas = ", ".join(str(c) for c in df.columns)
        raise ValueError(
            f"Coluna '{coluna_nome}' não encontrada na aba '{aba}'. "
            f"Colunas disponíveis: {colunas}"
        )

    return df


def encontrar_diferencas(
    df_maranguape: pd.DataFrame,
    df_controle: pd.DataFrame,
    coluna_nome: str,
) -> pd.DataFrame:
    """
    Retorna registros que existem em uma aba mas não na outra.
    Cada linha inclui a coluna Origem para indicar de qual aba veio.
    """
    df_m = df_maranguape.copy()
    df_c = df_controle.copy()

    df_m["_nome_norm"] = df_m[coluna_nome].map(normalizar_nome)
    df_c["_nome_norm"] = df_c[coluna_nome].map(normalizar_nome)

    nomes_controle = set(df_c["_nome_norm"]) - {""}
    nomes_maranguape = set(df_m["_nome_norm"]) - {""}

    apenas_maranguape = df_m[
        ~df_m["_nome_norm"].isin(nomes_controle) & (df_m["_nome_norm"] != "")
    ]
    apenas_controle = df_c[
        ~df_c["_nome_norm"].isin(nomes_maranguape) & (df_c["_nome_norm"] != "")
    ]

    apenas_maranguape = apenas_maranguape.drop(columns=["_nome_norm"])
    apenas_controle = apenas_controle.drop(columns=["_nome_norm"])

    apenas_maranguape.insert(0, "Origem", f"Somente em {ABA_MARANGUAPE}")
    apenas_controle.insert(0, "Origem", f"Somente em {ABA_CONTROLE}")

    resultado = pd.concat([apenas_maranguape, apenas_controle], ignore_index=True)

    if resultado.empty:
        return resultado

    return resultado.sort_values(
        by=coluna_nome,
        key=lambda s: s.map(normalizar_nome),
        ascending=True,
        kind="mergesort",
    ).reset_index(drop=True)


def comparar_abas(
    arquivo: Path,
    saida: Path,
    coluna_nome: str = COLUNA_NOME,
    aba_maranguape: str = ABA_MARANGUAPE,
    aba_controle: str = ABA_CONTROLE,
) -> pd.DataFrame:
    """Compara MARANGUAPE com CONTROLE e salva o resultado em Excel."""
    df_maranguape = carregar_aba(arquivo, aba_maranguape, coluna_nome)
    df_controle = carregar_aba(arquivo, aba_controle, coluna_nome)

    resultado = encontrar_diferencas(df_maranguape, df_controle, coluna_nome)

    saida.parent.mkdir(parents=True, exist_ok=True)
    resultado.to_excel(saida, index=False, sheet_name="Diferencas")

    return resultado


def criar_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Compara as abas MARANGUAPE e CONTROLE de um arquivo Excel "
            "pelos nomes e gera uma planilha com as diferenças."
        )
    )
    parser.add_argument(
        "arquivo",
        type=Path,
        help="Caminho do arquivo Excel com as abas MARANGUAPE e CONTROLE",
    )
    parser.add_argument(
        "-o",
        "--saida",
        type=Path,
        default=Path("resultado_diferencas.xlsx"),
        help="Caminho da planilha de saída (padrão: resultado_diferencas.xlsx)",
    )
    parser.add_argument(
        "-c",
        "--coluna",
        default=COLUNA_NOME,
        help=f"Coluna usada na comparação (padrão: {COLUNA_NOME})",
    )
    return parser


def main() -> int:
    parser = criar_parser()
    args = parser.parse_args()

    if not args.arquivo.exists():
        print(f"Erro: arquivo não encontrado: {args.arquivo}", file=sys.stderr)
        return 1

    try:
        resultado = comparar_abas(
            arquivo=args.arquivo,
            saida=args.saida,
            coluna_nome=args.coluna,
        )
    except ValueError as erro:
        print(f"Erro: {erro}", file=sys.stderr)
        return 1

    total = len(resultado)
    if total == 0:
        print(
            f"Nenhuma diferença encontrada entre {ABA_MARANGUAPE} e {ABA_CONTROLE}. "
            "Ambas possuem os mesmos nomes."
        )
    else:
        somente_m = (resultado["Origem"] == f"Somente em {ABA_MARANGUAPE}").sum()
        somente_c = (resultado["Origem"] == f"Somente em {ABA_CONTROLE}").sum()
        print(f"Comparação concluída: {total} registro(s) com diferença.")
        print(f"  - Somente em {ABA_MARANGUAPE}: {somente_m}")
        print(f"  - Somente em {ABA_CONTROLE}: {somente_c}")
        print(f"Arquivo gerado: {args.saida.resolve()}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
