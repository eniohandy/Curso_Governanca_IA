#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Classificador de Matriz de Risco Ético (Framework CertifAIEd)
Inspirado na estrutura do repositório 'Matriz-de-risco' do GitHub.

Coleta entradas do perfil ético, calcula pontuações flexíveis de risco
baseadas em Intensidade e Probabilidade, e gera uma Matriz de Risco 3x3
visual e em tabela no terminal.
"""

import csv
import argparse
import sys
import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.prompt import Prompt, IntPrompt, Confirm

console = Console()

# Lista oficial de Valores Éticos do CertifAIEd
VALORES_ETICOS = [
    "Transparência",
    "Justiça",
    "Privacidade",
    "Não-maleficência",
    "Responsabilidade",
    "Beneficência",
    "Liberdade e Autonomia",
    "Confiança",
    "Sustentabilidade",
    "Dignidade",
    "Solidariedade"
]

PARTES_INTERESSADAS_PADRAO = [
    "Clientes externos",
    "Clientes internos",
    "Desenvolvedores",
    "Sociedade / Geral",
    "Reguladores"
]

NIVEIS_MAPEADOS = {"Baixa": 1, "Média": 2, "Alta": 3}
MAPEAMENTO_INVERSO = {1: "Baixa", 2: "Média", 3: "Alta"}


def coletar_dados_interativo(output_csv="matriz_risco.csv"):
    """
    Coleta interativamente as entradas do perfil de avaliação ética e salva em CSV.
    """
    console.print(Panel.fit("[bold cyan]Coletor de Dados de Risco Ético (CertifAIEd)[/bold cyan]", border_style="cyan"))

    file_exists = os.path.exists(output_csv)
    
    # Determinar ID inicial
    existing_count = 0
    if file_exists:
        try:
            df_exist = pd.read_csv(output_csv)
            existing_count = len(df_exist)
        except Exception:
            existing_count = 0

    with open(output_csv, mode="a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not file_exists or os.path.getsize(output_csv) == 0:
            writer.writerow([
                "ID", "Valor_Etico", "Definicao", "Parte_Interessada",
                "Acao_Impacto", "Intensidade", "Probabilidade"
            ])

        count = existing_count + 1
        
        while True:
            console.print(f"\n[bold green]--- Nova Entrada de Risco (#{count}) ---[/bold green]")
            
            # 1. Valor Ético
            console.print("\n[yellow]Selecione o Valor Ético:[/yellow]")
            for i, val in enumerate(VALORES_ETICOS, 1):
                console.print(f"  [bold]{i}.[/bold] {val}")
            console.print(f"  [bold]{len(VALORES_ETICOS) + 1}.[/bold] Outro (digitar manualmente)")
            
            opcao_val = IntPrompt.ask("Escolha o número", choices=[str(i) for i in range(1, len(VALORES_ETICOS) + 2)])
            if opcao_val <= len(VALORES_ETICOS):
                valor_etico = VALORES_ETICOS[opcao_val - 1]
            else:
                valor_etico = Prompt.ask("Digite o valor ético")

            # 2. Definição / Contexto
            definicao = Prompt.ask("\nDigite a [yellow]Definição / Contexto do AIS[/yellow]")

            # 3. Parte Interessada
            console.print("\n[yellow]Selecione a Parte Interessada:[/yellow]")
            for i, pi in enumerate(PARTES_INTERESSADAS_PADRAO, 1):
                console.print(f"  [bold]{i}.[/bold] {pi}")
            console.print(f"  [bold]{len(PARTES_INTERESSADAS_PADRAO) + 1}.[/bold] Outra (digitar)")
            
            opcao_pi = IntPrompt.ask("Escolha o número", choices=[str(i) for i in range(1, len(PARTES_INTERESSADAS_PADRAO) + 2)])
            if opcao_pi <= len(PARTES_INTERESSADAS_PADRAO):
                parte_interessada = PARTES_INTERESSADAS_PADRAO[opcao_pi - 1]
            else:
                parte_interessada = Prompt.ask("Digite a parte interessada")

            # 4. Ação do Impacto (Minimiza ou Aumenta)
            opcao_acao = IntPrompt.ask(
                "\nO item [yellow]Minimiza ou Aumenta[/yellow] o impacto de risco?\n 1. Minimiza\n 2. Aumenta\nEscolha",
                choices=["1", "2"]
            )
            acao_impacto = "Minimiza" if opcao_acao == 1 else "Aumenta"

            # 5. Intensidade
            opcao_int = IntPrompt.ask(
                "\nQual a [yellow]Intensidade[/yellow]?\n 1. Baixa\n 2. Média\n 3. Alta\nEscolha",
                choices=["1", "2", "3"]
            )
            intensidade = MAPEAMENTO_INVERSO[opcao_int]

            # 6. Probabilidade
            opcao_prob = IntPrompt.ask(
                "\nQual a [yellow]Probabilidade[/yellow]?\n 1. Baixa\n 2. Média\n 3. Alta\nEscolha",
                choices=["1", "2", "3"]
            )
            probabilidade = MAPEAMENTO_INVERSO[opcao_prob]

            # Escrever linha no CSV
            item_id = f"R{count:03d}"
            writer.writerow([item_id, valor_etico, definicao, parte_interessada, acao_impacto, intensidade, probabilidade])
            console.print(f"[bold green]✓ Item {item_id} adicionado com sucesso a '{output_csv}'![/bold green]")
            
            count += 1
            if not Confirm.ask("\nDeseja adicionar outro item ao arquivo?"):
                break


def calcular_e_exibir_risco(input_csv="matriz_risco.csv", weight_int=1.0, weight_prob=1.0):
    """
    Lê o CSV, calcula a pontuação ponderada de risco e exibe tabela formatada no terminal.
    """
    if not os.path.exists(input_csv):
        console.print(f"[bold red]Erro: Arquivo '{input_csv}' não encontrado![/bold red]")
        return None

    try:
        df = pd.read_csv(input_csv)
    except Exception as e:
        console.print(f"[bold red]Erro ao ler o arquivo CSV '{input_csv}': {e}[/bold red]")
        return None

    # Validação de colunas
    colunas_requeridas = ["ID", "Valor_Etico", "Definicao", "Parte_Interessada", "Acao_Impacto", "Intensidade", "Probabilidade"]
    for col in colunas_requeridas:
        if col not in df.columns:
            console.print(f"[bold red]Erro: Coluna '{col}' ausente no arquivo CSV![/bold red]")
            return None

    # Mapeamento numérico
    df["Intensidade_Num"] = df["Intensidade"].map(NIVEIS_MAPEADOS).fillna(1)
    df["Probabilidade_Num"] = df["Probabilidade"].map(NIVEIS_MAPEADOS).fillna(1)

    # Cálculo flexível do score
    # Score Base = (weight_int * Intensidade) + (weight_prob * Probabilidade)
    def calcular_score(row):
        score_base = (weight_int * row["Intensidade_Num"]) + (weight_prob * row["Probabilidade_Num"])
        if str(row["Acao_Impacto"]).strip().capitalize() == "Minimiza":
            return round(-score_base / 2.0, 2)
        else:
            return round(score_base, 2)

    df["Score_Risco"] = df.apply(calcular_score, axis=1)

    # Classificação em faixas de cor
    def classificar_classe(row):
        score = row["Score_Risco"]
        acao = str(row["Acao_Impacto"]).strip().capitalize()
        if acao == "Minimiza" or score <= 2.0:
            return "Verde (Baixo / Mitigado)", "bold green"
        elif score <= 4.0:
            return "Amarela (Médio)", "bold yellow"
        else:
            return "Vermelha (Alto / Crítico)", "bold red"

    classes_e_estilos = df.apply(classificar_classe, axis=1)
    df["Classe_Risco"] = [c[0] for c in classes_e_estilos]
    df["Estilo_Cor"] = [c[1] for c in classes_e_estilos]

    # Exibição da tabela Rich
    table = Table(
        title=f"Matriz de Avaliação de Risco Ético — CertifAIEd\n(Ponderação Flexível: Peso Intensidade = {weight_int}, Peso Probabilidade = {weight_prob})",
        show_header=True,
        header_style="bold magenta",
        expand=True
    )
    table.add_column("ID", style="dim", width=6)
    table.add_column("Valor Ético", style="cyan", width=16)
    table.add_column("Definição / Contexto", style="white")
    table.add_column("Parte Interessada", style="blue", width=18)
    table.add_column("Ação", justify="center", width=10)
    table.add_column("Intensidade", justify="center", width=11)
    table.add_column("Probabilidade", justify="center", width=13)
    table.add_column("Score", justify="right", width=8)
    table.add_column("Classe de Risco", justify="center", width=22)

    for _, row in df.iterrows():
        acao_color = "green" if str(row["Acao_Impacto"]).strip().capitalize() == "Minimiza" else "red"
        table.add_row(
            str(row["ID"]),
            str(row["Valor_Etico"]),
            str(row["Definicao"]),
            str(row["Parte_Interessada"]),
            f"[{acao_color}]{row['Acao_Impacto']}[/{acao_color}]",
            str(row["Intensidade"]),
            str(row["Probabilidade"]),
            f"{row['Score_Risco']:.1f}",
            f"[{row['Estilo_Cor']}]{row['Classe_Risco']}[/{row['Estilo_Cor']}]"
        )

    console.print(table)
    return df


def plotar_matriz_risco_pdf_style(df, output_img="matriz_risco_certifaied.png"):
    """
    Gera um gráfico visual estilo 3x3 da Matriz de Risco do CertifAIEd.
    """
    if df is None or df.empty:
        console.print("[bold red]Nenhum dado fornecido para plotagem da matriz![/bold red]")
        return

    plt.figure(figsize=(9, 8), dpi=150)
    ax = plt.gca()

    # Cores de fundo por zona (3x3 Grid)
    cores_grid = {
        (1, 1): "#90ee90", (2, 1): "#90ee90", (1, 2): "#90ee90", # Verde (Baixo)
        (2, 2): "#ffeb3b", (3, 1): "#ffeb3b", (1, 3): "#ffeb3b", # Amarelo (Médio)
        (2, 3): "#ff6666", (3, 2): "#ff6666", (3, 3): "#ff4d4d", # Vermelho (Alto)
    }

    for (x, y), color in cores_grid.items():
        rect = patches.Rectangle((x - 0.5, y - 0.5), 1, 1, linewidth=1.5, edgecolor="white", facecolor=color, alpha=0.85)
        ax.add_patch(rect)

    # Plotar pontos / rótulos dos itens
    counts = {}
    for _, row in df.iterrows():
        px = row["Probabilidade_Num"]
        py = row["Intensidade_Num"]
        
        cell_key = (px, py)
        counts[cell_key] = counts.get(cell_key, 0) + 1
        offset = (counts[cell_key] - 1) * 0.12 - 0.08
        
        is_minimiza = str(row["Acao_Impacto"]).strip().capitalize() == "Minimiza"
        marker_color = "darkgreen" if is_minimiza else "darkred"
        marker_shape = "o" if is_minimiza else "s"
        
        ax.scatter(px + offset, py + offset, color=marker_color, marker=marker_shape, s=180, zorder=5, edgecolors="black")
        ax.text(
            px + offset, py + offset, f" {row['ID']}",
            fontsize=9, fontweight="bold", color="black",
            ha="left", va="bottom", zorder=6,
            bbox=dict(boxstyle="round,pad=0.2", facecolor="white", alpha=0.85, edgecolor="none")
        )

    # Eixos e títulos no visual CertifAIEd
    ax.set_xlim(0.5, 3.5)
    ax.set_ylim(0.5, 3.5)
    
    ax.set_xticks([1, 2, 3])
    ax.set_xticklabels(["Baixa", "Média", "Alta"], fontsize=12, fontweight="bold")
    ax.set_yticks([1, 2, 3])
    ax.set_yticklabels(["Baixa", "Média", "Alta"], fontsize=12, fontweight="bold")

    ax.set_xlabel("Probabilidade", fontsize=13, fontweight="bold", labelpad=10)
    ax.set_ylabel("Intensidade (Impacto)", fontsize=13, fontweight="bold", labelpad=10)
    
    plt.title("Matriz de Avaliação de Risco Ético (CertifAIEd)", fontsize=15, fontweight="bold", pad=15)

    # Legenda descritiva
    legend_elements = [
        patches.Patch(facecolor="#90ee90", edgecolor="gray", label="Classe Verde (Baixo Risco / Mitigado)"),
        patches.Patch(facecolor="#ffeb3b", edgecolor="gray", label="Classe Amarela (Médio Risco)"),
        patches.Patch(facecolor="#ff6666", edgecolor="gray", label="Classe Vermelha (Alto Risco)"),
        plt.Line2D([0], [0], marker="s", color="w", label="Aumenta Risco", markerfacecolor="darkred", markersize=10),
        plt.Line2D([0], [0], marker="o", color="w", label="Minimiza Risco (Mitiga)", markerfacecolor="darkgreen", markersize=10),
    ]
    ax.legend(handles=legend_elements, loc="upper left", bbox_to_anchor=(1.02, 1), borderaxespad=0.)

    plt.tight_layout()
    plt.savefig(output_img, bbox_inches="tight")
    plt.close()
    console.print(f"[bold green]✓ Gráfico da Matriz de Risco salvo como '{output_img}'![/bold green]")


def main():
    parser = argparse.ArgumentParser(description="Classificador de Matriz de Risco Ético (CertifAIEd)")
    parser.add_argument("--coletar", action="store_true", help="Executa o modo interativo para coletar novos dados e salvar em CSV")
    parser.add_argument("--csv", type=str, default="matriz_risco.csv", help="Caminho do arquivo CSV de dados (default: matriz_risco.csv)")
    parser.add_argument("--weight-int", type=float, default=1.0, help="Peso flexível atribuído à Intensidade (default: 1.0)")
    parser.add_argument("--weight-prob", type=float, default=1.0, help="Peso flexível atribuído à Probabilidade (default: 1.0)")
    parser.add_argument("--plot", action="store_true", help="Gera a imagem gráfica da Matriz de Risco 3x3")
    
    args = parser.parse_args()

    if args.coletar:
        coletar_dados_interativo(args.csv)

    if os.path.exists(args.csv):
        df_resultado = calcular_e_exibir_risco(args.csv, weight_int=args.weight_int, weight_prob=args.weight_prob)
        if df_resultado is not None:
            # Gera imagem por padrão se o arquivo existir ou se a flag --plot foi passada
            output_png = os.path.splitext(args.csv)[0] + "_certifaied_matrix.png"
            plotar_matriz_risco_pdf_style(df_resultado, output_img=output_png)
    else:
        if not args.coletar:
            console.print(f"[yellow]Arquivo '{args.csv}' não encontrado. Execute com '--coletar' para criar seus primeiros registros interativamente.[/yellow]")


if __name__ == "__main__":
    main()
