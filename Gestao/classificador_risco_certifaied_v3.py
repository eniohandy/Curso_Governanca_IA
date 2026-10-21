#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Classificador e Painel de Risco Ético (Framework CertifAIEd)
Inspirado na estrutura de triagem do repositório 'Matriz-de-risco' do GitHub.

Recursos principais:
 1. Coleta interativa de entradas no formato do perfil de avaliação ética do CertifAIEd.
 2. Cálculo flexível de score de risco com ponderação configurável de Intensidade e Probabilidade.
 3. Ordenação automática da lista do MAIOR RISCO (Crítico) ao MENOR RISCO (Mitigado).
 4. Exibição de tabela colorida no terminal com marcadores de risco (🔴, 🟡, 🟢).
 5. Geração do Painel em imagem PNG de alta resolução no formato de planilha com bolinhas coloridas.
 6. Exportação para arquivo Excel (.xlsx) formatado com estilos e cores condicionais.
"""

import csv
import argparse
import sys
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches

# Backend headless para matplotlib
import matplotlib
matplotlib.use('Agg')

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
    "Clientes externo",
    "Clientes interno",
    "Desenvolvedores",
    "Sociedade / Geral",
    "Reguladores"
]

NIVEIS_MAPEADOS = {"Baixa": 1, "Média": 2, "Alta": 3}
MAPEAMENTO_INVERSO = {1: "Baixa", 2: "Média", 3: "Alta"}


def coletar_dados_interativo(output_csv="matriz_risco.csv"):
    """
    Coleta interativamente as entradas do perfil de avaliação ética e salva no CSV.
    """
    console.print(Panel.fit("[bold cyan]Coletor de Dados de Risco Ético (CertifAIEd)[/bold cyan]", border_style="cyan"))

    file_exists = os.path.exists(output_csv)
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


def processar_e_ordenar_riscos(csv_path="matriz_risco.csv", weight_int=1.0, weight_prob=1.0):
    """
    Lê o CSV, calcula as notas de risco, classifica e ordena do MAIOR para o MENOR risco.
    """
    if not os.path.exists(csv_path):
        console.print(f"[bold red]Erro: Arquivo '{csv_path}' não encontrado![/bold red]")
        return None

    try:
        df = pd.read_csv(csv_path)
    except Exception as e:
        console.print(f"[bold red]Erro ao ler CSV '{csv_path}': {e}[/bold red]")
        return None

    # Validação de colunas
    colunas_requeridas = ["ID", "Valor_Etico", "Definicao", "Parte_Interessada", "Acao_Impacto", "Intensidade", "Probabilidade"]
    for col in colunas_requeridas:
        if col not in df.columns:
            console.print(f"[bold red]Erro: Coluna '{col}' ausente no arquivo CSV![/bold red]")
            return None

    df["Intensidade_Num"] = df["Intensidade"].map(NIVEIS_MAPEADOS).fillna(1)
    df["Probabilidade_Num"] = df["Probabilidade"].map(NIVEIS_MAPEADOS).fillna(1)

    def calcular_score(row):
        score_base = (weight_int * row["Intensidade_Num"]) + (weight_prob * row["Probabilidade_Num"])
        if str(row["Acao_Impacto"]).strip().capitalize() == "Minimiza":
            return round(-score_base / 2.0, 2)
        else:
            return round(score_base, 2)

    df["Score_Risco"] = df.apply(calcular_score, axis=1)

    def classificar_risco(row):
        score = row["Score_Risco"]
        acao = str(row["Acao_Impacto"]).strip().capitalize()
        if acao == "Minimiza" or score <= 2.0:
            return "Verde (Baixo / Mitigado)", "🟢", "#2e7d32", "#e8f5e9", "#4caf50"
        elif score <= 4.0:
            return "Amarela (Médio)", "🟡", "#f57f17", "#fffde7", "#fbc02d"
        else:
            return "Vermelha (Alto / Crítico)", "🔴", "#c62828", "#ffebee", "#e53935"

    res = df.apply(classificar_risco, axis=1)
    df["Classe_Risco"] = [r[0] for r in res]
    df["Bolinha"] = [r[1] for r in res]
    df["Cor_Texto"] = [r[2] for r in res]
    df["Cor_Fundo"] = [r[3] for r in res]
    df["Cor_Bolinha"] = [r[4] for r in res]

    # ORDENAÇÃO: DO MAIS ARRISCADO PARA O MENOS (Score Risco Decrescente)
    df = df.sort_values(by="Score_Risco", ascending=False).reset_index(drop=True)
    return df


def exibir_tabela_terminal(df, weight_int, weight_prob):
    """
    Exibe a planilha formatada no terminal com as bolinhas coloridas e ordem de risco.
    """
    table = Table(
        title=f"PAINEL DE AVALIAÇÃO DE RISCO ÉTICO (CERTIFAIEd)\nOrdens de Prioridade: Do Mais Arriscado ao Menos Arriscado (Pesos: Int={weight_int}, Prob={weight_prob})",
        show_header=True,
        header_style="bold white on blue",
        expand=True
    )

    table.add_column("Risco", justify="center", width=8)
    table.add_column("ID", style="dim", width=6)
    table.add_column("Valor Ético", style="cyan", width=16)
    table.add_column("Definição / Contexto do AIS", style="white")
    table.add_column("Parte Interessada", style="blue", width=18)
    table.add_column("Ação", justify="center", width=10)
    table.add_column("Intensidade", justify="center", width=11)
    table.add_column("Probabilidade", justify="center", width=13)
    table.add_column("Score", justify="right", width=8)
    table.add_column("Classe de Risco", justify="center", width=22)

    for _, row in df.iterrows():
        acao_color = "green" if str(row["Acao_Impacto"]).strip().capitalize() == "Minimiza" else "red"
        cor_classe = "bold red" if "Vermelha" in row["Classe_Risco"] else ("bold yellow" if "Amarela" in row["Classe_Risco"] else "bold green")
        
        table.add_row(
            row["Bolinha"],
            str(row["ID"]),
            str(row["Valor_Etico"]),
            str(row["Definicao"]),
            str(row["Parte_Interessada"]),
            f"[{acao_color}]{row['Acao_Impacto']}[/{acao_color}]",
            str(row["Intensidade"]),
            str(row["Probabilidade"]),
            f"[bold]{row['Score_Risco']:.1f}[/bold]",
            f"[{cor_classe}]{row['Classe_Risco']}[/{cor_classe}]"
        )

    console.print(table)


def gerar_imagem_painel_planilha(df, output_png="painel_risco_certifaied.png"):
    """
    Gera o Painel em imagem PNG de alta resolução no formato de planilha com as bolinhas de risco.
    """
    if df is None or df.empty:
        return

    num_rows = len(df)
    fig_height = max(5.0, 2.0 + num_rows * 0.75)
    fig, ax = plt.subplots(figsize=(16, fig_height), dpi=200)
    ax.axis('off')

    plt.suptitle("PAINEL DE AVALIAÇÃO DE RISCO ÉTICO (CERTIFAIEd)",
                 fontsize=16, fontweight='bold', y=0.97, color='#1a237e')
    plt.title("Visualização em Formato Planilha — Ordenada do Maior Risco (Crítico) ao Menor Risco (Mitigado)",
              fontsize=11, fontstyle='italic', pad=18, color='#424242')

    headers = ["Status Risco", "ID", "Valor Ético", "Definição / Contexto do AIS", "Parte Interessada", "Ação", "Intensidade", "Probabilidade", "Score Risco"]

    cell_text = []
    cell_colors = []

    for idx, row in df.iterrows():
        label_status = "      ALTO" if "Vermelha" in row['Classe_Risco'] else ("      MÉDIO" if "Amarela" in row['Classe_Risco'] else "      BAIXO")
        c_text = [
            label_status,
            row["ID"],
            row["Valor_Etico"],
            row["Definicao"],
            row["Parte_Interessada"],
            row["Acao_Impacto"],
            row["Intensidade"],
            row["Probabilidade"],
            f"{row['Score_Risco']:.1f}"
        ]
        cell_text.append(c_text)
        cell_colors.append([row["Cor_Fundo"]] * len(headers))

    widths = [0.11, 0.05, 0.13, 0.35, 0.13, 0.07, 0.08, 0.09, 0.07]

    table = ax.table(
        cellText=cell_text,
        colLabels=headers,
        cellColours=cell_colors,
        loc='center',
        cellLoc='center',
        colWidths=widths
    )

    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1.0, 2.3)

    for (i, j), cell in table.get_celld().items():
        cell.set_edgecolor('#cfd8dc')
        cell.set_linewidth(1.0)
        if i == 0:
            cell.set_facecolor('#1565c0')
            cell.set_text_props(weight='bold', color='white', fontsize=10.5)
        else:
            row_data = df.iloc[i - 1]
            if j == 0:
                cell.set_text_props(weight='bold', color=row_data['Cor_Texto'], ha='center')
            elif j == 3:
                cell.set_text_props(ha='left', fontsize=9.5)
            elif j == 8:
                cell.set_text_props(weight='bold', fontsize=11)

    fig.canvas.draw()

    # Desenhar as bolinhas coloridas de risco sobre a coluna Status
    for i in range(1, len(df) + 1):
        cell = table[(i, 0)]
        bbox = cell.get_window_extent(fig.canvas.get_renderer())
        bbox_data = ax.transData.inverted().transform(bbox)
        x0, y0 = bbox_data[0]
        x1, y1 = bbox_data[1]
        
        cx = x0 + (x1 - x0) * 0.22
        cy = y0 + (y1 - y0) * 0.50
        radius = min(x1 - x0, y1 - y0) * 0.18

        row_data = df.iloc[i - 1]
        circle = patches.Circle((cx, cy), radius, color=row_data["Cor_Bolinha"], zorder=10, ec="#ffffff", lw=1.5)
        ax.add_patch(circle)

    plt.tight_layout()
    plt.savefig(output_png, bbox_inches='tight', facecolor='white')
    plt.close()
    console.print(f"[bold green]✓ Imagem do painel salva com sucesso como '{output_png}'![/bold green]")


def gerar_excel_painel(df, output_xlsx="painel_risco_certifaied.xlsx"):
    """
    Exporta a planilha para Excel (.xlsx) com formatação condicional de cores e bolinhas.
    """
    if df is None or df.empty:
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Painel de Risco Ético"

    header_fill = PatternFill(start_color="1E88E5", end_color="1E88E5", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Calibri", size=14, bold=True, color="1A237E")

    thin_border = Border(
        left=Side(style='thin', color='B0BEC5'),
        right=Side(style='thin', color='B0BEC5'),
        top=Side(style='thin', color='B0BEC5'),
        bottom=Side(style='thin', color='B0BEC5')
    )

    ws.merge_cells("A1:I1")
    ws["A1"] = "PAINEL DE AVALIAÇÃO DE RISCO ÉTICO (CERTIFAIEd)"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A2:I2")
    ws["A2"] = "Lista ordenada do Maior Risco (Crítico) ao Menor Risco (Mitigado)"
    ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="555555")
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    headers = ["Status Risco", "ID", "Valor Ético", "Definição / Contexto do AIS", "Parte Interessada", "Ação", "Intensidade", "Probabilidade", "Score de Risco"]
    
    ws.append([])

    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    for row_idx, row in df.iterrows():
        r_num = row_idx + 5
        status_symbol = f"{row['Bolinha']}  ALTO" if "Vermelha" in row["Classe_Risco"] else (f"{row['Bolinha']}  MÉDIO" if "Amarela" in row["Classe_Risco"] else f"{row['Bolinha']}  BAIXO")
        
        ws.cell(row=r_num, column=1, value=status_symbol)
        ws.cell(row=r_num, column=2, value=row["ID"])
        ws.cell(row=r_num, column=3, value=row["Valor_Etico"])
        ws.cell(row=r_num, column=4, value=row["Definicao"])
        ws.cell(row=r_num, column=5, value=row["Parte_Interessada"])
        ws.cell(row=r_num, column=6, value=row["Acao_Impacto"])
        ws.cell(row=r_num, column=7, value=row["Intensidade"])
        ws.cell(row=r_num, column=8, value=row["Probabilidade"])
        ws.cell(row=r_num, column=9, value=row["Score_Risco"])

        fill_color = "FFEBEE" if "Vermelha" in row["Classe_Risco"] else ("FFFDE7" if "Amarela" in row["Classe_Risco"] else "E8F5E9")
        status_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

        for c_i in range(1, 10):
            c = ws.cell(row=r_num, column=c_i)
            c.border = thin_border
            c.alignment = Alignment(vertical="center", horizontal="left" if c_i in [3, 4, 5] else "center", wrap_text=True)
            if c_i == 1:
                c.fill = status_fill
                c.font = Font(bold=True)
            if c_i == 9:
                c.font = Font(bold=True)

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row >= 4:
                val_str = str(cell.value or '')
                max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
    ws.column_dimensions['D'].width = 45

    wb.save(output_xlsx)
    console.print(f"[bold green]✓ Planilha Excel salva com sucesso como '{output_xlsx}'![/bold green]")


def main():
    parser = argparse.ArgumentParser(description="Classificador e Painel de Risco Ético (CertifAIEd)")
    parser.add_argument("--coletar", action="store_true", help="Executa o modo interativo para coletar novos dados e salvar em CSV")
    parser.add_argument("--csv", type=str, default="matriz_risco.csv", help="Caminho do arquivo CSV de dados (default: matriz_risco.csv)")
    parser.add_argument("--weight-int", type=float, default=1.2, help="Peso flexível para Intensidade (default: 1.2)")
    parser.add_argument("--weight-prob", type=float, default=1.0, help="Peso flexível para Probabilidade (default: 1.0)")
    
    args = parser.parse_args()

    if args.coletar:
        coletar_dados_interativo(args.csv)

    if os.path.exists(args.csv):
        df_ordenado = processar_e_ordenar_riscos(args.csv, weight_int=args.weight_int, weight_prob=args.weight_prob)
        if df_ordenado is not None:
            # 1. Exibir no terminal
            exibir_tabela_terminal(df_ordenado, args.weight_int, args.weight_prob)
            
            # 2. Gerar Painel em imagem PNG
            output_png = "painel_risco_certifaied.png"
            gerar_imagem_painel_planilha(df_ordenado, output_png=output_png)
            
            # 3. Gerar Planilha Excel (.xlsx)
            output_xlsx = "painel_risco_certifaied.xlsx"
            gerar_excel_painel(df_ordenado, output_xlsx=output_xlsx)
    else:
        if not args.coletar:
            console.print(f"[yellow]Arquivo '{args.csv}' não encontrado. Execute com '--coletar' para cadastrar novos itens interativamente.[/yellow]")


if __name__ == "__main__":
    main()
