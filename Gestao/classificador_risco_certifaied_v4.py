#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Classificador e Painel de Risco Ético (Framework CertifAIEd)
Inspirado na estrutura do repositório 'Matriz-de-risco' do GitHub.

Gera:
1. Coleta interativa de riscos e salvamento em CSV (--coletar).
2. Tabela colorida no terminal organizada do maior risco para o menor.
3. Planilha Excel (.xlsx) estilizada e ordenada.
4. Imagem do Painel / Planilha em PNG (.png) com letras coloridas (sem preto ou branco).
"""

import csv
import argparse
import sys
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
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

            # 4. Ação do Impacto
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

            item_id = f"R{count:03d}"
            writer.writerow([item_id, valor_etico, definicao, parte_interessada, acao_impacto, intensidade, probabilidade])
            console.print(f"[bold green]✓ Item {item_id} adicionado com sucesso a '{output_csv}'![/bold green]")
            
            count += 1
            if not Confirm.ask("\nDeseja adicionar outro item ao arquivo?"):
                break


def calcular_e_processar_risco(input_csv="matriz_risco.csv", weight_int=1.0, weight_prob=1.0):
    """
    Lê o CSV, calcula os scores estritamente positivos (1.0 a 4.0) e ordena do maior para o menor risco.
    """
    if not os.path.exists(input_csv):
        # Exemplo padrão se o arquivo não existir
        df_default = pd.DataFrame([
            {"ID": "R001", "Valor_Etico": "Transparência", "Definicao": "Sistema de diálogo não informa detalhadamente o seu funcionamento", "Parte_Interessada": "Clientes externos", "Acao_Impacto": "Aumenta", "Intensidade": "Alta", "Probabilidade": "Alta"},
            {"ID": "R002", "Valor_Etico": "Justiça", "Definicao": "Nos diálogos trata as pessoas por sr/sra e evita linguagem ofensiva", "Parte_Interessada": "Clientes externos", "Acao_Impacto": "Minimiza", "Intensidade": "Baixa", "Probabilidade": "Baixa"},
            {"ID": "R003", "Valor_Etico": "Privacidade", "Definicao": "Pessoas podem informar dados pessoais no chatbot", "Parte_Interessada": "Clientes externos", "Acao_Impacto": "Aumenta", "Intensidade": "Média", "Probabilidade": "Média"},
            {"ID": "R004", "Valor_Etico": "Não-maleficência", "Definicao": "Sistema previne respostas com desinformação ou preconceito", "Parte_Interessada": "Desenvolvedores", "Acao_Impacto": "Minimiza", "Intensidade": "Média", "Probabilidade": "Baixa"},
            {"ID": "R005", "Valor_Etico": "Responsabilidade", "Definicao": "Ausência de logs de auditoria nas decisões do assistente", "Parte_Interessada": "Reguladores", "Acao_Impacto": "Aumenta", "Intensidade": "Alta", "Probabilidade": "Média"}
        ])
        df_default.to_csv(input_csv, index=False)
        df = df_default
    else:
        try:
            df = pd.read_csv(input_csv)
        except Exception as e:
            console.print(f"[bold red]Erro ao ler CSV '{input_csv}': {e}[/bold red]")
            return None

    df["Intensidade_Num"] = df["Intensidade"].map(NIVEIS_MAPEADOS).fillna(1)
    df["Probabilidade_Num"] = df["Probabilidade"].map(NIVEIS_MAPEADOS).fillna(1)

    # Cálculo proporcional de Score estritamente positivo (1.0 a 4.0)
    # Baixa / Baixa (1, 1) -> Raw Min = weight_int*1 + weight_prob*1
    # Alta / Alta (3, 3)   -> Raw Max = weight_int*3 + weight_prob*3
    min_raw = (weight_int * 1.0) + (weight_prob * 1.0)
    max_raw = (weight_int * 3.0) + (weight_prob * 3.0)

    def calc_score(row):
        raw = (weight_int * row["Intensidade_Num"]) + (weight_prob * row["Probabilidade_Num"])
        if max_raw != min_raw:
            score_base = 1.0 + 3.0 * (raw - min_raw) / (max_raw - min_raw)
        else:
            score_base = 1.0
        
        acao = str(row["Acao_Impacto"]).strip().capitalize()
        if acao == "Minimiza":
            score_final = max(1.0, score_base / 2.0)
        else:
            score_final = score_base
        return round(score_final, 2)

    df["Score_Risco"] = df.apply(calc_score, axis=1)

    # ORDENAR DO MAIS ARRISCADO PARA O MENOS ARRISCADO
    df = df.sort_values(by="Score_Risco", ascending=False).reset_index(drop=True)
    return df


def exibir_tabela_terminal(df, weight_int=1.0, weight_prob=1.0):
    """
    Exibe a tabela no terminal com texto colorido (sem letras brancas ou pretas).
    """
    table = Table(
        title=f"PAINEL DE AVALIAÇÃO DE RISCO ÉTICO (CERTIFAIEd)\nOrdens de Prioridade: Do Mais Arriscado ao Menos Arriscado (Pesos: Int={weight_int}, Prob={weight_prob})",
        show_header=True,
        header_style="bold magenta",
        expand=True
    )

    table.add_column("Sinal", justify="center", width=5)
    table.add_column("ID", style="bold cyan", width=6)
    table.add_column("Valor Ético", style="bold blue", width=16)
    table.add_column("Definição / Contexto", style="deep_sky_blue1")
    table.add_column("Parte Interessada", style="purple", width=18)
    table.add_column("Ação", justify="center", width=10)
    table.add_column("Intensidade", justify="center", width=11)
    table.add_column("Probabilidade", justify="center", width=13)
    table.add_column("Score", justify="right", width=8)

    for _, row in df.iterrows():
        score = row["Score_Risco"]
        if score >= 3.0:
            sinal = "🔴"
            score_style = "bold red"
        elif score > 1.0:
            sinal = "🟡"
            score_style = "bold yellow"
        else:
            sinal = "🟢"
            score_style = "bold green"

        acao = str(row["Acao_Impacto"]).strip().capitalize()
        acao_color = "green" if acao == "Minimiza" else "red"

        int_str = str(row["Intensidade"])
        int_color = "green" if int_str == "Baixa" else ("yellow" if int_str == "Média" else "red")

        prob_str = str(row["Probabilidade"])
        prob_color = "green" if prob_str == "Baixa" else ("yellow" if prob_str == "Média" else "red")

        table.add_row(
            sinal,
            str(row["ID"]),
            str(row["Valor_Etico"]),
            str(row["Definicao"]),
            str(row["Parte_Interessada"]),
            f"[{acao_color}]{acao}[/{acao_color}]",
            f"[{int_color}]{int_str}[/{int_color}]",
            f"[{prob_color}]{prob_str}[/{prob_color}]",
            f"[{score_style}]{score:.2f}[/{score_style}]"
        )

    console.print(table)


def gerar_planilha_excel(df, output_excel="painel_risco_certifaied.xlsx"):
    """
    Gera arquivo Excel (.xlsx) estilizado com cores específicas (sem texto preto/branco).
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Painel de Risco Ético"

    headers = ["Status", "ID", "Valor Ético", "Definição / Contexto", "Parte Interessada", "Ação", "Intensidade", "Probabilidade", "Score Risco"]
    ws.append(headers)

    # Estilos de Cabeçalho (Fundo Azul Claro, Fonte Navy)
    fill_header = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    font_header = Font(name="Calibri", size=11, bold=True, color="002060")

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center")

    thin_border = Border(
        left=Side(style='thin', color='B0C4DE'),
        right=Side(style='thin', color='B0C4DE'),
        top=Side(style='thin', color='B0C4DE'),
        bottom=Side(style='thin', color='B0C4DE')
    )

    for idx, row in df.iterrows():
        score = row["Score_Risco"]
        if score >= 3.0:
            status_symbol = "🔴 (Alto)"
            score_color = "C00000"
        elif score > 1.0:
            status_symbol = "🟡 (Médio)"
            score_color = "B25900"
        else:
            status_symbol = "🟢 (Baixo)"
            score_color = "008000"

        acao = str(row["Acao_Impacto"]).strip().capitalize()
        acao_color = "008000" if acao == "Minimiza" else "C00000"

        row_data = [
            status_symbol,
            str(row["ID"]),
            str(row["Valor_Etico"]),
            str(row["Definicao"]),
            str(row["Parte_Interessada"]),
            acao,
            str(row["Intensidade"]),
            str(row["Probabilidade"]),
            float(score)
        ]
        ws.append(row_data)
        row_idx = ws.max_row

        # Cores de fontes sem usar preto ou branco
        font_status = Font(name="Calibri", size=11, bold=True, color=score_color)
        font_id = Font(name="Calibri", size=11, bold=True, color="0044BB")
        font_vetico = Font(name="Calibri", size=11, bold=True, color="002060")
        font_def = Font(name="Calibri", size=11, color="1F497D")
        font_pi = Font(name="Calibri", size=11, color="7030A0")
        font_acao = Font(name="Calibri", size=11, bold=True, color=acao_color)
        font_level = Font(name="Calibri", size=11, color="4B0082")
        font_score = Font(name="Calibri", size=11, bold=True, color=score_color)

        ws.cell(row=row_idx, column=1).font = font_status
        ws.cell(row=row_idx, column=2).font = font_id
        ws.cell(row=row_idx, column=3).font = font_vetico
        ws.cell(row=row_idx, column=4).font = font_def
        ws.cell(row=row_idx, column=5).font = font_pi
        ws.cell(row=row_idx, column=6).font = font_acao
        ws.cell(row=row_idx, column=7).font = font_level
        ws.cell(row=row_idx, column=8).font = font_level
        ws.cell(row=row_idx, column=9).font = font_score

        for col_num in range(1, 10):
            c = ws.cell(row=row_idx, column=col_num)
            c.border = thin_border
            if col_num in [1, 2, 6, 7, 8, 9]:
                c.alignment = Alignment(horizontal="center", vertical="center")

    # Ajustar largura de colunas
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(output_excel)
    console.print(f"[bold green]✓ Planilha Excel salva com sucesso como '{output_excel}'![/bold green]")


def gerar_painel_imagem_png(df, output_png="painel_risco_certifaied.png", weight_int=1.0, weight_prob=1.0):
    """
    Gera a figura PNG do Painel/Planilha com letras estritamente em azul, verde, vermelho, roxo, etc. (NUNCA branco ou preto).
    """
    fig, ax = plt.subplots(figsize=(12, max(4, len(df) * 0.85 + 2)), dpi=150)
    fig.patch.set_facecolor('#F0F4F8')
    ax.axis('off')

    # Título do Painel em Azul Escuro
    ax.text(0.5, 0.96, "PAINEL DE AVALIAÇÃO DE RISCO ÉTICO (CERTIFAIEd)",
            fontsize=15, fontweight='bold', color='#002060', ha='center', va='top', transform=fig.transFigure)
    
    ax.text(0.5, 0.92, f"Prioridade: Do Mais Arriscado ao Menos Arriscado (Pesos: Intensidade={weight_int}, Probabilidade={weight_prob})",
            fontsize=11, fontweight='bold', color='#1F497D', ha='center', va='top', transform=fig.transFigure)

    col_widths = [0.08, 0.08, 0.16, 0.28, 0.15, 0.09, 0.08, 0.08]
    col_x = [0.03]
    for w in col_widths[:-1]:
        col_x.append(col_x[-1] + w)

    headers = ["Status", "ID", "Valor Ético", "Definição / Contexto", "Parte Interessada", "Ação", "Int/Prob", "Score"]
    
    y_start = 0.84
    row_height = 0.075

    # Desenhar Cabeçalho da Tabela
    for i, h in enumerate(headers):
        rect = patches.Rectangle((col_x[i], y_start), col_widths[i] - 0.005, 0.05,
                                 facecolor='#D9E1F2', edgecolor='#B0C4DE', transform=fig.transFigure)
        fig.patches.append(rect)
        ax.text(col_x[i] + (col_widths[i] - 0.005) / 2, y_start + 0.025, h,
                fontsize=10, fontweight='bold', color='#002060', ha='center', va='center', transform=fig.transFigure)

    curr_y = y_start - row_height

    for idx, row in df.iterrows():
        score = row["Score_Risco"]
        if score >= 3.0:
            circle_color = '#CC0000' # Vermelho
            score_color = '#B30000'
        elif score >= 2.0:
            circle_color = '#D68A00' # Amarelo/Dourado
            score_color = '#B25900'
        else:
            circle_color = '#008000' # Verde
            score_color = '#008000'

        acao = str(row["Acao_Impacto"]).strip().capitalize()
        acao_color = '#008000' if acao == "Minimiza" else '#CC0000'

        bg_color = '#F8FAFC' if idx % 2 == 0 else '#EEF2F7'

        # Fundo da Linha
        row_rect = patches.Rectangle((col_x[0], curr_y), sum(col_widths) - 0.005, 0.065,
                                     facecolor=bg_color, edgecolor='#CBD5E1', transform=fig.transFigure)
        fig.patches.append(row_rect)

        # 1. Status (Bolinha Colorida)
        circle_x = col_x[0] + (col_widths[0] - 0.005) / 2
        circle_y = curr_y + 0.0325
        circle = patches.Circle((circle_x, circle_y), 0.012, facecolor=circle_color, edgecolor=circle_color, transform=fig.transFigure)
        fig.patches.append(circle)

        # 2. ID
        ax.text(col_x[1] + (col_widths[1] - 0.005) / 2, circle_y, str(row["ID"]),
                fontsize=10, fontweight='bold', color='#0044BB', ha='center', va='center', transform=fig.transFigure)

        # 3. Valor Ético
        ax.text(col_x[2] + 0.005, circle_y, str(row["Valor_Etico"]),
                fontsize=9.5, fontweight='bold', color='#002060', ha='left', va='center', transform=fig.transFigure)

        # 4. Definição / Contexto (com quebra de linha se necessário)
        def_text = str(row["Definicao"])
        if len(def_text) > 42:
            def_text = def_text[:40] + "..."
        ax.text(col_x[3] + 0.005, circle_y, def_text,
                fontsize=8.5, fontweight='normal', color='#1F497D', ha='left', va='center', transform=fig.transFigure)

        # 5. Parte Interessada
        ax.text(col_x[4] + 0.005, circle_y, str(row["Parte_Interessada"]),
                fontsize=9, fontweight='normal', color='#7030A0', ha='left', va='center', transform=fig.transFigure)

        # 6. Ação
        ax.text(col_x[5] + (col_widths[5] - 0.005) / 2, circle_y, acao,
                fontsize=9, fontweight='bold', color=acao_color, ha='center', va='center', transform=fig.transFigure)

        # 7. Int/Prob
        int_prob_str = f"{row['Intensidade']}/{row['Probabilidade']}"
        ax.text(col_x[6] + (col_widths[6] - 0.005) / 2, circle_y, int_prob_str,
                fontsize=8.5, fontweight='bold', color='#4B0082', ha='center', va='center', transform=fig.transFigure)

        # 8. Score Risco
        ax.text(col_x[7] + (col_widths[7] - 0.005) / 2, circle_y, f"{score:.2f}",
                fontsize=11, fontweight='bold', color=score_color, ha='center', va='center', transform=fig.transFigure)

        curr_y -= row_height

    plt.savefig(output_png, bbox_inches='tight', facecolor='#F0F4F8')
    plt.close()
    console.print(f"[bold green]✓ Imagem do painel salva com sucesso como '{output_png}'![/bold green]")


def main():
    parser = argparse.ArgumentParser(description="Classificador e Painel de Risco Ético (CertifAIEd)")
    parser.add_argument("--coletar", action="store_true", help="Executa o modo interativo para coletar novos dados e salvar em CSV")
    parser.add_argument("--csv", type=str, default="matriz_risco.csv", help="Caminho do arquivo CSV de dados (default: matriz_risco.csv)")
    parser.add_argument("--weight-int", type=float, default=1.0, help="Peso flexível atribuído à Intensidade (default: 1.0)")
    parser.add_argument("--weight-prob", type=float, default=1.0, help="Peso flexível atribuído à Probabilidade (default: 1.0)")

    args = parser.parse_args()

    if args.coletar:
        coletar_dados_interativo(args.csv)

    df_resultado = calcular_e_processar_risco(args.csv, weight_int=args.weight_int, weight_prob=args.weight_prob)
    if df_resultado is not None:
        # 1. Exibir no terminal
        exibir_tabela_terminal(df_resultado, weight_int=args.weight_int, weight_prob=args.weight_prob)

        # 2. Gerar Excel
        gerar_planilha_excel(df_resultado, output_excel="painel_risco_certifaied.xlsx")

        # 3. Gerar Imagem Painel
        gerar_painel_imagem_png(df_resultado, output_png="painel_risco_certifaied.png", weight_int=args.weight_int, weight_prob=args.weight_prob)


if __name__ == "__main__":
    main()
