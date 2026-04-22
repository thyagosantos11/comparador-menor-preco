
import re
import threading
import os
import sys
from pathlib import Path
from tkinter import filedialog
import tkinter as tk

import customtkinter as ctk
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Tema ─────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

AZUL_ESCURO  = "#1F4E79"
AZUL_MED     = "#2E75B6"
VERDE        = "#1D9E75"
VERDE_CLARO  = "#C6EFCE"
CINZA_BG     = "#F4F7FB"
CINZA_CARD   = "#FFFFFF"
CINZA_BORDA  = "#D0DAE8"
TEXTO_ESCURO = "#1A2332"
TEXTO_MEDIO  = "#4A6080"
TEXTO_CLARO  = "#8099B3"

# ── Lógica de processamento ───────────────────────────────────────────────────
INVALIDOS = {"f", "F", "?"}
COR_HEADER    = "1F4E79"
COR_MENOR     = "C6EFCE"
COR_LINHA_PAR = "EEF3FB"

def limpar_preco(valor):
    if pd.isna(valor): return float("nan")
    s = str(valor).strip()
    if s in INVALIDOS or s == "": return float("nan")
    s = re.sub(r"\s+", "", s)
    try: return float(s)
    except ValueError: return float("nan")

def montar_nome_produto(df_raw, fornecedores):
    col_produto = df_raw.columns[1]
    nomes, prefixo = [], ""
    for _, row in df_raw.iterrows():
        nome = str(row[col_produto]).strip() if pd.notna(row[col_produto]) else ""
        if not nome or nome == "nan":
            nomes.append(None); continue
        fornecedor_cols = df_raw.columns[3:3+len(fornecedores)]
        sem_preco = all(pd.isna(row[c]) or str(row[c]).strip() in INVALIDOS | {""} for c in fornecedor_cols)
        if nome.endswith(":") and sem_preco:
            prefixo = nome + " "; nomes.append(None)
        else:
            venda = row[df_raw.columns[0]]
            eh_subitem = pd.isna(venda) and len(nome.split()) <= 2 and prefixo
            nomes.append(prefixo + nome if eh_subitem else nome)
            if not eh_subitem: prefixo = ""
    return nomes

def carregar_dados(arquivo):
    df_raw = pd.read_excel(arquivo, header=0)
    fornecedores = list(df_raw.columns[3:-1])
    df_raw.columns = ["venda", "produto_orig", "quant"] + fornecedores + ["obs"]
    df_raw = df_raw.dropna(how="all")
    df_raw["produto"] = montar_nome_produto(df_raw, fornecedores)
    df_raw = df_raw[df_raw["produto"].notna()].reset_index(drop=True)
    for f in fornecedores:
        df_raw[f] = df_raw[f].apply(limpar_preco)
    df_raw["quant"] = df_raw["quant"].astype(str).str.strip().replace("nan", "—")
    return df_raw[["produto", "quant"] + fornecedores], fornecedores

def calcular_menor(df, fornecedores):
    precos = df[fornecedores]
    sem_preco = precos.isna().all(axis=1)
    df["menor_preco"] = precos.min(axis=1)
    melhor = []
    for i, row in precos.iterrows():
        melhor.append("—" if sem_preco[i] else row.idxmin(skipna=True))
    df["melhor_forn"] = melhor
    df.loc[sem_preco, "menor_preco"] = float("nan")
    return df

def borda(cor="CCCCCC"):
    s = Side(style="thin", color=cor)
    return Border(left=s, right=s, top=s, bottom=s)

def cab(cell, texto, cor_bg=COR_HEADER):
    cell.value = texto
    cell.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    cell.fill = PatternFill("solid", start_color=cor_bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = borda("FFFFFF")

def cel(cell, valor, cor_fundo, bold=False, cor_fonte="000000", h_align="left", numero=False):
    cell.value = valor
    cell.font = Font(bold=bold, name="Arial", size=10, color=cor_fonte)
    cell.fill = PatternFill("solid", start_color=cor_fundo)
    cell.alignment = Alignment(horizontal=h_align, vertical="center")
    cell.border = borda()
    if numero and isinstance(valor, float) and not pd.isna(valor):
        cell.number_format = 'R$ #,##0.00'

def gerar_excel(df, fornecedores, arquivo_saida):
    df.to_excel(arquivo_saida, index=False)
    wb = load_workbook(arquivo_saida)
    ws = wb.active
    ws.title = "Menor Preço por Produto"
    ws.delete_rows(1, ws.max_row)
    colunas = ["Produto", "Quant."] + [f.capitalize() for f in fornecedores] + ["Menor Preço", "Melhor Fornecedor"]
    for ci, titulo in enumerate(colunas, 1):
        cab(ws.cell(row=1, column=ci), titulo)
    ws.row_dimensions[1].height = 30
    for ri, (_, row) in enumerate(df.iterrows(), start=2):
        fundo = COR_LINHA_PAR if ri % 2 == 0 else "FFFFFF"
        menor = row["menor_preco"]
        melhor = row["melhor_forn"]
        cel(ws.cell(ri, 1), row["produto"], fundo, bold=True)
        cel(ws.cell(ri, 2), row["quant"], fundo, h_align="center")
        for ci, forn in enumerate(fornecedores, start=3):
            preco = row[forn]
            eh_menor = pd.notna(preco) and pd.notna(menor) and abs(preco - menor) < 0.001
            fg = COR_MENOR if eh_menor else fundo
            cor_txt = "276221" if eh_menor else "000000"
            val = preco if pd.notna(preco) else "—"
            cel(ws.cell(ri, ci), val, fg, bold=eh_menor, cor_fonte=cor_txt, h_align="center", numero=True)
        ci_menor = len(fornecedores) + 3
        val_menor = menor if pd.notna(menor) else "—"
        cel(ws.cell(ri, ci_menor), val_menor, COR_MENOR, bold=True, cor_fonte="276221", h_align="center", numero=True)
        cel(ws.cell(ri, ci_menor+1), melhor.capitalize() if melhor != "—" else "—", fundo, h_align="center")
    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 10
    for ci in range(3, len(colunas)+1):
        ws.column_dimensions[get_column_letter(ci)].width = 13
    ws.freeze_panes = "C2"
    wb.save(arquivo_saida)

# ── Interface ─────────────────────────────────────────────────────────────────

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Comparador de Menor Preço")
        self.geometry("600x680")
        self.resizable(False, False)
        self.configure(fg_color=CINZA_BG)

        self.arquivo_entrada = None
        self.arquivo_saida   = None

        self._build_ui()

    def _build_ui(self):
        # ── Header ──
        header = ctk.CTkFrame(self, fg_color=AZUL_ESCURO, corner_radius=0, height=70)
        header.pack(fill="x")
        header.pack_propagate(False)

        ctk.CTkLabel(
            header,
            text="  📊  Comparador de Menor Preço",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color="white",
        ).pack(side="left", padx=24, pady=0)

        ctk.CTkLabel(
            header,
            text="Mercadinho Esperança",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#A8C8E8",
        ).pack(side="right", padx=24)

        # ── Corpo ──
        body = ctk.CTkFrame(self, fg_color=CINZA_BG, corner_radius=0)
        body.pack(fill="both", expand=True, padx=28, pady=24)

        # Card seleção de arquivo
        self._card_entrada(body)

        # Card saída
        self._card_saida(body)

        # Botão processar
        self.btn_processar = ctk.CTkButton(
            body,
            text="⚡  Gerar Relatório Comparativo",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            height=50,
            fg_color=AZUL_ESCURO,
            hover_color=AZUL_MED,
            corner_radius=10,
            command=self._iniciar_processamento,
            state="disabled",
        )
        self.btn_processar.pack(fill="x", pady=(4, 0))

        # Progresso
        self.progress = ctk.CTkProgressBar(body, height=6, fg_color="#D0DAE8", progress_color=VERDE)
        self.progress.pack(fill="x", pady=(12, 0))
        self.progress.set(0)

        # Log / status
        self.log_frame = ctk.CTkFrame(body, fg_color=CINZA_CARD, corner_radius=10, border_width=1, border_color=CINZA_BORDA)
        self.log_frame.pack(fill="both", expand=True, pady=(12, 0))

        ctk.CTkLabel(
            self.log_frame, text="Log de execução",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=TEXTO_CLARO,
        ).pack(anchor="w", padx=14, pady=(10, 0))

        self.log_box = ctk.CTkTextbox(
            self.log_frame,
            font=ctk.CTkFont(family="Consolas", size=12),
            fg_color=CINZA_CARD,
            text_color=TEXTO_MEDIO,
            border_width=0,
            activate_scrollbars=True,
            state="disabled",
        )
        self.log_box.pack(fill="both", expand=True, padx=8, pady=(4, 10))

        # Botão abrir resultado
        self.btn_abrir = ctk.CTkButton(
            body,
            text="📂  Abrir arquivo gerado",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            height=42,
            fg_color=VERDE,
            hover_color="#0F6E56",
            corner_radius=10,
            command=self._abrir_resultado,
            state="disabled",
        )
        self.btn_abrir.pack(fill="x", pady=(10, 0))

    def _card_entrada(self, parent):
        card = ctk.CTkFrame(parent, fg_color=CINZA_CARD, corner_radius=10, border_width=1, border_color=CINZA_BORDA)
        card.pack(fill="x", pady=(0, 12))

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=16, pady=14)

        ctk.CTkLabel(inner, text="1. Selecione a planilha de entrada",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=TEXTO_ESCURO).pack(anchor="w")
        ctk.CTkLabel(inner, text="Arquivo RELAÇÃO DE COMPRA.xlsx",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=TEXTO_CLARO).pack(anchor="w", pady=(1, 10))

        row = ctk.CTkFrame(inner, fg_color="transparent")
        row.pack(fill="x")

        self.lbl_entrada = ctk.CTkLabel(
            row, text="Nenhum arquivo selecionado",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=TEXTO_CLARO,
            fg_color="#F0F4F9", corner_radius=6,
            anchor="w",
        )
        self.lbl_entrada.pack(side="left", fill="x", expand=True, ipady=8, ipadx=10)

        ctk.CTkButton(
            row, text="Procurar...",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            width=100, height=36,
            fg_color=AZUL_ESCURO, hover_color=AZUL_MED,
            corner_radius=6,
            command=self._selecionar_entrada,
        ).pack(side="left", padx=(8, 0))

    def _card_saida(self, parent):
        card = ctk.CTkFrame(parent, fg_color=CINZA_CARD, corner_radius=10, border_width=1, border_color=CINZA_BORDA)
        card.pack(fill="x", pady=(0, 12))

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=16, pady=14)

        ctk.CTkLabel(inner, text="2. Escolha onde salvar o resultado",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=TEXTO_ESCURO).pack(anchor="w")
        ctk.CTkLabel(inner, text="O arquivo resultado_menor_preco.xlsx será salvo aqui",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=TEXTO_CLARO).pack(anchor="w", pady=(1, 10))

        row = ctk.CTkFrame(inner, fg_color="transparent")
        row.pack(fill="x")

        self.lbl_saida = ctk.CTkLabel(
            row, text="Mesma pasta do arquivo de entrada",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=TEXTO_CLARO,
            fg_color="#F0F4F9", corner_radius=6,
            anchor="w",
        )
        self.lbl_saida.pack(side="left", fill="x", expand=True, ipady=8, ipadx=10)

        ctk.CTkButton(
            row, text="Escolher...",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            width=100, height=36,
            fg_color="#5A7A9A", hover_color="#4A6A8A",
            corner_radius=6,
            command=self._selecionar_saida,
        ).pack(side="left", padx=(8, 0))

    # ── Ações ──────────────────────────────────────────────────────────────────

    def _selecionar_entrada(self):
        path = filedialog.askopenfilename(
            title="Selecionar planilha de entrada",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")]
        )
        if not path: return
        self.arquivo_entrada = path
        nome = Path(path).name
        self.lbl_entrada.configure(text=f"✓  {nome}", text_color=VERDE)

        # Sugere pasta de saída = mesma do arquivo
        pasta = str(Path(path).parent)
        saida = str(Path(path).parent / "resultado_menor_preco.xlsx")
        self.arquivo_saida = saida
        self.lbl_saida.configure(text=f"  {pasta}", text_color=TEXTO_MEDIO)

        self.btn_processar.configure(state="normal")
        self.btn_abrir.configure(state="disabled")
        self.progress.set(0)
        self._log_clear()

    def _selecionar_saida(self):
        path = filedialog.asksaveasfilename(
            title="Salvar resultado como",
            defaultextension=".xlsx",
            initialfile="resultado_menor_preco.xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if not path: return
        self.arquivo_saida = path
        self.lbl_saida.configure(text=f"  {Path(path).parent}", text_color=TEXTO_MEDIO)

    def _iniciar_processamento(self):
        if not self.arquivo_entrada or not self.arquivo_saida: return
        self.btn_processar.configure(state="disabled", text="Processando...")
        self.btn_abrir.configure(state="disabled")
        self.progress.set(0)
        self._log_clear()
        threading.Thread(target=self._processar, daemon=True).start()

    def _processar(self):
        try:
            self._log("📂 Lendo planilha...")
            self.progress.set(0.15)
            df, fornecedores = carregar_dados(self.arquivo_entrada)
            self._log(f"✓  {len(df)} produtos carregados")
            self._log(f"   Fornecedores detectados: {', '.join(f.capitalize() for f in fornecedores)}")
            self.progress.set(0.40)

            self._log("\n⚙️  Calculando menores preços...")
            df = calcular_menor(df, fornecedores)
            sem_cot = int(df["melhor_forn"].eq("—").sum())
            com_cot = len(df) - sem_cot
            self._log(f"✓  {com_cot} produtos com cotação  |  {sem_cot} sem cotação")
            self.progress.set(0.65)

            self._log("\n💾 Gerando Excel formatado...")
            gerar_excel(df, fornecedores, self.arquivo_saida)
            self.progress.set(0.95)

            self._log("\n🏆 Ranking de fornecedores:")
            contagem = df[df["melhor_forn"] != "—"]["melhor_forn"].value_counts()
            for forn, qtd in contagem.items():
                bar = "█" * min(int(qtd / max(contagem) * 20), 20)
                self._log(f"   {forn.capitalize():<14} {bar} {qtd} itens")

            self.progress.set(1.0)
            self._log(f"\n✅ Arquivo salvo em:\n   {self.arquivo_saida}")

            self.after(0, lambda: self.btn_abrir.configure(state="normal"))
        except Exception as e:
            self._log(f"\n❌ Erro: {e}")
            self.progress.set(0)
        finally:
            self.after(0, lambda: self.btn_processar.configure(
                state="normal", text="⚡  Gerar Relatório Comparativo"))

    def _abrir_resultado(self):
        if self.arquivo_saida and os.path.exists(self.arquivo_saida):
            os.startfile(self.arquivo_saida) if sys.platform == "win32" else os.system(f'xdg-open "{self.arquivo_saida}"')

    def _log(self, msg):
        def _do():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", msg + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.after(0, _do)

    def _log_clear(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")


if __name__ == "__main__":
    app = App()
    app.mainloop()
