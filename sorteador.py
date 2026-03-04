import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import random
import hashlib
import json
import os
import sys
from datetime import datetime
from pathlib import Path

# ─────────────────────────────────────────────────────────────────────────────
#  VERIFICAÇÃO DE DEPENDÊNCIAS — exibe erro amigável se faltar algo
# ─────────────────────────────────────────────────────────────────────────────
def _checar_deps():
    erros = []
    try:
        import pandas  # noqa: F401
    except ImportError:
        erros.append("pandas")
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        erros.append("openpyxl")
    try:
        import xlrd  # noqa: F401
    except ImportError:
        pass  # xlrd é opcional (para .xls antigos)
    if erros:
        import tkinter as _tk
        from tkinter import messagebox as _mb
        r = _tk.Tk(); r.withdraw()
        _mb.showerror(
            "Dependência faltando",
            f"Os pacotes a seguir não foram encontrados:\n\n"
            + "\n".join(f"  • {p}" for p in erros)
            + "\n\nAbra o Prompt de Comando e execute:\n\n"
              "  pip install pandas openpyxl xlrd\n\n"
              "Depois reabra o programa."
        )
        sys.exit(1)

_checar_deps()

import pandas as pd  # noqa: E402  (importado após checagem)

# ─────────────────────────────────────────────────────────────────────────────
#  UTILITÁRIOS DE AUDITORIA
# ─────────────────────────────────────────────────────────────────────────────

def calcular_hash_df(df: pd.DataFrame) -> str:
    """SHA-256 da planilha importada — prova que a base não foi alterada."""
    return hashlib.sha256(
        pd.util.hash_pandas_object(df, index=True).values.tobytes()
    ).hexdigest()

def gerar_semente() -> int:
    """Semente criptograficamente aleatória."""
    return int.from_bytes(os.urandom(8), "big") % (2 ** 32)

LOG_DIR = Path.home() / "Sorteador_Logs"
LOG_DIR.mkdir(exist_ok=True)

def salvar_log(registro: dict):
    log_file = LOG_DIR / "historico_sorteios.json"
    historico = []
    if log_file.exists():
        try:
            with open(log_file, "r", encoding="utf-8") as f:
                historico = json.load(f)
        except Exception:
            historico = []
    historico.append(registro)
    with open(log_file, "w", encoding="utf-8") as f:
        json.dump(historico, f, ensure_ascii=False, indent=2, default=str)

def ler_excel(path: str) -> pd.DataFrame:
    """
    Lê .xlsx ou .xls garantindo que o engine correto seja usado
    explicitamente, evitando erros de 'Import openpyxl failed'.
    """
    ext = Path(path).suffix.lower()
    if ext == ".xlsx":
        import openpyxl  # noqa: F401 — garante que o módulo está carregado
        return pd.read_excel(path, engine="openpyxl")
    elif ext == ".xls":
        try:
            import xlrd  # noqa: F401
            return pd.read_excel(path, engine="xlrd")
        except ImportError:
            raise ImportError(
                "Para abrir arquivos .xls antigos instale o xlrd:\n"
                "  pip install xlrd"
            )
    else:
        # Tenta openpyxl como padrão para extensões desconhecidas
        import openpyxl  # noqa: F401
        return pd.read_excel(path, engine="openpyxl")

# ─────────────────────────────────────────────────────────────────────────────
#  APLICAÇÃO PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

class SorteadorApp(tk.Tk):

    BG   = "#1e2130"
    CARD = "#252a3d"
    ACC  = "#4f8ef7"
    TXT  = "#e8eaf6"
    SUB  = "#8892b0"
    GRN  = "#43d18a"
    RED  = "#f76f6f"
    BORD = "#313650"

    def __init__(self):
        super().__init__()
        self.title("Sorteador de Base de Dados")
        self.resizable(False, False)
        self.configure(bg=self.BG)
        self._center(740, 690)

        self.df: pd.DataFrame | None = None
        self.hash_base = ""
        self.caminho = tk.StringVar(value="Nenhuma planilha carregada")

        self._set_icon()
        self._apply_styles()
        self._build_ui()

    def _set_icon(self):
        """Aplica icone na janela e barra de tarefas com transparencia correta."""
        base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        png_path = os.path.join(base, "icone.png")
        ico_path = os.path.join(base, "icone.ico")
        try:
            # iconphoto com PIL garante transparencia no Windows
            from PIL import Image, ImageTk
            if os.path.exists(png_path):
                pil_img = Image.open(png_path).convert("RGBA").resize((64, 64), Image.LANCZOS)
                self._icon_img = ImageTk.PhotoImage(pil_img)
                self.iconphoto(True, self._icon_img)
            # iconbitmap aplica o .ico na barra de tarefas do Windows
            if os.path.exists(ico_path):
                self.after(0, lambda: self.iconbitmap(ico_path))
        except Exception:
            pass

    # ── Helpers de layout ────────────────────────────────────────────────────

    def _center(self, w: int, h: int):
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw - w)//2}+{(sh - h)//2}")

    def _apply_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TCombobox",
            fieldbackground=self.BG, background=self.BG,
            foreground=self.TXT, selectbackground=self.ACC,
            selectforeground="white", bordercolor=self.BORD,
            arrowcolor=self.ACC)
        style.map("TCombobox",
            fieldbackground=[("readonly", self.BG)],
            foreground=[("readonly", self.TXT)])

    def _section(self, parent, titulo: str) -> tk.Frame:
        tk.Label(parent, text=titulo,
                 font=("Segoe UI", 10, "bold"),
                 bg=self.BG, fg=self.TXT
        ).pack(anchor="w", padx=24, pady=(12, 0))
        outer = tk.Frame(parent, bg=self.BORD)
        outer.pack(fill="x", padx=22, pady=(3, 0))
        inner = tk.Frame(outer, bg=self.CARD)
        inner.pack(fill="x", padx=1, pady=1, ipadx=14, ipady=10)
        return inner

    def _row(self, parent) -> tk.Frame:
        f = tk.Frame(parent, bg=self.CARD)
        f.pack(fill="x", pady=3)
        return f

    def _label(self, parent, text, width=24):
        tk.Label(parent, text=text, width=width, anchor="w",
                 bg=self.CARD, fg=self.TXT,
                 font=("Segoe UI", 9)
        ).pack(side="left")

    def _btn(self, parent, text, cmd, bg=None, fg="white",
             font_size=9, pady=6, padx=14):
        return tk.Button(parent, text=text, command=cmd,
                         bg=bg or self.ACC, fg=fg,
                         font=("Segoe UI", font_size, "bold"),
                         relief="flat", bd=0,
                         activebackground=self.ACC,
                         activeforeground="white",
                         cursor="hand2",
                         padx=padx, pady=pady)

    # ── Construção da UI ─────────────────────────────────────────────────────

    def _build_ui(self):
        # Cabeçalho
        hdr = tk.Frame(self, bg=self.BG)
        hdr.pack(fill="x", pady=(22, 4))
        tk.Label(hdr, text="🎲  Sorteador de Base de Dados",
                 font=("Segoe UI", 17, "bold"),
                 bg=self.BG, fg=self.ACC).pack()
        tk.Label(hdr, text="Sorteio auditável · Rastreabilidade completa · Reprodutível",
                 font=("Segoe UI", 9),
                 bg=self.BG, fg=self.SUB).pack(pady=(2, 0))

        # Seção 1 — Importar
        c1 = self._section(self, "📂  1. Importar Planilha")
        tk.Label(c1, textvariable=self.caminho,
                 font=("Segoe UI", 8), bg=self.CARD, fg=self.SUB,
                 wraplength=580, anchor="w"
        ).pack(fill="x", pady=(0, 6))
        btn_row = tk.Frame(c1, bg=self.CARD)
        btn_row.pack(fill="x")
        self._btn(btn_row, "  📁  Selecionar arquivo Excel (.xlsx / .xls)",
                  self._importar, pady=7, padx=18).pack(side="left")
        self.lbl_hash = tk.Label(c1, text="",
                                  font=("Courier", 7),
                                  bg=self.CARD, fg=self.SUB)
        self.lbl_hash.pack(anchor="w", pady=(5, 0))

        # Seção 2 — Configurar
        c2 = self._section(self, "⚙️  2. Configurar Sorteio")

        r1 = self._row(c2)
        self._label(r1, "Coluna denominadora:")
        self.cb_col = ttk.Combobox(r1, state="disabled", width=34, font=("Segoe UI", 9))
        self.cb_col.pack(side="left", padx=4)
        self.cb_col.bind("<<ComboboxSelected>>", self._atualizar_valores)

        r2 = self._row(c2)
        self._label(r2, "Filtrar por valor:")
        self.cb_val = ttk.Combobox(r2, state="disabled", width=34, font=("Segoe UI", 9))
        self.cb_val.pack(side="left", padx=4)
        self.cb_val.bind("<<ComboboxSelected>>", self._mostrar_total)

        self.lbl_total = tk.Label(c2, text="",
                                   bg=self.CARD, fg=self.SUB,
                                   font=("Segoe UI", 8))
        self.lbl_total.pack(anchor="w", padx=2, pady=(0, 2))

        r3 = self._row(c2)
        self._label(r3, "Qtd. pontos a sortear:")
        self.spin_qtd = tk.Spinbox(r3, from_=1, to=999999, width=12,
                                    font=("Segoe UI", 9),
                                    bg=self.BG, fg=self.TXT,
                                    insertbackground=self.TXT,
                                    relief="flat",
                                    buttonbackground=self.ACC)
        self.spin_qtd.pack(side="left", padx=4)

        r4 = self._row(c2)
        self._label(r4, "Semente (opcional):")
        self.ent_seed = tk.Entry(r4, width=22, font=("Segoe UI", 9),
                                  bg=self.BG, fg=self.TXT,
                                  insertbackground=self.TXT, relief="flat")
        self.ent_seed.pack(side="left", padx=4)
        tk.Label(r4, text="(vazio = gerada automaticamente)",
                 bg=self.CARD, fg=self.SUB,
                 font=("Segoe UI", 7)).pack(side="left", padx=4)

        # Seção 3 — Executar
        c3 = self._section(self, "🎯  3. Executar Sorteio")
        self._btn(c3, "  ▶  SORTEAR AGORA  ",
                  self._sortear, bg=self.GRN,
                  font_size=11, pady=10, padx=28
        ).pack(pady=6)
        self.lbl_status = tk.Label(c3, text="",
                                    font=("Segoe UI", 9),
                                    bg=self.CARD, fg=self.TXT,
                                    wraplength=600)
        self.lbl_status.pack(pady=(2, 0))

        # Rodapé — logs
        tk.Label(self,
                 text=f"📋 Logs de auditoria: {LOG_DIR}",
                 font=("Segoe UI", 7),
                 bg=self.BG, fg=self.SUB
        ).pack(pady=(10, 2))

        # Rodapé — LinkedIn
        sep = tk.Frame(self, bg=self.BORD, height=1)
        sep.pack(fill="x", padx=22, pady=(4, 4))

        linkedin_frame = tk.Frame(self, bg=self.BG)
        linkedin_frame.pack(pady=(0, 8))

        tk.Label(linkedin_frame,
                 text="Desenvolvido por  ",
                 font=("Segoe UI", 7),
                 bg=self.BG, fg=self.SUB
        ).pack(side="left")

        lnk = tk.Label(linkedin_frame,
                       text="Pablo Bernar",
                       font=("Segoe UI", 7, "underline"),
                       bg=self.BG, fg=self.ACC,
                       cursor="hand2")
        lnk.pack(side="left")
        lnk.bind("<Button-1>", lambda e: __import__("webbrowser").open(
            "https://www.linkedin.com/in/pablo-bernar/"))

    # ── Lógica: importar ─────────────────────────────────────────────────────

    def _importar(self):
        path = filedialog.askopenfilename(
            title="Selecionar planilha Excel",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        if not path:
            return
        try:
            df = ler_excel(path)          # usa engine explícito
            if df.empty:
                raise ValueError("A planilha está vazia.")
            self.df = df
            self.hash_base = calcular_hash_df(df)
            self.caminho.set(path)
            self.lbl_hash.config(
                text=f"🔒 SHA-256: {self.hash_base}",
                fg=self.GRN
            )
            cols = list(df.columns)
            self.cb_col.config(state="readonly", values=cols)
            self.cb_col.set("")
            self.cb_val.config(state="disabled", values=[])
            self.cb_val.set("")
            self.lbl_total.config(text="")
            self._status(
                f"✅  Planilha carregada — {len(df):,} linhas × {len(cols)} colunas",
                self.GRN
            )
        except Exception as exc:
            messagebox.showerror("Erro ao abrir planilha", str(exc))

    # ── Lógica: colunas / valores ─────────────────────────────────────────────

    def _atualizar_valores(self, _=None):
        col = self.cb_col.get()
        if not col or self.df is None:
            return
        vals = sorted(self.df[col].dropna().unique().tolist(), key=str)
        self.cb_val.config(state="readonly", values=[str(v) for v in vals])
        self.cb_val.set("")
        self.lbl_total.config(text="")

    def _mostrar_total(self, _=None):
        col, val = self.cb_col.get(), self.cb_val.get()
        if col and val and self.df is not None:
            n = len(self.df[self.df[col].astype(str) == val])
            self.lbl_total.config(
                text=f"   ➜  {n:,} registros disponíveis para sorteio",
                fg=self.SUB
            )

    # ── Lógica: sortear ───────────────────────────────────────────────────────

    def _sortear(self):
        if self.df is None:
            messagebox.showwarning("Atenção", "Importe uma planilha primeiro.")
            return
        col = self.cb_col.get()
        val = self.cb_val.get()
        if not col:
            messagebox.showwarning("Atenção", "Selecione a coluna denominadora.")
            return
        if not val:
            messagebox.showwarning("Atenção", "Selecione o valor para filtrar.")
            return
        try:
            qtd = int(self.spin_qtd.get())
            if qtd <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Atenção", "Informe uma quantidade válida (inteiro > 0).")
            return

        df_filtrado = self.df[self.df[col].astype(str) == val].copy()
        if df_filtrado.empty:
            messagebox.showerror("Erro", "Nenhum registro encontrado com esse filtro.")
            return
        if qtd > len(df_filtrado):
            messagebox.showerror(
                "Erro",
                f"Você pediu {qtd:,} pontos, mas só há {len(df_filtrado):,} "
                f"disponíveis para '{val}' na coluna '{col}'."
            )
            return

        # Semente
        s = self.ent_seed.get().strip()
        if s:
            try:
                semente = int(s)
            except ValueError:
                semente = int(hashlib.md5(s.encode()).hexdigest(), 16) % (2 ** 32)
        else:
            semente = gerar_semente()

        # Sorteio sem reposição
        rng = random.Random(semente)
        indices_sorteados = rng.sample(df_filtrado.index.tolist(), qtd)

        # Monta saídas
        df_completo = self.df.copy()
        df_completo["SORTEADO"] = ""
        df_completo.loc[indices_sorteados, "SORTEADO"] = "SORTEADO"
        df_apenas_sorteados = df_completo.loc[indices_sorteados].copy()

        # Pasta de destino
        pasta = filedialog.askdirectory(title="Escolha onde salvar os arquivos de resultado")
        if not pasta:
            return

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        path_completo  = os.path.join(pasta, f"sorteio_completo_{ts}.xlsx")
        path_sorteados = os.path.join(pasta, f"sorteio_selecionados_{ts}.xlsx")

        import openpyxl  # noqa: F401 — garante engine disponível na hora de salvar
        with pd.ExcelWriter(path_completo, engine="openpyxl") as w:
            df_completo.to_excel(w, index=False, sheet_name="Base Completa")
        with pd.ExcelWriter(path_sorteados, engine="openpyxl") as w:
            df_apenas_sorteados.to_excel(w, index=False, sheet_name="Sorteados")

        # Log de auditoria
        salvar_log({
            "data_hora"         : datetime.now().isoformat(timespec="seconds"),
            "arquivo_origem"    : self.caminho.get(),
            "hash_base_sha256"  : self.hash_base,
            "coluna_filtro"     : col,
            "valor_filtro"      : val,
            "total_disponivel"  : len(df_filtrado),
            "qtd_sorteada"      : qtd,
            "semente"           : semente,
            "indices_sorteados" : indices_sorteados,
            "arquivo_completo"  : path_completo,
            "arquivo_sorteados" : path_sorteados,
        })

        # ── Relatório .txt (mesmo conteúdo da tela) ──────────────────────────
        path_relatorio = os.path.join(pasta, f"relatorio_sorteio_{ts}.txt")
        relatorio = (
            f"============================================================\n"
            f"  RELATÓRIO DE SORTEIO\n"
            f"============================================================\n"
            f"\n"
            f"  Data e hora        : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n"
            f"\n"
            f"------------------------------------------------------------\n"
            f"  BASE DE DADOS\n"
            f"------------------------------------------------------------\n"
            f"  Arquivo            : {self.caminho.get()}\n"
            f"  SHA-256 da base    : {self.hash_base}\n"
            f"\n"
            f"------------------------------------------------------------\n"
            f"  CONFIGURAÇÃO DO SORTEIO\n"
            f"------------------------------------------------------------\n"
            f"  Coluna filtro      : {col}\n"
            f"  Valor filtro       : {val}\n"
            f"  Total disponível   : {len(df_filtrado):,} registros\n"
            f"  Quantidade sorteada: {qtd:,} pontos\n"
            f"\n"
            f"------------------------------------------------------------\n"
            f"  IDENTIFICAÇÃO DO SORTEIO\n"
            f"------------------------------------------------------------\n"
            f"  Semente            : {semente}\n"
            f"  (Use este número para reproduzir o sorteio identicamente)\n"
            f"\n"
            f"------------------------------------------------------------\n"
            f"  ARQUIVOS GERADOS\n"
            f"------------------------------------------------------------\n"
            f"  Base completa      : sorteio_completo_{ts}.xlsx\n"
            f"  Apenas sorteados   : sorteio_selecionados_{ts}.xlsx\n"
            f"  Este relatório     : relatorio_sorteio_{ts}.txt\n"
            f"\n"
            f"------------------------------------------------------------\n"
            f"  LOG DE AUDITORIA\n"
            f"------------------------------------------------------------\n"
            f"  {LOG_DIR / 'historico_sorteios.json'}\n"
            f"\n"
            f"============================================================\n"
        )
        with open(path_relatorio, "w", encoding="utf-8") as f:
            f.write(relatorio)

        self._status(
            f"✅  {qtd:,} pontos sorteados!   🔑 Semente: {semente}",
            self.GRN
        )
        messagebox.showinfo(
            "Sorteio concluído!",
            f"✅  {qtd:,} pontos sorteados de {len(df_filtrado):,} disponíveis\n"
            f"    Filtro: {col} = '{val}'\n\n"
            f"📁  Base completa:\n    sorteio_completo_{ts}.xlsx\n\n"
            f"📋  Apenas sorteados:\n    sorteio_selecionados_{ts}.xlsx\n\n"
            f"📄  Relatório:\n    relatorio_sorteio_{ts}.txt\n\n"
            f"──────────────────────────────────\n"
            f"🔑  Semente: {semente}\n"
            f"🔒  SHA-256: {self.hash_base[:40]}…\n\n"
            f"📋  Log: {LOG_DIR / 'historico_sorteios.json'}"
        )

    def _status(self, msg: str, cor: str = ""):
        self.lbl_status.config(text=msg, fg=cor or self.TXT)


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    SorteadorApp().mainloop()
