"""
PBIX Analyzer – Interface Gráfica
Selecione um arquivo .pbix, clique em Gerar e a documentação .docx
será salva na mesma pasta do executável (ou ao lado do .pbix).
"""

import json
import os
import sys
import shutil
import threading
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from datetime import datetime
from pathlib import Path

# ── Localiza a pasta do executável / script ────────────────────────────────────
if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).parent          # pasta do .exe
else:
    APP_DIR = Path(__file__).resolve().parent      # pasta do .py

CONFIG_FILE = APP_DIR / "config.json"

# ── Adiciona APP_DIR ao path para importar pbix_analyzer ─────────────────────
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

# Importa toda a lógica de análise
from pbix_analyzer import (
    extract_pbix,
    parse_layout,
    parse_diagram_layout,
    parse_connections,
    parse_metadata,
    build_analysis,
    generate_docx,
)

# Importa o analisador de notebooks
from ipynb_analyzer import (
    parse_notebook,
    analyze_notebook,
    generate_notebook_docx,
)

# ── Paletas de tema ───────────────────────────────────────────────────────────
THEME_PBIX = {
    "accent":    "#FFBE00",   # amarelo Construction
    "accent_hover": "#E6AB00",
    "text_on_accent": "#1A1A1A",
    "log_text":  "#FFBE00",
    "title":     "📊  PBIX Analyzer",
    "subtitle":  "Gerador automático de documentação para relatórios Power BI",
}
THEME_DATABRICKS = {
    "accent":    "#D14243",   # vermelho Databricks
    "accent_hover": "#B83535",
    "text_on_accent": "#FFFFFF",
    "log_text":  "#FF8080",
    "title":     "🧱  Databricks .ipynb",
    "subtitle":  "Gerador automático de documentação para notebooks Databricks",
}

JD_GREEN      = THEME_PBIX["accent"]
JD_GREEN_DARK = "#1A1A1A"
JD_TEXT_LIGHT = "#1A1A1A"
JD_TEXT_DARK  = "#1A1A1A"
BG_SURFACE    = "#F5F5F5"
BG_LOG        = "#1A1A1A"
TEXT_LOG      = "#FFBE00"
BTN_HOVER     = "#E6AB00"


# ══════════════════════════════════════════════════════════════════════════════
#  LÓGICA DE GERAÇÃO (roda em thread separada)
# ══════════════════════════════════════════════════════════════════════════════

def run_analysis(pbix_path: str, output_path: str, log_fn, done_fn, error_fn,
                 llm_config=None):
    """Executa toda a análise em uma thread separada para não travar a GUI."""
    extract_dir = None
    try:
        log_fn("🔓  Extraindo arquivo .pbix...")
        extract_dir = extract_pbix(pbix_path)

        log_fn("📋  Parseando Layout do Relatório...")
        report = parse_layout(extract_dir)

        log_fn("🗂️   Parseando tabelas do modelo...")
        tables = parse_diagram_layout(extract_dir)

        log_fn("🔌  Lendo informações de conexão...")
        connections = parse_connections(extract_dir)
        metadata    = parse_metadata(extract_dir)

        log_fn("🔍  Analisando e classificando estrutura...")
        analysis = build_analysis(report, tables, connections, metadata)

        n_pages   = len(report["pages"])
        n_main    = len(analysis["main_pages"])
        n_visuals = analysis["total_visuals"]
        n_fields  = len(analysis["all_fields"])
        n_tables  = len(tables)
        log_fn(f"     → {n_pages} páginas ({n_main} principais) | "
               f"{n_visuals} visuais | {n_fields} campos | {n_tables} tabelas")

        log_fn("📝  Gerando documentação Word...")
        pbix_name = Path(pbix_path).stem
        generate_docx(pbix_name, report, analysis, output_path,
                      llm_config=llm_config, log_callback=log_fn)

        done_fn(output_path)

    except Exception as exc:
        error_fn(str(exc))

    finally:
        if extract_dir:
            shutil.rmtree(extract_dir, ignore_errors=True)


def run_notebook_analysis(ipynb_path: str, output_path: str, log_fn, done_fn, error_fn,
                          llm_config=None):
    """Executa análise de notebook .ipynb em thread separada."""
    try:
        log_fn("📂  Lendo arquivo .ipynb...")
        notebook = parse_notebook(ipynb_path)

        log_fn(f"     → {len(notebook['cells'])} células  |  kernel: {notebook['kernel']}")

        log_fn("🔍  Analisando estrutura do notebook...")
        analysis = analyze_notebook(notebook)

        log_fn(f"     → {analysis['sql_cells']} SQL  |  "
               f"{len(analysis['imports'])} imports  |  "
               f"{len(analysis['data_sources'])} fontes de dados")

        log_fn("📝  Gerando documentação Word...")
        nb_name = Path(ipynb_path).stem
        generate_notebook_docx(nb_name, notebook, analysis, output_path,
                               llm_config=llm_config, log_callback=log_fn)

        done_fn(output_path)

    except Exception as exc:
        error_fn(str(exc))


# ══════════════════════════════════════════════════════════════════════════════
#  JANELA PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PBIX Analyzer  ·  Gerador de Documentação")
        self.resizable(False, False)
        self.configure(bg=BG_SURFACE)

        # Variáveis principais — PBIX
        self._pbix_var   = tk.StringVar()
        self._output_var = tk.StringVar(value=str(APP_DIR))
        self._running    = False

        # Variáveis principais — Notebook
        self._ipynb_var      = tk.StringVar()
        self._output_nb_var  = tk.StringVar(value=str(APP_DIR))
        self._running_nb     = False

        # Variáveis LLM
        self._llm_provider = tk.StringVar(value="Desabilitado")
        self._llm_endpoint = tk.StringVar()
        self._llm_apikey   = tk.StringVar()
        self._llm_model    = tk.StringVar(value="gpt-4o")

        self._load_config()
        self._build_ui()
        self._center_window(760, 620)

    # ── Construção da interface ───────────────────────────────────────────────
    def _build_ui(self):
        # ── Cabeçalho ────────────────────────────────────────────────────────
        self._header = tk.Frame(self, bg=THEME_PBIX["accent"], height=90)
        self._header.pack(fill="x")
        self._header.pack_propagate(False)

        self._header_title = tk.Label(
            self._header,
            text=THEME_PBIX["title"],
            font=("Segoe UI", 20, "bold"),
            bg=THEME_PBIX["accent"], fg=THEME_PBIX["text_on_accent"],
        )
        self._header_title.place(x=24, y=14)

        self._header_subtitle = tk.Label(
            self._header,
            text=THEME_PBIX["subtitle"],
            font=("Segoe UI", 10),
            bg=THEME_PBIX["accent"], fg=THEME_PBIX["text_on_accent"],
        )
        self._header_subtitle.place(x=26, y=52)

        # Faixa preta fina
        self._accent_bar = tk.Frame(self, bg="#1A1A1A", height=4)
        self._accent_bar.pack(fill="x")

        # ── Notebook com abas ────────────────────────────────────────────────
        style = ttk.Style()
        style.configure("TNotebook", background=BG_SURFACE)
        style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=[10, 4])

        self._nb = ttk.Notebook(self)
        self._nb.pack(fill="both", expand=True)

        tab_analyze  = tk.Frame(self._nb, bg=BG_SURFACE)
        tab_notebook = tk.Frame(self._nb, bg=BG_SURFACE)
        tab_llm      = tk.Frame(self._nb, bg=BG_SURFACE)
        self._nb.add(tab_analyze,  text="📄  PBIX")
        self._nb.add(tab_notebook, text="📓  Databricks .ipynb")
        self._nb.add(tab_llm,      text="⚙️  Config LLM")

        self._build_analyze_tab(tab_analyze)
        self._build_notebook_tab(tab_notebook)
        self._build_llm_tab(tab_llm)

        self._nb.bind("<<NotebookTabChanged>>", self._on_tab_change)

        # Rodapé
        self._footer = tk.Label(
            self,
            text=f"Data & Analytics Squad  ·  PBIX Analyzer  ·  {datetime.now().year}",
            font=("Segoe UI", 8),
            bg=JD_GREEN_DARK, fg=THEME_PBIX["accent"],
            pady=5,
        )
        self._footer.pack(fill="x", side="bottom")

    def _on_tab_change(self, event=None):
        """Muda o tema do app conforme a aba selecionada."""
        idx = self._nb.index(self._nb.select())
        theme = THEME_DATABRICKS if idx == 1 else THEME_PBIX

        # Cabeçalho
        self._header.config(bg=theme["accent"])
        self._header_title.config(
            bg=theme["accent"], fg=theme["text_on_accent"],
            text=theme["title"],
        )
        self._header_subtitle.config(
            bg=theme["accent"], fg=theme["text_on_accent"],
            text=theme["subtitle"],
        )

        # Faixa de acento e rodapé
        self._accent_bar.config(bg="#1A1A1A")
        self._footer.config(fg=theme["accent"])

        # Botão da aba de notebook
        if hasattr(self, "_btn_run_nb"):
            self._btn_run_nb.config(
                bg=theme["accent"],
                fg=theme["text_on_accent"],
                activebackground=theme["accent_hover"],
                activeforeground=theme["text_on_accent"],
            )

        # Log do notebook (cor do texto)
        if hasattr(self, "_log_text_nb"):
            self._log_text_nb.config(fg=theme["log_text"])

    def _build_analyze_tab(self, parent):
        """Conteúdo da aba principal de análise."""
        body = tk.Frame(parent, bg=BG_SURFACE, padx=24, pady=18)
        body.pack(fill="both", expand=True)

        # Arquivo PBIX
        self._section_label(body, "Arquivo .pbix")
        row1 = tk.Frame(body, bg=BG_SURFACE)
        row1.pack(fill="x", pady=(4, 12))

        self._entry_pbix = tk.Entry(
            row1, textvariable=self._pbix_var,
            font=("Segoe UI", 10), relief="solid", bd=1,
            bg="white", fg=JD_TEXT_DARK,
        )
        self._entry_pbix.pack(side="left", fill="x", expand=True)

        tk.Button(
            row1, text="  Procurar…  ",
            font=("Segoe UI", 9, "bold"),
            bg=JD_GREEN, fg=JD_TEXT_LIGHT,
            relief="flat", cursor="hand2",
            activebackground=BTN_HOVER, activeforeground=JD_TEXT_LIGHT,
            command=self._browse_pbix,
        ).pack(side="left", padx=(8, 0))

        # Pasta de saída
        self._section_label(body, "Pasta de destino do .docx")
        row2 = tk.Frame(body, bg=BG_SURFACE)
        row2.pack(fill="x", pady=(4, 16))

        self._entry_out = tk.Entry(
            row2, textvariable=self._output_var,
            font=("Segoe UI", 10), relief="solid", bd=1,
            bg="white", fg=JD_TEXT_DARK,
        )
        self._entry_out.pack(side="left", fill="x", expand=True)

        tk.Button(
            row2, text="  Alterar…  ",
            font=("Segoe UI", 9, "bold"),
            bg="#6c757d", fg=JD_TEXT_LIGHT,
            relief="flat", cursor="hand2",
            activebackground="#5a6268", activeforeground=JD_TEXT_LIGHT,
            command=self._browse_output,
        ).pack(side="left", padx=(8, 0))

        # Separador
        ttk.Separator(body, orient="horizontal").pack(fill="x", pady=(0, 14))

        # Botão GERAR
        self._btn_run = tk.Button(
            body,
            text="▶   GERAR DOCUMENTAÇÃO",
            font=("Segoe UI", 13, "bold"),
            bg=JD_GREEN, fg=JD_TEXT_LIGHT,
            relief="flat", cursor="hand2",
            activebackground=BTN_HOVER, activeforeground=JD_TEXT_LIGHT,
            pady=10,
            command=self._start,
        )
        self._btn_run.pack(fill="x")

        # Barra de progresso
        self._progress = ttk.Progressbar(body, mode="indeterminate", length=400)
        self._progress.pack(fill="x", pady=(10, 6))

        # Log
        self._section_label(body, "Log de execução")
        log_frame = tk.Frame(body, bg=BG_LOG, bd=1, relief="solid")
        log_frame.pack(fill="both", expand=True, pady=(4, 0))

        self._log_text = tk.Text(
            log_frame,
            font=("Consolas", 9),
            bg=BG_LOG, fg=TEXT_LOG,
            relief="flat", bd=0,
            state="disabled",
            height=8,
            wrap="word",
        )
        self._log_text.pack(fill="both", expand=True, padx=6, pady=6)

    def _build_notebook_tab(self, parent):
        """Conteúdo da aba de análise de notebooks .ipynb."""
        body = tk.Frame(parent, bg=BG_SURFACE, padx=24, pady=18)
        body.pack(fill="both", expand=True)

        # Arquivo .ipynb
        self._section_label(body, "Arquivo .ipynb  (Notebook Jupyter / Databricks)")
        row1 = tk.Frame(body, bg=BG_SURFACE)
        row1.pack(fill="x", pady=(4, 12))

        self._entry_ipynb = tk.Entry(
            row1, textvariable=self._ipynb_var,
            font=("Segoe UI", 10), relief="solid", bd=1,
            bg="white", fg=JD_TEXT_DARK,
        )
        self._entry_ipynb.pack(side="left", fill="x", expand=True)

        tk.Button(
            row1, text="  Procurar…  ",
            font=("Segoe UI", 9, "bold"),
            bg=JD_GREEN, fg=JD_TEXT_LIGHT,
            relief="flat", cursor="hand2",
            activebackground=BTN_HOVER, activeforeground=JD_TEXT_LIGHT,
            command=self._browse_ipynb,
        ).pack(side="left", padx=(8, 0))

        # Pasta de saída
        self._section_label(body, "Pasta de destino do .docx")
        row2 = tk.Frame(body, bg=BG_SURFACE)
        row2.pack(fill="x", pady=(4, 16))

        self._entry_out_nb = tk.Entry(
            row2, textvariable=self._output_nb_var,
            font=("Segoe UI", 10), relief="solid", bd=1,
            bg="white", fg=JD_TEXT_DARK,
        )
        self._entry_out_nb.pack(side="left", fill="x", expand=True)

        tk.Button(
            row2, text="  Alterar…  ",
            font=("Segoe UI", 9, "bold"),
            bg="#6c757d", fg=JD_TEXT_LIGHT,
            relief="flat", cursor="hand2",
            activebackground="#5a6268", activeforeground=JD_TEXT_LIGHT,
            command=self._browse_output_nb,
        ).pack(side="left", padx=(8, 0))

        # Separador
        ttk.Separator(body, orient="horizontal").pack(fill="x", pady=(0, 14))

        # Botão GERAR
        self._btn_run_nb = tk.Button(
            body,
            text="▶   GERAR DOCUMENTAÇÃO DO NOTEBOOK",
            font=("Segoe UI", 13, "bold"),
            bg=JD_GREEN, fg=JD_TEXT_LIGHT,
            relief="flat", cursor="hand2",
            activebackground=BTN_HOVER, activeforeground=JD_TEXT_LIGHT,
            pady=10,
            command=self._start_notebook,
        )
        self._btn_run_nb.pack(fill="x")

        # Barra de progresso
        self._progress_nb = ttk.Progressbar(body, mode="indeterminate", length=400)
        self._progress_nb.pack(fill="x", pady=(10, 6))

        # Log
        self._section_label(body, "Log de execução")
        log_frame = tk.Frame(body, bg=BG_LOG, bd=1, relief="solid")
        log_frame.pack(fill="both", expand=True, pady=(4, 0))

        self._log_text_nb = tk.Text(
            log_frame,
            font=("Consolas", 9),
            bg=BG_LOG, fg=TEXT_LOG,
            relief="flat", bd=0,
            state="disabled",
            height=8,
            wrap="word",
        )
        self._log_text_nb.pack(fill="both", expand=True, padx=6, pady=6)

    def _build_llm_tab(self, parent):
        """Conteúdo da aba de configuração LLM."""
        body = tk.Frame(parent, bg=BG_SURFACE, padx=24, pady=18)
        body.pack(fill="both", expand=True)

        self._section_label(body, "Configuração do Provedor LLM")
        tk.Label(
            body,
            text="Configure uma API de linguagem para enriquecer as narrativas com IA. "
                 "Deixe 'Desabilitado' para usar apenas as heurísticas automáticas.",
            font=("Segoe UI", 9), bg=BG_SURFACE, fg="#555555",
            wraplength=680, justify="left",
        ).pack(anchor="w", pady=(2, 12))

        LBL_W = 22

        # Provedor
        row_prov = tk.Frame(body, bg=BG_SURFACE)
        row_prov.pack(fill="x", pady=(0, 8))
        tk.Label(row_prov, text="Provedor:", font=("Segoe UI", 9), bg=BG_SURFACE,
                 fg=JD_GREEN_DARK, width=LBL_W, anchor="w").pack(side="left")
        providers = ["Desabilitado", "GitHub Models", "Ollama (Local)", "Azure OpenAI", "OpenAI"]
        cb = ttk.Combobox(row_prov, textvariable=self._llm_provider,
                          values=providers, state="readonly", width=26)
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", self._on_provider_change)

        # Endpoint URL (Azure only)
        self._row_endpoint = tk.Frame(body, bg=BG_SURFACE)
        tk.Label(self._row_endpoint, text="Endpoint URL:", font=("Segoe UI", 9),
                 bg=BG_SURFACE, fg=JD_GREEN_DARK, width=LBL_W, anchor="w").pack(side="left")
        tk.Entry(self._row_endpoint, textvariable=self._llm_endpoint,
                 font=("Segoe UI", 9), relief="solid", bd=1,
                 bg="white", fg=JD_TEXT_DARK, width=50).pack(side="left", fill="x", expand=True)

        # API Key (oculta para Ollama)
        self._row_key = tk.Frame(body, bg=BG_SURFACE)
        tk.Label(self._row_key, text="API Key:", font=("Segoe UI", 9), bg=BG_SURFACE,
                 fg=JD_GREEN_DARK, width=LBL_W, anchor="w").pack(side="left")
        tk.Entry(self._row_key, textvariable=self._llm_apikey, show="*",
                 font=("Segoe UI", 9), relief="solid", bd=1,
                 bg="white", fg=JD_TEXT_DARK, width=50).pack(side="left", fill="x", expand=True)

        # Modelo / Deployment
        row_model = tk.Frame(body, bg=BG_SURFACE)
        row_model.pack(fill="x", pady=(0, 8))
        tk.Label(row_model, text="Modelo/Deployment:", font=("Segoe UI", 9),
                 bg=BG_SURFACE, fg=JD_GREEN_DARK, width=LBL_W, anchor="w").pack(side="left")
        tk.Entry(row_model, textvariable=self._llm_model,
                 font=("Segoe UI", 9), relief="solid", bd=1,
                 bg="white", fg=JD_TEXT_DARK, width=28).pack(side="left")
        self._lbl_model_hint = tk.Label(
            row_model, text="", font=("Segoe UI", 8),
            bg=BG_SURFACE, fg="#888888")
        self._lbl_model_hint.pack(side="left", padx=(8, 0))

        # Dica Ollama
        self._frame_ollama_tip = tk.Frame(body, bg="#FFF8DC", bd=1, relief="solid")
        tk.Label(
            self._frame_ollama_tip,
            text=(
                "📦  Como usar o Ollama (LLM 100% local, grátis, sem internet):\n"
                "1. Instale o Ollama: baixe em https://ollama.com  (já foi baixado se você pediu)\n"
                "2. Abra o terminal e execute:  ollama pull llama3.2:3b\n"
                "   (baixa ~2 GB uma única vez — use llama3.1:8b para maior qualidade)\n"
                "3. O Ollama fica rodando em segundo plano automaticamente após a instalação.\n"
                "4. No campo Modelo acima, coloque:  llama3.2:3b  (ou o modelo que baixou)\n"
                "5. Clique 'Testar Conexão' para confirmar."
            ),
            font=("Segoe UI", 8), bg="#FFF8DC", fg="#5a4000",
            justify="left", wraplength=660, padx=10, pady=8,
        ).pack(anchor="w")

        # Dica GitHub Models
        self._frame_github_tip = tk.Frame(body, bg="#E6F4EA", bd=1, relief="solid")
        tk.Label(
            self._frame_github_tip,
            text=(
                "🐙  Como usar o GitHub Models (GPT-4o grátis com sua conta GitHub):\n"
                "• Acesse: https://github.com/marketplace/models  e aceite os termos.\n"
                "• Gere um Token em: github.com → Settings → Developer Settings → Personal Access Tokens\n"
                "  Permissão necessária: Models (read) — ou use um token clássico.\n"
                "• Cole o token no campo 'API Key' acima (começa com ghp_ ou github_pat_).\n"
                "• Se você já usa o GitHub Copilot CLI, o token pode já estar em GH_TOKEN — deixe vazio."
            ),
            font=("Segoe UI", 8), bg="#E6F4EA", fg="#1a4a2e",
            justify="left", wraplength=660, padx=10, pady=8,
        ).pack(anchor="w")

        # Botões
        btn_frame = tk.Frame(body, bg=BG_SURFACE)
        btn_frame.pack(fill="x", pady=(10, 0))
        tk.Button(
            btn_frame, text="💾 Salvar Configuração",
            font=("Segoe UI", 10, "bold"),
            bg=JD_GREEN, fg=JD_TEXT_LIGHT,
            relief="flat", cursor="hand2",
            activebackground=BTN_HOVER, activeforeground=JD_TEXT_LIGHT,
            padx=12, pady=6,
            command=self._save_config,
        ).pack(side="left", padx=(0, 10))

        tk.Button(
            btn_frame, text="🔍 Testar Conexão",
            font=("Segoe UI", 10, "bold"),
            bg="#6c757d", fg=JD_TEXT_LIGHT,
            relief="flat", cursor="hand2",
            activebackground="#5a6268", activeforeground=JD_TEXT_LIGHT,
            padx=12, pady=6,
            command=self._test_llm,
        ).pack(side="left")

        # Status
        self._llm_status = tk.Label(
            body, text="",
            font=("Segoe UI", 9),
            bg=BG_SURFACE, fg="#2e7d32",
            wraplength=680, justify="left",
        )
        self._llm_status.pack(anchor="w", pady=(6, 0))

        # Aplica visibilidade inicial
        self._on_provider_change()

    def _section_label(self, parent, text):
        tk.Label(
            parent, text=text,
            font=("Segoe UI", 9, "bold"),
            bg=BG_SURFACE, fg=JD_GREEN_DARK,
        ).pack(anchor="w")

    # ── Centraliza a janela na tela ──────────────────────────────────────────
    def _center_window(self, w, h):
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x  = (sw - w) // 2
        y  = (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ── Ações dos botões ─────────────────────────────────────────────────────
    def _browse_pbix(self):
        path = filedialog.askopenfilename(
            title="Selecione o arquivo .pbix",
            filetypes=[("Power BI Desktop", "*.pbix"), ("Todos os arquivos", "*.*")],
        )
        if path:
            self._pbix_var.set(path)
            # Sugere pasta de saída = pasta do pbix
            self._output_var.set(str(Path(path).parent))

    def _browse_output(self):
        folder = filedialog.askdirectory(title="Selecione a pasta de destino")
        if folder:
            self._output_var.set(folder)

    # ── LLM config helpers ───────────────────────────────────────────────────
    def _on_provider_change(self, event=None):
        """Mostra/oculta campos conforme o provedor selecionado."""
        prov = self._llm_provider.get()

        # Endpoint só para Azure
        if prov == "Azure OpenAI":
            self._row_endpoint.pack(fill="x", pady=(0, 8))
        else:
            self._row_endpoint.pack_forget()

        # API Key oculta para Ollama; opcional para GitHub Models
        if prov == "Ollama (Local)":
            self._row_key.pack_forget()
            self._frame_ollama_tip.pack(fill="x", pady=(4, 10))
            self._frame_github_tip.pack_forget()
            if not self._llm_model.get():
                self._llm_model.set("llama3.2:3b")
            self._lbl_model_hint.config(
                text="(recomendado: llama3.2:3b  ou  llama3.1:8b)")
        elif prov == "GitHub Models":
            self._row_key.pack(fill="x", pady=(0, 8))
            self._frame_ollama_tip.pack_forget()
            self._frame_github_tip.pack(fill="x", pady=(4, 10))
            if not self._llm_model.get() or self._llm_model.get() in ("llama3.2:3b", "llama3.1:8b"):
                self._llm_model.set("gpt-4o")
            self._lbl_model_hint.config(
                text="(sugestão: gpt-4o  ou  gpt-4o-mini  |  API Key opcional se GH_TOKEN definido)")
        else:
            self._row_key.pack(fill="x", pady=(0, 8))
            self._frame_ollama_tip.pack_forget()
            self._frame_github_tip.pack_forget()
            if prov in ("Azure OpenAI", "OpenAI"):
                if not self._llm_model.get() or self._llm_model.get() in ("llama3.2:3b","llama3.1:8b"):
                    self._llm_model.set("gpt-4o")
                self._lbl_model_hint.config(
                    text="(sugestão: gpt-4o  ou  gpt-4o-mini)")
            else:
                self._lbl_model_hint.config(text="")

    def _get_llm_config(self) -> dict:
        return {
            "provider": self._llm_provider.get(),
            "endpoint": self._llm_endpoint.get().strip(),
            "api_key":  self._llm_apikey.get().strip(),
            "model":    self._llm_model.get().strip() or "gpt-4o",
        }

    def _save_config(self):
        cfg = self._get_llm_config()
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2, ensure_ascii=False)
            self._llm_status.config(
                text="✅ Configuração salva com sucesso.", fg="#2e7d32")
        except Exception as e:
            self._llm_status.config(text=f"❌ Erro ao salvar: {e}", fg="#c62828")

    def _load_config(self):
        try:
            if CONFIG_FILE.exists():
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                self._llm_provider.set(cfg.get("provider", "Desabilitado"))
                self._llm_endpoint.set(cfg.get("endpoint", ""))
                self._llm_apikey.set(cfg.get("api_key", ""))
                self._llm_model.set(cfg.get("model", "gpt-4o"))
        except Exception:
            pass

    def _test_llm(self):
        from pbix_analyzer import call_llm
        cfg = self._get_llm_config()
        if cfg.get("provider") == "Desabilitado":
            self._llm_status.config(
                text="ℹ️ LLM desabilitado. Selecione um provedor para testar.", fg="#e65100")
            return
        self._llm_status.config(text="⏳ Testando conexão...", fg="#1565c0")
        self.update()

        def _do_test():
            result = call_llm("Responda apenas: OK", cfg)
            def _show():
                if result:
                    self._llm_status.config(
                        text=f"✅ Conexão OK. Resposta: {result[:80]}", fg="#2e7d32")
                else:
                    self._llm_status.config(
                        text="❌ Falha na conexão. Verifique endpoint, chave API e modelo.",
                        fg="#c62828")
            self.after(0, _show)

        threading.Thread(target=_do_test, daemon=True).start()

    def _browse_ipynb(self):
        path = filedialog.askopenfilename(
            title="Selecione o arquivo .ipynb",
            filetypes=[("Jupyter Notebook", "*.ipynb"), ("Todos os arquivos", "*.*")],
        )
        if path:
            self._ipynb_var.set(path)
            self._output_nb_var.set(str(Path(path).parent))

    def _browse_output_nb(self):
        folder = filedialog.askdirectory(title="Selecione a pasta de destino")
        if folder:
            self._output_nb_var.set(folder)

    def _start_notebook(self):
        ipynb = self._ipynb_var.get().strip()
        out_dir = self._output_nb_var.get().strip()

        if not ipynb:
            messagebox.showwarning("Atenção", "Selecione um arquivo .ipynb antes de continuar.")
            return
        if not os.path.isfile(ipynb):
            messagebox.showerror("Erro", f"Arquivo não encontrado:\n{ipynb}")
            return
        if not out_dir:
            out_dir = str(APP_DIR)

        nb_name     = Path(ipynb).stem
        output_path = str(Path(out_dir) / f"{nb_name}_Documentacao.docx")
        llm_config  = self._get_llm_config()

        self._log_nb_clear()
        self._log_nb(f"📂  Arquivo: {ipynb}")
        self._log_nb(f"💾  Saída:   {output_path}")
        if llm_config.get("provider") not in (None, "Desabilitado"):
            self._log_nb(f"🤖  LLM: {llm_config['provider']} / {llm_config.get('model','gpt-4o')}")
        self._log_nb("─" * 60)

        self._running_nb = True
        self._btn_run_nb.config(state="disabled", text="⏳  Processando...")
        self._progress_nb.start(12)

        t = threading.Thread(
            target=run_notebook_analysis,
            args=(ipynb, output_path, self._log_nb, self._on_done_nb, self._on_error_nb),
            kwargs={"llm_config": llm_config},
            daemon=True,
        )
        t.start()

    def _on_done_nb(self, output_path: str):
        self.after(0, self._finish_ok_nb, output_path)

    def _on_error_nb(self, msg: str):
        self.after(0, self._finish_err_nb, msg)

    def _finish_ok_nb(self, output_path: str):
        self._progress_nb.stop()
        self._running_nb = False
        self._btn_run_nb.config(state="normal", text="▶   GERAR DOCUMENTAÇÃO DO NOTEBOOK")
        self._log_nb("─" * 60)
        self._log_nb("✅  Documentação gerada com sucesso!")
        self._log_nb(f"📄  {output_path}")
        messagebox.showinfo("Concluído!", f"Documentação gerada com sucesso!\n\n📄 {output_path}")

    def _finish_err_nb(self, msg: str):
        self._progress_nb.stop()
        self._running_nb = False
        self._btn_run_nb.config(state="normal", text="▶   GERAR DOCUMENTAÇÃO DO NOTEBOOK")
        self._log_nb("─" * 60)
        self._log_nb(f"❌  Erro: {msg}")
        messagebox.showerror("Erro na geração", f"Ocorreu um erro:\n\n{msg}")

    def _log_nb(self, msg: str):
        def _insert():
            self._log_text_nb.config(state="normal")
            self._log_text_nb.insert("end", msg + "\n")
            self._log_text_nb.see("end")
            self._log_text_nb.config(state="disabled")
        if threading.current_thread() is threading.main_thread():
            _insert()
        else:
            self.after(0, _insert)

    def _log_nb_clear(self):
        self._log_text_nb.config(state="normal")
        self._log_text_nb.delete("1.0", "end")
        self._log_text_nb.config(state="disabled")

    # ── Iniciar geração PBIX ─────────────────────────────────────────────────
    def _start(self):
        pbix = self._pbix_var.get().strip()
        out_dir = self._output_var.get().strip()

        if not pbix:
            messagebox.showwarning("Atenção", "Selecione um arquivo .pbix antes de continuar.")
            return
        if not os.path.isfile(pbix):
            messagebox.showerror("Erro", f"Arquivo não encontrado:\n{pbix}")
            return
        if not out_dir:
            out_dir = str(APP_DIR)

        # Monta caminho de saída
        pbix_name   = Path(pbix).stem
        output_path = str(Path(out_dir) / f"{pbix_name}_Documentacao.docx")

        # Lê configuração LLM atual
        llm_config = self._get_llm_config()

        # Limpa log
        self._log_clear()
        self._log(f"📂  Arquivo: {pbix}")
        self._log(f"💾  Saída:   {output_path}")
        if llm_config.get("provider") not in (None, "Desabilitado"):
            self._log(f"🤖  LLM: {llm_config['provider']} / {llm_config.get('model','gpt-4o')}")
        self._log("─" * 60)

        # Bloqueia botão e inicia progress
        self._running = True
        self._btn_run.config(state="disabled", text="⏳  Processando...")
        self._progress.start(12)

        # Thread separada para não travar a GUI
        t = threading.Thread(
            target=run_analysis,
            args=(pbix, output_path, self._log, self._on_done, self._on_error),
            kwargs={"llm_config": llm_config},
            daemon=True,
        )
        t.start()

    # ── Callbacks de término ─────────────────────────────────────────────────
    def _on_done(self, output_path: str):
        """Chamado quando a geração termina com sucesso."""
        self.after(0, self._finish_ok, output_path)

    def _on_error(self, msg: str):
        """Chamado quando ocorre um erro."""
        self.after(0, self._finish_err, msg)

    def _finish_ok(self, output_path: str):
        self._progress.stop()
        self._running = False
        self._btn_run.config(state="normal", text="▶   GERAR DOCUMENTAÇÃO")
        self._log("─" * 60)
        self._log(f"✅  Documentação gerada com sucesso!")
        self._log(f"📄  {output_path}")
        messagebox.showinfo(
            "Concluído!",
            f"Documentação gerada com sucesso!\n\n📄 {output_path}",
        )

    def _finish_err(self, msg: str):
        self._progress.stop()
        self._running = False
        self._btn_run.config(state="normal", text="▶   GERAR DOCUMENTAÇÃO")
        self._log("─" * 60)
        self._log(f"❌  Erro: {msg}")
        messagebox.showerror("Erro na geração", f"Ocorreu um erro:\n\n{msg}")

    # ── Log helpers ──────────────────────────────────────────────────────────
    def _log(self, msg: str):
        """Adiciona linha ao log (thread-safe via after)."""
        def _insert():
            self._log_text.config(state="normal")
            self._log_text.insert("end", msg + "\n")
            self._log_text.see("end")
            self._log_text.config(state="disabled")
        if threading.current_thread() is threading.main_thread():
            _insert()
        else:
            self.after(0, _insert)

    def _log_clear(self):
        self._log_text.config(state="normal")
        self._log_text.delete("1.0", "end")
        self._log_text.config(state="disabled")


# ══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = App()
    app.mainloop()
