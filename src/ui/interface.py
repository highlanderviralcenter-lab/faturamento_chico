"""
interface.py
============
Interface gráfica Tkinter.
Coloca os arquivos em data/input/ e clica PROCESSAR.
"""

import os, sys, subprocess, threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.io.leitor       import validar_inputs
from src.services.processador  import processar


BASE_DIR   = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
INPUT_DIR  = os.path.join(BASE_DIR, 'data', 'input')
OUTPUT_DIR = os.path.join(BASE_DIR, 'data', 'output')

# ── Paleta ───────────────────────────────────────────────────────────────────
COR_BG      = '#F0F4FA'
COR_AZUL    = '#1F3864'
COR_AZUL2   = '#2E75B6'
COR_VERDE   = '#1E6B3C'
COR_VERMELHO= '#C00000'
COR_AMARELO = '#7F6000'
COR_BRANCO  = '#FFFFFF'


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Faturamento por Posto de Trabalho — BASIS')
        self.geometry('800x640')
        self.resizable(False, False)
        self.configure(bg=COR_BG)

        self._output_path = None
        self._build()
        self._validar()

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build(self):
        # Cabeçalho
        hdr = tk.Frame(self, bg=COR_AZUL, height=54)
        hdr.pack(fill='x')
        tk.Label(hdr, text='  Automação de Faturamento — Posto de Trabalho',
                 bg=COR_AZUL, fg=COR_BRANCO,
                 font=('Arial', 13, 'bold')).pack(side='left', pady=12)
        tk.Label(hdr, text='BASIS  ', bg=COR_AZUL2, fg=COR_BRANCO,
                 font=('Arial', 11)).pack(side='right', pady=12)

        # Arquivos
        frm_arq = tk.LabelFrame(self, text=' Arquivos em data/input/ ',
                                 bg=COR_BG, font=('Arial', 10, 'bold'),
                                 fg=COR_AZUL, padx=12, pady=8)
        frm_arq.pack(fill='x', padx=20, pady=(14, 6))

        self._lbl_bat = tk.Label(frm_arq, text='...', bg=COR_BG,
                                  font=('Arial', 10), anchor='w')
        self._lbl_bat.pack(fill='x')
        self._lbl_fat = tk.Label(frm_arq, text='...', bg=COR_BG,
                                  font=('Arial', 10), anchor='w')
        self._lbl_fat.pack(fill='x')

        tk.Button(frm_arq, text='↻  Revalidar', command=self._validar,
                  bg=COR_AZUL2, fg=COR_BRANCO, font=('Arial', 9),
                  relief='flat', padx=8).pack(anchor='e', pady=(4, 0))

        # Botões
        frm_btn = tk.Frame(self, bg=COR_BG)
        frm_btn.pack(fill='x', padx=20, pady=8)

        self._btn_proc = tk.Button(
            frm_btn, text='▶  PROCESSAR',
            command=self._processar,
            bg=COR_AZUL, fg=COR_BRANCO,
            font=('Arial', 12, 'bold'),
            relief='flat', padx=22, pady=10,
            state='disabled',
        )
        self._btn_proc.pack(side='left', padx=(0, 12))

        self._btn_abrir = tk.Button(
            frm_btn, text='📂  Abrir resultado',
            command=self._abrir_resultado,
            bg=COR_VERDE, fg=COR_BRANCO,
            font=('Arial', 11),
            relief='flat', padx=14, pady=10,
            state='disabled',
        )
        self._btn_abrir.pack(side='left')

        # Log
        frm_log = tk.LabelFrame(self, text=' Log ',
                                 bg=COR_BG, font=('Arial', 10, 'bold'),
                                 fg=COR_AZUL, padx=8, pady=6)
        frm_log.pack(fill='both', expand=True, padx=20, pady=(4, 14))

        self._log = scrolledtext.ScrolledText(
            frm_log, height=16, font=('Courier New', 9),
            bg='#1E1E1E', fg='#D4D4D4',
            state='disabled', relief='flat',
        )
        self._log.pack(fill='both', expand=True)

        # Rodapé
        tk.Label(self, text='Anderson Michel  |  BASIS  |  2026',
                 bg=COR_BG, fg='#999', font=('Arial', 8)).pack(pady=(0, 6))

    # ── Validação ─────────────────────────────────────────────────────────────

    def _validar(self):
        val = validar_inputs(INPUT_DIR)

        if val['batidas']:
            self._lbl_bat.config(
                text=f'✅  batidas.xlsx  →  {os.path.basename(val["batidas"])}',
                fg=COR_VERDE)
        else:
            self._lbl_bat.config(
                text='❌  batidas.xlsx  →  NÃO ENCONTRADO',
                fg=COR_VERMELHO)

        if val['faturamento']:
            self._lbl_fat.config(
                text=f'✅  Faturamento  →  {os.path.basename(val["faturamento"])}',
                fg=COR_VERDE)
        else:
            self._lbl_fat.config(
                text='❌  Faturamento_posto_de_trabalho_*.xlsx  →  NÃO ENCONTRADO',
                fg=COR_VERMELHO)

        self._btn_proc.config(state='normal' if val['ok'] else 'disabled')
        self._val = val

    # ── Processamento ─────────────────────────────────────────────────────────

    def _processar(self):
        self._btn_proc.config(state='disabled')
        self._btn_abrir.config(state='disabled')
        self._log_clear()

        def _run():
            os.makedirs(OUTPUT_DIR, exist_ok=True)
            ts  = datetime.now().strftime('%m_%Y_%H%M')
            out = os.path.join(OUTPUT_DIR, f'Faturamento_PROCESSADO_{ts}.xlsx')

            self._log(f'Iniciando  {datetime.now().strftime("%d/%m/%Y %H:%M:%S")}\n\n')

            try:
                res = processar(self._val['batidas'], self._val['faturamento'], out)
            except Exception as e:
                self._log(f'❌  ERRO INESPERADO: {e}\n')
                self.after(0, lambda: self._btn_proc.config(state='normal'))
                return

            self._output_path = out

            if res['ok']:
                p = res['params']
                self._log(f'✅  {p.nome}  |  PAR={p.max_par}  ÍMPAR={p.max_impar}  Total={p.total_dias} dias\n\n')

                # Resumo por funcionário
                self._log(f'{"FUNCIONÁRIO":<38} {"GRUPO":<6} {"LOGINS":>6}  {"DISP":>5}  {"INDISP":>6}\n')
                self._log('─' * 68 + '\n')
                for r in res['resultados']:
                    flag = '  ⚠' if r.indisponibilidade > 0 else ''
                    self._log(
                        f'{r.func.nome[:37]:<38} {r.func.grupo:<6} '
                        f'{r.logins:>6}  {r.disponibilidade:>5}  '
                        f'{r.indisponibilidade:>6}{flag}\n'
                    )

                if res['avisos']:
                    self._log('\n── AVISOS ──────────────────────────────\n')
                    for a in res['avisos']:
                        self._log(f'  ⚠  {a}\n')

                if res['erros']:
                    self._log('\n── ERROS ───────────────────────────────\n')
                    for e in res['erros']:
                        self._log(f'  ❌  {e}\n')

                self._log(f'\nSalvo em: {out}\n')
                self.after(0, lambda: self._btn_abrir.config(state='normal'))

            else:
                self._log('❌  ERRO CRÍTICO — processamento cancelado:\n')
                for e in res['erros']:
                    self._log(f'  {e}\n')

            self.after(0, lambda: self._btn_proc.config(state='normal'))

        threading.Thread(target=_run, daemon=True).start()

    def _abrir_resultado(self):
        if self._output_path and os.path.exists(self._output_path):
            if sys.platform == 'win32':
                os.startfile(self._output_path)
            elif sys.platform == 'darwin':
                subprocess.call(['open', self._output_path])
            else:
                subprocess.call(['xdg-open', self._output_path])

    # ── Log helpers ───────────────────────────────────────────────────────────

    def _log(self, texto: str):
        def _w():
            self._log_widget.config(state='normal')
            self._log_widget.insert('end', texto)
            self._log_widget.see('end')
            self._log_widget.config(state='disabled')
        self.after(0, _w)

    def _log_clear(self):
        self._log_widget.config(state='normal')
        self._log_widget.delete('1.0', 'end')
        self._log_widget.config(state='disabled')

    @property
    def _log_widget(self):
        return self._log.__func__  # workaround — pega o widget diretamente

    def _log(self, texto: str):
        """Escreve no log de forma thread-safe."""
        def _w():
            self._txt.config(state='normal')
            self._txt.insert('end', texto)
            self._txt.see('end')
            self._txt.config(state='disabled')
        self.after(0, _w)

    def _log_clear(self):
        self._txt.config(state='normal')
        self._txt.delete('1.0', 'end')
        self._txt.config(state='disabled')

    def _build(self):
        # Cabeçalho
        hdr = tk.Frame(self, bg=COR_AZUL, height=54)
        hdr.pack(fill='x')
        tk.Label(hdr, text='  Automação de Faturamento — Posto de Trabalho',
                 bg=COR_AZUL, fg=COR_BRANCO,
                 font=('Arial', 13, 'bold')).pack(side='left', pady=12)
        tk.Label(hdr, text='BASIS  ', bg=COR_AZUL2, fg=COR_BRANCO,
                 font=('Arial', 11)).pack(side='right', pady=12)

        # Arquivos
        frm_arq = tk.LabelFrame(self, text=' Arquivos em data/input/ ',
                                 bg=COR_BG, font=('Arial', 10, 'bold'),
                                 fg=COR_AZUL, padx=12, pady=8)
        frm_arq.pack(fill='x', padx=20, pady=(14, 6))

        self._lbl_bat = tk.Label(frm_arq, text='...', bg=COR_BG,
                                  font=('Arial', 10), anchor='w')
        self._lbl_bat.pack(fill='x')
        self._lbl_fat = tk.Label(frm_arq, text='...', bg=COR_BG,
                                  font=('Arial', 10), anchor='w')
        self._lbl_fat.pack(fill='x')

        tk.Button(frm_arq, text='↻  Revalidar', command=self._validar,
                  bg=COR_AZUL2, fg=COR_BRANCO, font=('Arial', 9),
                  relief='flat', padx=8).pack(anchor='e', pady=(4, 0))

        # Botões
        frm_btn = tk.Frame(self, bg=COR_BG)
        frm_btn.pack(fill='x', padx=20, pady=8)

        self._btn_proc = tk.Button(
            frm_btn, text='▶  PROCESSAR',
            command=self._processar,
            bg=COR_AZUL, fg=COR_BRANCO,
            font=('Arial', 12, 'bold'),
            relief='flat', padx=22, pady=10,
            state='disabled',
        )
        self._btn_proc.pack(side='left', padx=(0, 12))

        self._btn_abrir = tk.Button(
            frm_btn, text='📂  Abrir resultado',
            command=self._abrir_resultado,
            bg=COR_VERDE, fg=COR_BRANCO,
            font=('Arial', 11),
            relief='flat', padx=14, pady=10,
            state='disabled',
        )
        self._btn_abrir.pack(side='left')

        # Log
        frm_log = tk.LabelFrame(self, text=' Log ',
                                 bg=COR_BG, font=('Arial', 10, 'bold'),
                                 fg=COR_AZUL, padx=8, pady=6)
        frm_log.pack(fill='both', expand=True, padx=20, pady=(4, 14))

        self._txt = scrolledtext.ScrolledText(
            frm_log, height=16, font=('Courier New', 9),
            bg='#1E1E1E', fg='#D4D4D4',
            state='disabled', relief='flat',
        )
        self._txt.pack(fill='both', expand=True)

        # Rodapé
        tk.Label(self, text='Anderson Michel  |  BASIS  |  2026',
                 bg=COR_BG, fg='#999', font=('Arial', 8)).pack(pady=(0, 6))


def iniciar():
    App().mainloop()


if __name__ == '__main__':
    iniciar()
