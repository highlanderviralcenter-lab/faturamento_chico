import os
import sys
import subprocess
import threading
from datetime import datetime
import tkinter as tk
from tkinter import scrolledtext

BASE_DIR   = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
INPUT_DIR  = os.path.join(BASE_DIR, 'data', 'input')
OUTPUT_DIR = os.path.join(BASE_DIR, 'data', 'output')

sys.path.insert(0, BASE_DIR)

from src.io.leitor            import validar_inputs
from src.services.processador import processar

AZUL     = '#1F3864'
AZUL2    = '#2E75B6'
VERDE    = '#1E6B3C'
VERMELHO = '#C00000'
BRANCO   = '#FFFFFF'
BG       = '#F0F4FA'


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Faturamento — BASIS')
        self.geometry('820x640')
        self.resizable(False, False)
        self.configure(bg=BG)
        self._output_path = None
        self._val = {}
        self._build()
        self._validar()

    def _build(self):
        hdr = tk.Frame(self, bg=AZUL, height=52)
        hdr.pack(fill='x')
        tk.Label(hdr, text='  Faturamento por Posto de Trabalho',
                 bg=AZUL, fg=BRANCO, font=('Arial', 13, 'bold')).pack(side='left', pady=12)
        tk.Label(hdr, text='BASIS  ',
                 bg=AZUL2, fg=BRANCO, font=('Arial', 11)).pack(side='right', pady=12)

        frm = tk.LabelFrame(self, text=' Arquivos — data/input/ ',
                            bg=BG, fg=AZUL, font=('Arial', 10, 'bold'),
                            padx=12, pady=8)
        frm.pack(fill='x', padx=20, pady=(14, 6))

        self._lbl_bat = tk.Label(frm, text='...', bg=BG, font=('Arial', 10), anchor='w')
        self._lbl_bat.pack(fill='x')
        self._lbl_fat = tk.Label(frm, text='...', bg=BG, font=('Arial', 10), anchor='w')
        self._lbl_fat.pack(fill='x')

        tk.Button(frm, text='↻  Revalidar', command=self._validar,
                  bg=AZUL2, fg=BRANCO, font=('Arial', 9),
                  relief='flat', padx=8).pack(anchor='e', pady=(6, 0))

        frm_btn = tk.Frame(self, bg=BG)
        frm_btn.pack(fill='x', padx=20, pady=8)

        self._btn_proc = tk.Button(frm_btn, text='▶  PROCESSAR',
                                   command=self._processar,
                                   bg=AZUL, fg=BRANCO,
                                   font=('Arial', 12, 'bold'),
                                   relief='flat', padx=22, pady=10,
                                   state='disabled')
        self._btn_proc.pack(side='left', padx=(0, 12))

        self._btn_abrir = tk.Button(frm_btn, text='Abrir resultado',
                                    command=self._abrir_resultado,
                                    bg=VERDE, fg=BRANCO,
                                    font=('Arial', 11),
                                    relief='flat', padx=14, pady=10,
                                    state='disabled')
        self._btn_abrir.pack(side='left')

        frm_log = tk.LabelFrame(self, text=' Log ',
                                bg=BG, fg=AZUL, font=('Arial', 10, 'bold'),
                                padx=8, pady=6)
        frm_log.pack(fill='both', expand=True, padx=20, pady=(4, 14))

        self._txt = scrolledtext.ScrolledText(
            frm_log, height=16, font=('Courier New', 9),
            bg='#1E1E1E', fg='#D4D4D4', state='disabled', relief='flat')
        self._txt.pack(fill='both', expand=True)

        tk.Label(self, text='Anderson Michel  |  BASIS  |  2026',
                 bg=BG, fg='#999', font=('Arial', 8)).pack(pady=(0, 6))

    def _validar(self):
        val = validar_inputs(INPUT_DIR)
        self._val = val

        if val['batidas']:
            self._lbl_bat.config(fg=VERDE,
                text=f'ok  batidas  ->  {os.path.basename(val["batidas"])}')
        else:
            self._lbl_bat.config(fg=VERMELHO,
                text='nao encontrado  ->  batidas.xlsx')

        if val['faturamento']:
            self._lbl_fat.config(fg=VERDE,
                text=f'ok  faturamento  ->  {os.path.basename(val["faturamento"])}')
        else:
            self._lbl_fat.config(fg=VERMELHO,
                text='nao encontrado  ->  Faturamento_posto_*.xlsx')

        self._btn_proc.config(state='normal' if val['ok'] else 'disabled')

    def _processar(self):
        self._btn_proc.config(state='disabled')
        self._btn_abrir.config(state='disabled')
        self._limpar()

        def _run():
            os.makedirs(OUTPUT_DIR, exist_ok=True)
            ts  = datetime.now().strftime('%m_%Y_%H%M')
            out = os.path.join(OUTPUT_DIR, f'Faturamento_PROCESSADO_{ts}.xlsx')

            self._w(f'inicio: {datetime.now().strftime("%d/%m/%Y %H:%M:%S")}\n\n')

            try:
                res = processar(self._val['batidas'], self._val['faturamento'], out)
            except Exception as e:
                self._w(f'erro: {e}\n')
                self.after(0, lambda: self._btn_proc.config(state='normal'))
                return

            self._output_path = out

            if res['ok']:
                p = res['params']
                self._w(f'{p.nome}  {p.total_dias} dias  PAR={p.max_par}  IMPAR={p.max_impar}\n\n')
                self._w(f'{"nome":<38} {"grupo":<6} {"logins":>6}  {"disp":>5}  {"indisp":>6}\n')
                self._w('-' * 68 + '\n')
                for r in res['resultados']:
                    av = '  *' if r.indisponibilidade > 0 else ''
                    self._w(
                        f'{r.func.nome[:37]:<38} {r.func.grupo:<6} '
                        f'{r.logins:>6}  {r.disponibilidade:>5}  '
                        f'{r.indisponibilidade:>6}{av}\n'
                    )
                if res['avisos']:
                    self._w('\navisos:\n')
                    for a in res['avisos']:
                        self._w(f'  {a}\n')
                if res['erros']:
                    self._w('\nerros:\n')
                    for e in res['erros']:
                        self._w(f'  {e}\n')
                self._w(f'\nsalvo: {out}\n')
                self.after(0, lambda: self._btn_abrir.config(state='normal'))
            else:
                self._w('erro critico:\n')
                for e in res['erros']:
                    self._w(f'  {e}\n')

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

    def _w(self, texto):
        def _escreve():
            self._txt.config(state='normal')
            self._txt.insert('end', texto)
            self._txt.see('end')
            self._txt.config(state='disabled')
        self.after(0, _escreve)

    def _limpar(self):
        self._txt.config(state='normal')
        self._txt.delete('1.0', 'end')
        self._txt.config(state='disabled')


def iniciar():
    App().mainloop()


if __name__ == '__main__':
    iniciar()
