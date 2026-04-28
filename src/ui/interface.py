import os
import sys
import subprocess
import threading
from datetime import datetime, timedelta
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
AMARELO  = '#7F6000'
BRANCO   = '#FFFFFF'
BG       = '#F0F4FA'

TURNO_DIURNO  = (7, 19)   # 07:00 - 19:00
TURNO_NOTURNO = (19, 7)   # 19:00 - 07:00
JORNADA_HORAS = 12.0


def calcular_horas_dia(login, logout):
    if logout < login:
        logout = logout + timedelta(days=1)
    return (logout - login).total_seconds() / 3600


def auditar_jornadas(df_bat):
    import pandas as pd
    registros = []
    for (nome, data), g in df_bat.groupby(['NOME', 'DATA']):
        login_row  = g[g['TIPO_EVENTO'] == 'LOGIN']
        logout_row = g[g['TIPO_EVENTO'] == 'LOGOUT']
        if login_row.empty or logout_row.empty:
            continue
        login  = login_row['DATA_HORA'].iloc[0]
        logout = logout_row['DATA_HORA'].iloc[-1]
        horas  = calcular_horas_dia(login, logout)
        registros.append({
            'NOME':   nome,
            'DATA':   data,
            'LOGIN':  login,
            'LOGOUT': logout,
            'HORAS':  round(horas, 2),
            'OK':     horas >= JORNADA_HORAS,
        })
    return pd.DataFrame(registros)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Faturamento - BASIS')
        self.geometry('860x700')
        self.resizable(False, False)
        self.configure(bg=BG)
        self._output_path  = None
        self._val          = {}
        self._resultado    = None
        self._revisao_var  = tk.BooleanVar(value=False)
        self._build()
        self._validar()

    def _build(self):
        # cabecalho
        hdr = tk.Frame(self, bg=AZUL, height=52)
        hdr.pack(fill='x')
        tk.Label(hdr, text='  Faturamento por Posto de Trabalho',
                 bg=AZUL, fg=BRANCO, font=('Arial', 13, 'bold')).pack(side='left', pady=12)
        tk.Label(hdr, text='BASIS  ',
                 bg=AZUL2, fg=BRANCO, font=('Arial', 11)).pack(side='right', pady=12)

        # arquivos
        frm = tk.LabelFrame(self, text=' Arquivos - data/input/ ',
                            bg=BG, fg=AZUL, font=('Arial', 10, 'bold'),
                            padx=12, pady=8)
        frm.pack(fill='x', padx=20, pady=(14, 6))
        self._lbl_bat = tk.Label(frm, text='...', bg=BG, font=('Arial', 10), anchor='w')
        self._lbl_bat.pack(fill='x')
        self._lbl_fat = tk.Label(frm, text='...', bg=BG, font=('Arial', 10), anchor='w')
        self._lbl_fat.pack(fill='x')
        tk.Button(frm, text='Revalidar', command=self._validar,
                  bg=AZUL2, fg=BRANCO, font=('Arial', 9),
                  relief='flat', padx=8).pack(anchor='e', pady=(6, 0))

        # opcao de revisao
        frm_opt = tk.LabelFrame(self, text=' Opcoes ',
                                bg=BG, fg=AZUL, font=('Arial', 10, 'bold'),
                                padx=12, pady=8)
        frm_opt.pack(fill='x', padx=20, pady=(0, 6))
        self._chk = tk.Checkbutton(frm_opt,
                                    text='  Pretendo revisar manualmente antes de finalizar',
                                    variable=self._revisao_var,
                                    command=self._atualizar_botoes,
                                    bg=BG, font=('Arial', 10),
                                    activebackground=BG)
        self._chk.pack(anchor='w')

        # botoes
        self._frm_btn = tk.Frame(self, bg=BG)
        self._frm_btn.pack(fill='x', padx=20, pady=8)

        self._btn_proc = tk.Button(self._frm_btn, text='PROCESSAR',
                                   command=self._processar,
                                   bg=AZUL, fg=BRANCO,
                                   font=('Arial', 12, 'bold'),
                                   relief='flat', padx=22, pady=10,
                                   state='disabled')
        self._btn_proc.pack(side='left', padx=(0, 12))

        # botao dinamico - muda conforme checkbox
        self._btn_acao = tk.Button(self._frm_btn, text='FINALIZAR AUTO',
                                   command=self._finalizar_auto,
                                   bg=VERDE, fg=BRANCO,
                                   font=('Arial', 11, 'bold'),
                                   relief='flat', padx=16, pady=10,
                                   state='disabled')
        self._btn_acao.pack(side='left', padx=(0, 12))

        # botao finalizar manual - aparece so apos abrir planilha de revisao
        self._btn_final_manual = tk.Button(self._frm_btn, text='FINALIZAR MANUAL',
                                            command=self._finalizar_manual,
                                            bg=VERDE, fg=BRANCO,
                                            font=('Arial', 11, 'bold'),
                                            relief='flat', padx=16, pady=10)

        # log
        frm_log = tk.LabelFrame(self, text=' Log ',
                                bg=BG, fg=AZUL, font=('Arial', 10, 'bold'),
                                padx=8, pady=6)
        frm_log.pack(fill='both', expand=True, padx=20, pady=(4, 14))
        self._txt = scrolledtext.ScrolledText(
            frm_log, height=18, font=('Courier New', 9),
            bg='#1E1E1E', fg='#D4D4D4', state='disabled', relief='flat')
        self._txt.pack(fill='both', expand=True)

        tk.Label(self, text='Anderson Michel  |  BASIS  |  2026',
                 bg=BG, fg='#999', font=('Arial', 8)).pack(pady=(0, 6))

    def _atualizar_botoes(self):
        if self._resultado is None:
            return
        if self._revisao_var.get():
            self._btn_acao.config(text='Abrir planilha de revisao',
                                  command=self._abrir_revisao,
                                  bg=AMARELO, state='normal')
            self._btn_final_manual.pack_forget()
        else:
            self._btn_acao.config(text='FINALIZAR AUTO',
                                  command=self._finalizar_auto,
                                  bg=VERDE, state='normal')
            self._btn_final_manual.pack_forget()

    def _validar(self):
        val = validar_inputs(INPUT_DIR)
        self._val = val
        if val['batidas']:
            self._lbl_bat.config(fg=VERDE,
                text='ok  batidas  ->  ' + os.path.basename(val['batidas']))
        else:
            self._lbl_bat.config(fg=VERMELHO, text='nao encontrado  ->  batidas.xlsx')
        if val['faturamento']:
            self._lbl_fat.config(fg=VERDE,
                text='ok  faturamento  ->  ' + os.path.basename(val['faturamento']))
        else:
            self._lbl_fat.config(fg=VERMELHO,
                text='nao encontrado  ->  Faturamento_posto_*.xlsx')
        self._btn_proc.config(state='normal' if val['ok'] else 'disabled')

    def _processar(self):
        self._btn_proc.config(state='disabled')
        self._btn_acao.config(state='disabled')
        self._btn_final_manual.pack_forget()
        self._limpar()

        def _run():
            os.makedirs(OUTPUT_DIR, exist_ok=True)
            ts  = datetime.now().strftime('%m_%Y_%H%M')
            out = os.path.join(OUTPUT_DIR, 'Faturamento_PROCESSADO_' + ts + '.xlsx')

            self._w('inicio: ' + datetime.now().strftime('%d/%m/%Y %H:%M:%S') + '\n\n')

            try:
                import pandas as pd
                df_bat = __import__('src.io.leitor', fromlist=['ler_batidas']).ler_batidas(self._val['batidas'])
                auditoria = auditar_jornadas(df_bat)
                n_abaixo  = (~auditoria['OK']).sum()

                if n_abaixo > 0:
                    self._w('jornadas abaixo de 12h:\n')
                    for _, row in auditoria[~auditoria['OK']].iterrows():
                        self._w('  ' + str(row['NOME'])[:30].ljust(30) +
                                '  ' + str(row['DATA']) +
                                '  login ' + row['LOGIN'].strftime('%H:%M') +
                                '  logout ' + row['LOGOUT'].strftime('%H:%M') +
                                '  ' + str(round(row['HORAS'], 1)) + 'h\n')
                    self._w('\n')

                res = processar(self._val['batidas'], self._val['faturamento'], out)
            except Exception as e:
                self._w('erro: ' + str(e) + '\n')
                self.after(0, lambda: self._btn_proc.config(state='normal'))
                return

            self._output_path = out
            self._resultado   = res

            if res['ok']:
                p = res['params']
                self._w(p.nome + '  ' + str(p.total_dias) + ' dias  PAR=' +
                        str(p.max_par) + '  IMPAR=' + str(p.max_impar) + '\n\n')
                self._w('nome'.ljust(38) + ' grupo  logins   disp  indisp\n')
                self._w('-' * 68 + '\n')
                for r in res['resultados']:
                    av = '  *' if r.indisponibilidade > 0 else ''
                    self._w(r.func.nome[:37].ljust(38) + ' ' + r.func.grupo.ljust(6) +
                            str(r.logins).rjust(6) + '  ' +
                            str(r.disponibilidade).rjust(5) + '  ' +
                            str(r.indisponibilidade).rjust(6) + av + '\n')
                if res['avisos']:
                    self._w('\navisos:\n')
                    for a in res['avisos']:
                        self._w('  ' + a + '\n')
                if res['erros']:
                    self._w('\nerros:\n')
                    for e in res['erros']:
                        self._w('  ' + e + '\n')
                self._w('\npronto. salvo em: ' + out + '\n')
                self.after(0, self._atualizar_botoes)
            else:
                self._w('erro critico:\n')
                for e in res['erros']:
                    self._w('  ' + e + '\n')

            self.after(0, lambda: self._btn_proc.config(state='normal'))

        threading.Thread(target=_run, daemon=True).start()

    def _finalizar_auto(self):
        self._btn_acao.config(state='disabled')
        self._w('\nfinalizado. arquivo em data/output/\n')

    def _abrir_revisao(self):
        if self._output_path and os.path.exists(self._output_path):
            if sys.platform == 'win32':
                os.startfile(self._output_path)
            elif sys.platform == 'darwin':
                subprocess.call(['open', self._output_path])
            else:
                subprocess.call(['xdg-open', self._output_path])
            self._btn_acao.config(state='disabled')
            self._w('\nplanilha aberta. revise e clique em FINALIZAR MANUAL.\n')
            self.after(0, lambda: self._btn_final_manual.pack(side='left'))

    def _finalizar_manual(self):
        self._btn_final_manual.pack_forget()
        self._w('\nfinalizado manualmente. arquivo em data/output/\n')

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
