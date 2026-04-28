"""
main.py
=======
Ponto de entrada único da ferramenta.

Uso via interface gráfica:
    python main.py

Uso via linha de comando (sem interface):
    python main.py --cli batidas.xlsx Faturamento.xlsx [saida.xlsx]
"""

import sys
import os

BASE_DIR = os.path.dirname(__file__)
sys.path.insert(0, BASE_DIR)


def cli():
    """Modo terminal."""
    args = sys.argv[2:]
    if len(args) < 2:
        print('Uso: python main.py --cli batidas.xlsx Faturamento.xlsx [saida.xlsx]')
        sys.exit(1)

    from src.services.processador import processar

    batidas_path     = args[0]
    faturamento_path = args[1]
    saida_path       = args[2] if len(args) > 2 else 'Faturamento_PROCESSADO.xlsx'

    for f in [batidas_path, faturamento_path]:
        if not os.path.exists(f):
            print(f'ERRO: arquivo não encontrado: {f}')
            sys.exit(1)

    res = processar(batidas_path, faturamento_path, saida_path)

    if res['ok']:
        p = res['params']
        print(f'\n[OK] {p.nome} — {p.total_dias} dias | PAR={p.max_par} ÍMPAR={p.max_impar}\n')
        print(f'{"FUNCIONÁRIO":<38} {"GRUPO":<6} {"LOGINS":>6}  {"DISP":>5}  {"INDISP":>6}')
        print('─' * 68)
        for r in res['resultados']:
            flag = '  ⚠' if r.indisponibilidade > 0 else ''
            print(f'{r.func.nome[:37]:<38} {r.func.grupo:<6} '
                  f'{r.logins:>6}  {r.disponibilidade:>5}  '
                  f'{r.indisponibilidade:>6}{flag}')
        if res['erros']:
            print('\n!! ERROS:')
            for e in res['erros']: print(f'  {e}')
        if res['avisos']:
            print('\nAVISOS:')
            for a in res['avisos']: print(f'  {a}')
        print(f'\nConcluído: {saida_path}')
    else:
        print('ERRO CRÍTICO:')
        for e in res['erros']: print(f'  {e}')
        sys.exit(1)


def gui():
    """Modo interface gráfica."""
    from src.ui.interface import iniciar
    iniciar()


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == '--cli':
        cli()
    else:
        gui()
