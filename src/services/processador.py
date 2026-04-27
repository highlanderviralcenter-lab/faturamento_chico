"""
src/processador.py
==================
Orquestra: lê → calcula → valida → grava.
Não sabe nada de Excel internamente — delega para leitor/escritor.
"""

import shutil
from datetime import date
from typing import Optional

import openpyxl
import pandas as pd

from src.core.regras import (
    ParamMes, Funcionario, ResultadoFuncionario,
    validar_params_mes, calcular_disponibilidade,
    calcular_max_admissao_parcial, validar_resultado,
)
from src.io.leitor import (
    ler_batidas, ler_params_mes, ler_admissoes,
    ler_funcionarios_faturamento, extrair_ausencias,
)
from src.services.escritor import (
    gravar_resultado_faturamento,
    preencher_aba_indisponibilidade,
)


def processar(batidas_path: str, faturamento_path: str,
              output_path: str) -> dict:
    """
    Processa um mês completo.

    Fluxo:
      1. Lê batidas e detecta o mês
      2. Valida parâmetros PAR/ÍMPAR do mês (C1)
      3. Para cada funcionário no faturamento:
         - Conta LOGINs nas batidas (fonte da verdade)
         - Calcula disponibilidade e indisponibilidade
         - Valida a equação disp + indisp = max (C2)
      4. Grava colunas S e T no faturamento
      5. Preenche aba INDISPONIBILIDADE

    Retorna dict com: ok, params, resultados, avisos, erros
    """
    resultado = {
        'ok': False, 'params': None,
        'resultados': [], 'avisos': [], 'erros': [],
    }

    # ── 1. Leitura ────────────────────────────────────────────────────────────
    df_bat  = ler_batidas(batidas_path)
    mes_num = int(pd.to_datetime(df_bat['DATA_HORA']).dt.month.mode()[0])

    shutil.copy2(faturamento_path, output_path)
    wb_ro = openpyxl.load_workbook(output_path, data_only=True)
    wb    = openpyxl.load_workbook(output_path)

    params   = ler_params_mes(wb_ro, mes_num)
    admissoes = ler_admissoes(wb_ro)
    funcs    = ler_funcionarios_faturamento(wb_ro, admissoes)
    ausencias_map = extrair_ausencias(df_bat)

    # ── 2. Valida C1 ──────────────────────────────────────────────────────────
    ok_c1, msg_c1 = validar_params_mes(params)
    if not ok_c1:
        resultado['erros'].append(msg_c1)
        return resultado

    resultado['params'] = params

    # Conta LOGINs por funcionário (fonte da verdade — cada LOGIN = 1 plantão)
    logins_por_nome = (
        df_bat[df_bat['TIPO_EVENTO'] == 'LOGIN']
        .groupby(df_bat['NOME'].str.lower().str.strip())['DATA']
        .nunique()
        .to_dict()
    )

    max_grupo = {'PAR': params.max_par, 'ÍMPAR': params.max_impar}

    ws_fat     = wb['FATURAMENTO']
    resultados = []

    # ── 3. Calcula por funcionário ────────────────────────────────────────────
    for func in funcs:
        logins = logins_por_nome.get(func.nome_norm)

        if logins is None:
            resultado['avisos'].append(f'Sem batidas: {func.nome}')
            continue

        max_std = max_grupo[func.grupo]

        # Detecta admissão no meio do mês
        inicio_mes = date(2026, mes_num, 1)
        admissao_parcial = (
            func.admissao is not None and func.admissao > inicio_mes
        )

        if admissao_parcial:
            max_dias = calcular_max_admissao_parcial(max_std, func.admissao, params)
        else:
            max_dias = max_std

        # Alerta grupo possivelmente errado (só para funcionários não-parciais)
        if not admissao_parcial and logins > max_std:
            resultado['avisos'].append(
                f'GRUPO ERRADO? {func.nome}: {logins} logins > max {max_std} [{func.grupo}]'
            )

        disponib, indisp = calcular_disponibilidade(logins, max_dias, admissao_parcial)

        ausencias = ausencias_map.get(func.nome_norm, [])

        res = ResultadoFuncionario(
            func             = func,
            logins           = logins,
            max_dias         = max_dias,
            disponibilidade  = disponib,
            indisponibilidade= indisp,
            ausencias        = ausencias,
        )

        # C2 — equação deve fechar
        ok_c2, msg_c2 = validar_resultado(res)
        if not ok_c2:
            res.erro = msg_c2
            resultado['erros'].append(f'{func.nome}: {msg_c2}')

        # Aviso quando há indisponibilidade mas poucos registros de motivo
        if indisp > 0 and indisp > len(ausencias):
            res.aviso = f'{indisp} dias indisp., {len(ausencias)} motivo(s) no ponto'
            resultado['avisos'].append(
                f'SEM MOTIVO COMPLETO: {func.nome} — '
                f'{indisp} dias indisponível, {len(ausencias)} evento(s) registrado(s)'
            )

        resultados.append(res)

        # ── 4. Grava na planilha ──────────────────────────────────────────────
        gravar_resultado_faturamento(ws_fat, res)

    # ── 5. Preenche aba INDISPONIBILIDADE ─────────────────────────────────────
    preencher_aba_indisponibilidade(wb, resultados)

    wb.save(output_path)

    resultado['ok']         = True
    resultado['resultados'] = resultados
    return resultado
