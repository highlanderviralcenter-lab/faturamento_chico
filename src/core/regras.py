"""
src/regras.py
=============
Regras de negócio puras.
Sem I/O, sem Excel, sem interface — só lógica.
"""

from dataclasses import dataclass, field
from datetime import date
from typing import Optional


MESES_PT = {
    'JANEIRO':1,'FEVEREIRO':2,'MARÇO':3,'ABRIL':4,'MAIO':5,'JUNHO':6,
    'JULHO':7,'AGOSTO':8,'SETEMBRO':9,'OUTUBRO':10,'NOVEMBRO':11,'DEZEMBRO':12,
}


@dataclass
class ParamMes:
    numero: int
    nome: str
    max_par: int
    max_impar: int

    @property
    def total_dias(self) -> int:
        return self.max_par + self.max_impar


@dataclass
class Funcionario:
    nome: str
    nome_norm: str          # lowercase strip para comparação
    matricula: Optional[int]
    grupo: str              # 'PAR' ou 'ÍMPAR'
    admissao: Optional[date]
    linha_excel: int        # linha na aba FATURAMENTO


@dataclass
class ResultadoFuncionario:
    func: Funcionario
    logins: int
    max_dias: int           # pode ser menor que max_grupo se admissão parcial
    disponibilidade: int
    indisponibilidade: int
    ausencias: list = field(default_factory=list)   # lista de {'data', 'motivo'}
    aviso: str = ''
    erro: str = ''

    @property
    def ok(self) -> bool:
        return not self.erro


# ─── Validação dos parâmetros do mês (C1) ────────────────────────────────────

def validar_params_mes(p: ParamMes) -> tuple[bool, str]:
    """
    PAR + ÍMPAR deve fechar o total de dias do mês.
    Em mês ímpar (31 dias): ÍMPAR = PAR + 1
    Em mês par  (30/28 dias): ÍMPAR = PAR
    """
    total = p.total_dias
    diff  = p.max_impar - p.max_par
    esperado_diff = 1 if total % 2 == 1 else 0

    if diff != esperado_diff:
        return False, (
            f'ERRO: mês {p.nome} tem {total} dias mas '
            f'PAR={p.max_par} e ÍMPAR={p.max_impar} '
            f'(diferença {diff}, esperado {esperado_diff})'
        )
    return True, f'OK: {p.nome} — {total} dias | PAR={p.max_par} ÍMPAR={p.max_impar}'


# ─── Cálculo de disponibilidade ──────────────────────────────────────────────

def calcular_max_admissao_parcial(max_grupo: int, admissao: date,
                                  params: ParamMes) -> int:
    """
    Para quem entrou no meio do mês, o máximo proporcional é calculado
    mas nunca pode ser MENOR que os logins reais — isso é tratado no
    processador. Aqui apenas calculamos o máximo esperado.
    """
    inicio_mes = date(admissao.year, params.numero, 1)
    if admissao <= inicio_mes:
        return max_grupo
    # Quantos plantões cabem entre admissão e fim do mês
    import math
    dias_corridos = params.total_dias - admissao.day + 1
    return math.ceil(max_grupo * dias_corridos / params.total_dias)


def calcular_disponibilidade(logins: int, max_dias: int,
                              admissao_parcial: bool = False) -> tuple[int, int]:
    """
    Retorna (disponibilidade, indisponibilidade).
    Para admissão parcial: se logins >= max_dias, sem indisponibilidade.
    Regra central: disp + indisp = max_dias
    """
    if admissao_parcial and logins >= max_dias:
        # Trabalhou tudo que era esperado (ou mais) — sem falta
        return logins, 0
    disponib = logins
    indisp   = max(max_dias - logins, 0)
    return disponib, indisp


# ─── Validações por funcionário (C2 + C3) ────────────────────────────────────

def validar_resultado(r: ResultadoFuncionario) -> tuple[bool, str]:
    """
    C2: disp + indisp == max_dias
    C3: para admissão parcial com logins > max_calculado, disp = logins é aceito
    """
    d, i, m = r.disponibilidade, r.indisponibilidade, r.max_dias
    if d + i != m and not (r.func.admissao and d > m):
        return False, f'C2 falhou: {d}+{i}={d+i} != {m}'
    return True, 'OK'
