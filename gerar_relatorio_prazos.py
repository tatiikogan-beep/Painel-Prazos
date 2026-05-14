"""
Gerador de Relatório de Prazos Preclusivos — Versão Web (PAINEL_WEB)
"""

import io, re, sys, unicodedata
from datetime import date, datetime, timedelta

import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

try:
    import pandas as pd
    _USE_PANDAS = True
except ImportError:
    _USE_PANDAS = False

try:
    import holidays as holidays_lib
    _USE_HOLIDAYS = True
except ImportError:
    _USE_HOLIDAYS = False

# ─── CORES (hexadecimais do arquivo de referência Relatorio_734) ─────────────
F_VERDE    = PatternFill('solid', fgColor='C6EFCE')
F_AMARELO  = PatternFill('solid', fgColor='FFEB9C')
F_ROSA     = PatternFill('solid', fgColor='FFC7CE')
F_VERMELHO = PatternFill('solid', fgColor='CC0000')
F_LARANJA  = PatternFill('solid', fgColor='FFCC99')
F_CABEC    = PatternFill('solid', fgColor='2F4F8F')
F_TOTAL    = PatternFill('solid', fgColor='D9E1F2')
F_DATA_ERR = PatternFill('solid', fgColor='FF0000')

FONTE_BRANCA  = Font(color='FFFFFF', bold=True)
FONTE_CABEC   = Font(color='FFFFFF', bold=True)
FONTE_NEGRITO = Font(bold=True)
AL_C      = Alignment(horizontal='center', vertical='center', wrap_text=False)
AL_E      = Alignment(horizontal='left',   vertical='center', wrap_text=False)
AL_C_WRAP = Alignment(horizontal='center', vertical='center', wrap_text=True)

# ─── FERIADOS NACIONAIS (sem Carnaval e Corpus Christi) ──────────────────────
def _feriados_np(anos):
    if not _USE_HOLIDAYS:
        return np.array([], dtype='datetime64[D]')
    br = holidays_lib.Brazil(years=anos)
    excluir = [d for d, n in br.items()
               if 'carnaval' in n.lower() or 'corpus' in n.lower()]
    for d in excluir:
        del br[d]
    return np.array([d.strftime('%Y-%m-%d') for d in sorted(br.keys())],
                    dtype='datetime64[D]')

# ─── MAPEAMENTO EXECUTOR → SUPERVISOR ────────────────────────────────────────
RESPONSAVEL_MAP = {
    # HELANZIA
    "HELANZIA DE ARAUJO XAVIER WICHMANN":          "HELANZIA DE ARAUJO XAVIER WICHMANN",
    "HELANZIA DE ARAUJO XAVIER WICHAMNN":          "HELANZIA DE ARAUJO XAVIER WICHMANN",
    "GUSTAVO LOPES ALENCAR FILHO":                 "HELANZIA DE ARAUJO XAVIER WICHMANN",
    "KELIANE DE OLIVEIRA":                         "HELANZIA DE ARAUJO XAVIER WICHMANN",
    "NATALIA PAIVA DE PAULA":                      "HELANZIA DE ARAUJO XAVIER WICHMANN",
    "ROBERTA RAYANNE VASCONCELOS BOTO":            "HELANZIA DE ARAUJO XAVIER WICHMANN",
    "WELLINGTON SANTOS PINHEIRO":                  "HELANZIA DE ARAUJO XAVIER WICHMANN",
    "ARTUR MAIA SILVA":                            "HELANZIA DE ARAUJO XAVIER WICHMANN",
    "VICTOR HUGO LIMA":                            "HELANZIA DE ARAUJO XAVIER WICHMANN",
    # LUCIANE
    "LUCIANE MODERNEL MENDES":                     "LUCIANE MODERNEL MENDES",
    "MATHEUS CAVALCANTI DE ARAUJO":                "LUCIANE MODERNEL MENDES",
    "ANTONIO EDUARDO GOES AGUIAR FILHO":           "LUCIANE MODERNEL MENDES",
    "SANE BORGES BORGOMONI":                       "LUCIANE MODERNEL MENDES",
    "ERIKA PAULA SANTOS LIMA":                     "LUCIANE MODERNEL MENDES",
    "LAYLA LIMA BORGES":                           "LUCIANE MODERNEL MENDES",
    "EDUARDO BLASQUES":                            "LUCIANE MODERNEL MENDES",
    # GABRIEL
    "GABRIEL GIORGIO CICCHELERO":                  "GABRIEL GIORGIO CICCHELERO",
    "JULIANA DE OLIVEIRA ROCHA":                   "GABRIEL GIORGIO CICCHELERO",
    "ALYSSON NARBAL DE OLIVEIRA SOMBRA":           "GABRIEL GIORGIO CICCHELERO",
    "JAMILE ALVES":                                "GABRIEL GIORGIO CICCHELERO",
    "IRENE FLAVIA SERENARIO":                      "GABRIEL GIORGIO CICCHELERO",
    "IRENE FLÁVIA SERENÁRIO":                      "GABRIEL GIORGIO CICCHELERO",
    "ARMANDO HELIO ALMEIDA MONTEIRO DE MORAES":    "GABRIEL GIORGIO CICCHELERO",
    "ARMANDO HÉLIO ALMEIDA MONTEIRO DE MORAES":    "GABRIEL GIORGIO CICCHELERO",
    "RAFAEL CAVALCANTE BARSOSA":                   "GABRIEL GIORGIO CICCHELERO",
    # JENIFFER
    "JENIFFER ROSA BARBOSA DE SALES":              "JENIFFER ROSA BARBOSA DE SALES",
    "PAULO MARCIO SOARES DE CARVALHO FILHO":       "JENIFFER ROSA BARBOSA DE SALES",
    # JULIANA MIRELLA
    "JULIANA MIRELLA ALVES RODRIGUES":             "JULIANA MIRELLA ALVES RODRIGUES",
    "THALLYS ANDERSON FERREIRA DE LIMA":           "JULIANA MIRELLA ALVES RODRIGUES",
    # NAYANDERSON
    "NAYANDERSON LUAN MELLO PINHEIRO":             "NAYANDERSON LUAN MELLO PINHEIRO",
    "ANDRE VIANA GARRIDO":                         "NAYANDERSON LUAN MELLO PINHEIRO",
    "YURI GONDIM DE AMORIM":                       "NAYANDERSON LUAN MELLO PINHEIRO",
    "EMERSON LIMA SOUSA":                          "NAYANDERSON LUAN MELLO PINHEIRO",
    # YURI
    "YURI ALVES BARROS DOS SANTOS":                "YURI ALVES BARROS DOS SANTOS",
    "LUIZ GUILHERME PEREIRA":                      "YURI ALVES BARROS DOS SANTOS",
    # MARCELLE
    "MARCELLE LEITE RENTROIA":                     "MARCELLE LEITE RENTROIA",
    "MARIANA MOTA SANTOS":                         "MARCELLE LEITE RENTROIA",
    # SUZANA
    "SUZANA MARIA CAMPOS MARANHAO DE LIMA":        "SUZANA MARIA CAMPOS MARANHAO DE LIMA",
    "SUZANA MARIA CAMPOS MARANHÃO DE LIMA":        "SUZANA MARIA CAMPOS MARANHAO DE LIMA",
    "GIOVANNA CAMPOS PEREIRA":                     "SUZANA MARIA CAMPOS MARANHAO DE LIMA",
    "GIOVANNA CESAR FERREIRA":                     "SUZANA MARIA CAMPOS MARANHAO DE LIMA",
    "DANIEL BARROS DE OLIVEIRA":                   "SUZANA MARIA CAMPOS MARANHAO DE LIMA",
    "LETICIA OLIVEIRA DA SILVA":                   "SUZANA MARIA CAMPOS MARANHAO DE LIMA",
    "TATIANE CARMO SANTA ROSA":                    "SUZANA MARIA CAMPOS MARANHAO DE LIMA",
    "EVILANY GABRIELA BRAGA PONTES":               "SUZANA MARIA CAMPOS MARANHAO DE LIMA",
    "FRANCOISE LIMA SANTOS":                       "SUZANA MARIA CAMPOS MARANHAO DE LIMA",
    # TICIANNA
    "TICIANNA PIRES DE SOUZA":                     "TICIANNA PIRES DE SOUZA",
    # RONALD
    "RONALD FEITOSA AGUIAR FILHO":                 "RONALD FEITOSA AGUIAR FILHO",
    "ALEXIA ALENCAR CAPIBARIBE":                   "RONALD FEITOSA AGUIAR FILHO",
}

ORDEM_SUPERVISORES = [
    "GABRIEL GIORGIO CICCHELERO",
    "HELANZIA DE ARAUJO XAVIER WICHMANN",
    "JENIFFER ROSA BARBOSA DE SALES",
    "JULIANA MIRELLA ALVES RODRIGUES",
    "LUCIANE MODERNEL MENDES",
    "MARCELLE LEITE RENTROIA",
    "NAYANDERSON LUAN MELLO PINHEIRO",
    "RONALD FEITOSA AGUIAR FILHO",
    "SUZANA MARIA CAMPOS MARANHAO DE LIMA",
    "TICIANNA PIRES DE SOUZA",
    "YURI ALVES BARROS DOS SANTOS",
]

# ─── HELPERS ─────────────────────────────────────────────────────────────────
def _norm(s):
    s = str(s).strip().upper()
    nfd = unicodedata.normalize('NFD', s)
    return ''.join(c for c in nfd if unicodedata.category(c) != 'Mn')

def _mapear(nome_raw):
    if not nome_raw or str(nome_raw).strip() in ('', 'nan', 'None'):
        return None
    nome = str(nome_raw).split('\n')[0].strip()
    if nome in RESPONSAVEL_MAP:
        return RESPONSAVEL_MAP[nome]
    nome_n = _norm(nome)
    for k, v in RESPONSAVEL_MAP.items():
        if _norm(k) == nome_n:
            return v
    return None

def _para_date(valor):
    if valor is None or (isinstance(valor, float) and str(valor) == 'nan'):
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    s = str(valor).strip().split()[0]
    for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%d.%m.%Y'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None

_PAT_DATA = re.compile(r'(\d{1,2})[/\.](\d{1,2})[/\.](\d{2,4})')

def _data_fatal(descricao, ano_ref):
    res = []
    for m in _PAT_DATA.finditer(str(descricao or '')):
        d, mo, a_s = int(m.group(1)), int(m.group(2)), m.group(3)
        a = int(a_s)
        if a < 100:
            a = (ano_ref // 100) * 100 + a
        elif len(a_s) == 3:
            a = int(str(ano_ref)[:1] + a_s)
        try:
            res.append(date(a, mo, d))
        except ValueError:
            pass
    return max(res) if res else None

def _d1_util(fatal, fer_np):
    prev = np.datetime64(fatal.strftime('%Y-%m-%d'), 'D') - np.timedelta64(1, 'D')
    return date.fromisoformat(str(
        np.busday_offset(prev, 0, roll='backward', holidays=fer_np)))

def _proxima_util(ref, fer_np):
    d_np = np.datetime64(ref.strftime('%Y-%m-%d'), 'D')
    return date.fromisoformat(str(
        np.busday_offset(d_np, 1, roll='forward', holidays=fer_np)))

TST_EXCECAO = 'protocolar recurso sobre decisao tst ou justificar nao recurso'

def _cor(descricao, justificativa, data_cad, hoje, fer_np):
    desc = str(descricao or '')
    just = str(justificativa or '').strip()

    # Verde forçado: audiências
    if 'AUDIENCIA DE JULGAMENTO' in _norm(desc) or 'ACOMPANHAR JULGAMENTO' in _norm(desc):
        return F_VERDE, Font(), 'verde', False, ''

    # Laranja
    tem_nao = bool(re.search(r'\bn[aã]o\b', desc, re.IGNORECASE))
    eh_tst  = TST_EXCECAO in _norm(desc)
    laranja  = bool(just) or (tem_nao and not eh_tst)

    # Data fatal
    ano_ref = data_cad.year if data_cad else hoje.year
    fatal   = _data_fatal(desc, ano_ref)
    amanha  = hoje + timedelta(days=1)
    data_err, div_str = False, ''

    if fatal:
        d1_calc = _d1_util(fatal, fer_np)
        if data_cad and data_cad != d1_calc:
            data_err = True
            div_str  = (f'Cadastro {data_cad.strftime("%d/%m/%Y")} '
                        f'<> D-1 util {d1_calc.strftime("%d/%m/%Y")}')
        cor_base = ('vermelho' if fatal < hoje else
                    'rosa'     if fatal == hoje else
                    'amarelo'  if fatal == amanha else 'verde')
    elif data_cad:
        fatal_inf = _proxima_util(data_cad, fer_np)
        cor_base  = ('vermelho' if fatal_inf < hoje else
                     'rosa'     if fatal_inf == hoje else
                     'amarelo'  if fatal_inf == amanha else 'verde')
    else:
        cor_base = 'verde'

    if laranja:
        return F_LARANJA, Font(), cor_base, data_err, div_str

    FILLS  = {'verde': F_VERDE, 'amarelo': F_AMARELO,
              'rosa': F_ROSA,   'vermelho': F_VERMELHO}
    FONTES = {'vermelho': FONTE_BRANCA}
    return FILLS[cor_base], FONTES.get(cor_base, Font()), cor_base, data_err, div_str

def _cel(ws, row, col, v=None, fill=None, font=None, align=None):
    c = ws.cell(row, col, v)
    if fill: c.fill = fill
    if font: c.font = font
    c.alignment = align or AL_C
    return c

def _cabec(ws, row, headers):
    for j, h in enumerate(headers, 1):
        _cel(ws, row, j, h, F_CABEC, FONTE_CABEC, AL_C_WRAP)

# ─── FUNÇÃO PRINCIPAL ────────────────────────────────────────────────────────
def gerar_relatorio(input_source):
    # Leitura
    if isinstance(input_source, bytes):
        buf = io.BytesIO(input_source)
    else:
        buf = input_source

    if _USE_PANDAS:
        df = pd.read_excel(buf, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        rename = {}
        for c in df.columns:
            cn = _norm(c)
            if cn == 'DATA': rename[c] = 'Data'
            elif cn == 'EXECUTANTE': rename[c] = 'Executante'
            elif cn == 'RESPONSAVEL': rename[c] = 'Responsavel'
            elif cn == 'DESCRICAO': rename[c] = 'Descricao'
            elif cn == 'JUSTIFICATIVA': rename[c] = 'Justificativa'
            elif cn == 'PROCESSO': rename[c] = 'Processo'
        df = df.rename(columns=rename)
        header   = list(df.columns)
        rows_in  = df.values.tolist()
    else:
        wb_in = openpyxl.load_workbook(buf, data_only=True)
        ws_in = wb_in.active
        all_rows = list(ws_in.iter_rows(values_only=True))
        header   = [str(h or '').strip() for h in all_rows[0]]
        rows_in  = [list(r) for r in all_rows[1:]]

    # Índices
    def _idx(nome):
        try: return header.index(nome)
        except ValueError:
            hn = _norm(nome)
            for i, h in enumerate(header):
                if _norm(h) == hn: return i
            return None

    idx_data = _idx('Data')
    idx_exec = _idx('Executante')
    idx_resp = _idx('Responsavel')
    idx_desc = _idx('Descricao')
    idx_just = _idx('Justificativa')
    idx_proc = _idx('Processo')
    idx_div  = _idx('Divergencia de Data')

    if idx_div is None:
        header.append('Divergencia de Data')
        idx_div = len(header) - 1
        rows_in = [list(r) + [''] for r in rows_in]

    ncols = len(header)
    hoje  = date.today()
    fer_np = _feriados_np([hoje.year, hoje.year + 1])

    # Processar
    processadas, nao_mapeados = [], []

    for row in rows_in:
        if all(str(v or '').strip() in ('', 'nan') for v in row): continue

        def _v(i):
            if i is None or i >= len(row): return None
            v = row[i]
            return None if str(v or '').strip() in ('', 'nan', 'None') else v

        exec_raw = _v(idx_exec)
        desc     = _v(idx_desc) or ''
        just     = _v(idx_just) or ''
        data_cad = _para_date(_v(idx_data))

        supervisor = _mapear(exec_raw)

        fill, font, cor_base, data_err, div_str = _cor(
            desc, just, data_cad, hoje, fer_np)

        row = list(row)
        while len(row) < ncols: row.append(None)
        row[idx_div] = div_str or None
        if supervisor and idx_resp is not None:
            row[idx_resp] = supervisor

        if supervisor:
            processadas.append((row, supervisor, str(exec_raw or ''),
                                 fill, font, cor_base, data_err, div_str))
        else:
            nao_mapeados.append((row, str(exec_raw or '')))

    # Agrupamentos
    por_sup, por_exec = {}, {}
    for it in processadas:
        _, sup, exe, *_ = it
        por_sup.setdefault(sup, []).append(it)
        por_exec.setdefault(exe.split('\n')[0].strip(), []).append(it)

    def _cnt(lista):
        c = dict(total=0, amarelo=0, rosa=0, vermelho=0, verde=0)
        for it in lista:
            c['total'] += 1; c[it[5]] = c.get(it[5], 0) + 1
        c['atraso'] = c['rosa'] + c['vermelho']
        return c

    # Workbook
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── RESUMO ────────────────────────────────────────────────────────────────
    ws_r = wb.create_sheet('RESUMO', 0)
    ws_r.column_dimensions['A'].width = 44
    ws_r.column_dimensions['B'].width = 44
    for col in 'CDEFGH': ws_r.column_dimensions[col].width = 16

    L = 1
    ws_r.cell(L, 1, f'Relatorio de Prazos Preclusivos - {hoje.strftime("%d/%m/%Y")}').font = FONTE_NEGRITO
    L += 2

    H_SUP  = ['Responsavel','Total de Prazos','D-1','Fatal',
              'Prazo Vencido - Pendente de Baixa','Dentro do Prazo',
              'Total em Atraso (Fatal + Prazo Vencido - Pendente de Baixa)']
    H_EXEC = ['Executante','Responsavel','Total de Prazos','D-1','Fatal',
              'Prazo Vencido - Pendente de Baixa','Dentro do Prazo',
              'Total em Atraso (Fatal + Prazo Vencido - Pendente de Baixa)']

    def _linha_resumo_sup(ws, row, sup_label, c, is_total=False):
        vals = [sup_label, c['total'],
                c.get('amarelo') or None, c.get('rosa') or None,
                c.get('vermelho') or None, c.get('verde') or None,
                c.get('atraso') or None]
        fills_c = [None, None, F_AMARELO, F_ROSA, F_VERMELHO, F_VERDE, None]
        fontes_c = [None, None, None, None, FONTE_BRANCA, None, None]
        for j, (v, fc, ff) in enumerate(zip(vals, fills_c, fontes_c), 1):
            cell = ws.cell(row, j, v)
            if is_total: cell.fill = F_TOTAL; cell.font = FONTE_NEGRITO
            elif fc and v: cell.fill = fc
            if not is_total and ff and v: cell.font = ff
            cell.alignment = AL_E if j == 1 else AL_C

    # Tabela por supervisor
    _cabec(ws_r, L, H_SUP); L += 1
    tot = dict(total=0, amarelo=0, rosa=0, vermelho=0, verde=0, atraso=0)
    for sup in ORDEM_SUPERVISORES:
        if sup not in por_sup: continue
        c = _cnt(por_sup[sup])
        _linha_resumo_sup(ws_r, L, sup, c); L += 1
        for k in tot: tot[k] += c.get(k, 0)
    _linha_resumo_sup(ws_r, L, 'TOTAL GERAL', tot, is_total=True); L += 2

    # Tabela por executante
    ws_r.cell(L, 1, 'Prazos por Executante:').font = FONTE_NEGRITO; L += 1
    _cabec(ws_r, L, H_EXEC); L += 1
    tot_e = dict(total=0, amarelo=0, rosa=0, vermelho=0, verde=0, atraso=0)
    for exe in sorted(por_exec.keys()):
        c = _cnt(por_exec[exe])
        sup = por_exec[exe][0][1]
        vals = [exe, sup, c['total'],
                c.get('amarelo') or None, c.get('rosa') or None,
                c.get('vermelho') or None, c.get('verde') or None,
                c.get('atraso') or None]
        fills_c  = [None,None,None,F_AMARELO,F_ROSA,F_VERMELHO,F_VERDE,None]
        fontes_c = [None,None,None,None,None,FONTE_BRANCA,None,None]
        for j, (v, fc, ff) in enumerate(zip(vals, fills_c, fontes_c), 1):
            cell = ws_r.cell(L, j, v)
            if fc and v: cell.fill = fc
            if ff and v: cell.font = ff
            cell.alignment = AL_E if j <= 2 else AL_C
        L += 1
        for k in tot_e: tot_e[k] += c.get(k, 0)
    # Total executor
    tot_e_vals = [None,'TOTAL GERAL', tot_e['total'],
                  tot_e.get('amarelo') or None, tot_e.get('rosa') or None,
                  tot_e.get('vermelho') or None, tot_e.get('verde') or None,
                  tot_e.get('atraso') or None]
    for j, v in enumerate(tot_e_vals, 1):
        cell = ws_r.cell(L, j, v)
        cell.fill = F_TOTAL; cell.font = FONTE_NEGRITO
        cell.alignment = AL_E if j <= 2 else AL_C
    L += 2

    # Laranjas para revisão
    laranjas = [(it[0], it[1], it[2]) for it in processadas if it[3] == F_LARANJA]
    if laranjas:
        ws_r.cell(L, 1,
            f'ATENCAO: {len(laranjas)} linha(s) em laranja para revisao'
        ).font = Font(bold=True, color='FF6600'); L += 1
        _cabec(ws_r, L, ['Responsavel','Executante','Numero do Processo','Motivo'])
        for j in range(1, 5):
            ws_r.cell(L, j).fill = F_LARANJA
            ws_r.cell(L, j).font = FONTE_NEGRITO
        L += 1
        for row_v, sup, exe in laranjas:
            proc  = row_v[idx_proc] if idx_proc is not None else ''
            just_v = row_v[idx_just] if idx_just is not None else ''
            motivo = (f'Justificativa: {str(just_v or "")[:60]}'
                      if just_v else 'Descricao contem NAO')
            for j, v in enumerate([sup, exe, proc, motivo], 1):
                ws_r.cell(L, j, v).alignment = AL_E
            L += 1
        L += 1

    # Datas fora do D-1
    erros_data = [(it[0], it[1], it[2], it[7]) for it in processadas if it[6]]
    if erros_data:
        ws_r.cell(L, 1,
            f'ATENCAO: {len(erros_data)} data(s) fora do D-1 util:'
        ).font = Font(bold=True, color='CC0000'); L += 1
        _cabec(ws_r, L, ['Responsavel','Executante','Numero do Processo','Divergencia'])
        for j in range(1, 5):
            ws_r.cell(L, j).fill = F_DATA_ERR
            ws_r.cell(L, j).font = FONTE_BRANCA
        L += 1
        for row_v, sup, exe, div in erros_data:
            proc = row_v[idx_proc] if idx_proc is not None else ''
            for j, v in enumerate([sup, exe, proc, div], 1):
                ws_r.cell(L, j, v).alignment = AL_E
            L += 1
        L += 1

    # Legenda
    ws_r.cell(L, 1, 'LEGENDA:').font = FONTE_NEGRITO; L += 1
    for fill, font, txt in [
        (F_VERDE,    Font(),       'Dentro do prazo (D-1 ainda nao chegou)'),
        (F_AMARELO,  Font(),       'D-1 (hoje e o dia de cumprir, fatal amanha)'),
        (F_ROSA,     Font(),       'Fatal hoje (atrasado - prazo vence hoje)'),
        (F_VERMELHO, FONTE_BRANCA, 'Fatal ja passou (perda de prazo)'),
        (F_LARANJA,  Font(),       'Com justificativa ou NAO na descricao'),
        (F_DATA_ERR, FONTE_BRANCA, 'Celula Data: cadastro fora do D-1 util'),
    ]:
        c2 = ws_r.cell(L, 1, txt)
        c2.fill = fill; c2.font = font; c2.alignment = AL_E; L += 1

    # ── ABAS POR SUPERVISOR ───────────────────────────────────────────────────
    for sup in ORDEM_SUPERVISORES:
        if sup not in por_sup: continue
        ws = wb.create_sheet(sup[:31])
        ws.column_dimensions['A'].width = 22
        for ci in range(2, ncols + 1):
            ws.column_dimensions[get_column_letter(ci)].width = 24
        # Cabeçalho
        for j, h in enumerate(header, 1):
            _cel(ws, 1, j, h, F_CABEC, FONTE_CABEC)
        # Dados
        for li, it in enumerate(por_sup[sup], 2):
            row_v, _, _, fill, font, _, data_err, _ = it
            for j, v in enumerate(row_v[:ncols], 1):
                disp = None if str(v or '').strip() in ('nan','None','') else v
                c2 = ws.cell(li, j, disp)
                c2.fill = fill; c2.font = font; c2.alignment = AL_C
                if idx_data is not None and j == idx_data + 1 and data_err:
                    c2.fill = F_DATA_ERR; c2.font = FONTE_BRANCA

    # Resumo de retorno
    cnt_t = _cnt(processadas)
    resumo = {
        'total': cnt_t['total'],
        'nao_mapeados': len(nao_mapeados),
        'data_erros': len(erros_data),
        'laranjas': len(laranjas),
        'por_supervisor': {s: _cnt(v) for s, v in por_sup.items()},
    }

    out = io.BytesIO()
    wb.save(out); out.seek(0)
    return out.getvalue(), resumo


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Uso: python gerar_relatorio_prazos.py <entrada.xlsx> [saida.xlsx]')
        sys.exit(1)
    entrada = sys.argv[1]
    saida   = sys.argv[2] if len(sys.argv) > 2 else entrada.replace('.xlsx','_RELATORIO.xlsx')
    out_bytes, res = gerar_relatorio(entrada)
    with open(saida, 'wb') as f: f.write(out_bytes)
    print(f"Salvo: {saida} | Total: {res['total']} | "
          f"Nao mapeados: {res['nao_mapeados']} | Erros data: {res['data_erros']}")
