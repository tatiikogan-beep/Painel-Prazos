"""
Gerador de Relatório de Prazos Preclusivos — Versão Web (PAINEL_WEB)
Pode ser usado como módulo (import) ou executado diretamente.
"""

import re
import io
import sys
import unicodedata
from datetime import date, timedelta

import numpy as np
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

try:
    import holidays as holidays_lib
    _USE_HOLIDAYS = True
except ImportError:
    _USE_HOLIDAYS = False

# ─── FERIADOS NACIONAIS (sem Carnaval e Corpus Christi) ──────────────────────

def _feriados(anos):
    if not _USE_HOLIDAYS:
        return []
    br = holidays_lib.Brazil(years=anos)
    excluir = [d for d, n in br.items() if 'carnaval' in n.lower() or 'corpus' in n.lower()]
    for d in excluir:
        del br[d]
    return sorted(br.keys())

# ─── NORMALIZAÇÃO ────────────────────────────────────────────────────────────

def _norm(texto):
    t = str(texto).lower().strip()
    nfd = unicodedata.normalize('NFD', t)
    return ''.join(c for c in nfd if unicodedata.category(c) != 'Mn')

# ─── SUPERVISORES ────────────────────────────────────────────────────────────

SUPERVISORES = [
    'HELANZIA DE ARAUJO XAVIER WICHMANN',
    'LUCIANE MODERNEL MENDES',
    'GABRIEL GIORGIO CICCHELERO',
    'JENIFFER ROSA BARBOSA DE SALES',
    'JULIANA MIRELLA ALVES RODRIGUES',
    'NAYANDERSON LUAN MELLO PINHEIRO',
    'YURI ALVES BARROS DOS SANTOS',
    'MARCELLE LEITE RENTROIA',
    'SUZANA MARIA CAMPOS MARANHAO DE LIMA',
    'TICIANNA PIRES DE SOUZA',
    'RONALD FEITOSA AGUIAR FILHO',
]

# Fragmento normalizado → supervisor (mais específico primeiro)
_MAPA = [
    ('roberta rayanne',           'HELANZIA DE ARAUJO XAVIER WICHMANN'),
    ('wichamnn',                  'HELANZIA DE ARAUJO XAVIER WICHMANN'),
    ('wichmann',                  'HELANZIA DE ARAUJO XAVIER WICHMANN'),
    ('wellington',                'HELANZIA DE ARAUJO XAVIER WICHMANN'),
    ('keliane',                   'HELANZIA DE ARAUJO XAVIER WICHMANN'),
    ('natalia',                   'HELANZIA DE ARAUJO XAVIER WICHMANN'),
    ('gustavo',                   'HELANZIA DE ARAUJO XAVIER WICHMANN'),
    ('artur',                     'HELANZIA DE ARAUJO XAVIER WICHMANN'),
    ('victor',                    'HELANZIA DE ARAUJO XAVIER WICHMANN'),
    ('helanzia',                  'HELANZIA DE ARAUJO XAVIER WICHMANN'),
    ('erika paula',               'LUCIANE MODERNEL MENDES'),
    ('antonio eduardo',           'LUCIANE MODERNEL MENDES'),
    ('eduardo blasques',          'LUCIANE MODERNEL MENDES'),
    ('sane borges',               'LUCIANE MODERNEL MENDES'),
    ('matheus',                   'LUCIANE MODERNEL MENDES'),
    ('layla',                     'LUCIANE MODERNEL MENDES'),
    ('luciane',                   'LUCIANE MODERNEL MENDES'),
    ('juliana de oliveira rocha', 'GABRIEL GIORGIO CICCHELERO'),
    ('irene flavia',              'GABRIEL GIORGIO CICCHELERO'),
    ('armando helio',             'GABRIEL GIORGIO CICCHELERO'),
    ('gabriel giorgio',           'GABRIEL GIORGIO CICCHELERO'),
    ('cicchelero',                'GABRIEL GIORGIO CICCHELERO'),
    ('alysson',                   'GABRIEL GIORGIO CICCHELERO'),
    ('jamile',                    'GABRIEL GIORGIO CICCHELERO'),
    ('paulo marcio',              'JENIFFER ROSA BARBOSA DE SALES'),
    ('jeniffer',                  'JENIFFER ROSA BARBOSA DE SALES'),
    ('juliana mirella',           'JULIANA MIRELLA ALVES RODRIGUES'),
    ('thallys',                   'JULIANA MIRELLA ALVES RODRIGUES'),
    ('yuri gondim',               'NAYANDERSON LUAN MELLO PINHEIRO'),
    ('andre viana',               'NAYANDERSON LUAN MELLO PINHEIRO'),
    ('emerson',                   'NAYANDERSON LUAN MELLO PINHEIRO'),
    ('nayanderson',               'NAYANDERSON LUAN MELLO PINHEIRO'),
    ('yuri alves',                'YURI ALVES BARROS DOS SANTOS'),
    ('yuri barros',               'YURI ALVES BARROS DOS SANTOS'),
    ('luiz guilherme',            'YURI ALVES BARROS DOS SANTOS'),
    ('yuri santos',               'YURI ALVES BARROS DOS SANTOS'),
    ('mariana mota',              'MARCELLE LEITE RENTROIA'),
    ('marcelle',                  'MARCELLE LEITE RENTROIA'),
    ('giovanna cesar',            'SUZANA MARIA CAMPOS MARANHAO DE LIMA'),
    ('giovanna',                  'SUZANA MARIA CAMPOS MARANHAO DE LIMA'),
    ('francoise',                 'SUZANA MARIA CAMPOS MARANHAO DE LIMA'),
    ('leticia',                   'SUZANA MARIA CAMPOS MARANHAO DE LIMA'),
    ('tatiane',                   'SUZANA MARIA CAMPOS MARANHAO DE LIMA'),
    ('daniel',                    'SUZANA MARIA CAMPOS MARANHAO DE LIMA'),
    ('suzana',                    'SUZANA MARIA CAMPOS MARANHAO DE LIMA'),
    ('ticianna',                  'TICIANNA PIRES DE SOUZA'),
    ('alexia alencar',            'RONALD FEITOSA AGUIAR FILHO'),
    ('alexia capibaribe',         'RONALD FEITOSA AGUIAR FILHO'),
    ('ronald',                    'RONALD FEITOSA AGUIAR FILHO'),
    # Fallback: "yuri" sozinho → supervisor Yuri (após yuri gondim acima)
    ('yuri',                      'YURI ALVES BARROS DOS SANTOS'),
    # Fallback: "gabriel" sozinho → Gabriel (após gabriel giorgio acima)
    ('gabriel',                   'GABRIEL GIORGIO CICCHELERO'),
]

_MAPA_NORM = [(_norm(k), v) for k, v in _MAPA]
_SUP_NORM  = {_norm(s): s for s in SUPERVISORES}


def _mapear_supervisor(executor_raw):
    if not executor_raw or str(executor_raw).strip() in ('', 'None', 'nan'):
        return None
    nome = str(executor_raw).split('\n')[0].strip()
    n = _norm(nome)
    # Correspondência exata com supervisor
    if n in _SUP_NORM:
        return _SUP_NORM[n]
    # Supervisor contido no nome do executor (nome completo no campo)
    for sup_n, sup in _SUP_NORM.items():
        if sup_n in n:
            return sup
    # Fragmento do mapa (ordenado, mais específico primeiro)
    for frag, sup in _MAPA_NORM:
        if frag in n:
            return sup
    return None

# ─── CORES ───────────────────────────────────────────────────────────────────

FILLS = {
    'verde':   PatternFill('solid', fgColor='C6EFCE'),
    'amarelo': PatternFill('solid', fgColor='FFEB9C'),
    'rosa':    PatternFill('solid', fgColor='FFB6C1'),
    'vermelho':PatternFill('solid', fgColor='9C0006'),
    'laranja': PatternFill('solid', fgColor='FF9900'),
}
FILL_DATA_ERR = PatternFill('solid', fgColor='9C0006')
FILL_TOTAL    = PatternFill('solid', fgColor='BDD7EE')
FILL_HEADER   = PatternFill('solid', fgColor='2F5496')
FONT_WHITE    = Font(color='FFFFFF', bold=True)
FONT_HEADER   = Font(bold=True, color='FFFFFF')
FONT_BOLD     = Font(bold=True)
AL_CENTER     = Alignment(horizontal='center', vertical='center', wrap_text=True)
AL_LEFT       = Alignment(horizontal='left',   vertical='center', wrap_text=True)

TST_EXCECAO = 'protocolar recurso sobre decisao tst ou justificar nao recurso'

# ─── DATA FATAL DA DESCRIÇÃO ─────────────────────────────────────────────────

_PAT_DATA = re.compile(r'(\d{1,2})[/\.](\d{1,2})[/\.](\d{2,4})')


def _extrair_data_fatal(descricao, ano_ref):
    """Retorna a data máxima encontrada na descrição, ou None."""
    resultados = []
    for m in _PAT_DATA.finditer(str(descricao)):
        d, mo, a_str = int(m.group(1)), int(m.group(2)), m.group(3)
        a = int(a_str)
        if a < 100:
            a = (ano_ref // 100) * 100 + a
        elif len(a_str) == 3:
            # ano truncado ex: "206" → inferir do ano de referência
            a = int(str(ano_ref)[:1] + a_str)
        try:
            resultados.append(date(a, mo, d))
        except ValueError:
            pass
    return max(resultados) if resultados else None


def _d1(data_fatal, fer_np):
    """Dia útil imediatamente anterior à data_fatal."""
    prev = np.datetime64(data_fatal.strftime('%Y-%m-%d'), 'D') - np.timedelta64(1, 'D')
    d1_np = np.busday_offset(prev, 0, roll='backward', holidays=fer_np)
    return date.fromisoformat(str(d1_np))


def _proxima_util(ref_date, fer_np):
    """Próximo dia útil após ref_date."""
    d_np = np.datetime64(ref_date.strftime('%Y-%m-%d'), 'D')
    prox = np.busday_offset(d_np, 1, roll='forward', holidays=fer_np)
    return date.fromisoformat(str(prox))

# ─── LÓGICA DE COR ───────────────────────────────────────────────────────────

def _cor_linha(descricao, justificativa, data_cadastrada, hoje, fer_np):
    """
    Retorna (cor_display, cor_base, data_err_flag).
    cor_base: categoria para o RESUMO (verde/amarelo/rosa/vermelho)
    cor_display: cor visual da linha (pode ser laranja ou verde)
    data_err_flag: True se data cadastrada ≠ D-1 calculado
    """
    desc = str(descricao or '')
    just = str(justificativa or '').strip()
    desc_upper = desc.upper()

    # ── VERDE FORÇADO: audiências não são prazos preclusivos ──
    if 'AUDIENCIA DE JULGAMENTO' in _norm(desc_upper) or 'ACOMPANHAR JULGAMENTO' in _norm(desc_upper):
        return 'verde', 'verde', False

    # ── LARANJA: justificativa preenchida OU "não/nao" na descrição ──
    tem_nao = bool(re.search(r'\bn[aã]o\b', desc, re.IGNORECASE))
    excecao_tst = TST_EXCECAO in _norm(desc)
    laranja = bool(just) or (tem_nao and not excecao_tst)

    # ── DATA FATAL ──
    ano_ref = hoje.year
    if data_cadastrada and hasattr(data_cadastrada, 'year'):
        ano_ref = data_cadastrada.year

    data_fatal = _extrair_data_fatal(desc, ano_ref)
    data_err = False
    amanha = hoje + timedelta(days=1)

    if data_fatal:
        d1_calc = _d1(data_fatal, fer_np)
        dc = data_cadastrada
        if dc and hasattr(dc, 'date'):
            dc = dc.date()
        if dc and dc != d1_calc:
            data_err = True
        if data_fatal < hoje:
            cor_base = 'vermelho'
        elif data_fatal == hoje:
            cor_base = 'rosa'
        elif data_fatal == amanha:
            cor_base = 'amarelo'
        else:
            cor_base = 'verde'
    else:
        # Sem data na descrição: data cadastrada é o D-1; fatal = próximo útil
        if data_cadastrada:
            dc = data_cadastrada
            if hasattr(dc, 'date'):
                dc = dc.date()
            data_fatal_inf = _proxima_util(dc, fer_np)
            if data_fatal_inf < hoje:
                cor_base = 'vermelho'
            elif data_fatal_inf == hoje:
                cor_base = 'rosa'
            elif data_fatal_inf == amanha:
                cor_base = 'amarelo'
            else:
                cor_base = 'verde'
        else:
            cor_base = 'verde'

    cor_display = 'laranja' if laranja else cor_base
    return cor_display, cor_base, data_err

# ─── DETECÇÃO DE COLUNAS ─────────────────────────────────────────────────────

def _detectar_colunas(headers):
    """Retorna dict com índices (0-based) das colunas-chave."""
    cols = {}
    for i, h in enumerate(headers):
        hn = _norm(str(h or ''))
        if 'execut' in hn or 'responsav' in hn:
            cols.setdefault('executor', i)
        if 'descri' in hn:
            cols.setdefault('descricao', i)
        if 'justif' in hn:
            cols.setdefault('justificativa', i)
        if hn in ('data', 'data cadastro', 'data de cadastro', 'dt cadastro'):
            cols.setdefault('data', i)
        elif 'data' in hn and 'cadas' in hn:
            cols.setdefault('data', i)
        if 'process' in hn or 'numero' in hn or 'nro' in hn:
            cols.setdefault('processo', i)
    return cols

# ─── GERAÇÃO DO RELATÓRIO ────────────────────────────────────────────────────

def gerar_relatorio(input_source):
    """
    Parâmetro:
        input_source: caminho (str) ou bytes do arquivo Excel de entrada.
    Retorna:
        (output_bytes: bytes, resumo: dict)
    """
    if isinstance(input_source, (str, bytes)):
        if isinstance(input_source, str):
            wb_in = load_workbook(input_source, data_only=True)
        else:
            wb_in = load_workbook(io.BytesIO(input_source), data_only=True)
    else:
        wb_in = load_workbook(input_source, data_only=True)

    ws_in = wb_in.active
    rows_raw = list(ws_in.iter_rows(values_only=True))
    if not rows_raw:
        raise ValueError("Planilha vazia.")

    # Header
    header = list(rows_raw[0])
    dados = rows_raw[1:]
    cols = _detectar_colunas(header)

    # Feriados para os anos presentes nos dados
    hoje = date.today()
    anos = {hoje.year, hoje.year + 1}
    fer_list = _feriados(list(anos))
    fer_np = np.array([d.strftime('%Y-%m-%d') for d in fer_list], dtype='datetime64[D]')

    # ── Processar cada linha ─────────────────────────────────────────────────
    processadas = []   # (row_values, supervisor, executor, cor_display, cor_base, data_err)
    nao_mapeados = []  # (row_values, executor_raw)

    for row in dados:
        if all(v is None or str(v).strip() == '' for v in row):
            continue

        executor_raw = row[cols['executor']] if 'executor' in cols else ''
        descricao    = row[cols['descricao']] if 'descricao' in cols else ''
        justificativa = row[cols['justificativa']] if 'justificativa' in cols else ''
        data_cad     = row[cols['data']] if 'data' in cols else None

        supervisor = _mapear_supervisor(executor_raw)
        cor_d, cor_b, data_err = _cor_linha(descricao, justificativa, data_cad, hoje, fer_np)

        if supervisor:
            processadas.append((list(row), supervisor, str(executor_raw or ''), cor_d, cor_b, data_err))
        else:
            nao_mapeados.append((list(row), str(executor_raw or '')))

    # ── Contagens para RESUMO ────────────────────────────────────────────────
    def _contar(linhas_filtradas):
        c = {'total': 0, 'amarelo': 0, 'rosa': 0, 'vermelho': 0, 'verde': 0}
        for _, _, _, _, cor_b, _ in linhas_filtradas:
            c['total'] += 1
            c[cor_b] = c.get(cor_b, 0) + 1
        c['atraso'] = c['rosa'] + c['vermelho']
        return c

    por_supervisor = {}
    por_executor   = {}

    for item in processadas:
        sup = item[1]
        exe = item[2].split('\n')[0].strip()
        por_supervisor.setdefault(sup, []).append(item)
        por_executor.setdefault(exe, []).append(item)

    # ── Montar workbook de saída ─────────────────────────────────────────────
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)  # remove sheet padrão

    # ── ABA RESUMO ──────────────────────────────────────────────────────────
    ws_r = wb_out.create_sheet('RESUMO', 0)
    col_headers = ['Responsável', 'Total de Prazos', 'D-1 🟡', 'Fatal 🌸',
                   'Prazo Vencido 🔴', 'Dentro do Prazo 🟢', 'Total em Atraso']

    def _escrever_tabela_resumo(ws, linha_inicio, titulo, dados_dict):
        # Título
        ws.cell(linha_inicio, 1, titulo).font = Font(bold=True, size=12)
        linha_inicio += 1
        # Header
        for j, h in enumerate(col_headers, 1):
            c = ws.cell(linha_inicio, j, h)
            c.fill = FILL_HEADER
            c.font = FONT_HEADER
            c.alignment = AL_CENTER
        linha_inicio += 1
        # Dados
        total_geral = {k: 0 for k in ['total','amarelo','rosa','vermelho','verde','atraso']}
        for chave in sorted(dados_dict.keys()):
            cont = _contar(dados_dict[chave])
            valores = [chave, cont['total'], cont['amarelo'], cont['rosa'],
                       cont['vermelho'], cont['verde'], cont['atraso']]
            for j, v in enumerate(valores, 1):
                c = ws.cell(linha_inicio, j, v)
                c.alignment = AL_CENTER if j > 1 else AL_LEFT
            linha_inicio += 1
            for k in total_geral:
                total_geral[k] += cont[k]
        # Total
        tot_vals = ['TOTAL GERAL', total_geral['total'], total_geral['amarelo'],
                    total_geral['rosa'], total_geral['vermelho'],
                    total_geral['verde'], total_geral['atraso']]
        for j, v in enumerate(tot_vals, 1):
            c = ws.cell(linha_inicio, j, v)
            c.fill = FILL_TOTAL
            c.font = FONT_BOLD
            c.alignment = AL_CENTER if j > 1 else AL_LEFT
        return linha_inicio + 2  # pula 1 linha entre tabelas

    prox_linha = 1
    prox_linha = _escrever_tabela_resumo(ws_r, prox_linha, 'RESUMO POR SUPERVISOR', por_supervisor)
    prox_linha = _escrever_tabela_resumo(ws_r, prox_linha, 'RESUMO POR EXECUTOR', por_executor)

    # Seção de datas com erro
    if any(item[5] for item in processadas):
        ws_r.cell(prox_linha, 1, 'DATAS FORA DO D-1 ÚTIL').font = Font(bold=True, size=12, color='9C0006')
        prox_linha += 1
        for h_i, h in enumerate(['Processo', 'Executor', 'Supervisor', 'Data Cadastrada', 'Descrição'], 1):
            c = ws_r.cell(prox_linha, h_i, h)
            c.fill = FILL_HEADER
            c.font = FONT_HEADER
            c.alignment = AL_CENTER
        prox_linha += 1
        for item in processadas:
            if item[5]:
                row_v, sup, exe, _, _, _ = item
                proc = row_v[cols.get('processo', 0)] if 'processo' in cols else ''
                data = row_v[cols.get('data', 0)] if 'data' in cols else ''
                desc = row_v[cols.get('descricao', 0)] if 'descricao' in cols else ''
                for j, v in enumerate([proc, exe, sup, data, desc], 1):
                    ws_r.cell(prox_linha, j, v).alignment = AL_LEFT
                prox_linha += 1
        prox_linha += 1

    # Não mapeados
    if nao_mapeados:
        ws_r.cell(prox_linha, 1, 'EXECUTORES NÃO MAPEADOS').font = Font(bold=True, size=12, color='FF0000')
        prox_linha += 1
        for item in nao_mapeados:
            row_v, exe = item
            proc = row_v[cols.get('processo', 0)] if 'processo' in cols else ''
            ws_r.cell(prox_linha, 1, exe).alignment = AL_LEFT
            ws_r.cell(prox_linha, 2, str(proc)).alignment = AL_LEFT
            prox_linha += 1

    # Largura das colunas do RESUMO
    ws_r.column_dimensions['A'].width = 40
    for col_l in 'BCDEFG':
        ws_r.column_dimensions[col_l].width = 18

    # ── ABAS POR SUPERVISOR ─────────────────────────────────────────────────
    idx_data = cols.get('data')
    for sup in SUPERVISORES:
        if sup not in por_supervisor:
            continue
        linhas_sup = por_supervisor[sup]
        # Nome da aba: primeiras palavras do supervisor (máx 31 chars)
        nome_aba = sup[:31]
        ws = wb_out.create_sheet(nome_aba)

        # Header
        for j, h in enumerate(header, 1):
            c = ws.cell(1, j, h)
            c.fill = FILL_HEADER
            c.font = FONT_HEADER
            c.alignment = AL_CENTER

        # Dados
        for linha_i, item in enumerate(linhas_sup, 2):
            row_v, _, _, cor_d, _, data_err = item
            fill = FILLS[cor_d]
            font = FONT_WHITE if cor_d == 'vermelho' else Font()

            for j, v in enumerate(row_v, 1):
                c = ws.cell(linha_i, j, v)
                c.fill = fill
                c.font = font
                c.alignment = AL_CENTER

                # Flag de data divergente
                if idx_data is not None and j == idx_data + 1 and data_err:
                    c.fill = FILL_DATA_ERR
                    c.font = FONT_WHITE

        # Larguras
        ws.column_dimensions['A'].width = 20
        for col_i in range(2, len(header) + 1):
            ws.column_dimensions[get_column_letter(col_i)].width = 22

    # ── Estatísticas para retorno ────────────────────────────────────────────
    total_c = _contar(processadas)
    resumo = {
        'total':      total_c['total'],
        'nao_mapeados': len(nao_mapeados),
        'data_erros': sum(1 for item in processadas if item[5]),
        'por_supervisor': {s: _contar(v) for s, v in por_supervisor.items()},
    }

    buf = io.BytesIO()
    wb_out.save(buf)
    buf.seek(0)
    return buf.getvalue(), resumo


# ─── EXECUÇÃO DIRETA ────────────────────────────────────────────────────────

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Uso: python gerar_relatorio_prazos.py <arquivo_entrada.xlsx>")
        sys.exit(1)
    entrada = sys.argv[1]
    saida = sys.argv[2] if len(sys.argv) > 2 else entrada.replace('.xlsx', '_RELATORIO.xlsx')
    print(f"Processando: {entrada}")
    output_bytes, resumo = gerar_relatorio(entrada)
    with open(saida, 'wb') as f:
        f.write(output_bytes)
    print(f"Salvo em: {saida}")
    print(f"Total: {resumo['total']} | Não mapeados: {resumo['nao_mapeados']} | Erros de data: {resumo['data_erros']}")
