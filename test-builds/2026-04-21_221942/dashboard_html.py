import json
﻿import json
import os
import locale
import re
from datetime import datetime
from text_utils import write_text_utf8

# Importações para integração com waterfy_engine e cruzamento de dados
try:
    import waterfy_engine as wfy_eng
    HAS_WATERFY = True
except ImportError:
    HAS_WATERFY = False

# ── Compilar regex uma vez (reutilizada em múltiplas funções) ────────────────
_REGEX_TELEFONE = re.compile(r"^(?:55)?([1-9]\d{9,10})$")

def _fmt_brl(v: float) -> str:
    """Formata valor em BRL com locale otimizado."""
    try:
        # Usar locale.currency() é mais rápido que 3 .replace()
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
        return locale.currency(v, symbol=True, grouping=True)
    except (locale.Error, AttributeError):
        # Fallback mantém compatibilidade
        return f'R$ {v:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')

def _mini_insight(emoji: str, cor: str, texto: str) -> str:
    """Gera HTML de um mini-insight para exibir abaixo de um gráfico."""
    cor_map = {
        'green':  ('rgba(46,125,50,0.08)',  '#2E7D32', 'rgba(46,125,50,0.35)'),
        'red':    ('rgba(198,40,40,0.08)',   '#C62828', 'rgba(198,40,40,0.35)'),
        'orange': ('rgba(230,115,0,0.08)',   '#E65100', 'rgba(230,115,0,0.35)'),
        'blue':   ('rgba(21,101,192,0.08)',  '#1565C0', 'rgba(21,101,192,0.35)'),
        '':       ('rgba(100,116,139,0.07)', '#475569', 'rgba(100,116,139,0.25)'),
    }
    bg, txt_cor, border = cor_map.get(cor, cor_map[''])
    return (
        f'<div style="display:flex;align-items:flex-start;gap:8px;padding:9px 13px;'
        f'background:{bg};border-left:3px solid {border};border-radius:0 7px 7px 0;'
        f'margin-top:6px;font-size:0.82rem;line-height:1.45;color:{txt_cor};">'
        f'<span style="font-size:1rem;flex-shrink:0;margin-top:1px">{emoji}</span>'
        f'<span>{texto}</span></div>'
    )

def _mini_insights_html(items: list) -> str:
    """Converte lista de dicts {emoji, cor, texto} em HTML."""
    return ''.join(_mini_insight(i['emoji'], i['cor'], i['texto']) for i in items)


def _insights_cards_html(items: list) -> str:
  """Converte insights em cards com tom semântico (positive/warning/info/neutral)."""
  if not items:
    return (
      '<div class="insights-cards">'
      '<div class="insight-card tone-neutral">'
      '<div class="insight-title">Sem insights suficientes</div>'
      '<div class="insight-text">Não há dados suficientes para gerar recomendações no período.</div>'
      '</div></div>'
    )

  cards = []
  for it in items:
    tone = str(it.get('tone', 'neutral')).strip().lower()
    if tone not in {'positive', 'warning', 'info', 'neutral'}:
      tone = 'neutral'
    title = str(it.get('title', 'Insight')).strip()
    text = str(it.get('text', '')).strip()
    cards.append(
      f'<div class="insight-card tone-{tone}">'
      f'<div class="insight-title">{title}</div>'
      f'<div class="insight-text">{text}</div>'
      '</div>'
    )

  return '<div class="insights-cards">' + ''.join(cards) + '</div>'

def _calcular_mini_insights_8091(d91: dict) -> dict:
    """
    Gera mini-insights (melhor/pior/média) para cada gráfico da aba 8091.
    Retorna dict com chave = id do gráfico, valor = HTML dos mini-insights.
    """
    kpis       = d91.get('kpis', {})
    sit91      = d91.get('situacao', [])
    fase91     = d91.get('fase', [])
    cat91      = d91.get('categoria', [])
    bairros91  = d91.get('bairros', [])
    grupos91   = d91.get('grupos', [])
    grp_rec    = d91.get('grupos_receita', [])
    leituristas= d91.get('leiturista', [])
    czero      = d91.get('consumo_zero', [])
    faixas_vol = d91.get('faixas_vol', {})
    faixas     = d91.get('faixas', {})
    serie      = d91.get('serie_meses', [])
    leit_dia   = d91.get('leit_dia', [])

    total_fat  = kpis.get('total_fat', 0)
    n_lig      = kpis.get('n_ligacoes', 0)
    vol_total  = kpis.get('vol_total', 0)
    result     = {}

    # ── Situação das Faturas ──────────────────────────────────────────────
    if sit91:
        sit_total = sum(s.get('qtd', 0) for s in sit91)
        sit_sorted_pct = sorted(sit91, key=lambda x: x.get('qtd', 0) / sit_total if sit_total else 0, reverse=True)
        melhor = sit_sorted_pct[0]
        pior   = sit_sorted_pct[-1]
        pct_m  = melhor['qtd'] / sit_total * 100 if sit_total else 0
        pct_p  = pior['qtd']   / sit_total * 100 if sit_total else 0
        items = [
            {'emoji': '', 'cor': 'green',
             'texto': f'<strong>Situação mais frequente:</strong> {melhor["label"]} com {melhor["qtd"]:,} faturas ({pct_m:.1f}% da base)'},
            {'emoji': '', 'cor': 'red' if 'DEBITO' in pior['label'].upper() or 'DÉBITO' in pior['label'].upper() else 'orange',
             'texto': f'<strong>Situação menos frequente:</strong> {pior["label"]} com {pior["qtd"]:,} faturas ({pct_p:.1f}%)'},
        ]
        result['chart-91-situacao'] = _mini_insights_html(items)

    # ── Fase de Cobrança ──────────────────────────────────────────────────
    if fase91:
        fase_s = sorted(fase91, key=lambda x: x.get('qtd', 0), reverse=True)
        top    = fase_s[0]
        bot    = fase_s[-1]
        tot_f  = sum(f.get('qtd', 0) for f in fase91)
        criticas = [f for f in fase91 if any(x in f.get('label','').upper() for x in ['COBRAN','JUDICI','PROTESTO'])]
        tot_crit = sum(f['qtd'] for f in criticas)
        items = [
            {'emoji': '', 'cor': '',
             'texto': f'<strong>Fase predominante:</strong> {top["label"]} com {top["qtd"]:,} ligações ({top["qtd"]/tot_f*100:.1f}% da base)'},
            {'emoji': '', 'cor': 'red' if tot_crit/tot_f*100 > 10 else 'orange',
             'texto': f'<strong>Cobrança crítica:</strong> {tot_crit:,} ligações em fases avançadas ({tot_crit/tot_f*100:.1f}%)  {_fmt_brl(sum(f.get("valor",0) for f in criticas))} em risco'},
        ]
        result['chart-91-fase'] = _mini_insights_html(items)

    # ── Composição do Faturamento ─────────────────────────────────────────
    comp = d91.get('comp', {})
    if comp and total_fat > 0:
        agua  = comp.get('agua', 0)
        multa = comp.get('multa', 0)
        outros= comp.get('outros', 0)
        pct_a = agua  / total_fat * 100 if total_fat else 0
        pct_m = multa / total_fat * 100 if total_fat else 0
        obs_m = '  nível elevado, sinal de inadimplência histórica' if pct_m > 5 else '  dentro do esperado'
        items = [
            {'emoji': '', 'cor': 'blue',
             'texto': f'<strong>Água</strong> é o maior componente: {_fmt_brl(agua)} ({pct_a:.1f}% do total faturado)'},
            {'emoji': '', 'cor': 'orange' if pct_m > 5 else '',
             'texto': f'<strong>Multas e juros:</strong> {_fmt_brl(multa)} ({pct_m:.1f}%){obs_m}'},
        ]
        result['chart-91-comp'] = _mini_insights_html(items)

    # ── Faturamento por Categoria ─────────────────────────────────────────
    if cat91 and total_fat > 0:
        cat_s  = sorted(cat91, key=lambda x: x.get('valor', 0), reverse=True)
        top_c  = cat_s[0]
        bot_c  = cat_s[-1]
        media_c = total_fat / len(cat91)
        items = [
            {'emoji': '', 'cor': 'green',
             'texto': f'<strong>{top_c["label"]}</strong> lidera com {_fmt_brl(top_c["valor"])} ({top_c["valor"]/total_fat*100:.1f}% do faturamento, {top_c["qtd"]:,} ligações)'},
            {'emoji': '', 'cor': '',
             'texto': f'<strong>{bot_c["label"]}</strong> tem menor participação: {_fmt_brl(bot_c["valor"])} ({bot_c["valor"]/total_fat*100:.1f}%, {bot_c["qtd"]:,} lig.)'},
            {'emoji': '', 'cor': '',
             'texto': f'<strong>Média por categoria:</strong> {_fmt_brl(media_c)}  {len(cat91)} categorias ativas na base'},
        ]
        result['chart-91-categoria'] = _mini_insights_html(items)

    # ── Top Bairros ───────────────────────────────────────────────────────
    if bairros91 and total_fat > 0:
        top_b  = bairros91[0]
        bot_b  = bairros91[-1]
        media_b = total_fat / len(bairros91)
        ticket_top = top_b['valor'] / top_b['qtd'] if top_b.get('qtd', 0) > 0 else 0
        ticket_bot = bot_b['valor'] / bot_b['qtd'] if bot_b.get('qtd', 0) > 0 else 0
        items = [
            {'emoji': '', 'cor': 'green',
             'texto': f'<strong>{top_b["label"].title()}</strong> lidera: {_fmt_brl(top_b["valor"])} ({top_b["qtd"]:,} lig., ticket médio {_fmt_brl(ticket_top)})'},
            {'emoji': '', 'cor': '',
             'texto': f'<strong>{bot_b["label"].title()}</strong> tem menor faturamento: {_fmt_brl(bot_b["valor"])} ({bot_b["qtd"]:,} lig., ticket {_fmt_brl(ticket_bot)})'},
            {'emoji': '', 'cor': '',
             'texto': f'<strong>Média por bairro:</strong> {_fmt_brl(media_b)} entre os {len(bairros91)} bairros listados'},
        ]
        result['chart-91-bairros'] = _mini_insights_html(items)

    # ── Receita por m³ por Grupo ──────────────────────────────────────────
    if grp_rec and len(grp_rec) >= 2:
        validos = [g for g in grp_rec if g.get('vol', 0) > 0]
        if validos:
            top_g  = max(validos, key=lambda x: x['receita_m3'])
            bot_g  = min(validos, key=lambda x: x['receita_m3'])
            media_rm3 = sum(g['receita_m3'] for g in validos) / len(validos)
            diff   = (top_g['receita_m3'] - bot_g['receita_m3']) / bot_g['receita_m3'] * 100 if bot_g['receita_m3'] > 0 else 0
            obs = f'  diferença de {diff:.0f}% entre grupos, investigar perfil tarifário' if diff > 30 else ''
            items = [
                {'emoji': '', 'cor': 'green',
                 'texto': f'<strong>{top_g["label"]}</strong> tem maior eficiência: {_fmt_brl(top_g["receita_m3"])}/m³ ({top_g["vol"]:,.0f} m³ consumidos)'},
                {'emoji': '', 'cor': 'orange' if diff > 30 else '',
                 'texto': f'<strong>{bot_g["label"]}</strong> tem menor retorno: {_fmt_brl(bot_g["receita_m3"])}/m³{obs}'},
                {'emoji': '', 'cor': '',
                 'texto': f'<strong>Média geral:</strong> {_fmt_brl(media_rm3)}/m³ entre os {len(validos)} grupos de leitura'},
            ]
            result['chart-91-grupos'] = _mini_insights_html(items)

    # ── Evolução Mensal ───────────────────────────────────────────────────
    if len(serie) >= 2:
        top_s  = max(serie, key=lambda x: x.get('total', 0))
        bot_s  = min(serie, key=lambda x: x.get('total', 0))
        media_s = sum(s.get('total', 0) for s in serie) / len(serie)
        ult    = serie[-1]
        ante   = serie[-2]
        var    = (ult['total'] - ante['total']) / ante['total'] * 100 if ante['total'] else 0
        items = [
            {'emoji': '', 'cor': 'green',
             'texto': f'<strong>Melhor mês:</strong> {top_s["periodo"]} com {_fmt_brl(top_s["total"])} faturados'},
            {'emoji': '', 'cor': 'red',
             'texto': f'<strong>Pior mês:</strong> {bot_s["periodo"]} com {_fmt_brl(bot_s["total"])}  diferença de {_fmt_brl(top_s["total"]-bot_s["total"])} em relação ao melhor'},
            {'emoji': '', 'cor': 'green' if var >= 0 else 'red',
             'texto': f'<strong>Última variação:</strong> {var:+.1f}% de {ante["periodo"]} para {ult["periodo"]} (média do período: {_fmt_brl(media_s)})'},
        ]
        result['chart-8091-linha'] = _mini_insights_html(items)

    # ── Distribuição de Consumo por Faixa (m³) ────────────────────────────
    if faixas_vol:
        total_fv = sum(faixas_vol.values())
        top_fv  = max(faixas_vol.items(), key=lambda x: x[1])
        bot_fv  = min(faixas_vol.items(), key=lambda x: x[1])
        zeros   = next((v for k, v in faixas_vol.items() if '0' in k), 0)
        pct_zero = zeros / total_fv * 100 if total_fv else 0
        items = [
            {'emoji': '', 'cor': 'green',
             'texto': f'<strong>Faixa mais frequente:</strong> {top_fv[0]} com {top_fv[1]:,} ligações ({top_fv[1]/total_fv*100:.1f}% da base)'},
            {'emoji': '', 'cor': '',
             'texto': f'<strong>Faixa menos frequente:</strong> {bot_fv[0]} com {bot_fv[1]:,} ligações ({bot_fv[1]/total_fv*100:.1f}%)'},
        ]
        if pct_zero > 2:
            items.append({'emoji': '', 'cor': 'orange' if pct_zero <= 8 else 'red',
             'texto': f'<strong>{pct_zero:.1f}% das ligações com consumo zero</strong> ({zeros:,} lig.)  verificar hidrômetros e imóveis fechados'})
        result['chart-91-faixas-vol'] = _mini_insights_html(items)

    # ── Desempenho por Leiturista ─────────────────────────────────────────
    if leituristas and len(leituristas) >= 2:
        validos_l = [l for l in leituristas if l.get('qtd', 0) > 0]
        if validos_l:
            top_l   = max(validos_l, key=lambda x: x['vol'])
            bot_l   = min(validos_l, key=lambda x: x['vol'])
            media_vm = sum(l['vol_medio'] for l in validos_l) / len(validos_l)
            top_vm  = max(validos_l, key=lambda x: x['vol_medio'])
            items = [
                {'emoji': '', 'cor': 'green',
                 'texto': f'<strong>{top_l["label"].split()[0]}</strong> leu mais volume: {top_l["vol"]:,.0f} m³ em {top_l["qtd"]:,} ligações (média {top_l["vol_medio"]:.1f} m³/lig.)'},
                {'emoji': '', 'cor': '',
                 'texto': f'<strong>{bot_l["label"].split()[0]}</strong> teve menor volume: {bot_l["vol"]:,.0f} m³ em {bot_l["qtd"]:,} ligações'},
                {'emoji': '', 'cor': 'orange' if top_vm['vol_medio'] > media_vm * 1.35 else '',
                 'texto': f'<strong>Vol. médio da equipe: {media_vm:.1f} m³/lig.</strong>  {top_vm["label"].split()[0]} tem o maior índice ({top_vm["vol_medio"]:.1f} m³/lig.){"   acima do esperado" if top_vm["vol_medio"] > media_vm * 1.35 else ""}'},
            ]
            result['chart-91-leiturista'] = _mini_insights_html(items)

    # ── Consumo Zero ──────────────────────────────────────────────────────
    if czero:
        tot_z  = sum(c.get('qtd', 0) for c in czero)
        top_cz = max(czero, key=lambda x: x.get('qtd', 0))
        bot_cz = min(czero, key=lambda x: x.get('qtd', 0))
        pct_z  = tot_z / n_lig * 100 if n_lig else 0
        items = [
            {'emoji': '', 'cor': 'red' if top_cz['qtd']/tot_z*100 > 50 else 'orange',
             'texto': f'<strong>{top_cz["label"]}</strong> concentra mais zeros: {top_cz["qtd"]:,} ligações ({top_cz["qtd"]/tot_z*100:.1f}% dos zeros)'},
            {'emoji': '', 'cor': 'green',
             'texto': f'<strong>{bot_cz["label"]}</strong> tem menos zeros: {bot_cz["qtd"]:,} ligações ({bot_cz["qtd"]/tot_z*100:.1f}% dos zeros)'},
        ]
        if pct_z > 3:
            items.append({'emoji': '', 'cor': 'orange',
             'texto': f'<strong>Total de {tot_z:,} zeros ({pct_z:.1f}% da base)</strong>  acima de 3% indica necessidade de vistoria em campo'})
        result['chart-91-consumo-zero'] = _mini_insights_html(items)

    # ── Volume por Dia de Leitura ─────────────────────────────────────────
    if leit_dia:
        validos_d = [d for d in leit_dia if d.get('qtd', 0) > 0]
        if validos_d:
            top_d   = max(validos_d, key=lambda x: x['vol'])
            bot_d   = min(validos_d, key=lambda x: x['vol'])
            media_d = sum(d['vol'] for d in validos_d) / len(validos_d)
            top_vm_d = max(validos_d, key=lambda x: x['vol']/x['qtd'] if x['qtd'] > 0 else 0)
            items = [
                {'emoji': '', 'cor': 'green',
                 'texto': f'<strong>Dia com maior volume:</strong> {top_d["data"]}  {top_d["vol"]:,.0f} m³ em {top_d["qtd"]:,} ligações'},
                {'emoji': '', 'cor': '',
                 'texto': f'<strong>Dia com menor volume:</strong> {bot_d["data"]}  {bot_d["vol"]:,.0f} m³ em {bot_d["qtd"]:,} ligações'},
                {'emoji': '', 'cor': '',
                 'texto': f'<strong>Média diária:</strong> {media_d:,.0f} m³/dia de leitura  pico em {top_vm_d["data"]} com maior consumo médio por ligação'},
            ]
            result['chart-91-leit-dia'] = _mini_insights_html(items)

    return result



def _safe_num(v, default=0):
    """Sanitiza floats para JS: converte nan/inf/None em 0 (ou default)."""
    import math
    if v is None:
        return default
    try:
        f = float(v)
        if math.isnan(f) or math.isinf(f):
            return default
        return f
    except (TypeError, ValueError):
        return default


def _safe_json(v):
    """json.dumps seguro: converte nan/inf em 0 para evitar SyntaxError no JS."""
    import json, math
    def _fix(x):
        if isinstance(x, float):
            if math.isnan(x) or math.isinf(x):
                return 0
        if isinstance(x, list):
            return [_fix(i) for i in x]
        if isinstance(x, dict):
            return {k: _fix(val) for k, val in x.items()}
        return x
    return json.dumps(_fix(v))

def gerar_pn_html(dados: dict, caminho_saida: str, _return_parts: bool = False) -> str:
    """Gera um relatório PN mensal completo, com drilldown por dias, KPIs e comparativos."""
    import os
    import base64
    import mimetypes
    from datetime import datetime

    periodo = dados.get('periodo', {'inicial': 'N/A', 'final': 'N/A'})
    cidade = dados.get('cidade', 'ITAPOÁ')
    d17 = dados.get('macro_8117', {})
    d21 = dados.get('macro_8121', {})
    d91 = dados.get('macro_8091', {})

    # ── Identidade visual Itapoá (logo embutido em base64) ──
    logo_data_uri = ""
    logo_candidates = [
        "Itapoa_branca.png",
        "itapoa_informa.png",
        "Logo mathey tk 1.png",
    ]
    roots = [
        os.path.dirname(os.path.abspath(caminho_saida)),
        os.path.dirname(os.path.dirname(os.path.abspath(caminho_saida))),
        os.path.dirname(os.path.abspath(__file__)),
        os.getcwd(),
    ]

    for root in roots:
        if logo_data_uri:
            break
        for name in logo_candidates:
            path_logo = os.path.join(root, name)
            if not os.path.exists(path_logo):
                continue
            try:
                mime = mimetypes.guess_type(path_logo)[0] or "image/png"
                with open(path_logo, "rb") as f_logo:
                    b64 = base64.b64encode(f_logo.read()).decode("ascii")
                logo_data_uri = f"data:{mime};base64,{b64}"
                break
            except Exception:
                continue

    logo_html = (
        f'<img class="brand-logo" src="{logo_data_uri}" alt="Logo Itapoá Saneamento" style="max-width:64px;max-height:64px">'
        if logo_data_uri else
        '<div class="brand-logo-fallback" title="Itapoá Saneamento">IS</div>'
    )
    # Fallback seguro se logo_html ficar vazio
    if not logo_html or logo_html.isspace():
        logo_html = '<div class="brand-logo-fallback">IS</div>'

    def _fmt_k(v: float) -> str:
        if not v:
            return "-"
        val_k = v / 1000.0
        return f"R$ {val_k:,.1f}k".replace(',', 'X').replace('.', ',').replace('X', '.')

    # ── Dados base (8121) ──
    datas21 = d21.get('datas', [])
    fat_d = d21.get('fat_diario', [])
    fat_ac = d21.get('fat_acumulado', [])
    arr21_d = d21.get('arr_diario', [])
    arr21_ac = d21.get('arr_acumulado', [])

    # ── Consolidação mensal + drilldown diário ──
    meses = {}
    datas_validas = []
    for i, dt_str in enumerate(datas21):
        try:
            dt = datetime.strptime(str(dt_str), "%d/%m/%Y")
            datas_validas.append(dt)
        except Exception:
            continue

        chave = (dt.year, dt.month)
        mes_key = f"{dt.year:04d}-{dt.month:02d}"
        if chave not in meses:
            meses[chave] = {
                "key": mes_key,
                "label": dt.strftime("%m/%Y"),
                "fat": 0.0,
                "arr": 0.0,
                "dias": [],
            }

        fd = fat_d[i] if i < len(fat_d) else 0.0
        ad = arr21_d[i] if i < len(arr21_d) else 0.0
        meses[chave]["fat"] += fd
        meses[chave]["arr"] += ad
        meses[chave]["dias"].append({"dt": dt, "fd": fd, "ad": ad})

    meses_ordenados = [meses[k] for k in sorted(meses.keys(), reverse=True)]

    total_fat = sum(m["fat"] for m in meses_ordenados) if meses_ordenados else (fat_ac[-1] if fat_ac else 0.0)
    total_arr = sum(m["arr"] for m in meses_ordenados) if meses_ordenados else (arr21_ac[-1] if arr21_ac else 0.0)
    pct_realizado = (total_arr / total_fat * 100) if total_fat > 0 else 0.0
    gap = total_fat - total_arr
    inad_est = (gap / total_fat * 100) if total_fat > 0 else 0.0

    if datas_validas:
        periodo = {
            'inicial': min(datas_validas).strftime('%d/%m/%Y'),
            'final': max(datas_validas).strftime('%d/%m/%Y'),
        }

    # ── Metas/Orçamento (se disponível no payload) ──
    metas_mensais = {}
    metas_raw = dados.get("metas_mensais") or dados.get("metas") or {}
    if isinstance(metas_raw, dict):
        for k, v in metas_raw.items():
            ks = str(k).strip()
            try:
                if "/" in ks and len(ks) == 7:
                    mm, yyyy = ks.split("/")
                    metas_mensais[f"{int(yyyy):04d}-{int(mm):02d}"] = float(v)
                elif "-" in ks and len(ks) == 7:
                    yyyy, mm = ks.split("-")
                    metas_mensais[f"{int(yyyy):04d}-{int(mm):02d}"] = float(v)
            except Exception:
                continue

    meta_mensal_unica = dados.get("meta_mensal")
    if not metas_mensais and meta_mensal_unica is not None:
        try:
            meta_base = float(meta_mensal_unica)
            for m in meses_ordenados:
                metas_mensais[m["key"]] = meta_base
        except Exception:
            pass

    total_meta = sum(metas_mensais.get(m["key"], 0.0) for m in meses_ordenados)
    tem_meta = total_meta > 0
    desvio_meta = (total_arr - total_meta) if tem_meta else None
    ating_meta = (total_arr / total_meta * 100) if tem_meta and total_meta > 0 else None

    # ── Comparativos M-1 e M-12 ──
    idx_meses = {(int(m["key"][:4]), int(m["key"][5:7])): m for m in meses_ordenados}
    m_ref = meses_ordenados[-1] if meses_ordenados else None
    m_1 = meses_ordenados[-2] if len(meses_ordenados) >= 2 else None
    m_12 = None
    if m_ref is not None:
        yr = int(m_ref["key"][:4])
        mo = int(m_ref["key"][5:7])
        m_12 = idx_meses.get((yr - 1, mo))

    def _pct_var(a, b):
        return ((a - b) / b * 100) if b else None

    var_m1_arr = _pct_var(m_ref["arr"], m_1["arr"]) if (m_ref and m_1) else None
    var_m12_arr = _pct_var(m_ref["arr"], m_12["arr"]) if (m_ref and m_12) else None

    # ── Fechamento e YTD ──
    now = datetime.now()
    mes_em_andamento = None
    mes_fechado = None
    ytd_fat = 0.0
    ytd_arr = 0.0
    if m_ref is not None:
        ref_year = int(m_ref["key"][:4])
        ref_month = int(m_ref["key"][5:7])
        if ref_year == now.year and ref_month == now.month:
            mes_em_andamento = m_ref["label"]
            mes_fechado = m_1["label"] if m_1 else "-"
        else:
            mes_fechado = m_ref["label"]
            mes_em_andamento = "-"

        for m in meses_ordenados:
            if int(m["key"][:4]) == ref_year:
                ytd_fat += m["fat"]
                ytd_arr += m["arr"]

    # ── KPIs operacionais de cobrança ──
    formas = d17.get('formas', {})
    total_q_formas = 0.0
    total_v_formas = 0.0
    if isinstance(formas, dict):
        for _, val in formas.items():
            if not isinstance(val, dict):
                continue
            total_q_formas += float(val.get('qtde', 0) or 0)
            total_v_formas += float(val.get('valor', 0) or 0)
    ticket_medio_arr = (total_v_formas / total_q_formas) if total_q_formas > 0 else 0.0
    recuperacao_gap = (total_arr / gap * 100) if gap > 0 else 100.0

    # ── Qualidade de dados ──
    alertas = []
    for m in meses_ordenados:
        mes_alerta_ano = int(m["key"][:4])
        mes_alerta_num = int(m["key"][5:7])
        mes_alerta_atual = mes_alerta_ano == now.year and mes_alerta_num == now.month
        if m["fat"] == 0 and m["arr"] > 0:
            alertas.append(f"{m['label']}: arrecadação sem faturamento no mês.")
        if (not mes_alerta_atual) and m["fat"] > 0 and m["arr"] > m["fat"] * 2:
            alertas.append(f"{m['label']}: arrecadação muito acima do faturamento mensal (>200%).")
    if not datas_validas:
        alertas.append("Base 8121 sem datas válidas no formato DD/MM/AAAA.")
    if len(meses_ordenados) <= 1:
        alertas.append("Somente um mês consolidado foi detectado na base.")

    # ── Insights automáticos ──
    insights_cards = []
    if meses_ordenados:
        melhor_arr = max(meses_ordenados, key=lambda x: x["arr"])
        pior_arr = min(meses_ordenados, key=lambda x: x["arr"])
        insights_cards.append({
            'tone': 'positive',
            'title': 'Melhor mês de arrecadação',
            'text': f"{melhor_arr['label']} com {_fmt_k(melhor_arr['arr'])}."
        })
        insights_cards.append({
            'tone': 'warning',
            'title': 'Menor mês de arrecadação',
            'text': f"{pior_arr['label']} com {_fmt_k(pior_arr['arr'])}."
        })
    if var_m12_arr is not None:
        insights_cards.append({
            'tone': 'info' if var_m12_arr >= 0 else 'warning',
            'title': 'Variação versus o mesmo mês do ano anterior',
            'text': f"A arrecadação variou {var_m12_arr:+.1f}% em relação ao mesmo mês do ano anterior."
        })
    if tem_meta and ating_meta is not None:
        insights_cards.append({
            'tone': 'positive' if ating_meta >= 100 else 'warning',
            'title': 'Atingimento da meta no período',
            'text': f"{ating_meta:.1f}% ({_fmt_k(total_arr)} de {_fmt_k(total_meta)})."
        })

    # ── Tabela Mensal com drilldown diário ──
    tabela_mensal = ""
    if meses_ordenados:
        fat_acum = 0.0
        arr_acum = 0.0
        for m in meses_ordenados:
            fm = m["fat"]
            am = m["arr"]
            fat_acum += fm
            arr_acum += am
            dif_fat_pct = ((fm - am) / fm) if fm > 0 else 0.0
            dif_arr_pct = ((am - fm) / am) if am > 0 else 0.0

            tabela_mensal += f"""
            <tr class="month-row" data-month="{m['key']}">
              <td class="month-cell" data-month="{m['key']}" style="text-align: center; font-weight: 700;">
                    <button class="month-toggle" type="button" data-month="{m['key']}">▸</button>
                    <span>{m['label']}</span>
                </td>
                <td style="text-align: center;">{_fmt_k(fm)}</td>
                <td style="text-align: center; background: #F5F8FF; font-weight: 700;">{_fmt_k(fat_acum)}</td>
                <td style="text-align: center; font-weight: 700; color: {'#C03A2B' if dif_fat_pct >= 0 else '#0B8A6F'};">{dif_fat_pct:+.1%}</td>
                <td style="text-align: center;">{_fmt_k(am)}</td>
                <td style="text-align: center; background: #F5F8FF; font-weight: 700;">{_fmt_k(arr_acum)}</td>
                <td style="text-align: center; font-weight: 700; color: {'#0B8A6F' if dif_arr_pct >= 0 else '#C03A2B'};">{dif_arr_pct:+.1%}</td>
            </tr>
            """

            def _resumo_mes_ate_dia(year_ref, month_ref, ultimo_dia):
                ref_mes = idx_meses.get((year_ref, month_ref))
                if not ref_mes:
                    return None
                ref_dias = [x for x in ref_mes["dias"] if x["dt"].day <= ultimo_dia]
                if not ref_dias:
                    return None
                ref_fat = sum(x["fd"] for x in ref_dias)
                ref_arr = sum(x["ad"] for x in ref_dias)
                ref_dif_fat = ((ref_fat - ref_arr) / ref_fat) if ref_fat > 0 else 0.0
                ref_dif_arr = ((ref_arr - ref_fat) / ref_arr) if ref_arr > 0 else 0.0
                return {
                    "scope": f"ate {ultimo_dia:02d}/{month_ref:02d}/{year_ref:04d}",
                    "fat": ref_fat,
                    "arr": ref_arr,
                    "dif_fat": ref_dif_fat,
                    "dif_arr": ref_dif_arr,
                }

            _ano_mes = int(m["key"][:4])
            _mes_num = int(m["key"][5:7])
            _ultimo_dia = max((x["dt"].day for x in m["dias"]), default=0)
            _ref_prev = _resumo_mes_ate_dia(_ano_mes - 1, _mes_num, _ultimo_dia)
            _ref_next = _resumo_mes_ate_dia(_ano_mes + 1, _mes_num, _ultimo_dia)

            if _ref_prev:
                tabela_mensal += f"""
                <tr class="day-row day-{m['key']}" data-month="{m['key']}" style="display:none; background:#F8FBFF;">
                    <td style="text-align:left; padding-left: 60px;">
                        <div style="font-size:0.78rem; font-weight:700; color:#5C6E82; text-transform:uppercase; letter-spacing:.05em;">Mesmo periodo do ano passado</div>
                        <div>{_ref_prev['scope']}</div>
                    </td>
                    <td style="text-align: center; font-weight:600;">{_fmt_k(_ref_prev['fat'])}</td>
                    <td style="text-align: center; font-weight:700;">{_fmt_k(_ref_prev['fat'])}</td>
                    <td style="text-align: center; font-weight:700; color: {'#C03A2B' if _ref_prev['dif_fat'] >= 0 else '#0B8A6F'};">{_ref_prev['dif_fat']:+.1%}</td>
                    <td style="text-align: center; font-weight:600;">{_fmt_k(_ref_prev['arr'])}</td>
                    <td style="text-align: center; font-weight:700;">{_fmt_k(_ref_prev['arr'])}</td>
                    <td style="text-align: center; font-weight:700; color: {'#0B8A6F' if _ref_prev['dif_arr'] >= 0 else '#C03A2B'};">{_ref_prev['dif_arr']:+.1%}</td>
                </tr>
                """

            if _ref_next:
                tabela_mensal += f"""
                <tr class="day-row day-{m['key']}" data-month="{m['key']}" style="display:none; background:#F8FBFF;">
                    <td style="text-align:left; padding-left: 60px;">
                        <div style="font-size:0.78rem; font-weight:700; color:#5C6E82; text-transform:uppercase; letter-spacing:.05em;">Mesmo periodo do ano seguinte</div>
                        <div>{_ref_next['scope']}</div>
                    </td>
                    <td style="text-align: center; font-weight:600;">{_fmt_k(_ref_next['fat'])}</td>
                    <td style="text-align: center; font-weight:700;">{_fmt_k(_ref_next['fat'])}</td>
                    <td style="text-align: center; font-weight:700; color: {'#C03A2B' if _ref_next['dif_fat'] >= 0 else '#0B8A6F'};">{_ref_next['dif_fat']:+.1%}</td>
                    <td style="text-align: center; font-weight:600;">{_fmt_k(_ref_next['arr'])}</td>
                    <td style="text-align: center; font-weight:700;">{_fmt_k(_ref_next['arr'])}</td>
                    <td style="text-align: center; font-weight:700; color: {'#0B8A6F' if _ref_next['dif_arr'] >= 0 else '#C03A2B'};">{_ref_next['dif_arr']:+.1%}</td>
                </tr>
                """

            fat_mes_ac = 0.0
            arr_mes_ac = 0.0
            for d in sorted(m["dias"], key=lambda x: x["dt"]):
                fat_mes_ac += d["fd"]
                arr_mes_ac += d["ad"]
                dif_fat_dia = ((d["fd"] - d["ad"]) / d["fd"]) if d["fd"] > 0 else 0.0
                dif_arr_dia = ((d["ad"] - d["fd"]) / d["ad"]) if d["ad"] > 0 else 0.0
                tabela_mensal += f"""
                <tr class="day-row day-{m['key']}" data-month="{m['key']}" style="display:none;">
                    <td style="text-align: center; padding-left: 34px; color: #5C6E82;">{d['dt'].strftime('%d/%m/%Y')}</td>
                    <td style="text-align: center;">{_fmt_k(d['fd'])}</td>
                    <td style="text-align: center; font-weight: 600;">{_fmt_k(fat_mes_ac)}</td>
                    <td style="text-align: center; font-weight: 600; color: {'#C03A2B' if dif_fat_dia >= 0 else '#0B8A6F'};">{dif_fat_dia:+.1%}</td>
                    <td style="text-align: center;">{_fmt_k(d['ad'])}</td>
                    <td style="text-align: center; font-weight: 600;">{_fmt_k(arr_mes_ac)}</td>
                    <td style="text-align: center; font-weight: 600; color: {'#0B8A6F' if dif_arr_dia >= 0 else '#C03A2B'};">{dif_arr_dia:+.1%}</td>
                </tr>
                """
    else:
        tabela_mensal = "<tr><td colspan='7' style='text-align:center;'>Nenhum dado mensal encontrado no período.</td></tr>"

    # ── Tabela 8091 (água/esgoto/leitura/economias) ──
    tabela_8091 = ""
    serie_91 = d91.get('serie_meses', []) if isinstance(d91, dict) else []
    if isinstance(serie_91, list) and serie_91:
        def _mes_key(lbl):
            try:
                mm, yyyy = str(lbl).split('/')
                return int(yyyy), int(mm)
            except Exception:
                return (0, 0)

        serie_91_ord = sorted(serie_91, key=lambda s: _mes_key(s.get('periodo', '')), reverse=True)

        # ── Índice leit_dia por mês: média dos dias >= 26 com qtd > 0 ──────────
        from datetime import datetime as _dt
        _leit_dia_91 = d91.get('leit_dia', []) if isinstance(d91, dict) else []
        _leit_dia_por_mes = {}  # chave: "MM/YYYY" → list de números de dias >= 26
        for _ld in _leit_dia_91:
            _data_str = str(_ld.get('data', ''))
            try:
                _d = _dt.strptime(_data_str, "%d/%m/%Y")
                if _d.day >= 26 and _ld.get('qtd', 0) > 0:
                    _chave_mes = f"{_d.month:02d}/{_d.year:04d}"
                    _leit_dia_por_mes.setdefault(_chave_mes, []).append(_d.day)
            except Exception:
                pass

        fat_agua_acum = 0.0
        fat_esgoto_acum = 0.0
        for s in serie_91_ord:
            fat_agua = float(s.get('agua', 0) or 0)
            vol_fat_agua = float(s.get('vol_fat_agua', 0) or 0)
            fat_esgoto = float(s.get('esgoto', 0) or 0)
            # D. Lidos: média dos dias >= 26 com leitura; fallback ao valor original
            _periodo_s = s.get('periodo', '')
            _dias_26 = _leit_dia_por_mes.get(_periodo_s, [])
            if _dias_26:
                dias_lidos = round(sum(_dias_26) / len(_dias_26), 1)
            else:
                dias_lidos = int(float(s.get('dias_lidos', 0) or 0))
            eco_agua_fat = float(s.get('eco_agua_faturadas', 0) or 0)
            eco_esc_fat = float(s.get('eco_esgoto_faturadas', 0) or 0)
            fat_agua_acum += fat_agua
            fat_esgoto_acum += fat_esgoto

            tabela_8091 += f"""
            <tr>
              <td style="text-align:center;font-weight:700;">{s.get('periodo', '-')}</td>
              <td style="text-align:center;">{_fmt_k(fat_agua)}</td>
              <td style="text-align:center;font-weight:700;">{_fmt_k(fat_agua_acum)}</td>
              <td style="text-align:center;">{vol_fat_agua:,.1f} m³</td>
              <td style="text-align:center;">{int(round(eco_agua_fat)):,}</td>
              <td style="text-align:center;">{_fmt_k(fat_esgoto)}</td>
              <td style="text-align:center;font-weight:700;">{_fmt_k(fat_esgoto_acum)}</td>
              <td style="text-align:center;">{int(round(eco_esc_fat)):,}</td>
              <td style="text-align:center;">{dias_lidos:.1f}</td>
            </tr>
            """
    else:
      tabela_8091 = "<tr><td colspan='9' style='text-align:center;'>Macro 8091 não disponível neste PN.</td></tr>"

    # ── Tabela Formas consolidada ──
    tabela_formas = ""
    if isinstance(formas, dict) and formas:
        formas_sorted = sorted(
            [(k, v if isinstance(v, dict) else {}) for k, v in formas.items()],
            key=lambda x: float(x[1].get('valor', 0) or 0),
            reverse=True,
        )
        total_v_formas = sum(float(v.get('valor', 0) or 0) for _, v in formas_sorted)
        total_q_formas = sum(float(v.get('qtde', 0) or 0) for _, v in formas_sorted)

        for nome, val in formas_sorted:
            valor = float(val.get('valor', 0) or 0)
            qtde = float(val.get('qtde', 0) or 0)
            pct_v = (valor / total_v_formas * 100) if total_v_formas > 0 else 0
            tabela_formas += f"""
            <tr>
                <td style="text-align: center;">{nome}</td>
                <td style="text-align: center;">{int(qtde):,}</td>
                <td style="text-align: center; font-weight: 600;">{_fmt_k(valor)}</td>
                <td style="text-align: center;">{pct_v:.1f}%</td>
            </tr>
            """

        tabela_formas += f"""
            <tr class="total-row">
                <td style="text-align: center;">TOTAL</td>
                <td style="text-align: center;">{int(total_q_formas):,}</td>
                <td style="text-align: center;">{_fmt_k(total_v_formas)}</td>
                <td style="text-align: center;">100.0%</td>
            </tr>
        """
    else:
        tabela_formas = "<tr><td colspan='4' style='text-align:center;'>Dados da Macro 8117 não fornecidos.</td></tr>"

    # ── Matriz mensal por forma de pagamento ──
    formas_mensais_html = "<tr><td colspan='2' style='text-align:center;'>Dados por mês/forma não disponíveis na macro 8117 processada.</td></tr>"
    formas_mes_raw = d17.get("formas_mensais") or d17.get("formas_por_mes") or {}
    matrix = {}
    if isinstance(formas_mes_raw, dict) and formas_mes_raw:
        def _eh_mes(lbl):
            s = str(lbl).strip()
            return bool(re.match(r"^(0[1-9]|1[0-2])/\d{4}$", s))

        # Caso A: mes -> {forma -> valor|{valor,qtde}}
        if any(isinstance(v, dict) and _eh_mes(k) for k, v in formas_mes_raw.items()):
            for mes_lbl, bloco in formas_mes_raw.items():
                if not isinstance(bloco, dict):
                    continue
                matrix[str(mes_lbl)] = {}
                for forma, val in bloco.items():
                    if isinstance(val, dict):
                        matrix[str(mes_lbl)][str(forma)] = float(val.get('valor', 0) or 0)
                    else:
                        try:
                            matrix[str(mes_lbl)][str(forma)] = float(val or 0)
                        except Exception:
                            matrix[str(mes_lbl)][str(forma)] = 0.0
        else:
            # Caso B: forma -> {mes -> valor}
            for forma, bloco in formas_mes_raw.items():
                if not isinstance(bloco, dict):
                    continue
                for mes_lbl, val in bloco.items():
                    matrix.setdefault(str(mes_lbl), {})
                    if isinstance(val, dict):
                        matrix[str(mes_lbl)][str(forma)] = float(val.get('valor', 0) or 0)
                    else:
                        try:
                            matrix[str(mes_lbl)][str(forma)] = float(val or 0)
                        except Exception:
                            matrix[str(mes_lbl)][str(forma)] = 0.0

    if matrix:
        formas_cols = {}
        for _, bloc in matrix.items():
            for f, vv in bloc.items():
                formas_cols[f] = formas_cols.get(f, 0.0) + float(vv or 0.0)
        formas_ordenadas = [f for f, _ in sorted(formas_cols.items(), key=lambda x: x[1], reverse=True)][:8]

        def _key_mes(lbl):
            try:
                mm, yyyy = str(lbl).split('/')
                return int(yyyy), int(mm)
            except Exception:
                return (0, 0)

        head = ''.join([f"<th>{f}</th>" for f in formas_ordenadas])
        body = ""
        for mes_lbl in sorted(matrix.keys(), key=_key_mes):
            linha_vals = []
            total_l = 0.0
            for f in formas_ordenadas:
                v = float(matrix.get(mes_lbl, {}).get(f, 0.0) or 0.0)
                total_l += v
                linha_vals.append(f"<td style='text-align:center'>{_fmt_k(v)}</td>")
            body += f"<tr><td style='font-weight:700;text-align:center'>{mes_lbl}</td>{''.join(linha_vals)}<td style='text-align:center;font-weight:700'>{_fmt_k(total_l)}</td></tr>"

        formas_mensais_html = f"""
        <tr>
            <td colspan="2" style="padding:0;border:none;">
                <div class="table-wrap">
                    <table id="pn-formas-mes" style="min-width: 940px;">
                        <thead>
                            <tr><th>Mês/Ano</th>{head}<th>Total</th></tr>
                        </thead>
                        <tbody>{body}</tbody>
                    </table>
                </div>
            </td>
        </tr>
        """

    status_real = "Alta" if pct_realizado >= 90 else "Atenção" if pct_realizado >= 80 else "Crítica"
    status_class = "ok" if pct_realizado >= 90 else "warn" if pct_realizado >= 80 else "bad"

    m1_txt = f"{var_m1_arr:+.1f}%" if var_m1_arr is not None else "N/D"
    m12_txt = f"{var_m12_arr:+.1f}%" if var_m12_arr is not None else "N/D"
    meta_txt = _fmt_k(total_meta) if tem_meta else "N/D"
    desvio_txt = _fmt_k(desvio_meta) if desvio_meta is not None else "N/D"
    ating_meta_txt = f"{ating_meta:.1f}%" if ating_meta is not None else "N/D"
    qualidade_html = ''.join([f"<li>{x}</li>" for x in alertas]) if alertas else "<li>Nenhuma inconsistência relevante detectada.</li>"
    insights_html = _insights_cards_html(insights_cards)

    ano_atual_ref = int(meses_ordenados[0]["key"][:4]) if meses_ordenados else now.year
    meses_ano_atual = [m for m in meses_ordenados if int(m["key"][:4]) == ano_atual_ref]
    meses_ano_anterior = [m for m in meses_ordenados if int(m["key"][:4]) == (ano_atual_ref - 1)]
    total_fat_ano = sum(m["fat"] for m in meses_ano_atual)
    total_arr_ano = sum(m["arr"] for m in meses_ano_atual)
    gap_ano = total_fat_ano - total_arr_ano
    conv_ano = (total_arr_ano / total_fat_ano * 100) if total_fat_ano > 0 else 0.0

    ano_cmp_cards = []
    ano_destaque_cards = []
    if meses_ano_atual:
        melhor_fat_ano = max(meses_ano_atual, key=lambda x: x["fat"])
        melhor_arr_ano = max(meses_ano_atual, key=lambda x: x["arr"])
        maior_gap_ano = max(meses_ano_atual, key=lambda x: (x["fat"] - x["arr"]))
        ano_destaque_cards = [
            {'tone': 'info', 'title': f'Faturamento {ano_atual_ref}', 'text': _fmt_k(total_fat_ano)},
            {'tone': 'positive', 'title': f'Arrecadacao {ano_atual_ref}', 'text': _fmt_k(total_arr_ano)},
            {'tone': 'info', 'title': 'Maior faturamento', 'text': f"{melhor_fat_ano['label']} | {_fmt_k(melhor_fat_ano['fat'])}"},
            {'tone': 'positive', 'title': 'Maior arrecadacao', 'text': f"{melhor_arr_ano['label']} | {_fmt_k(melhor_arr_ano['arr'])}"},
            {'tone': 'positive' if conv_ano >= 90 else 'warning', 'title': 'Conversao em caixa', 'text': f'{conv_ano:.1f}% no acumulado'},
            {'tone': 'warning' if gap_ano > 0 else 'neutral', 'title': 'Maior gap do ano', 'text': f"{maior_gap_ano['label']} | {_fmt_k(maior_gap_ano['fat'] - maior_gap_ano['arr'])}"},
        ]

    if meses_ano_atual and meses_ano_anterior:
        mapa_atual = {int(m["key"][5:7]): m for m in meses_ano_atual}
        mapa_anterior = {int(m["key"][5:7]): m for m in meses_ano_anterior}
        meses_comuns = sorted(set(mapa_atual.keys()) & set(mapa_anterior.keys()))
        if meses_comuns:
            fat_atual_cmp = sum(mapa_atual[mes]["fat"] for mes in meses_comuns)
            fat_ant_cmp = sum(mapa_anterior[mes]["fat"] for mes in meses_comuns)
            arr_atual_cmp = sum(mapa_atual[mes]["arr"] for mes in meses_comuns)
            arr_ant_cmp = sum(mapa_anterior[mes]["arr"] for mes in meses_comuns)
            var_fat_ano = _pct_var(fat_atual_cmp, fat_ant_cmp)
            var_arr_ano = _pct_var(arr_atual_cmp, arr_ant_cmp)
            if var_fat_ano is not None:
                ano_cmp_cards.append({'tone': 'positive' if var_fat_ano >= 0 else 'warning', 'title': f'Vs {ano_atual_ref - 1} no faturamento', 'text': f'{var_fat_ano:+.1f}% no mesmo recorte'})
            if var_arr_ano is not None:
                ano_cmp_cards.append({'tone': 'positive' if var_arr_ano >= 0 else 'warning', 'title': f'Vs {ano_atual_ref - 1} na arrecadacao', 'text': f'{var_arr_ano:+.1f}% no mesmo recorte'})

    ano_resumo_cards = []
    if len(ano_destaque_cards) >= 2:
        ano_resumo_cards.extend(ano_destaque_cards[:2])
    if ano_cmp_cards:
        ano_resumo_cards.extend(ano_cmp_cards[:2])
    if len(ano_destaque_cards) >= 4:
        ano_resumo_cards.extend(ano_destaque_cards[2:4])
    if len(ano_destaque_cards) >= 6:
        ano_resumo_cards.extend(ano_destaque_cards[4:6])

    verao_temporadas = []
    anos_base_disponiveis = sorted({ano for ano, _mes in idx_meses.keys()})
    anos_inicio_temporada = sorted(
        {
            ano for ano in anos_base_disponiveis
            if (ano, 12) in idx_meses or (ano + 1, 1) in idx_meses or (ano + 1, 2) in idx_meses
        },
        reverse=True,
    )
    for ano_inicio in anos_inicio_temporada:
        refs_temporada = [(ano_inicio, 12), (ano_inicio + 1, 1), (ano_inicio + 1, 2)]
        meses_temporada = [idx_meses[k] for k in refs_temporada if k in idx_meses]
        if not meses_temporada:
            continue

        ano_corrente_temporada = ano_inicio + 1
        total_fat_ano_base = sum(m["fat"] for m in meses_ordenados if int(m["key"][:4]) == ano_corrente_temporada)
        total_arr_ano_base = sum(m["arr"] for m in meses_ordenados if int(m["key"][:4]) == ano_corrente_temporada)
        fat_temporada = sum(m["fat"] for m in meses_temporada)
        arr_temporada = sum(m["arr"] for m in meses_temporada)
        gap_temporada = fat_temporada - arr_temporada
        conv_temporada = (arr_temporada / fat_temporada * 100) if fat_temporada > 0 else 0.0
        peso_fat_temporada = (fat_temporada / total_fat_ano_base * 100) if total_fat_ano_base > 0 else 0.0
        peso_arr_temporada = (arr_temporada / total_arr_ano_base * 100) if total_arr_ano_base > 0 else 0.0
        rotulo_periodo = " + ".join(m["label"] for m in meses_temporada)

        cards_temporada = [
            {'tone': 'info', 'title': 'Faturamento do verao', 'text': f"{_fmt_k(fat_temporada)} | {rotulo_periodo}"},
            {'tone': 'positive', 'title': 'Arrecadacao do verao', 'text': f"{_fmt_k(arr_temporada)} | {rotulo_periodo}"},
        ]

        refs_temporada_anterior = [(ano_inicio - 1, 12), (ano_inicio, 1), (ano_inicio, 2)]
        meses_temporada_anterior = [idx_meses[k] for k in refs_temporada_anterior if k in idx_meses]
        if meses_temporada_anterior:
            fat_ant = sum(m["fat"] for m in meses_temporada_anterior)
            arr_ant = sum(m["arr"] for m in meses_temporada_anterior)
            var_fat_temporada = _pct_var(fat_temporada, fat_ant)
            var_arr_temporada = _pct_var(arr_temporada, arr_ant)
            if var_fat_temporada is not None:
                cards_temporada.append({'tone': 'positive' if var_fat_temporada >= 0 else 'warning', 'title': 'Vs verao anterior no faturamento', 'text': f'{var_fat_temporada:+.1f}%'})
            if var_arr_temporada is not None:
                cards_temporada.append({'tone': 'positive' if var_arr_temporada >= 0 else 'warning', 'title': 'Vs verao anterior na arrecadacao', 'text': f'{var_arr_temporada:+.1f}%'})

        cards_temporada.extend([
            {'tone': 'info', 'title': 'Peso no ano atual', 'text': f'{peso_fat_temporada:.1f}% do faturamento de {ano_corrente_temporada}'},
            {'tone': 'positive', 'title': 'Peso na arrecadacao do ano', 'text': f'{peso_arr_temporada:.1f}% da arrecadacao de {ano_corrente_temporada}'},
            {'tone': 'positive' if conv_temporada >= 90 else 'warning', 'title': 'Conversao da temporada', 'text': f'{conv_temporada:.1f}% em caixa'},
            {'tone': 'warning' if gap_temporada > 0 else 'neutral', 'title': 'Gap do verao', 'text': _fmt_k(gap_temporada)},
        ])

        verao_temporadas.append({
            "id": f"verao-{ano_inicio}-{ano_inicio + 1}",
            "label": f"{ano_inicio}/{ano_inicio + 1}",
            "html": _insights_cards_html(cards_temporada),
        })

    if verao_temporadas:
        verao_tabs_html = ''.join(
            f'<button class="season-btn{" active" if i == 0 else ""}" type="button" data-season-target="{item["id"]}">{item["label"]}</button>'
            for i, item in enumerate(verao_temporadas)
        )
        verao_paineis_html = ''.join(
            f'<div class="season-panel{" active" if i == 0 else ""}" id="{item["id"]}">{item["html"]}</div>'
            for i, item in enumerate(verao_temporadas)
        )
    else:
        verao_tabs_html = ''
        verao_paineis_html = _insights_cards_html([])

    ano_resumo_html = _insights_cards_html(ano_resumo_cards)

    html = f"""<!DOCTYPE html>
  <html lang="pt-BR">
  <head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PN - Painel de Negócios | {cidade}</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@500;700;800&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');

    :root {{
      --navy: #06203D;
      --teal: #0D5C63;
      --gold: #D98E04;
      --mint: #0B8A6F;
      --danger: #C03A2B;
      --card: #FFFFFF;
      --ink: #12263A;
      --muted: #5C6E82;
      --line: #D7E0EC;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      color: var(--ink);
      font-family: 'IBM Plex Sans', 'Segoe UI', sans-serif;
      background:
        radial-gradient(1000px 500px at -10% -20%, #DCEBFF 0%, transparent 65%),
        radial-gradient(900px 450px at 110% -10%, #D9F7EE 0%, transparent 60%),
        linear-gradient(180deg, #F8FAFD 0%, #F2F6FB 100%);
    }}
    .container {{ max-width: 1220px; margin: 0 auto; padding: 24px; }}
    .hero {{
      position: relative; overflow: hidden; border-radius: 18px;
      border: 1px solid rgba(255,255,255,0.35);
      background: linear-gradient(120deg, var(--navy) 0%, #0D325A 55%, #125173 100%);
      color: #fff; padding: 26px 24px; box-shadow: 0 16px 40px rgba(6,32,61,0.22);
    }}
    .hero:after {{
      content: ""; position: absolute; width: 340px; height: 340px;
      right: -100px; top: -140px; border-radius: 50%; background: rgba(255,255,255,0.08);
    }}
    .hero-top {{ display:flex; flex-wrap:wrap; gap:16px; justify-content:space-between; align-items:flex-start; position:relative; z-index:1; }}
    .brand-row {{ display:flex; align-items:center; gap:12px; }}
    .brand-logo {{ width:64px; height:64px; object-fit:contain; border-radius:12px; background:rgba(255,255,255,0.14); padding:8px; border:1px solid rgba(255,255,255,0.2); }}
    .brand-logo-fallback {{ width:64px; height:64px; border-radius:12px; display:grid; place-items:center; font-family:'Montserrat',sans-serif; font-size:24px; font-weight:800; color:#fff; background:linear-gradient(145deg, rgba(255,255,255,0.22), rgba(255,255,255,0.08)); border:1px solid rgba(255,255,255,0.22); }}
    .hero h1 {{ margin:0; font-family:'Montserrat',sans-serif; font-size:30px; line-height:1.1; letter-spacing:0.3px; text-transform:uppercase; }}
    .hero p {{ margin:8px 0 0; color:rgba(255,255,255,0.86); font-size:14px; }}
    .meta {{ text-align:right; min-width:260px; font-size:13px; color:rgba(255,255,255,0.92); line-height:1.5; }}
    .actions {{ margin-top:18px; display:flex; gap:10px; flex-wrap:wrap; position:relative; z-index:1; }}
    .btn {{ border:0; border-radius:10px; padding:10px 14px; font-size:12px; font-weight:700; letter-spacing:0.2px; cursor:pointer; color:#fff; background:rgba(255,255,255,0.16); }}
    .btn:hover {{ background: rgba(255,255,255,0.24); }}
    .status-band {{ margin-top:14px; display:flex; align-items:center; gap:8px; position:relative; z-index:1; }}
    .chip {{ border-radius:999px; padding:5px 11px; font-size:11px; font-weight:800; text-transform:uppercase; letter-spacing:0.35px; }}
    .chip.ok {{ background:#D5F5E3; color:#146C43; }}
    .chip.warn {{ background:#FDEBD0; color:#9A6700; }}
    .chip.bad {{ background:#FADBD8; color:#A93226; }}
    .status-band small {{ color:rgba(255,255,255,0.86); font-size:12px; }}

    .kpi-grid {{ display:grid; grid-template-columns: repeat(4, minmax(180px, 1fr)); gap:12px; margin-top:18px; }}
    .kpi-group-title {{
      grid-column: 1 / -1;
      font-family:'Montserrat',sans-serif;
      font-size:12px;
      text-transform:uppercase;
      letter-spacing:0.45px;
      color:#24415f;
      background: linear-gradient(180deg, #F0F5FC 0%, #E7EFF9 100%);
      border:1px solid var(--line);
      border-radius:10px;
      padding:10px 12px;
      font-weight:700;
      box-shadow: inset 0 1px 0 rgba(255,255,255,0.6);
    }}
    .kpi {{
      background:linear-gradient(180deg, #FFFFFF 0%, #F8FBFF 100%);
      border:1px solid var(--line);
      border-radius:14px;
      padding:14px;
      box-shadow:0 8px 20px rgba(11,22,41,0.05);
      position:relative;
      overflow:hidden;
    }}
    .kpi:before {{
      content:"";
      position:absolute;
      left:0;
      top:0;
      width:100%;
      height:4px;
      background:linear-gradient(90deg, #0E3159 0%, #138A72 100%);
      opacity:0.95;
    }}
    .kpi .label {{
      font-size:11px;
      text-transform:uppercase;
      letter-spacing:0.45px;
      color:var(--muted);
      margin-bottom:10px;
      font-weight:700;
      display:flex;
      align-items:center;
      gap:6px;
    }}
    .kpi .label .ico {{
      width:20px;
      height:20px;
      border-radius:6px;
      display:inline-grid;
      place-items:center;
      font-size:11px;
      font-weight:800;
      color:#11314f;
      background:#E5F0FF;
      border:1px solid #C8DAF0;
    }}
    .kpi-help {{
      width:16px;
      height:16px;
      border-radius:50%;
      display:inline-grid;
      place-items:center;
      font-size:10px;
      font-weight:800;
      line-height:1;
      cursor:help;
      color:#0A2748;
      background:#EAF2FF;
      border:1px solid #C8DAF0;
      margin-left:2px;
    }}
    .kpi .value {{ font-family:'Montserrat',sans-serif; font-size:26px; line-height:1.1; font-weight:800; color:var(--navy); }}
    .kpi .hint {{ margin-top:8px; font-size:12px; color:var(--muted); border-top:1px dashed #D9E3F1; padding-top:6px; }}
    .value.good {{ color: var(--mint); }}
    .value.bad {{ color: var(--danger); }}
    .value.neutral {{ color: var(--gold); }}

    .row-grid {{ display:grid; grid-template-columns: 2fr 1fr; gap:12px; margin-top:12px; align-items:start; }}
    .panel {{ background:var(--card); border:1px solid var(--line); border-radius:14px; padding:16px; box-shadow:0 8px 20px rgba(11,22,41,0.05); }}
    .panel h3 {{ margin:0; font-family:'Montserrat',sans-serif; font-size:14px; color:var(--navy); }}
    .panel ul {{ margin:0; padding-left:18px; color:#334155; line-height:1.5; }}
    .panel-header-row {{ display:flex; align-items:center; gap:10px; margin-bottom:14px; }}
    .panel-icon {{ width:32px; height:32px; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; font-size:15px; flex-shrink:0; }}
    .insights-cards {{ display:grid; grid-template-columns: 1fr 1fr; gap:10px; }}
    .season-tabs {{ display:flex; gap:8px; flex-wrap:wrap; margin:-2px 0 14px; }}
    .season-btn {{
      border:1px solid #C9D7EA; background:#F5F8FE; color:#1A3D66; border-radius:999px;
      padding:7px 12px; font-size:12px; font-weight:800; cursor:pointer; transition:all .18s ease;
    }}
    .season-btn:hover {{ background:#EAF1FB; border-color:#AFC5E3; }}
    .season-btn.active {{ background:#0E3159; color:#fff; border-color:#0E3159; }}
    .season-panel {{ display:none; }}
    .season-panel.active {{ display:block; }}
    .insight-card {{ border:1px solid var(--line); border-left:4px solid #93A7C3; border-radius:10px; padding:10px 11px; background:#F8FAFE; }}
    .insight-title {{ font-size:12px; font-weight:800; color:#153354; margin-bottom:4px; text-transform:uppercase; letter-spacing:0.3px; }}
    .insight-text {{ font-size:13px; color:#31465E; line-height:1.45; }}
    .insight-card.tone-positive {{ background:#F1FBF6; border-left-color:#1E8E5A; }}
    .insight-card.tone-positive .insight-title {{ color:#0D6A41; }}
    .insight-card.tone-warning {{ background:#FFF7EE; border-left-color:#D07A00; }}
    .insight-card.tone-warning .insight-title {{ color:#955700; }}
    .insight-card.tone-info {{ background:#EFF5FF; border-left-color:#2C6AC5; }}
    .insight-card.tone-info .insight-title {{ color:#1E4F97; }}
    .insight-card.tone-neutral {{ background:#F5F7FB; border-left-color:#8193AA; }}
    .insight-card.tone-neutral .insight-title {{ color:#3D5674; }}
    .quality-panel {{ display:flex; flex-direction:column; gap:0; }}
    .quality-metrics {{ display:grid; grid-template-columns:1fr 1fr; gap:8px; margin-bottom:12px; }}
    .quality-metric {{
      background: linear-gradient(135deg,#F8FAFE 0%,#F0F5FF 100%);
      border:1px solid var(--line);
      border-radius:10px;
      padding:10px 12px;
      display:flex;
      flex-direction:column;
      gap:3px;
    }}
    .qm-label {{ font-size:10.5px; text-transform:uppercase; letter-spacing:0.4px; font-weight:700; color:var(--muted); }}
    .qm-value {{ font-family:'Montserrat',sans-serif; font-size:15px; font-weight:800; color:var(--navy); }}
    .qm-good {{ color:var(--mint); }}
    .qm-warn {{ color:var(--gold); }}
    .qm-bad  {{ color:var(--danger); }}
    .quality-alerts {{ display:flex; flex-direction:column; gap:6px; }}
    .quality-alert-item {{
      font-size:12px; color:#7A4A00; background:#FFF8EC;
      border:1px solid #FDDCA0; border-radius:8px; padding:8px 10px; line-height:1.4;
    }}
    .quality-ok {{
      font-size:12.5px; color:#0D6A41; background:#F0FBF5;
      border:1px solid #A7E3C2; border-radius:8px; padding:10px 12px; font-weight:600;
    }}

    .section {{ margin-top:20px; }}
    .section h2 {{ font-family:'Montserrat',sans-serif; font-size:16px; margin:0 0 12px; color:var(--navy); display:flex; align-items:center; gap:12px; font-weight:700; letter-spacing:0.3px; }}
    .section h2::before {{ content:""; width:12px; height:12px; border-radius:50%; background:var(--gold); box-shadow:0 0 0 4px rgba(217,142,4,0.16); flex-shrink:0; }}
    .section.monthly h2::before {{ background:#1565C0; box-shadow:0 0 0 4px rgba(21,101,192,0.16); content:""; width:20px; height:20px; display:flex; align-items:center; justify-content:center; border-radius:50%; font-size:12px; box-shadow:none; }}
    .section.water h2::before {{ background:#00838F; box-shadow:0 0 0 4px rgba(0,131,143,0.16); content:""; width:20px; height:20px; display:flex; align-items:center; justify-content:center; border-radius:50%; font-size:12px; box-shadow:none; }}
    .sub {{ font-size:12px; color:var(--muted); margin:-2px 0 10px; }}
    .table-card {{ background:var(--card); border:1px solid var(--line); border-radius:14px; overflow:hidden; box-shadow:0 8px 20px rgba(11,22,41,0.05); }}
    .table-wrap {{ width:100%; overflow-x:auto; }}
    table {{ width:100%; border-collapse:collapse; font-size:12px; }}
    th {{ background:linear-gradient(180deg,#0E3159 0%, #0A2748 100%); color:#fff; padding:8px 6px; text-align:center; border:1px solid rgba(255,255,255,0.2); font-size:10px; text-transform:uppercase; letter-spacing:0.35px; position:sticky; top:0; z-index:2; word-break: break-word; white-space:normal; transition:background 0.2s ease; }}
    th:hover {{ background:linear-gradient(180deg,#1a4166 0%, #154057 100%); text-shadow:0 0 8px rgba(255,255,255,0.3); }}
    .th-wrap {{ display:inline-flex; align-items:center; justify-content:center; gap:4px; flex-wrap:wrap; }}
    .help-icon {{
      width:14px;
      height:14px;
      border-radius:50%;
      display:inline-grid;
      place-items:center;
      font-size:9px;
      font-weight:800;
      line-height:1;
      cursor:help;
      color:#0A2748;
      background:rgba(255,255,255,0.92);
      border:1px solid rgba(255,255,255,0.45);
      box-shadow:0 1px 0 rgba(10,39,72,0.15);
      flex-shrink:0;
    }}
    td {{ border:1px solid var(--line); padding:7px 6px; text-align:center; vertical-align:middle; white-space:normal; word-break:break-word; background:#fff; }}
    tbody tr:nth-child(even) td {{ background:#F9FBFF; text-align:center; }}
    tbody tr:hover td {{ background:#FFF6E8; text-align:center; }}
    .total-row td {{ background:linear-gradient(180deg,#F3F8FF 0%, #EAF2FF 100%) !important; font-weight:800; text-align:center; }}
    .month-toggle {{ border:1px solid #C7D5E8; border-radius:6px; background:#fff; color:#123; font-weight:700; width:22px; height:22px; cursor:pointer; margin-right:6px; }}
    .month-row {{ cursor: pointer; text-align:center; }}
    .month-row td:first-child {{ user-select: none; text-align:center; }}
    .month-row:hover td:first-child {{ background:#F2F7FF; text-align:center; }}
    .day-row td {{ background:#FBFDFF !important; font-size:11px; text-align:center; }}
    #pn-tabela-8091 th, #pn-tabela-8091 td {{ text-align:center !important; }}
    .muted-mini {{ font-size:11px; color:#6A7B91; }}
    .footnote {{ margin-top:16px; font-size:11px; color:#6A7B91; line-height:1.5; text-align:center; }}

    @media (max-width: 1100px) {{
      .container {{ padding: 14px; }}
      .hero {{ padding: 18px 16px; }}
      .hero h1 {{ font-size: 24px; }}
      .meta {{ text-align: left; }}
      .kpi-grid {{ grid-template-columns: repeat(2, minmax(150px, 1fr)); }}
      .row-grid {{ grid-template-columns: 1fr; }}
      .insights-cards {{ grid-template-columns: 1fr; }}
      .quality-metrics {{ grid-template-columns: 1fr 1fr; }}
    }}
    @media (max-width: 600px) {{ .kpi-grid {{ grid-template-columns: 1fr; }} }}
    @media print {{
      body {{ background:#fff; }}
      .container {{ max-width:100%; padding:0; }}
      .hero {{ box-shadow:none; border-radius:0; }}
      .actions {{ display:none; }}
      .table-card, .panel {{ box-shadow:none; }}
      tr {{ break-inside: avoid; page-break-inside: avoid; }}
      thead {{ display: table-header-group; }}
    }}
  </style>
  </head>
  <body>
  <div class="container">
    <div class="hero">
      <div class="hero-top">
        <div>
          <div class="brand-row">
            {logo_html}
            <div>
              <h1>Painel de Negócios (PN)</h1>
              <p>Itapoá Saneamento | Unidade {cidade} | Consolidado mensal em milhares (R$ k)</p>
            </div>
          </div>
        </div>
        <div class="meta">
          Período real da base<br>
          <strong>{periodo['inicial']} a {periodo['final']}</strong><br>
          Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}
        </div>
      </div>

      <div class="actions">
        <button class="btn" onclick="window.print()">Imprimir / Salvar em PDF</button>
        <button class="btn" onclick="exportarPnCsv()">Exportar CSV</button>
        <button class="btn" onclick="exportarExcelHtml()">Exportar Excel</button>
      </div>

      <div class="status-band">
        <span class="chip {status_class}">{status_real}</span>
        <small>Performance consolidada: {pct_realizado:.1f}% de realização sobre o faturamento do período.</small>
      </div>
    </div>

    <div class="kpi-grid">
      <div class="kpi-group-title">Resumo Financeiro do Período</div>
      <div class="kpi"><div class="label"><span class="ico">R$</span>Total Faturado <span class="kpi-help" title="Soma do faturamento do período selecionado (macro 8121).">?</span></div><div class="value">{_fmt_k(total_fat)}</div><div class="hint">Período consolidado</div></div>
      <div class="kpi"><div class="label"><span class="ico">Cx</span>Total Arrecadado <span class="kpi-help" title="Soma da arrecadação efetivamente recebida no período (macro 8121/8117).">?</span></div><div class="value good">{_fmt_k(total_arr)}</div><div class="hint">Receita em caixa</div></div>
      <div class="kpi"><div class="label"><span class="ico">Df</span>Gap (Fat - Arr) <span class="kpi-help" title="Diferença entre faturamento e arrecadação: quanto maior, maior o valor ainda não convertido em caixa.">?</span></div><div class="value bad">{_fmt_k(gap)}</div><div class="hint">Risco de inadimplência</div></div>
      <div class="kpi"><div class="label"><span class="ico">%</span>Índice de Realização <span class="kpi-help" title="Percentual de conversão em caixa: (Arrecadado / Faturado) x 100 no período.">?</span></div><div class="value {'good' if pct_realizado >= 90 else 'bad' if pct_realizado < 80 else 'neutral'}">{pct_realizado:.1f}%</div><div class="hint">Conversão em caixa</div></div>
    </div>

    <div class="row-grid">
      <div class="panel">
        <div class="panel-header-row">
          <span class="panel-icon" style="background:#EAF4EE;color:#0B6E4A"></span>
          <h3>Insights Automáticos</h3>
        </div>
        {ano_resumo_html}
      </div>
      <div class="panel quality-panel">
        <div class="panel-header-row">
          <span class="panel-icon" style="background:#EEF2FF;color:#3730A3"></span>
          <h3>Diagnóstico do Período</h3>
        </div>
        <div class="season-tabs">{verao_tabs_html}</div>
        {verao_paineis_html}
        <div class="quality-metrics" style="display:none;">
          <div class="quality-metric">
            <span class="qm-label">Mês fechado</span>
            <span class="qm-value">{mes_fechado}</span>
          </div>
          <div class="quality-metric">
            <span class="qm-label">Em andamento</span>
            <span class="qm-value">{mes_em_andamento}</span>
          </div>
          <div class="quality-metric">
            <span class="qm-label">Meta atingida</span>
            <span class="qm-value {'qm-good' if ating_meta is not None and ating_meta >= 100 else 'qm-warn' if ating_meta is not None else ''}">{ating_meta_txt}</span>
          </div>
          <div class="quality-metric">
            <span class="qm-label">Inadimplência est.</span>
            <span class="qm-value {'qm-bad' if inad_est > 15 else 'qm-warn' if inad_est > 8 else 'qm-good'}">{inad_est:.1f}%</span>
          </div>
          <div class="quality-metric">
            <span class="qm-label">Recuperação / gap</span>
            <span class="qm-value">{recuperacao_gap:.1f}%</span>
          </div>
          <div class="quality-metric">
            <span class="qm-label">Ticket médio arr.</span>
            <span class="qm-value">{_fmt_brl(ticket_medio_arr)}</span>
          </div>
        </div>
        <div class="quality-alerts">
          {''.join(f'<div class="quality-alert-item"> {a}</div>' for a in alertas) if alertas else '<div class="quality-ok"> Nenhuma inconsistência relevante detectada.</div>'}
        </div>
      </div>
    </div>

    <section class="section monthly">
      <h2>Acompanhamento Mensal</h2>
      <p class="sub">Clique em um mês para abrir os dias daquele mês com os dados detalhados.</p>
      <div class="table-card">
        <div class="table-wrap">
          <table id="pn-tabela-mensal">
            <thead>
              <tr>
                <th rowspan="2"><span class="th-wrap">Mês/Ano <span class="help-icon" title="Competência mensal consolidada com opção de abrir os dias.">?</span></span></th>
                <th colspan="3">Faturamento (R$ k)</th>
                <th colspan="3">Arrecadação (R$ k)</th>
              </tr>
              <tr>
                <th><span class="th-wrap">PN <span class="help-icon" title="Faturamento do mês (competência mensal, R$ k).">?</span></span></th>
                <th><span class="th-wrap">REALIZADO <span class="help-icon" title="Faturamento acumulado realizado até o mês (R$ k).">?</span></span></th>
                <th><span class="th-wrap">Dif. FAT x ARR (%) <span class="help-icon" title="Fórmula: ((Faturamento - Arrecadação) / Faturamento) x 100. Valor positivo indica arrecadação abaixo do faturamento no mês.">?</span></span></th>
                <th><span class="th-wrap">PN <span class="help-icon" title="Arrecadação do mês (competência mensal, R$ k).">?</span></span></th>
                <th><span class="th-wrap">REALIZADO <span class="help-icon" title="Arrecadação acumulada realizada até o mês (R$ k).">?</span></span></th>
                <th><span class="th-wrap">Dif. ARR x FAT (%) <span class="help-icon" title="Fórmula: ((Arrecadação - Faturamento) / Arrecadação) x 100. Valor positivo indica arrecadação acima do faturamento no mês.">?</span></span></th>
              </tr>
            </thead>
            <tbody>
              {tabela_mensal}
            </tbody>
          </table>
        </div>
      </div>
    </section>

    <section class="section water">
      <h2>Detalhamento 8091  Faturamento de Água e Esgoto</h2>
      <p class="sub">Resumo mensal com faturamento, volume faturado de água, acumulados e economias faturadas em esgoto.</p>
      <div class="table-card">
        <div class="table-wrap">
          <table id="pn-tabela-8091">
            <thead>
              <tr>
                <th rowspan="2"><span class="th-wrap">Mês/Ano <span class="help-icon" title="Competência mensal da macro 8091.">?</span></span></th>
                <th colspan="4">Água</th>
                <th colspan="3">Esgoto</th>
                <th colspan="1">Outros</th>
              </tr>
              <tr>
                <th><span class="th-wrap">Fat. Água <span class="help-icon" title="Somatório de VL. AGUA no mês.">?</span></span></th>
                <th><span class="th-wrap">Fat. Água Ac. <span class="help-icon" title="Acumulado mensal de VL. AGUA no período.">?</span></span></th>
                <th><span class="th-wrap">Vol. Fat. Água <span class="help-icon" title="Somatório de VOL.FAT.AGUA no mês.">?</span></span></th>
                <th><span class="th-wrap">Econ. Fat. Água <span class="help-icon" title="Soma de ECO* nas linhas com VL. AGUA > 0.">?</span></span></th>
                <th><span class="th-wrap">Fat. Esgoto <span class="help-icon" title="Somatório de VL. ESGOTO no mês.">?</span></span></th>
                <th><span class="th-wrap">Fat. Esgoto Ac. <span class="help-icon" title="Acumulado mensal de VL. ESGOTO no período.">?</span></span></th>
                <th><span class="th-wrap">Econ. Fat. Esgoto <span class="help-icon" title="Soma de ECO* nas linhas com VL. ESGOTO > 0.">?</span></span></th>
                <th><span class="th-wrap">D. Lidos <span class="help-icon" title="Quantidade de dias distintos com leitura no mês (DATA LEITURA).">?</span></span></th>
              </tr>
            </thead>
            <tbody>
              {tabela_8091}
            </tbody>
          </table>
        </div>
      </div>
    </section>

    <div class="footnote">
      Relatório Gerencial PN - Mathools 1.0<br>
      Valores monetários representados em milhares (k).
    </div>
  </div>

  <script>
  function _csvEscape(value) {{
    return '"' + String(value ?? '').replace(/"/g, '""').trim() + '"';
  }}

  function _tableToCsv(table, onlyVisible=false) {{
    if (!table) return '';
    const rows = [...table.querySelectorAll('tr')].filter((row) => {{
      if (!onlyVisible) return true;
      return row.style.display !== 'none';
    }});
    return rows.map((row) => {{
      const cols = [...row.querySelectorAll('th,td')];
      return cols.map((col) => _csvEscape(col.innerText || '')).join(';');
    }}).join('\\n');
  }}

  function exportarPnCsv() {{
    const tabelaMensal = document.getElementById('pn-tabela-mensal');
    const tabela8091 = document.getElementById('pn-tabela-8091');
    if (!tabelaMensal || !tabela8091) {{
      alert('As duas tabelas do PN precisam estar disponíveis para exportação.');
      return;
    }}
    const linhas = [];
    linhas.push(_csvEscape('Painel de Negócios (PN)'));
    linhas.push(_csvEscape('Unidade') + ';' + _csvEscape('{cidade}'));
    linhas.push(_csvEscape('Período da base') + ';' + _csvEscape('{periodo['inicial']} a {periodo['final']}'));
    linhas.push(_csvEscape('Gerado em') + ';' + _csvEscape('{datetime.now().strftime('%d/%m/%Y %H:%M')}'));
    linhas.push('');
    linhas.push(_csvEscape('Acompanhamento Mensal'));
    linhas.push(_tableToCsv(tabelaMensal, true));
    linhas.push('');
    linhas.push(_csvEscape('Detalhamento 8091'));
    linhas.push(_tableToCsv(tabela8091, true));

    const blob = new Blob(["\ufeff" + linhas.join('\\n')], {{ type: 'text/csv;charset=utf-8;' }});
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'PN_Executivo_{periodo['final'].replace('/', '-')}.csv';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }}

  function exportarExcelHtml() {{
    const tabelaMensal = document.getElementById('pn-tabela-mensal');
    const tabela8091 = document.getElementById('pn-tabela-8091');

    if (!tabelaMensal || !tabela8091) {{
      alert('As duas tabelas do PN precisam estar disponíveis para exportação conjunta.');
      return;
    }}

    const metaHtml = `
      <table style="width:100%; margin-bottom:16px;">
        <tr><td style="font-weight:700; background:#EEF4FD;">Painel</td><td>Painel de Negócios (PN)</td></tr>
        <tr><td style="font-weight:700; background:#EEF4FD;">Unidade</td><td>{cidade}</td></tr>
        <tr><td style="font-weight:700; background:#EEF4FD;">Período da base</td><td>{periodo['inicial']} a {periodo['final']}</td></tr>
        <tr><td style="font-weight:700; background:#EEF4FD;">Gerado em</td><td>{datetime.now().strftime('%d/%m/%Y %H:%M')}</td></tr>
      </table>`;
    const partes = [];
    partes.push(metaHtml);
    partes.push('<h2>Acompanhamento Mensal</h2>' + tabelaMensal.outerHTML);
    partes.push('<h2>Detalhamento 8091</h2>' + tabela8091.outerHTML);

    const excelHtml = `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:x="urn:schemas-microsoft-com:office:excel"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="Content-Type" content="application/vnd.ms-excel; charset=UTF-8" />
  <style>
    body {{ font-family: Calibri, Arial, sans-serif; font-size: 11pt; }}
    h2 {{ margin: 16px 0 8px; font-size: 12pt; }}
    table {{ border-collapse: collapse; margin-bottom: 14px; width: 100%; }}
    th, td {{ border: 1px solid #BFC8D8; padding: 6px 8px; }}
    th {{ background: #DCE6F4; font-weight: 700; }}
    td {{ mso-number-format: "\\@"; }}
  </style>
</head>
<body>
  ${{partes.join('<br/>')}}
</body>
</html>`;

    const blob = new Blob(["\ufeff" + excelHtml], {{ type: 'application/vnd.ms-excel;charset=utf-8;' }});
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.href = url;
    link.download = 'PN_Executivo_{periodo['final'].replace('/', '-')}.xls';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }}

  (() => {{
    const seasonButtons = document.querySelectorAll('.season-btn');
    const seasonPanels = document.querySelectorAll('.season-panel');
    if (seasonButtons.length && seasonPanels.length) {{
      seasonButtons.forEach((btn) => {{
        btn.addEventListener('click', () => {{
          const target = btn.getAttribute('data-season-target');
          seasonButtons.forEach((b) => b.classList.remove('active'));
          seasonPanels.forEach((p) => p.classList.remove('active'));
          btn.classList.add('active');
          const panel = document.getElementById(target);
          if (panel) panel.classList.add('active');
        }});
      }});
    }}

    function toggleMonth(month) {{
      const rows = document.querySelectorAll('tr.day-row[data-month="' + month + '"]');
      if (!rows.length) return;
      const isOpen = rows[0].style.display !== 'none';
      rows.forEach((r) => {{ r.style.display = isOpen ? 'none' : 'table-row'; }});
      document.querySelectorAll('.month-toggle[data-month="' + month + '"]').forEach((b) => {{
        b.textContent = isOpen ? '▸' : '▾';
      }});
    }}

    document.querySelectorAll('.month-toggle').forEach((btn) => {{
      btn.addEventListener('click', (ev) => {{
        ev.preventDefault();
        ev.stopPropagation();
        toggleMonth(btn.getAttribute('data-month'));
      }});
    }});

    document.querySelectorAll('tr.month-row').forEach((row) => {{
      row.addEventListener('click', () => {{
        toggleMonth(row.getAttribute('data-month'));
      }});
    }});
  }})();
  </script>
  </body>
  </html>"""

    if _return_parts:
        import re as _re
        # Extrai TODOS os <style> e limpa raízes/body da CSS
        _all_styles = _re.findall(r'<style[^>]*>(.*?)</style>', html, _re.DOTALL)
        _css_raw = (_all_styles[-1] if _all_styles else '').strip()  # Pega último (PN)
        _css_raw = _re.sub(r'@import[^\n]+\n?', '', _css_raw)
        _css_raw = _re.sub(r':root\s*\{[^}]*\}', '', _css_raw, flags=_re.DOTALL)
        _css_raw = _re.sub(r'\bbody\s*\{[^}]*\}', '', _css_raw, flags=_re.DOTALL)
        # Extrai conteúdo entre <body> e primeiro <script>
        _body_m = _re.search(r'<body>(.*?)<script', html, _re.DOTALL)
        _body_html = (_body_m.group(1) if _body_m else '').strip()
        # Fallback: pega o último <script> se houver múltiplos
        _all_scripts = _re.findall(r'<script>(.*?)</script>', html, _re.DOTALL)
        _js_raw = (_all_scripts[-1] if _all_scripts else '').strip()
        return _css_raw, _body_html, _js_raw

    write_text_utf8(caminho_saida, html)
    return caminho_saida

def integrar_combo_mapa(usuario: str, senha: str, log_fn=None, base_dir=None,
                        cache_mgr=None) -> dict:
    """
    Integra o fluxo completo de mapeamento de OSs com lat/long (macro 50012).
    
    Fluxo automático:
      1. Login no Waterfy
      2. Download das Ordens de Serviço (OSs ativas)
      3. Download da macro 50012 (lat/long por matrícula)
      4. Cruzamento: OSs + lat/long → pontos no mapa
      5. Distribui por funcionário e gera estatísticas
    
    Args:
        usuario   : usuário Waterfy
        senha     : senha Waterfy
        log_fn    : função de log (str → None), opcional
        base_dir  : diretório base (default: script dir)
    
    
    Returns:
        dict com combo_mapa pronto para gerar_html()
    """
    if not HAS_WATERFY:
        return {'status': 'error', 'msg': 'waterfy_engine não disponível'}
    
    try:
        if log_fn:
            log_fn(" Iniciando integração de mapeamento de OSs com lat/long...")
        
        resultado = wfy_eng.gerar_combo_mapa(usuario, senha, log_fn, base_dir,
                                             cache_mgr=cache_mgr)
        
        if resultado.get('status') == 'error':
            if log_fn:
                log_fn(f" Erro: {resultado.get('msg', 'Desconhecido')}")
            return resultado
        
        if log_fn:
            log_fn(f" Mapeamento concluído: {resultado.get('total_cruzadas', 0)} OSs cruzadas")
        
        return resultado
    
    except Exception as e:
        if log_fn:
            log_fn(f" Exceção: {str(e)}")
        return {'status': 'error', 'msg': str(e)}


def _calcular_insights(dados_8117: dict, dados_8121: dict) -> dict:
    # Arrecadação: prefere dados da 8121 (que tem arr_diario próprio),
    # cai para 8117 se 8121 não tiver
    arr_d   = dados_8121.get('arr_diario', []) or dados_8117.get('arr_diario', [])
    arr_ac  = dados_8121.get('arr_acumulado', []) or dados_8117.get('arr_acumulado', [])
    datas   = dados_8121.get('datas', []) or dados_8117.get('datas', [])

    fat_d   = dados_8121.get('fat_diario', [])
    fat_ac  = dados_8121.get('fat_acumulado', [])
    datas21 = dados_8121.get('datas', datas)

    # Otimização: usar zip + filter em vez de enumerate + list comprehension
    if arr_d and datas:
        validos_arr = [(d, v) for d, v in zip(datas, arr_d) if v > 0]
    else:
        validos_arr = []
    
    if fat_d and datas21:
        validos_fat = [(d, v) for d, v in zip(datas21, fat_d) if v > 0]
    else:
        validos_fat = []

    # Total: usar último acumulado se disponível, senão soma
    total_arr  = arr_ac[-1] if arr_ac else (sum(arr_d) if arr_d else 0)
    total_fat  = fat_ac[-1] if fat_ac else (sum(fat_d) if fat_d else 0)
    
    # Calcular valores em uma única passagem
    if validos_arr:
        arr_valores = [v for _, v in validos_arr]
        media_arr = sum(arr_valores) / len(arr_valores)
        maior_arr = max(validos_arr, key=lambda x: x[1])
    else:
        media_arr = 0
        maior_arr = ('--', 0)
    
    if validos_fat:
        fat_valores = [v for _, v in validos_fat]
        media_fat = sum(fat_valores) / len(fat_valores)
        maior_fat = max(validos_fat, key=lambda x: x[1])
    else:
        media_fat = 0
        maior_fat = ('--', 0)
    
    dias_sem_arr    = sum(1 for v in arr_d if v == 0) if arr_d else 0
    pct_realizado   = (total_arr / total_fat * 100) if total_fat > 0 else 0
    acima_media_arr = sum(1 for _, v in validos_arr if v > media_arr) if validos_arr else 0

    if len(arr_d) >= 3:
        tendencia = 'alta' if arr_d[-1] > arr_d[-3] else 'baixa'
    elif len(arr_d) >= 2:
        tendencia = 'alta' if arr_d[-1] > arr_d[0] else 'baixa'
    else:
        tendencia = 'neutro'

    return {
        'total_arr':       total_arr,
        'total_fat':       total_fat,
        'media_arr':       media_arr,
        'media_fat':       media_fat,
        'maior_arr':       maior_arr,
        'maior_fat':       maior_fat,
        'dias_sem_arr':    dias_sem_arr,
        'pct_realizado':   pct_realizado,
        'acima_media_arr': acima_media_arr,
        'total_dias':      len(validos_arr),
        'tendencia':       tendencia,
    }


def _calcular_insights_8091(d91: dict) -> list:
    """Gera análise interpretada automática para a Macro 8091."""
    kpis        = d91.get('kpis', {})
    sit91       = d91.get('situacao', [])
    cat91       = d91.get('categoria', [])
    fase91      = d91.get('fase', [])
    bairros91   = d91.get('bairros', [])
    faixas91    = d91.get('faixas', {})
    faixas_vol  = d91.get('faixas_vol', {})
    comp91      = d91.get('comp', {})
    serie       = d91.get('serie_meses', [])
    grupos91    = d91.get('grupos', [])
    grp_rec     = d91.get('grupos_receita', [])
    leituristas = d91.get('leiturista', [])
    czero       = d91.get('consumo_zero', [])
    leit_dia    = d91.get('leit_dia', [])

    total_fat   = kpis.get('total_fat', 0)
    total_agua  = kpis.get('total_agua', 0)
    total_mult  = kpis.get('total_multa', 0)
    n_lig       = kpis.get('n_ligacoes', 0)
    vol_total   = kpis.get('vol_total', 0)
    media_fat   = kpis.get('media_fatura', 0)
    meses_count = d91.get('meses_count', 1)

    insights = []

    def _p(txt, emoji='', cor=''):
        insights.append({'emoji': emoji, 'cor': cor, 'texto': txt})

    # ── Pré-processar dicts ────────────────────────────────────────────────
    sit_dict  = {s.get('label','').upper(): s for s in sit91}
    sit_total = sum(s.get('qtd', 0) for s in sit91)
    fases_criticas = {f.get('label','').upper(): f for f in fase91
                      if any(x in f.get('label','').upper() for x in ['COBRAN','JUDICI','PROTESTO'])}

    # ── 1. VISÃO GERAL DO FATURAMENTO ────────────────────────────────────
    if total_fat > 0:
        media_por_mes = total_fat / meses_count if meses_count > 1 else total_fat
        periodo_txt = f"em {meses_count} meses (média de {_fmt_brl(media_por_mes)}/mês)" if meses_count > 1 else "no mês"
        _p(f'<strong>Faturamento total de {_fmt_brl(total_fat)}</strong> {periodo_txt}, '
           f'distribuído entre {n_lig:,} ligações ativas.', '', 'green')

    # Composição do faturamento
    if total_fat > 0 and total_agua > 0:
        pct_agua = total_agua / total_fat * 100
        pct_mult = total_mult / total_fat * 100 if total_mult else 0
        pct_outros = 100 - pct_agua - pct_mult
        obs = ''
        if pct_mult > 5:
            obs = f' O percentual de multas ({pct_mult:.1f}%) está elevado  indica alta inadimplência histórica.'
        elif pct_mult < 1:
            obs = f' Multas representam apenas {pct_mult:.1f}%  base com bom histórico de pagamento.'
        _p(f'<strong>Composição do faturamento:</strong> {pct_agua:.1f}% água ({_fmt_brl(total_agua)}), '
           f'{pct_mult:.1f}% multas ({_fmt_brl(total_mult)}) e {pct_outros:.1f}% outros encargos.{obs}',
           '', 'orange' if pct_mult > 5 else '')

    # ── 2. INADIMPLÊNCIA E SITUAÇÃO DAS CONTAS ───────────────────────────
    if sit_total > 0:
        pagos  = next((v for k, v in sit_dict.items() if 'PAGO' in k), None)
        debito = next((v for k, v in sit_dict.items() if 'DEBITO' in k or 'DÉBITO' in k), None)

        if pagos and debito:
            pct_pago = pagos['qtd'] / sit_total * 100
            pct_deb  = debito['qtd'] / sit_total * 100
            valor_deb = debito.get('valor', 0)

            if pct_pago >= 80:
                avaliacao = f'índice de adimplência <strong>excelente</strong> ({pct_pago:.1f}% pagos)'
                cor = 'green'
            elif pct_pago >= 65:
                avaliacao = f'índice de adimplência <strong>satisfatório</strong> ({pct_pago:.1f}% pagos)'
                cor = 'green'
            elif pct_pago >= 50:
                avaliacao = f'índice de adimplência <strong>preocupante</strong> ({pct_pago:.1f}% pagos)'
                cor = 'orange'
            else:
                avaliacao = f'índice de adimplência <strong>crítico</strong> ({pct_pago:.1f}% pagos)'
                cor = 'red'

            ticket_medio_deb = valor_deb / debito['qtd'] if debito['qtd'] > 0 else 0
            _p(f'A base apresenta {avaliacao}. '
               f'Há <strong>{debito["qtd"]:,} faturas em débito</strong> ({pct_deb:.1f}%), '
               f'representando {_fmt_brl(valor_deb)} em aberto '
               f'(ticket médio de débito: {_fmt_brl(ticket_medio_deb)}).',
               '' if pct_pago >= 65 else '', cor)

    # Cobrança avançada
    if fases_criticas:
        tot_crit  = sum(f['qtd'] for f in fases_criticas.values())
        val_crit  = sum(f.get('valor', 0) for f in fases_criticas.values())
        pct_crit  = tot_crit / n_lig * 100 if n_lig else 0
        nomes = [k.title() for k in list(fases_criticas.keys())[:3]]
        cor = 'red' if pct_crit > 10 else 'orange'
        acao = 'Recomenda-se priorizar ações de cobrança judicial e negativação.' if pct_crit > 10 else \
               'Acompanhar evolução para evitar escalada para fases judiciais.'
        _p(f'<strong>{pct_crit:.1f}% das ligações ({tot_crit:,}) estão em cobrança avançada</strong> '
           f'({", ".join(nomes)}), totalizando {_fmt_brl(val_crit)}. {acao}',
           '', cor)

    # ── 3. ANÁLISE POR CATEGORIA ─────────────────────────────────────────
    if cat91 and total_fat > 0:
        top = cat91[0]
        pct_top = top['valor'] / total_fat * 100
        outras = [(c['label'], c['valor'] / total_fat * 100) for c in cat91[1:3] if c.get('valor', 0) > 0]
        outras_txt = ', '.join(f'{l} ({p:.1f}%)' for l, p in outras)
        obs = ''
        if pct_top > 95:
            obs = ' A base é praticamente homogênea  ações tarifárias têm impacto uniforme.'
        elif pct_top < 70:
            obs = ' A diversidade de categorias exige estratégias de cobrança diferenciadas por perfil.'
        _p(f'<strong>{top["label"]} concentra {pct_top:.1f}%</strong> do faturamento '
           f'({top["qtd"]:,} ligações, {_fmt_brl(top["valor"])}).{" Demais: " + outras_txt + "." if outras_txt else ""}{obs}',
           '')

    # ── 4. ANÁLISE DE VOLUME E CONSUMO ───────────────────────────────────
    if vol_total > 0 and n_lig > 0:
        media_vol = vol_total / n_lig
        receita_m3 = total_fat / vol_total if vol_total > 0 else 0

        nivel = 'dentro do esperado'
        if media_vol < 6:
            nivel = 'abaixo do padrão residencial típico (615 m³)  possível submedição ou muitos imóveis fechados'
            cor = 'orange'
        elif media_vol > 20:
            nivel = 'acima do padrão residencial  base com perfil comercial/industrial significativo'
            cor = 'orange'
        else:
            cor = ''

        _p(f'<strong>Consumo médio de {media_vol:.1f} m³/ligação</strong> ({vol_total:,.0f} m³ total), '
           f'{nivel}. Eficiência tarifária: <strong>{_fmt_brl(receita_m3)}/m³</strong> arrecadado.',
           '', cor)

    # Faixa de consumo volumétrico dominante
    if faixas_vol:
        top_fv = max(faixas_vol.items(), key=lambda x: x[1])
        pct_fv = top_fv[1] / n_lig * 100 if n_lig else 0
        zeros = faixas_vol.get('0 m³', 0) or faixas_vol.get('Zero', 0)
        pct_zero = zeros / n_lig * 100 if n_lig and zeros else 0
        obs_zero = f' <strong>Atenção: {pct_zero:.1f}% das ligações ({zeros:,}) registraram consumo zero</strong>  investigar hidrômetros com defeito ou imóveis desocupados.' if pct_zero > 3 else ''
        _p(f'A faixa de consumo mais frequente é <strong>{top_fv[0]}</strong>, '
           f'com {top_fv[1]:,} ligações ({pct_fv:.1f}% da base).{obs_zero}',
           '', 'orange' if pct_zero > 3 else '')

    # ── 5. ANÁLISE POR GRUPOS DE LEITURA ─────────────────────────────────
    if grp_rec and len(grp_rec) >= 2:
        recs = [(g['label'], g['receita_m3'], g['vol'], g['valor']) for g in grp_rec if g.get('vol', 0) > 0]
        if recs:
            media_rm3 = sum(r[1] for r in recs) / len(recs)
            top_rm3   = max(recs, key=lambda x: x[1])
            bot_rm3   = min(recs, key=lambda x: x[1])
            diff_pct  = (top_rm3[1] - bot_rm3[1]) / bot_rm3[1] * 100 if bot_rm3[1] > 0 else 0

            obs = ''
            if diff_pct > 30:
                obs = (f' A diferença de {diff_pct:.0f}% entre o grupo mais e menos eficiente sugere '
                       f'desigualdade tarifária entre rotas  vale verificar perfil de clientes e categorias.')
            _p(f'<strong>Eficiência tarifária por grupo:</strong> '
               f'{top_rm3[0]} lidera com {_fmt_brl(top_rm3[1])}/m³ '
               f'({_fmt_brl(top_rm3[3])} faturados), enquanto '
               f'{bot_rm3[0]} tem menor retorno ({_fmt_brl(bot_rm3[1])}/m³). '
               f'Média geral: {_fmt_brl(media_rm3)}/m³.{obs}',
               '', 'orange' if diff_pct > 30 else '')

    # ── 6. ANÁLISE DE BAIRROS ────────────────────────────────────────────
    if bairros91 and total_fat > 0:
        top3 = bairros91[:3]
        val_top3 = sum(b['valor'] for b in top3)
        pct_top3 = val_top3 / total_fat * 100
        lista = ', '.join(f'<strong>{b["label"].title()}</strong> ({_fmt_brl(b["valor"])}, {b["qtd"]:,} lig.)' for b in top3)
        obs = ' Alta concentração geográfica do faturamento.' if pct_top3 > 60 else ''
        _p(f'Os 3 bairros com maior faturamento respondem por <strong>{pct_top3:.1f}%</strong> do total: '
           f'{lista}.{obs}', '')

    # ── 7. ANÁLISE DE LEITURISTAS ────────────────────────────────────────
    if leituristas and len(leituristas) >= 2:
        vols_l  = [(l['label'], l['vol'], l['qtd'], l['vol_medio']) for l in leituristas if l.get('qtd', 0) > 0]
        if vols_l:
            media_vm = sum(l[3] for l in vols_l) / len(vols_l)
            top_vol  = max(vols_l, key=lambda x: x[1])
            anomalos_alto = [l for l in vols_l if l[3] > media_vm * 1.35]
            anomalos_bx   = [l for l in vols_l if l[3] < media_vm * 0.65 and l[1] > 0]

            txt = (f'<strong>Desempenho da equipe de leitura:</strong> '
                   f'{top_vol[0].split()[0]} liderou em volume ({top_vol[1]:,.0f} m³, {top_vol[2]:,} lig.). '
                   f'Volume médio geral da equipe: {media_vm:.1f} m³/ligação.')
            cor = ''
            if anomalos_alto:
                nomes_a = ', '.join(l[0].split()[0] for l in anomalos_alto[:2])
                txt += (f' <strong>{nomes_a} apresentam volume médio acima de 35% da média</strong> '
                        f'({", ".join(f"{l[3]:.1f} m³/lig" for l in anomalos_alto[:2])})  '
                        f'pode indicar rotas com clientes comerciais ou leituras estimadas.')
                cor = 'orange'
            if anomalos_bx:
                nomes_b = ', '.join(l[0].split()[0] for l in anomalos_bx[:2])
                txt += (f' <strong>{nomes_b} ficaram abaixo de 65% da média</strong>  '
                        f'verificar rotas incompletas ou alto índice de consumo zero.')
                cor = 'orange'
            _p(txt, '', cor)

    # ── 8. CONSUMO ZERO ──────────────────────────────────────────────────
    if czero:
        tot_zero = sum(c.get('qtd', 0) for c in czero)
        pct_zero = tot_zero / n_lig * 100 if n_lig else 0
        top_czero = max(czero, key=lambda x: x.get('qtd', 0))
        cor = 'red' if pct_zero > 8 else ('orange' if pct_zero > 3 else '')
        acao = ('Nível crítico  recomenda-se vistoria em campo.' if pct_zero > 8 else
                'Nível de atenção  monitorar evolução.' if pct_zero > 3 else
                'Nível dentro do esperado.')
        _p(f'<strong>{tot_zero:,} ligações ({pct_zero:.1f}%) registraram consumo zero.</strong> '
           f'{acao} Categoria com mais zeros: <strong>{top_czero["label"]}</strong> '
           f'({top_czero["qtd"]:,} ligações).',
           '', cor)

    # ── 9. TENDÊNCIA MULTI-MÊS ───────────────────────────────────────────
    if len(serie) >= 2:
        ult  = serie[-1]
        ante = serie[-2]
        var  = (ult['total'] - ante['total']) / ante['total'] * 100 if ante['total'] else 0
        sinal = '' if var >= 0 else ''
        cor   = 'green' if var >= 2 else ('red' if var <= -5 else '')

        if len(serie) >= 3:
            # Tendência de 3 meses
            vars3 = [(serie[i]['total'] - serie[i-1]['total']) / serie[i-1]['total'] * 100
                     for i in range(1, len(serie)) if serie[i-1]['total'] > 0]
            tend = 'crescimento consistente' if all(v > 0 for v in vars3[-2:]) else \
                   'queda consistente' if all(v < 0 for v in vars3[-2:]) else 'oscilação'
            _p(f'<strong>Tendência de {len(serie)} meses: {tend}.</strong> '
               f'Última variação: {var:+.1f}% ({ante["periodo"]} → {ult["periodo"]}). '
               f'Faturamento atual: {_fmt_brl(ult["total"])}.',
               sinal, cor)
        else:
            _p(f'<strong>Variação no período:</strong> {var:+.1f}% '
               f'({ante["periodo"]} → {ult["periodo"]}), '
               f'de {_fmt_brl(ante["total"])} para {_fmt_brl(ult["total"])}.',
               sinal, cor)

    # ── 10. FAIXA DE VALOR DOMINANTE ─────────────────────────────────────
    if faixas91 and n_lig > 0:
        top_fx = max(faixas91.items(), key=lambda x: x[1])
        pct_fx = top_fx[1] / n_lig * 100
        # Detecta concentração na faixa mínima
        faixa_min = next(((k, v) for k, v in faixas91.items()
                          if any(x in k for x in ['0', 'mín', 'min', 'R$0', 'R$ 0'])), None)
        obs_min = ''
        if faixa_min and faixa_min[1] / n_lig * 100 > 20:
            obs_min = (f' {faixa_min[1] / n_lig * 100:.1f}% das faturas estão na faixa mínima  '
                       f'potencial de receita travado pela tarifa social.')
        _p(f'A faixa de faturamento mais comum é <strong>{top_fx[0]}</strong> '
           f'({top_fx[1]:,} faturas, {pct_fx:.1f}% da base).{obs_min}',
           '', 'orange' if obs_min else '')

    return insights


# ── MINI-INSIGHTS 8117 ────────────────────────────────────────────────────────

def _calcular_mini_insights_8117(d17: dict) -> dict:
    result = {}
    datas  = d17.get('datas', [])
    arr_d  = d17.get('arr_diario', [])
    arr_ac = d17.get('arr_acumulado', [])
    formas = d17.get('formas', {})
    total_arr = arr_ac[-1] if arr_ac else sum(arr_d) if arr_d else 0
    validos = [(d, v) for d, v in zip(datas, arr_d) if v > 0]

    if validos:
        media_dia = sum(v for _, v in validos) / len(validos)
        melhor = max(validos, key=lambda x: x[1])
        pior   = min(validos, key=lambda x: x[1])
        dias_sem = len(arr_d) - len(validos)
        acima = sum(1 for _, v in validos if v > media_dia)
        # Tendência: só faz sentido entre meses completos  dentro de um único
        # mês a arrecadação varia por ciclo de vencimentos (mais no início),
        # não por desempenho real. O card JS atualiza dinamicamente ao trocar
        # para view Mensal quando houver >= 2 meses na série.
        items = [
            {'emoji': '', 'cor': 'green',
             'texto': f'<strong>Melhor dia:</strong> {melhor[0]} com {_fmt_brl(melhor[1])} ({melhor[1]/total_arr*100:.1f}% do total do período)'},
            {'emoji': '', 'cor': '',
             'texto': f'<strong>Média diária:</strong> {_fmt_brl(media_dia)}  {acima} de {len(validos)} dias acima da média ({acima/len(validos)*100:.0f}%)'},
            {'emoji': '', 'cor': 'red',
             'texto': f'<strong>Menor dia:</strong> {pior[0]} com {_fmt_brl(pior[1])} ({pior[1]/total_arr*100:.1f}% do total)'},
        ]
        if dias_sem > 0:
            items.append({'emoji': 'ℹ', 'cor': '',
             'texto': f'<strong>{dias_sem} dia(s) sem arrecadação</strong> no período (possíveis feriados ou falhas de importação)'})
        result['chart-arrec-diario'] = _mini_insights_html(items)

    if formas:
        total_f = sum(v['valor'] for v in formas.values()) or 1
        sorted_f = sorted(formas.items(), key=lambda x: -x[1]['valor'])
        top_f = sorted_f[0]; bot_f = sorted_f[-1]
        pct_top = top_f[1]['valor'] / total_f * 100
        pct_bot = bot_f[1]['valor'] / total_f * 100
        top3_pct = sum(v['valor'] for _, v in sorted_f[:3]) / total_f * 100 if len(sorted_f) >= 3 else pct_top
        items = [
            {'emoji': '', 'cor': 'green',
             'texto': f'<strong>{top_f[0]}</strong> é a principal forma: {_fmt_brl(top_f[1]["valor"])} ({pct_top:.1f}% do total)' + ('  alta concentração em um único canal' if pct_top > 50 else '')},
            {'emoji': '', 'cor': '',
             'texto': f'<strong>Top 3 formas concentram {top3_pct:.1f}%</strong> de toda a arrecadação do período'},
            {'emoji': '', 'cor': '',
             'texto': f'<strong>{bot_f[0]}</strong> tem menor participação: {_fmt_brl(bot_f[1]["valor"])} ({pct_bot:.1f}%)'},
        ]
        result['chart-formas-pizza'] = _mini_insights_html(items)
        result['chart-formas-barra'] = _mini_insights_html(items)

    return result


# ── MINI-INSIGHTS 8121 ────────────────────────────────────────────────────────

def _calcular_mini_insights_8121(d17: dict, d21: dict) -> dict:
    result = {}
    datas21 = d21.get('datas', [])
    # Arrecadação: prefere dados próprios da 8121, cai para 8117 se necessário
    arr_d   = d21.get('arr_diario', []) or d17.get('arr_diario', [])
    arr_ac  = d21.get('arr_acumulado', []) or d17.get('arr_acumulado', [])
    fat_d   = d21.get('fat_diario', [])
    fat_ac  = d21.get('fat_acumulado', [])
    total_arr = arr_ac[-1] if arr_ac else sum(arr_d) if arr_d else 0
    total_fat = fat_ac[-1] if fat_ac else sum(fat_d) if fat_d else 0
    pct_real  = total_arr / total_fat * 100 if total_fat else 0

    if fat_d and arr_d:
        pares = [(d, a, f) for d, a, f in zip(datas21, arr_d[:len(datas21)], fat_d) if f > 0]
        if pares:
            efic = [(d, a/f*100) for d, a, f in pares]
            melhor_e = max(efic, key=lambda x: x[1])
            pior_e   = min(efic, key=lambda x: x[1])
            gaps = [(d, f-a) for d, a, f in pares]
            maior_gap = max(gaps, key=lambda x: x[1])
            cor_pct = 'green' if pct_real >= 90 else ('orange' if pct_real >= 70 else 'red')
            items = [
                {'emoji': '', 'cor': 'green',
                 'texto': f'<strong>Dia mais eficiente:</strong> {melhor_e[0]} com {melhor_e[1]:.1f}% de realização diária'},
                {'emoji': '', 'cor': cor_pct,
                 'texto': f'<strong>Realização acumulada: {pct_real:.1f}%</strong>  {_fmt_brl(total_arr)} arrecadados de {_fmt_brl(total_fat)} faturados'},
                {'emoji': '', 'cor': 'orange' if pior_e[1] < 60 else '',
                 'texto': f'<strong>Maior gap:</strong> {maior_gap[0]} com {_fmt_brl(maior_gap[1])} entre faturamento e arrecadação'},
            ]
            result['chart-arrec-fat'] = _mini_insights_html(items)

    if total_fat > 0:
        gap_acum = total_fat - total_arr
        cor_pct  = 'green' if pct_real >= 90 else ('orange' if pct_real >= 70 else 'red')
        tend = ''
        if len(arr_d) >= 5:
            media_ini = sum(arr_d[:3]) / 3
            media_fim = sum(arr_d[-3:]) / 3
            if media_fim > media_ini * 1.05:
                tend = 'aceleração'
            elif media_fim < media_ini * 0.95:
                tend = 'desaceleração'
        items = [
            {'emoji': '', 'cor': 'green' if pct_real >= 90 else '',
             'texto': f'<strong>Total arrecadado:</strong> {_fmt_brl(total_arr)} | <strong>Total faturado:</strong> {_fmt_brl(total_fat)}'},
            {'emoji': '', 'cor': cor_pct,
             'texto': f'<strong>{pct_real:.1f}% do faturamento realizado</strong>  gap de {_fmt_brl(gap_acum)} ainda em aberto'},
        ]
        if tend:
            items.append({'emoji': '', 'cor': 'green' if tend == 'aceleração' else 'orange',
             'texto': f'<strong>Tendência de {tend} na arrecadação</strong> nos últimos 3 dias do período'})
        result['chart-acumulado'] = _mini_insights_html(items)

    return result


# ── MINI-INSIGHTS 50012 ───────────────────────────────────────────────────────

def _calcular_mini_insights_50012(d50: dict) -> dict:
    result = {}
    if not d50:
        return result
    kpis        = d50.get('kpis', {})
    criticas    = d50.get('criticas', [])
    faixas      = d50.get('faixas_consumo', {})
    leituristas = d50.get('leituristas', [])
    grupos      = d50.get('grupos', [])
    cobranca    = d50.get('cobranca', [])
    bairros_fat = d50.get('bairros_fat', [])
    desvio      = d50.get('desvio', [])
    leit_dia    = d50.get('leit_dia', [])
    categorias  = d50.get('categorias', [])
    hidrometros = d50.get('hidrometros', [])
    equipes     = d50.get('equipes', [])
    bairros     = d50.get('bairros_top15', [])
    n_total  = kpis.get('total_ligacoes', 0) or 1
    zeros    = kpis.get('zeros_vol', 0)
    fat_total= kpis.get('total_fat', 0)
    media_vol= kpis.get('media_vol', 0)
    pct_normal = kpis.get('pct_normal', 0)

    # Críticas
    if criticas:
        tot_c = sum(c['qtd'] for c in criticas) or 1
        normal = next((c for c in criticas if 'NORMAL' in c['label'].upper()), None)
        crit_reais = [c for c in criticas if 'NORMAL' not in c['label'].upper()]
        top_crit = crit_reais[0] if crit_reais else None
        pct_n = normal['qtd'] / tot_c * 100 if normal else pct_normal
        cor_n = 'green' if pct_n >= 90 else ('orange' if pct_n >= 75 else 'red')
        items = [
            {'emoji': '', 'cor': cor_n,
             'texto': f'<strong>{pct_n:.1f}% das leituras são normais</strong> ({(normal["qtd"] if normal else 0):,} ligações)  {"nível excelente" if pct_n >= 90 else "nível aceitável" if pct_n >= 75 else "requer atenção"}'},
        ]
        if top_crit:
            pct_tc = top_crit['qtd'] / tot_c * 100
            items.append({'emoji': '', 'cor': 'red' if pct_tc > 10 else 'orange',
             'texto': f'<strong>Crítica mais frequente:</strong> {top_crit["label"]}  {top_crit["qtd"]:,} ligações ({pct_tc:.1f}%)'})
        if zeros > 0:
            pct_z = zeros / n_total * 100
            items.append({'emoji': '', 'cor': 'red' if pct_z > 5 else 'orange',
             'texto': f'<strong>{zeros:,} leituras com volume zero</strong> ({pct_z:.1f}%)  verificar hidrômetros parados ou imóveis fechados'})
        result['chart-50-criticas'] = _mini_insights_html(items)

    # Faixas consumo
    if faixas:
        tot_f = sum(faixas.values()) or 1
        top_fx = max(faixas.items(), key=lambda x: x[1])
        bot_fx = min(faixas.items(), key=lambda x: x[1])
        items = [
            {'emoji': '', 'cor': 'green',
             'texto': f'<strong>Faixa mais frequente: {top_fx[0]} m³</strong> com {top_fx[1]:,} ligações ({top_fx[1]/tot_f*100:.1f}% da base)'},
            {'emoji': '', 'cor': '',
             'texto': f'<strong>Volume médio por ligação: {media_vol:.1f} m³</strong>  ' + ('abaixo do padrão residencial (6-15 m³)' if media_vol < 6 else 'acima do padrão  perfil comercial/industrial' if media_vol > 20 else 'dentro do padrão residencial esperado')},
            {'emoji': '', 'cor': '',
             'texto': f'<strong>Faixa menos frequente: {bot_fx[0]} m³</strong> com {bot_fx[1]:,} ligações ({bot_fx[1]/tot_f*100:.1f}%)'},
        ]
        result['chart-50-faixas'] = _mini_insights_html(items)

    # Cobrança
    if cobranca:
        tot_cob = sum(c['qtd'] for c in cobranca) or 1
        critica_kws = ('COBRAN', 'JUDICI', 'PROTESTO', 'CORTE', 'BLOQUEIO')
        crit_cob = [c for c in cobranca if any(k in c['label'].upper() for k in critica_kws)]
        n_crit_cob = sum(c['qtd'] for c in crit_cob)
        pct_crit_cob = n_crit_cob / tot_cob * 100
        top_cob = max(cobranca, key=lambda x: x['qtd'])
        bot_cob = min(cobranca, key=lambda x: x['qtd'])
        items = [
            {'emoji': '', 'cor': '',
             'texto': f'<strong>Status predominante: {top_cob["label"]}</strong> com {top_cob["qtd"]:,} ligações ({top_cob["qtd"]/tot_cob*100:.1f}%)'},
        ]
        if n_crit_cob > 0:
            top_cc = max(crit_cob, key=lambda x: x['qtd']) if crit_cob else None
            items.append({'emoji': '', 'cor': 'red' if pct_crit_cob > 10 else 'orange',
             'texto': f'<strong>{n_crit_cob:,} ligações em cobrança crítica</strong> ({pct_crit_cob:.1f}%)' + (f'  {top_cc["label"]} lidera' if top_cc else '')})
        items.append({'emoji': '', 'cor': '',
         'texto': f'<strong>Menor status: {bot_cob["label"]}</strong> com {bot_cob["qtd"]:,} ligações ({bot_cob["qtd"]/tot_cob*100:.1f}%)'})
        result['chart-50-cobranca'] = _mini_insights_html(items)

    # Leituristas
    if leituristas and len(leituristas) >= 2:
        validos_l = [l for l in leituristas if l.get('qtd', 0) > 0]
        if validos_l:
            media_tc   = sum(l['taxa_crit'] for l in validos_l) / len(validos_l)
            top_vol    = max(validos_l, key=lambda x: x.get('vol', 0))
            top_crit_l = max(validos_l, key=lambda x: x['taxa_crit'])
            bot_crit_l = min(validos_l, key=lambda x: x['taxa_crit'])
            items = [
                {'emoji': '', 'cor': 'green',
                 'texto': f'<strong>{top_vol["label"].split()[0]}</strong> registrou maior volume: {top_vol.get("vol",0):,.0f} m³ em {top_vol["qtd"]:,} leituras'},
                {'emoji': '', 'cor': 'orange' if top_crit_l['taxa_crit'] > media_tc * 1.5 else '',
                 'texto': f'<strong>Taxa crítica média da equipe: {media_tc:.1f}%</strong>  ' + (f'{top_crit_l["label"].split()[0]} tem a maior taxa ({top_crit_l["taxa_crit"]:.1f}%) ' if top_crit_l['taxa_crit'] > media_tc * 1.5 else 'equipe dentro do padrão esperado')},
                {'emoji': '', 'cor': 'green',
                 'texto': f'<strong>{bot_crit_l["label"].split()[0]}</strong> tem menor taxa crítica: {bot_crit_l["taxa_crit"]:.1f}%'},
            ]
            result['chart-50-leituristas'] = _mini_insights_html(items)

    # Grupos
    if grupos and len(grupos) >= 2:
        top_v  = max(grupos, key=lambda x: x.get('vol', 0))
        bot_v  = min(grupos, key=lambda x: x.get('vol', 0))
        top_f2 = max(grupos, key=lambda x: x.get('valor', 0))
        tot_gvol = sum(g.get('vol', 0) for g in grupos) or 1
        items = [
            {'emoji': '', 'cor': 'green',
             'texto': f'<strong>{top_v["label"]}</strong> lidera em volume: {top_v.get("vol",0):,.0f} m³ ({top_v.get("vol",0)/tot_gvol*100:.1f}% do total)'},
            {'emoji': '', 'cor': '',
             'texto': f'<strong>Maior receita:</strong> {top_f2["label"]} com {_fmt_brl(top_f2.get("valor",0))}'},
            {'emoji': '', 'cor': '',
             'texto': f'<strong>Menor volume:</strong> {bot_v["label"]} com {bot_v.get("vol",0):,.0f} m³  verificar se rota está completa'},
        ]
        result['chart-50-grupos'] = _mini_insights_html(items)

    # Desvio
    if desvio:
        tot_dev = sum(d['qtd'] for d in desvio) or 1
        alto  = next((d for d in desvio if '+50%' in d['label']), None)
        queda = next((d for d in desvio if '< -50%' in d['label']), None)
        normal_d = max(desvio, key=lambda x: x['qtd'])
        items = [
            {'emoji': '', 'cor': '',
             'texto': f'<strong>Faixa mais comum: {normal_d["label"]}</strong> com {normal_d["qtd"]:,} ligações ({normal_d["qtd"]/tot_dev*100:.1f}%)'},
        ]
        if alto:
            pct_a = alto['qtd'] / tot_dev * 100
            items.append({'emoji': '', 'cor': 'orange' if pct_a > 15 else '',
             'texto': f'<strong>{alto["qtd"]:,} ligações com consumo > +50%</strong> da média histórica ({pct_a:.1f}%)  investigar vazamentos ou fraudes'})
        if queda:
            pct_q = queda['qtd'] / tot_dev * 100
            items.append({'emoji': '', 'cor': 'orange' if pct_q > 15 else '',
             'texto': f'<strong>{queda["qtd"]:,} ligações com queda > 50%</strong> ({pct_q:.1f}%)  verificar hidrômetros ou imóveis desocupados'})
        if not alto and not queda:
            items.append({'emoji': '', 'cor': 'green',
             'texto': 'Desvios dentro do esperado  consumo alinhado com histórico da base'})
        result['chart-50-desvio'] = _mini_insights_html(items)

    # Bairros faturamento
    if bairros_fat and fat_total > 0:
        top_b = bairros_fat[0]; bot_b = bairros_fat[-1]
        top3_fat = sum(b['valor'] for b in bairros_fat[:3])
        pct_top3 = top3_fat / fat_total * 100
        ticket_top = top_b['valor'] / top_b['qtd'] if top_b['qtd'] > 0 else 0
        items = [
            {'emoji': '', 'cor': 'green',
             'texto': f'<strong>{top_b["bairro"].title()}</strong> lidera: {_fmt_brl(top_b["valor"])} ({top_b["qtd"]:,} lig., ticket médio {_fmt_brl(ticket_top)})'},
            {'emoji': '', 'cor': 'orange' if pct_top3 > 60 else '',
             'texto': f'<strong>Top 3 bairros concentram {pct_top3:.1f}%</strong> do faturamento total' + ('  alta concentração geográfica' if pct_top3 > 60 else '')},
            {'emoji': '', 'cor': '',
             'texto': f'<strong>{bot_b["bairro"].title()}</strong> tem menor faturamento: {_fmt_brl(bot_b["valor"])} ({bot_b["qtd"]:,} lig.)'},
        ]
        result['chart-50-bairros-fat'] = _mini_insights_html(items)

    # Categorias
    if categorias:
        tot_cat = sum(c['qtd'] for c in categorias) or 1
        top_cat = max(categorias, key=lambda x: x['qtd'])
        bot_cat = min(categorias, key=lambda x: x['qtd'])
        pct_top_cat = top_cat['qtd'] / tot_cat * 100
        items = [
            {'emoji': '', 'cor': '',
             'texto': f'<strong>{top_cat["label"]}</strong> é a categoria dominante: {top_cat["qtd"]:,} ligações ({pct_top_cat:.1f}% da base)'},
            {'emoji': '', 'cor': '',
             'texto': 'Base praticamente homogênea  ações tarifárias têm impacto uniforme' if pct_top_cat > 90 else 'Demais: ' + ', '.join(f'{c["label"]} ({c["qtd"]:,})' for c in categorias[1:3])},
            {'emoji': '', 'cor': '',
             'texto': f'<strong>Menor categoria: {bot_cat["label"]}</strong> com {bot_cat["qtd"]:,} ligações ({bot_cat["qtd"]/tot_cat*100:.1f}%)'},
        ]
        result['chart-50-categorias'] = _mini_insights_html(items)

    # Leituras por dia
    if leit_dia:
        validos_d = [d for d in leit_dia if d.get('qtd', 0) > 0]
        if validos_d:
            top_d = max(validos_d, key=lambda x: x['qtd'])
            bot_d = min(validos_d, key=lambda x: x['qtd'])
            media_d = sum(d['qtd'] for d in validos_d) / len(validos_d)
            items = [
                {'emoji': '', 'cor': 'green',
                 'texto': f'<strong>Dia de maior produção: {top_d["data"]}</strong>  {top_d["qtd"]:,} leituras e {top_d.get("vol",0):,.0f} m³ medidos'},
                {'emoji': '', 'cor': '',
                 'texto': f'<strong>Média diária: {media_d:.0f} leituras/dia</strong>  {top_d["qtd"]:,} no pico vs {bot_d["qtd"]:,} no dia mais baixo'},
                {'emoji': '', 'cor': '',
                 'texto': f'<strong>Dia mais baixo: {bot_d["data"]}</strong> com {bot_d["qtd"]:,} leituras'},
            ]
            result['chart-50-dias'] = _mini_insights_html(items)

    # Equipes
    if equipes and len(equipes) >= 2:
        top_eq = max(equipes, key=lambda x: x['total'])
        bot_eq = min(equipes, key=lambda x: x['total'])
        tot_eq = sum(e['total'] for e in equipes) or 1
        items = [
            {'emoji': '', 'cor': 'green',
             'texto': f'<strong>{top_eq["equipe"]}</strong> tem mais OSs: {top_eq["total"]:,} ({top_eq["total"]/tot_eq*100:.1f}% do total, cobertura {top_eq.get("cobertura_pct",0):.1f}%)'},
            {'emoji': '', 'cor': '',
             'texto': f'<strong>Cobertura média de coordenadas:</strong> {sum(e.get("cobertura_pct",0) for e in equipes)/len(equipes):.1f}% entre as {len(equipes)} equipes'},
            {'emoji': '', 'cor': 'orange' if bot_eq["total"] < top_eq["total"] * 0.3 else '',
             'texto': f'<strong>{bot_eq["equipe"]}</strong> tem menos OSs: {bot_eq["total"]:,} ({bot_eq["total"]/tot_eq*100:.1f}%)'},
        ]
        result['chart-50-equipes'] = _mini_insights_html(items)

    # Bairros (ligações)
    if bairros:
        top_b2 = bairros[0]; bot_b2 = bairros[-1]
        tot_b2 = sum(b['total'] for b in bairros) or 1
        top3_b = sum(b['total'] for b in bairros[:3])
        items = [
            {'emoji': '', 'cor': 'green',
             'texto': f'<strong>{top_b2["bairro"].title()}</strong> concentra mais ligações: {top_b2["total"]:,} ({top_b2["total"]/tot_b2*100:.1f}% dos top 15)'},
            {'emoji': '', 'cor': '',
             'texto': f'<strong>Top 3 bairros representam {top3_b/tot_b2*100:.1f}%</strong> das ligações listadas neste ranking'},
            {'emoji': '', 'cor': '',
             'texto': f'<strong>{bot_b2["bairro"].title()}</strong> tem menos ligações no ranking: {bot_b2["total"]:,}'},
        ]
        result['chart-50-bairros'] = _mini_insights_html(items)

    # Hidrômetros
    if hidrometros:
        tot_hid = sum(h['qtd'] for h in hidrometros) or 1
        top_hid = hidrometros[0]; bot_hid = hidrometros[-1]
        pct_top_hid = top_hid['qtd'] / tot_hid * 100
        items = [
            {'emoji': '', 'cor': '',
             'texto': f'<strong>Modelo mais comum: {top_hid["modelo"][:3]}</strong> com {top_hid["qtd"]:,} hidrômetros ({pct_top_hid:.1f}% da base)'},
            {'emoji': '', 'cor': '',
             'texto': f'<strong>{len(hidrometros)} modelos diferentes</strong> identificados  ' + ('parque homogêneo' if len(hidrometros) <= 5 else 'diversidade de equipamentos requer atenção em manutenção')},
            {'emoji': '', 'cor': '',
             'texto': f'<strong>Modelo menos frequente: {bot_hid["modelo"][:3]}</strong> com {bot_hid["qtd"]:,} unidades ({bot_hid["qtd"]/tot_hid*100:.1f}%)'},
        ]
        result['chart-50-hidrometros'] = _mini_insights_html(items)

    return result


def _gerar_mapa_folium(pontos: list, centro: dict) -> str:
    """
    Gera HTML de mapa Folium com os pontos do combo_mapa.
    Retorna string HTML completa (sem <html>/<body>) para embutir como iframe srcdoc.
    Cada equipe recebe uma cor diferente. Retorna string vazia se Folium não disponível.
    """
    try:
        import folium
        from collections import defaultdict
        from datetime import datetime

        lat_c = centro.get('latitude')  or -26.1085
        lng_c = centro.get('longitude') or -48.6161

        def _fmt_data_popup(valor):
            valor = (valor or '').strip()
            if not valor:
                return ''
            for fmt_in, fmt_out in (
                ('%Y-%m-%d %H:%M:%S', '%d/%m/%Y %H:%M:%S'),
                ('%Y-%m-%d %H:%M', '%d/%m/%Y %H:%M'),
                ('%Y-%m-%d', '%d/%m/%Y'),
            ):
                try:
                    return datetime.strptime(valor, fmt_in).strftime(fmt_out)
                except Exception:
                    continue
            return valor

        # Mapa com visualização de satélite (Esri World Imagery)
        m = folium.Map(location=[lat_c, lng_c], zoom_start=13,
                       tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
                       attr='Tiles &copy; Esri')

        # Adiciona camada de mapa padrão como alternativa (clicável no mapa)
        folium.TileLayer('CartoDB positron', name='Mapa de Rua').add_to(m)

        # Paleta de cores por equipe (SEM verde/vermelho para evitar conflito com status de leitura)
        CORES = ['#1565C0', '#FF8F00', '#6A1B9A', '#1976D2', '#00838F',
                 '#AD1457', '#F57C00', '#5E35B1', '#0097A7', '#7B1FA2']
        equipes = list({p['equipe'] for p in pontos if p.get('equipe')})
        cor_map = {eq: CORES[i % len(CORES)] for i, eq in enumerate(sorted(equipes))}

        # Grupos por equipe para o controle de camadas
        grupos = {}
        for eq in equipes:
            grupos[eq] = folium.FeatureGroup(name=eq, show=True)

        for p in pontos:
            eq  = p.get('equipe', 'Indefinida')
            cor = cor_map.get(eq, '#37474F')

            # ── Cor do marcador por criticidade de medição ───────────────
            critica_raw = p.get('critica', '') or ''
            if '5900' in critica_raw or 'NORMAL' in critica_raw.upper():
                fill_cor = '#2E7D32'       # verde  leitura normal
            elif critica_raw:
                fill_cor = '#E53935'       # vermelho  leitura crítica
            else:
                fill_cor = '#2E7D32'       # padrão: verde

            # ── Badge de cobrança ────────────────────────────────────────
            cob = p.get('cobranca', '') or ''
            cob_badge = ''
            if any(kw in cob.upper() for kw in ('COBRAN', 'JUDICI', 'PROTESTO', 'CORTE')):
                cob_badge = f'<span style="background:#E53935;color:#fff;padding:1px 5px;border-radius:3px;font-size:0.75em">{cob}</span>'
            elif cob:
                cob_badge = f'<span style="background:#2196F3;color:#fff;padding:1px 5px;border-radius:3px;font-size:0.75em">{cob}</span>'

            # ── Monta popup ──────────────────────────────────────────────
            vol_str   = p.get('volume', '') or ''
            med_str   = p.get('media_consumo', '') or ''
            val_str   = p.get('valor_fatura', '') or ''
            crit_str  = p.get('critica', '') or 'Sem info'
            data_os_fmt = _fmt_data_popup(p.get('data_os', '') or '')
            prazo_os_fmt = _fmt_data_popup(p.get('prazo_os', '') or '')
            leitura_fmt = _fmt_data_popup(p.get('data_leitura', '') or '')
            popup_txt = (
                f"<b style='font-size:0.95em'>Matrícula: {p.get('matricula','')}</b><br>"
                f"<b>Equipe:</b> {eq}<br>"
                f"<b>Bairro:</b> {p.get('bairro','')} &nbsp;|&nbsp; <b>Setor:</b> {p.get('setor','') or ''}<br>"
                f"<b>Endereço:</b> {p.get('logradouro','') or ''}, {p.get('nr_imovel','') or 's/n'}<br>"
                f"<hr style='margin:4px 0'>"
                f"<b>N° OS:</b> {p.get('num_os','') or ''} &nbsp;&nbsp; "
                f"<b>Tipo:</b> {p.get('tipo_servico','') or ''}<br>"
                f"<b>Status OS:</b> {p.get('status_os','') or ''} &nbsp;|&nbsp; "
                f"<b>Prioridade:</b> {p.get('prioridade','') or ''}<br>"
                f"<b>Data OS:</b> {data_os_fmt} &nbsp;|&nbsp; "
                f"<b>Prazo:</b> {prazo_os_fmt}<br>"
                f"<hr style='margin:4px 0'>"
                f"<b>Crítica medição:</b> {crit_str}<br>"
                f"<b>Volume:</b> {vol_str or ''} m³ &nbsp;|&nbsp; <b>Média:</b> {med_str or ''} m³<br>"
                f"<b>Valor fatura:</b> R$ {val_str or ''}<br>"
                f"<b>Categoria:</b> {p.get('categoria','') or ''} &nbsp;|&nbsp; "
                f"<b>Grupo:</b> {p.get('grupo','') or ''}<br>"
                f"<b>Cobrança:</b> {cob_badge if cob_badge else (cob or '')}<br>"
                f"<b>Leitura:</b> {leitura_fmt}<br>"
                f"<small style='color:#777'>"
                f"Lat: {p.get('lat',''):.6f} | Lng: {p.get('lng',''):.6f}</small>"
            )
            tooltip_txt = (
                f"{eq}  {p.get('bairro','')} | OS:{p.get('num_os','') or '?'} | "
                f"Vol:{vol_str or '?'}m³"
            )
            folium.CircleMarker(
                location=[p['lat'], p['lng']],
                radius=6,
                color=cor,                # ← CONTORNO: cor da equipe
                fill=True,
                fill_color=fill_cor,      # ← PREENCHIMENTO: cor da leitura (crítica ou normal)
                fill_opacity=0.85,
                popup=folium.Popup(popup_txt, max_width=320),
                tooltip=tooltip_txt,
            ).add_to(grupos.get(eq, m))

        for g in grupos.values():
            g.add_to(m)

        folium.LayerControl(collapsed=False).add_to(m)

        return m.get_root().render()

    except ImportError:
        return (
            "<div style=\"padding:40px;text-align:center;font-family:sans-serif;color:#666\">"
            "<div style=\"font-size:2rem;margin-bottom:12px\"></div>"
            "<b>Folium não instalado</b><br>"
            "<small>Execute: <code>pip install folium</code></small>"
            "</div>"
        )
    except Exception as ex:
        return f"<div style=\"padding:20px;color:#C62828\">Erro ao gerar mapa: {ex}</div>"


def _mini_insights_distribuicao(pessoas: list) -> str:
    """Mini-insights para o grafico de distribuicao por funcionario (painel Mapeamento)."""
    if not pessoas:
        return ''
    tot = sum(p.get('n_registros', 0) for p in pessoas) or 1
    top_p = max(pessoas, key=lambda x: x.get('n_registros', 0))
    bot_p = min(pessoas, key=lambda x: x.get('n_registros', 0))
    media_p = tot / len(pessoas)
    items = [
        {'emoji': '', 'cor': 'green',
         'texto': f'<strong>{top_p["nome"]}</strong> tem mais OSs mapeadas: {top_p.get("n_registros",0):,} ({top_p.get("n_registros",0)/tot*100:.1f}% do total)'},
        {'emoji': '', 'cor': '',
         'texto': f'<strong>Media por funcionario: {media_p:.0f} OSs</strong> entre os {len(pessoas)} mapeados ({tot:,} no total)'},
        {'emoji': '', 'cor': 'orange' if bot_p.get("n_registros",0) < media_p * 0.4 else '',
         'texto': f'<strong>{bot_p["nome"]}</strong> tem menos OSs: {bot_p.get("n_registros",0):,} ({bot_p.get("n_registros",0)/tot*100:.1f}%)'},
    ]
    return _mini_insights_html(items)

def _build_panel_mapeamento(tem_combo, d_combo, _combo_total, _combo_pessoas, _combo_sem, _combo_pontos, _mapa_folium_html, dist_func_L_list, dist_func_Q_list, _hid_retroativo=None):
        if not tem_combo:
            return ''
        _hid_por_tipo   = d_combo.get('hidrometro_por_tipo', [])
        _top_tipos_os   = d_combo.get('top_tipos_os', [])
        _leit_os        = d_combo.get('leiturista_os', [])

        _stats         = d_combo.get('stats_cruzamento', {})
        _tipos_os      = _stats.get('tipos_os', [])
        _status_os     = _stats.get('status_os', [])
        _criticas      = _stats.get('criticas', [])
        _n_cob_critica = _stats.get('n_cobranca_critica', 0)

        h = []
        h.append('<div class="tab-panel" id="panel-mapeamento">')

        # ── KPIs ────────────────────────────────────────────────────────────
        h.append('<div class="kpi-grid">')
        _status_abertos = ('ABERTA', 'ABERTO', 'PRODUCAO', 'PRODUÇÃO', 'EM PRODUCAO', 'EM PRODUÇÃO')
        _n_os_abertas = sum(
            s['qtd']
            for s in _status_os
            if any(k in (s.get('status') or '').upper() for k in _status_abertos)
        )
        h.append(f'<div class="kpi-card blue" data-delay="0"><div class="kpi-label">OSs Abertas (COM)</div><div class="kpi-value">{_n_os_abertas:,}</div><div class="kpi-sub">ordens abertas da equipe COM-*</div></div>')
        h.append(f'<div class="kpi-card teal" data-delay="80"><div class="kpi-label">Funcionários Mapeados</div><div class="kpi-value">{len(_combo_pessoas)}</div><div class="kpi-sub">equipes COM-*</div></div>')
        _hid_top = _hid_por_tipo[0] if _hid_por_tipo else None
        if _hid_top:
            _hid_modelo_label = (_hid_top['modelo'] or '')[:3]
            h.append(f'<div class="kpi-card green" data-delay="160"><div class="kpi-label">Hidrômetro Problemático</div><div class="kpi-value">{_hid_top["total"]:,}</div><div class="kpi-sub">{_hid_modelo_label} · mais OSs abertas</div></div>')
        else:
            h.append(f'<div class="kpi-card green" data-delay="160"><div class="kpi-label">Hidrômetro Problemático</div><div class="kpi-value"></div><div class="kpi-sub">sem dados de hidrômetro</div></div>')
        _n_sem_coord_com = len(d_combo.get('os_sem_coordenada_com', []))
        if _n_sem_coord_com:
            h.append(f'<div class="kpi-card orange" data-delay="200"><div class="kpi-label">COM-* sem Coord</div><div class="kpi-value">{_n_sem_coord_com:,}</div><div class="kpi-sub">OSs COM-* sem lat/long na 50012</div></div>')
        if _n_cob_critica:
            h.append(f'<div class="kpi-card red" data-delay="240"><div class="kpi-label">Cobrança Crítica</div><div class="kpi-value">{_n_cob_critica:,}</div><div class="kpi-sub">OSs em cobrança avançada</div></div>')
        if _combo_sem:
            _n_sem_coord_com = len(d_combo.get('os_sem_coordenada_com', []))
            _sub_sem = f'{len(_combo_sem)} funcionário(s) sem OS'
            if _n_sem_coord_com:
                _sub_sem += f' · {_n_sem_coord_com} OSs COM-* sem coord'
            h.append(f'<div class="kpi-card red" data-delay="320"><div class="kpi-label">Sem Correspondência</div><div class="kpi-value">{len(_combo_sem)}</div><div class="kpi-sub">{_sub_sem}</div></div>')
        h.append('</div>')  # /kpi-grid

        # ── Grid: Tipos OS | Status OS | Criticidade 50012 ───────────────────
        if _tipos_os or _status_os or _criticas:
            h.append('<div class="grid-3" style="margin-bottom:22px">')

            h.append('</div>')  # /grid-3

        # ── Gráfico de pizza: Distribuição por Funcionário ───────────────────────────────────────────
        if _combo_pessoas:
            h.append(
                '<div class="chart-card" style="margin-bottom:22px">'
                '<div class="chart-header"><div>'
                '<div class="chart-title">Distribuição por Funcionário</div>'
                '<div class="chart-subtitle">OSs cruzadas (OS × lat/long 50012) - Análise de carga de trabalho</div>'
                '</div><span class="badge badge-green">Equipe-COM</span></div>'
                '<div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;padding:0 16px 16px">'
                # ── Coluna 1: Pizza ────────────────────────────────────
                '<div>'
                '<div style="position:relative;height:320px;margin-bottom:8px">'
                '<canvas id="chart-distribuicao-pie"></canvas>'
                '</div>'
                '<div style="font-size:0.8rem;color:var(--muted);text-align:center">Proporção de OSs</div>'
                '</div>'
                # ── Coluna 2: Barras ────────────────────────────────────
                '<div>'
                '<div style="position:relative;height:320px;margin-bottom:8px">'
                '<canvas id="chart-funcionarios-barras"></canvas>'
                '</div>'
                '<div style="font-size:0.8rem;color:var(--muted);text-align:center">Quantidade de OSs</div>'
                '</div>'
                '</div>'
                '<!-- charts de distribuição inicializados via buildChartDistribuicao() -->'
                '<div class="mini-insights-block">'
                + _mini_insights_distribuicao(_combo_pessoas) +
                '</div>'
                '</div>'
            )

        # ── Tabela de OSs cruzadas com DataTables (TODOS os registros) ────────────────────────────
        if _combo_pontos:
            _linhas = _combo_pontos  # Usar TODOS os registros (DataTables fará paginação)
            _has_num     = any(p.get('num_os')       for p in _linhas)
            _has_tipo    = any(p.get('tipo_servico')  for p in _linhas)
            _has_critica = any(p.get('critica')       for p in _linhas)
            _has_vol     = any(p.get('volume')        for p in _linhas)
            _has_cob     = any(p.get('cobranca')      for p in _linhas)

            _cor_st = {'ABERTA': '#2196F3', 'PENDENTE': '#FF9800', 'ENCERRADA': '#4CAF50',
                       'CANCELADA': '#9E9E9E', 'EM ATENDIMENTO': '#00BCD4'}

            def _badge_st(s):
                s = s or ''
                bg = _cor_st.get(s.upper(), '#607D8B')
                return f'<span style="background:{bg};color:#fff;padding:2px 7px;border-radius:4px;font-size:0.78em">{s or ""}</span>'

            def _badge_crit(c):
                c = c or ''
                if '5900' in c or 'NORMAL' in c.upper():
                    return f'<span style="color:#388E3C;font-size:0.82em"> Normal</span>'
                elif c:
                    return f'<span style="color:#E53935;font-weight:600;font-size:0.82em"> {c[:30]}</span>'
                return ''

            def _badge_cob(c):
                c = c or ''
                if any(kw in c.upper() for kw in ('COBRAN','JUDICI','PROTESTO','CORTE')):
                    return (f'<span style="background:#FFEBEE;color:#C62828;padding:2px 6px;'
                            f'border-radius:3px;font-size:0.78em">{c[:22]}</span>')
                return f'<span style="font-size:0.82em">{c or ""}</span>'

            th = 'padding:8px 10px;text-align:left;color:var(--muted);font-size:0.82em;white-space:nowrap'
            td = 'padding:8px 10px;font-size:0.83em;white-space:nowrap'

            cols = ['Matrícula', 'Funcionário', 'Bairro', 'Hidrômetro']
            if _has_num:     cols.append('N° OS')
            if _has_tipo:    cols.append('Tipo Serviço')
            cols.append('Status OS')
            if _has_critica: cols.append('Crítica (50012)')
            if _has_vol:     cols.append('Vol m³')
            if _has_cob:     cols.append('Cobrança')

            thead = ''.join(f'<th style="{th}">{c}</th>' for c in cols)
            rows_os = []
            for p in _linhas:
                _cob_raw = (p.get('cobranca') or '').upper()
                _bg = 'background:#FFF8F8' if any(kw in _cob_raw for kw in ('COBRAN','JUDICI','PROTESTO','CORTE')) else ''
                tds = [
                    f'<td style="{td}"><b>{p.get("matricula","")}</b></td>',
                    f'<td style="{td}">{p.get("equipe","") or ""}</td>',
                    f'<td style="{td}">{p.get("bairro","") or ""}</td>',
                    f'<td style="{td};font-family:monospace;font-weight:600;color:#1565C0">{(p.get("hidrometro","") or "")[:20]}</td>',
                ]
                if _has_num:     tds.append(f'<td style="{td}">{p.get("num_os","") or ""}</td>')
                if _has_tipo:    tds.append(f'<td style="{td};max-width:130px;overflow:hidden;text-overflow:ellipsis">{(p.get("tipo_servico","") or "")[:28]}</td>')
                tds.append(f'<td style="{td}">{_badge_st(p.get("status_os",""))}</td>')
                if _has_critica: tds.append(f'<td style="{td}">{_badge_crit(p.get("critica",""))}</td>')
                if _has_vol:     tds.append(f'<td style="{td};text-align:right">{p.get("volume","") or ""}</td>')
                if _has_cob:     tds.append(f'<td style="{td}">{_badge_cob(p.get("cobranca",""))}</td>')
                rows_os.append(f'<tr style="border-bottom:1px solid var(--border);{_bg}">{"".join(tds)}</tr>')

            _badge_cnt = (f'<span class="badge" style="background:#E3F2FD;color:#1565C0">'
                          f'{len(_combo_pontos):,} OSs COM-* mapeadas</span>')
            h.append(
                '<div class="chart-card" style="margin-bottom:22px">'
                '<div class="chart-header"><div>'
                '<div class="chart-title"> Ordens de Serviço × Ligações 50012</div>'
                '<div class="chart-subtitle">'
                'Cruzamento OS (ligação) ↔ 50012 (lat/long + analítico)  com filtros por coluna'
                f'</div></div>{_badge_cnt}</div>'
                '<div style="overflow-x:auto">'
                '<table id="tabela-os-cruzadas" class="dataTable" style="width:100%;border-collapse:collapse">'
                f'<thead><tr style="background:var(--surface);border-bottom:2px solid var(--border)">{thead}</tr></thead>'
                f'<tbody>{"".join(rows_os)}</tbody>'
                '</table></div>'
                '</div>'
            )

        # ── Mapa Folium ──────────────────────────────────────────────────────
        if _mapa_folium_html:
            mapa_escaped = _mapa_folium_html.replace('&', '&amp;').replace('"', '&quot;')
            iframe = (
                f'<iframe srcdoc="{mapa_escaped}" '
                'style="width:100%;height:600px;border:none;border-radius:8px;margin-top:8px">'
                '</iframe>'
            )
        else:
            iframe = (
                '<div style="padding:40px;text-align:center;color:var(--muted)">'
                '<div style="font-size:2rem;margin-bottom:8px"></div>'
                'Nenhuma coordenada válida encontrada para exibir no mapa.</div>'
            )
        
        # ── Legenda de cores das equipes ──────────────────────────────────────
        # Paleta sem vermelho (crítica) nem verde (normal) para evitar conflito visual
        CORES_LEGENDA = ['#1565C0', '#FF8F00', '#6A1B9A', '#1976D2', '#00838F',
                         '#AD1457', '#F57C00', '#5E35B1', '#0097A7', '#7B1FA2']
        
        # Extrai equipes únicas
        equipes_set = {p.get('equipe', 'Indefinida') for p in _combo_pontos if p.get('equipe')}
        equipes_sorted = sorted(list(equipes_set))
        
        # Monta legenda com 2 seções: Equipes à esquerda, Criticidade à direita
        legenda_items_equipes = []
        for i, eq in enumerate(equipes_sorted):
            cor = CORES_LEGENDA[i % len(CORES_LEGENDA)]
            legenda_items_equipes.append(
                f'<div style="display:flex;align-items:center;gap:6px">'
                f'<div style="width:12px;height:12px;border-radius:50%;background:{cor};border:2px solid #fff;box-shadow:0 0 3px rgba(0,0,0,0.2)"></div>'
                f'<span style="font-size:0.85rem;color:var(--text);font-weight:500">{eq}</span>'
                f'</div>'
            )
        
        legenda_html = (
            '<div style="margin-bottom:12px;padding:12px 16px;background:var(--surface);border-radius:8px;border:1px solid var(--border)">'
            '<div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">'
            # Seção ESQUERDA: Equipes
            f'<div style="display:flex;flex-wrap:wrap;gap:10px;align-content:flex-start">'
            f'{"".join(legenda_items_equipes)}'
            f'</div>'
            # Seção DIREITA: Criticidade
            f'<div style="display:flex;flex-direction:column;gap:8px;justify-content:center">'
            f'<div style="display:flex;align-items:center;gap:6px">'
            f'<div style="width:12px;height:12px;border-radius:50%;background:#2E7D32;border:2px solid #fff;box-shadow:0 0 3px rgba(0,0,0,0.2)"></div>'
            f'<span style="font-size:0.85rem;color:var(--text);font-weight:500"> Leitura Normal</span>'
            f'</div>'
            f'<div style="display:flex;align-items:center;gap:6px">'
            f'<div style="width:12px;height:12px;border-radius:50%;background:#E53935;border:2px solid #fff;box-shadow:0 0 3px rgba(0,0,0,0.2)"></div>'
            f'<span style="font-size:0.85rem;color:var(--text);font-weight:500"> Leitura Crítica</span>'
            f'</div>'
            f'</div>'
            f'</div>'
            f'</div>'
        )
        
        h.append(
            '<div class="chart-card" style="margin-bottom:22px">'
            '<div class="chart-header"><div>'
            '<div class="chart-title">Mapa Geográfico das Equipes</div>'
            '<div class="chart-subtitle">'
            'Distribuição espacial das OSs  clique em um ponto para detalhes completos'
            '</div></div>'
            '<span class="badge" style="background:#E0F2F1;color:#00695C"> OS × 50012</span></div>'
            f'{legenda_html}'
            f'{iframe}'
            '</div>'
        )


        # ── Gráfico: Hidrômetro × Tipo de OS (espelho) ───────────────────────
        if _hid_por_tipo:
            n_mod = len(_hid_por_tipo)
            # rowH=36, PAD_T=40, PAD_B=40 (deve bater com o JS)
            altura_hid = n_mod * 36 + 80 + 20  # margem extra p/ legenda
            _mini_hid_tipos = ''
            if _tipos_os:
                _tot_tipos = sum(x['qtd'] for x in _tipos_os) or 1
                _mini_hid_tipos = _mini_insights_html([
                    {'emoji': '', 'cor': 'green',
                     'texto': f'<strong>Tipo OS mais comum: {_tipos_os[0]["tipo"]}</strong> -- {_tipos_os[0]["qtd"]:,} ordens ({_tipos_os[0]["qtd"]/_tot_tipos*100:.1f}% do total)'},
                    {'emoji': '', 'cor': '',
                     'texto': f'<strong>Media por tipo: {_tot_tipos//len(_tipos_os):,} OSs</strong> entre os {len(_tipos_os)} tipos distintos identificados'},
                    {'emoji': '', 'cor': '',
                     'texto': f'<strong>Tipo menos frequente: {_tipos_os[-1]["tipo"]}</strong> com {_tipos_os[-1]["qtd"]:,} ordens ({_tipos_os[-1]["qtd"]/_tot_tipos*100:.1f}%)'},
                ])
            h.append(
                '<div class="chart-card" style="margin-bottom:22px">'
                '<div class="chart-header"><div>'
                '<div class="chart-title"> Ordens de Serviço por Modelo de Hidrômetro</div>'
                '<div class="chart-subtitle">Lado esquerdo: total de OSs | Lado direito: tipos de OS por modelo (nome dentro da barra)</div>'
                '</div><span class="badge" style="background:#E3F2FD;color:#1565C0">Hidrômetro × OS</span></div>'
                '<div style="position:relative;width:100%;overflow:hidden">'
                '<canvas id="chart-hid-espelho" style="display:block;width:100%"></canvas>'
                '</div>'
                + ('<div class="mini-insights-block">' + _mini_hid_tipos + '</div>' if _mini_hid_tipos else '')
                + '</div>'
            )

        # ── Gráfico: Hidrômetro retroativo (Histórico Completo) ───────────────
        if _hid_retroativo:
            _n_meses_ret   = len(list(dict.fromkeys(r['mes'] for r in _hid_retroativo))) or 12
            _n_modelos_ret = len(list(dict.fromkeys(r['modelo'] for r in _hid_retroativo))) or 1
            # Uma linha por modelo (agregado no período), altura baseada no nº de modelos
            h.append(
                '<div class="chart-card" style="margin-bottom:22px">'
                '<div class="chart-header"><div>'
                '<div class="chart-title"> OSs de Hidrômetro  Histórico Completo</div>'
                '<div class="chart-subtitle">Substituições corretivas e preventivas de todo o período</div>'
                '</div><span class="badge" style="background:#E8F5E9;color:#2E7D32">Histórico Total</span></div>'
                '<div style="position:relative;width:100%;overflow:hidden">'
                '<canvas id="chart-hid-retroativo" style="display:block;width:100%"></canvas>'
                '</div>'
                '</div>'
            )

        # ── Gráfico: Leiturista × Total de OSs ───────────────────────────────
        # ── Tabela: OSs SEM coordenada ───────────────────────────────────────
        _os_sem_coord     = d_combo.get('os_sem_coordenada', [])
        _os_sem_com       = d_combo.get('os_sem_coordenada_com', [])
        _os_sem_outras    = d_combo.get('os_sem_coordenada_outras', [])
        # fallback: se engine antigo (sem separação), usa lista geral
        if not _os_sem_com and not _os_sem_outras and _os_sem_coord:
            _os_sem_com = [r for r in _os_sem_coord if r.get('LEITURISTA','').upper().startswith('COM-')]
            _os_sem_outras = [r for r in _os_sem_coord if not r.get('LEITURISTA','').upper().startswith('COM-')]

        def _tabela_sem_coord(lista, table_id, titulo, badge_cor_bg, badge_cor_txt, subtitulo):
            if not lista:
                return ''
            _has_tipo_sc  = any(p.get('TIPO_SERVICO') for p in lista)
            _has_log_sc   = any(p.get('LOGRADOURO')   for p in lista)
            _has_prazo_sc = any(p.get('PRAZO_OS')      for p in lista)
            _has_prio_sc  = any(p.get('PRIORIDADE')    for p in lista)

            cols_sc = ['Matrícula', 'Funcionário']
            if _has_tipo_sc:  cols_sc.append('Tipo Serviço')
            cols_sc.append('Status OS')
            cols_sc.append('Data OS')
            if _has_prazo_sc: cols_sc.append('Prazo')
            if _has_prio_sc:  cols_sc.append('Prioridade')
            if _has_log_sc:   cols_sc.append('Logradouro')
            cols_sc.append('N° Imóvel')
            cols_sc.append('N° OS')

            thead_sc = ''.join(f'<th style="{th}">{c}</th>' for c in cols_sc)
            rows_sc  = []
            for p in lista:
                tds = [
                    f'<td style="{td}"><b>{p.get("MATRICULA","")}</b></td>',
                    f'<td style="{td}">{p.get("LEITURISTA","") or ""}</td>',
                ]
                if _has_tipo_sc:  tds.append(f'<td style="{td};max-width:140px;overflow:hidden;text-overflow:ellipsis">{(p.get("TIPO_SERVICO","") or "")[:30]}</td>')
                tds.append(f'<td style="{td}">{_badge_st(p.get("STATUS_OS",""))}</td>')
                tds.append(f'<td style="{td}">{p.get("DATA_OS","") or ""}</td>')
                if _has_prazo_sc: tds.append(f'<td style="{td}">{p.get("PRAZO_OS","") or ""}</td>')
                if _has_prio_sc:  tds.append(f'<td style="{td}">{p.get("PRIORIDADE","") or ""}</td>')
                if _has_log_sc:   tds.append(f'<td style="{td}">{(p.get("LOGRADOURO","") or "")[:28]}</td>')
                tds.append(f'<td style="{td}">{p.get("NR_IMOVEL","") or ""}</td>')
                tds.append(f'<td style="{td}">{p.get("NUM_OS","") or ""}</td>')
                rows_sc.append(f'<tr style="border-bottom:1px solid var(--border)">{"".join(tds)}</tr>')

            _badge = (f'<span class="badge" style="background:{badge_cor_bg};color:{badge_cor_txt}">'
                      f'{len(lista):,} sem coordenada</span>')
            return (
                '<div class="chart-card" style="margin-bottom:22px">'
                '<div class="chart-header"><div>'
                f'<div class="chart-title">{titulo}</div>'
                f'<div class="chart-subtitle">{subtitulo}</div>'
                f'</div>{_badge}</div>'
                '<div style="overflow-x:auto">'
                f'<table id="{table_id}" class="dataTable" style="width:100%;border-collapse:collapse">'
                f'<thead><tr style="background:var(--surface);border-bottom:2px solid var(--border)">{thead_sc}</tr></thead>'
                f'<tbody>{"".join(rows_sc)}</tbody>'
                '</table></div>'
                '</div>'
            )

        h.append(_tabela_sem_coord(
            _os_sem_com,
            'tabela-os-sem-coord-com',
            ' OSs COM-* sem Coordenada na 50012',
            '#FFF3E0', '#E65100',
            'Equipes COM-* com OS ativa mas sem lat/long correspondente na macro 50012  '
            'estas OSs não aparecem no mapa e não foram distribuídas aos funcionários'
        ))
        h.append(_tabela_sem_coord(
            _os_sem_outras,
            'tabela-os-sem-coord-outras',
            'OSs de Outras Equipes sem Coordenada',
            '#F3F4F6', '#6B7280',
            'OSs de equipes não-COM-* sem lat/long na 50012  informativo'
        ))

        h.append('</div><!-- /panel-mapeamento -->')
        return '\n'.join(h)


def formatar_payload_bruta(payload_bruta, periodo: str = None, cidade: str = None) -> dict:
    """
    Converte payload bruta (pura, sem filtro) para formato esperado por gerar_html.
    
    A payload pode ser:
    - Dicionário Python direto
    - String JSON
    - Qualquer estrutura que contenha dados brutos dos macros
    - Raw GWT-RPC pipe-separated data (formato --data-raw $'...')
    
    Normaliza automaticamente:
    - Nomes de chaves (macro_8091, macro_8121, macro_8117, macro_50012, combo_mapa)
    - Períodos e datas ausentes
    - Estrutura de dados incompleta
    - Raw GWT-RPC data
    
    Args:
        payload_bruta : dict ou str JSON com dados brutos, ou raw GWT-RPC data
        periodo       : período (ex: '03-2026' ou '03/2026'). Se None, usa data atual
        cidade        : cidade (ex: 'ITAPOÁ'). Se None, usa 'ITAPOÁ' como padrão
    
    Returns:
        dict : dados formatados prontos para gerar_html()
    
    Exemplos:
        # JSON bruto do arquivo/API
        payload_json = '{...dados...}'
        dados_formatos = formatar_payload_bruta(payload_json, '03-2026', 'ITAPOÁ')
        resultado = gerar_html(dados_formatados, 'dashboard.html')
        
        # Dict bruto
        payload_dict = {'macro_8091': {...}, ...}
        dados_formatados = formatar_payload_bruta(payload_dict)
        resultado = gerar_html(dados_formatados, 'dashboard.html')
        
        # Raw GWT-RPC (pipe-separated)
        raw_gwt = '7|0|527|https://...|...'
        dados_formatados = formatar_payload_bruta({'macro_8091_raw': raw_gwt})
    """
    from datetime import datetime as dt
    
    # 1. Converter payload bruta para dict se necessário
    if isinstance(payload_bruta, str):
        try:
            payload_dict = json.loads(payload_bruta)
        except json.JSONDecodeError as e:
            raise ValueError(f"Payload bruta inválida (não é JSON válido): {e}")
    elif isinstance(payload_bruta, dict):
        payload_dict = payload_bruta.copy()
    else:
        raise TypeError(f"Payload deve ser dict ou string JSON. Recebido: {type(payload_bruta)}")
    
    # 2. Definir período padrão se não fornecido
    if periodo is None:
        periodo = dt.now().strftime('%m-%Y')
    else:
        # Normalizar formato do período (aceita '03-2026' ou '03/2026')
        periodo = str(periodo).replace('/', '-')
    
    # 3. Definir cidade padrão
    if cidade is None:
        cidade = 'ITAPOÁ'
    else:
        cidade = str(cidade).strip().upper()
    
    # 4. Criar estrutura base do dados formatado
    dados_formatados = {
        'periodo': {
            'inicial': periodo,
            'final': periodo
        },
        'cidade': cidade,
    }
    
    # 5. Mapear chaves possíveis da payload bruta para formato padrão
    mapa_chaves = {
        'macro_8091': ['macro_8091', '8091', 'macro8091', 'dados_8091', 'data_8091', 'macro_8091_raw'],
        'macro_8117': ['macro_8117', '8117', 'macro8117', 'dados_8117', 'data_8117', 'macro_8117_raw'],
        'macro_8121': ['macro_8121', '8121', 'macro8121', 'dados_8121', 'data_8121', 'macro_8121_raw'],
        'macro_50012': ['macro_50012', '50012', 'macro50012', 'dados_50012', 'data_50012', 'macro_50012_raw'],
        'combo_mapa': ['combo_mapa', 'mapa', 'mapeamento', 'dados_mapa'],
    }
    
    # 6. Processar cada macro: procurar pela chave e copiar dados
    for chave_padrao, chaves_possiveis in mapa_chaves.items():
        for chave_busca in chaves_possiveis:
            if chave_busca in payload_dict:
                valor = payload_dict[chave_busca]
                if valor:  # Qualquer valor não-vazio é aceito (dict ou string raw)
                    if isinstance(valor, dict) and len(valor) > 0:
                        dados_formatados[chave_padrao] = valor
                        break
                    elif isinstance(valor, str):
                        # Raw GWT-RPC ou similiar - cria estrutura minimal
                        dados_formatados[chave_padrao] = {'_raw_data': valor}
                        break
    
    # 7. Garantir que período está em todos os macros (se forem dicts)
    for chave in ['macro_8091', 'macro_8117', 'macro_8121', 'macro_50012']:
        if chave in dados_formatados:
            if isinstance(dados_formatados[chave], dict):
                if 'periodo' not in dados_formatados[chave]:
                    dados_formatados[chave]['periodo'] = {
                        'inicial': periodo,
                        'final': periodo
                    }
    
    # 8. Validação básica - aceita qualquer macro válido (mesmo se raw)
    if 'macro_8091' not in dados_formatados and 'macro_8121' not in dados_formatados and \
       'macro_8117' not in dados_formatados and 'macro_50012' not in dados_formatados and \
       'combo_mapa' not in dados_formatados:
        raise ValueError("Payload bruta não contém nenhum macro válido (8091, 8117, 8121, 50012, combo_mapa)")
    
    return dados_formatados


def gerar_html_com_payload_bruta(payload_bruta, caminho_saida: str, periodo: str = None, 
                                  cidade: str = None, incluir_pn: bool = False) -> str:
    """
    Wrapper completo que formata payload bruta e gera HTML em um passo.
    
    Equivalente a:
        dados = formatar_payload_bruta(payload_bruta, periodo, cidade)
        return gerar_html(dados, caminho_saida, incluir_pn)
    
    Args:
        payload_bruta : dict ou string JSON com dados brutos
        caminho_saida : caminho do arquivo HTML a gerar
        periodo       : período (ex: '03-2026'). Se None, usa data atual
        cidade        : cidade. Se None, usa 'ITAPOÁ'
        incluir_pn    : se True, inclui aba de PN Financeiro
    
    Returns:
        str : caminho do arquivo gerado
    """
    dados_formatados = formatar_payload_bruta(payload_bruta, periodo, cidade)
    return gerar_html(dados_formatados, caminho_saida, incluir_pn)


def gerar_html(dados: dict, caminho_saida: str, incluir_pn: bool = False) -> str:
    periodo  = dados['periodo']
    cidade   = dados['cidade']
    tem_pn_aba = bool(
        incluir_pn or
        dados.get('pn_html_arquivo') or
        dados.get('macro_8117_pn') or
        dados.get('macro_8121_pn')
    )

    pn_html_arquivo = dados.get('pn_html_arquivo')
    if not pn_html_arquivo:
        _nome_dashboard = os.path.basename(caminho_saida)
        if _nome_dashboard.startswith('dashboard_'):
            pn_html_arquivo = 'PN_Financeiro_' + _nome_dashboard[len('dashboard_'):]
        else:
            pn_html_arquivo = 'PN_Financeiro.html'

    macros = []
    if 'macro_8117' in dados: macros.append('8117')
    if 'macro_8121' in dados: macros.append('8121')
    if 'macro_8091' in dados: macros.append('8091')
    if 'macro_50012' in dados: macros.append('50012')
    if 'combo_mapa'  in dados: macros.append('Mapeamento')
    if tem_pn_aba: macros.append('PN')

    tem_8117      = 'macro_8117' in dados
    tem_8121      = 'macro_8121' in dados
    tem_8091      = 'macro_8091' in dados
    tem_50012     = 'macro_50012' in dados
    tem_combo     = 'combo_mapa'  in dados
    d_combo       = dados.get('combo_mapa', {})

    # Gera mapa Folium para aba de mapeamento
    _combo_pontos       = d_combo.get('pontos_mapa', [])
    _combo_centro       = d_combo.get('centro_geografico', {})
    _combo_pessoas      = d_combo.get('pessoas', [])
    _combo_total        = d_combo.get('total_registros', 0)
    _combo_sem          = d_combo.get('sem_correspondencia', [])
    _hid_por_tipo       = d_combo.get('hidrometro_por_tipo', [])
    _top_tipos_os       = d_combo.get('top_tipos_os', [])
    _leit_os            = d_combo.get('leiturista_os', [])
    _hid_retroativo     = d_combo.get('hid_retroativo', [])
    _mapa_folium_html = _gerar_mapa_folium(_combo_pontos, _combo_centro) if tem_combo and _combo_pontos else ''
    
    # Dados para gráfico de pizza: Distribuição por Funcionário (desde combo_mapa)
    dist_func_L_list = [p.get('nome', '?') for p in _combo_pessoas]
    dist_func_Q_list = [p.get('n_registros', 0) for p in _combo_pessoas]

    # ── Gera HTML do painel Mapeamento em Python puro (sem f-string aninhada) ──
    

    _html_panel_mapeamento = _build_panel_mapeamento(tem_combo, d_combo, _combo_total, _combo_pessoas, _combo_sem, _combo_pontos, _mapa_folium_html, dist_func_L_list, dist_func_Q_list, _hid_retroativo)

    d17 = dados.get('macro_8117', {})
    d21 = dados.get('macro_8121', {})
    d91 = dados.get('macro_8091', {})
    d50 = dados.get('macro_50012', {})

    _empty = {'datas': [], 'arr_diario': [], 'arr_acumulado': [],
              'fat_diario': [], 'fat_acumulado': [], 'formas': {}}
    if not d17: d17 = _empty
    if not d21: d21 = _empty
    if not d91: d91 = {}

    # ── Dados 8091 ──────────────────────────────────────────────────────────
    kpis91        = d91.get('kpis', {})
    total_91      = d91.get('total', kpis91.get('total_fat', 0.0))
    total_91_fmt  = _fmt_brl(total_91)
    meses_count   = d91.get('meses_count', 1)

    sit91    = d91.get('situacao', [])
    cat91    = d91.get('categoria', [])
    fase91   = d91.get('fase', [])
    emis91   = d91.get('emissao', [])
    bairros91= d91.get('bairros', [])
    grupos91 = d91.get('grupos', [])
    faixas91 = d91.get('faixas', {})
    comp91   = d91.get('comp', {})
    serie_meses  = d91.get('serie_meses', [])

    k91_fat   = kpis91.get('total_fat',    total_91)
    k91_agua  = kpis91.get('total_agua',   0.0)
    k91_mult  = kpis91.get('total_multa',  0.0)
    k91_lig   = kpis91.get('n_ligacoes',   0)
    k91_vol   = kpis91.get('vol_total',    0.0)
    k91_media = kpis91.get('media_fatura', 0.0)

    sit91_L   = _safe_json([i['label'] for i in sit91])
    sit91_Q   = _safe_json([i['qtd']   for i in sit91])
    sit91_V   = _safe_json([i['valor'] for i in sit91])
    cat91_L   = _safe_json([i['label'] for i in cat91])
    cat91_Q   = _safe_json([i['qtd']   for i in cat91])
    cat91_V   = _safe_json([i['valor'] for i in cat91])
    fase91_L  = _safe_json([i['label'] for i in fase91])
    fase91_Q  = _safe_json([i['qtd']   for i in fase91])
    fase91_V  = _safe_json([i['valor'] for i in fase91])
    emis91_L  = _safe_json([i['data']  for i in emis91])
    emis91_V  = _safe_json([i['valor'] for i in emis91])
    bar91_L   = _safe_json([i['label'] for i in bairros91])
    bar91_V   = _safe_json([i['valor'] for i in bairros91])
    bar91_Q   = _safe_json([i['qtd']   for i in bairros91])
    grp91_L   = _safe_json([i['label'] for i in grupos91])
    grp91_V   = _safe_json([i['valor'] for i in grupos91])
    grp91_VOL = _safe_json([i['vol']   for i in grupos91])
    faixas91_L= _safe_json(list(faixas91.keys()))
    faixas91_V= _safe_json(list(faixas91.values()))

    # Novos dados
    faixas_vol91   = d91.get('faixas_vol', {})
    leiturista91   = d91.get('leiturista', [])
    consumo_zero91 = d91.get('consumo_zero', [])
    fvol91_L  = _safe_json(list(faixas_vol91.keys()))
    fvol91_V  = _safe_json(list(faixas_vol91.values()))
    leit91_L  = _safe_json([i['label']     for i in leiturista91])
    leit91_Q  = _safe_json([i['qtd']       for i in leiturista91])
    leit91_VOL= _safe_json([i['vol']       for i in leiturista91])
    leit91_VM = _safe_json([i['vol_medio'] for i in leiturista91])
    czero91_L = _safe_json([i['label'] for i in consumo_zero91])
    czero91_Q = _safe_json([i['qtd']   for i in consumo_zero91])

    leit_dia91     = d91.get('leit_dia', [])
    grupos_receita91 = d91.get('grupos_receita', [])
    ldia91_L  = _safe_json([i['data'] for i in leit_dia91])
    ldia91_Q  = _safe_json([i['qtd']  for i in leit_dia91])
    ldia91_V  = _safe_json([i['vol']  for i in leit_dia91])
    grp_r91_L = _safe_json([i['label']      for i in grupos_receita91])
    grp_r91_RM= _safe_json([i['receita_m3'] for i in grupos_receita91])
    grp_r91_V = _safe_json([i['valor']      for i in grupos_receita91])
    grp_r91_VOL=_safe_json([i['vol']        for i in grupos_receita91])
    comp91_L  = _safe_json(['Água', 'Multas', 'Outros'])
    comp91_V  = _safe_json([round(comp91.get('agua',0),2), round(comp91.get('multa',0),2), round(comp91.get('outros',0),2)])
    serie_labels = _safe_json([s['periodo'] for s in serie_meses])
    serie_totais = _safe_json([s['total']   for s in serie_meses])

    # Tabela bairros
    _bairros_rows = ''
    _max_b = max((i['valor'] for i in bairros91), default=1)
    for idx, b in enumerate(bairros91):
        pct = int(b['valor'] / _max_b * 80)
        _bairros_rows += (
            f'<tr><td style="color:var(--muted);font-weight:700">{idx+1}</td>'
            f'<td><div style="display:flex;align-items:center;gap:8px">'
            f'<div style="height:6px;border-radius:3px;background:linear-gradient(90deg,var(--navy),#42A5F5);width:{pct}px;min-width:4px"></div>'
            f'{b["label"]}</div></td>'
            f'<td style="color:var(--muted)">{b["qtd"]:,}</td>'
            f'<td>R$\xa0{b["valor"]:,.2f}</td></tr>'.replace(',', 'X').replace('.', ',').replace('X', '.')
        )
    _total_top10 = sum(b['valor'] for b in bairros91)
    _bairros_rows += f'<tr><td></td><td>TOTAL TOP 10</td><td></td><td>R$\xa0{_total_top10:,.2f}</td></tr>'.replace(',','X').replace('.',',').replace('X','.')

    # Legado
    contas91      = d91.get('contas', {})
    contas_sorted = sorted(contas91.items(), key=lambda x: x[1]['valor'], reverse=True)
    contas_nomes  = _safe_json([c[0] for c in contas_sorted])
    contas_vals   = _safe_json([c[1]['valor'] for c in contas_sorted])

    ins        = _calcular_insights(d17, d21)
    ins91      = _calcular_insights_8091(d91) if tem_8091 else []
    mini91     = _calcular_mini_insights_8091(d91) if tem_8091 else {}
    mini17     = _calcular_mini_insights_8117(d17) if tem_8117 else {}
    mini21     = _calcular_mini_insights_8121(d17, d21) if tem_8121 else {}
    mini50     = _calcular_mini_insights_50012(d50) if tem_50012 else {}
    gerado_em  = datetime.now().strftime('%d/%m/%Y às %H:%M')

    macros_badges = ''.join([
        f'<span style="background:rgba(255,255,255,0.15);border:1px solid rgba(255,255,255,0.25);'
        f'border-radius:6px;padding:3px 10px;font-size:0.75rem;font-weight:600;color:#fff">'
        f'Macro {m}</span> '
        for m in macros
    ])

    datas_js      = _safe_json(d17.get('datas', []))
    arr_d_js      = _safe_json(d17.get('arr_diario', []))
    arr_ac_js     = _safe_json(d17.get('arr_acumulado', []))
    fat_d_js      = _safe_json(d21.get('fat_diario', []))
    fat_ac_js     = _safe_json(d21.get('fat_acumulado', []))
    datas_21_js   = _safe_json(d21.get('datas', []))
    # Arrecadação própria da 8121 (independente da 8117)
    arr21_d_js    = _safe_json(d21.get('arr_diario', []))
    arr21_ac_js   = _safe_json(d21.get('arr_acumulado', []))

    # Período anterior da 8121 (quando disponível  múltiplos CSVs)
    d21_ant        = dados.get('macro_8121_anterior', {})
    fat_ac_ant_js  = _safe_json(d21_ant.get('fat_acumulado', []))
    arr_ac_ant_js  = _safe_json(d21_ant.get('arr_acumulado', []))
    dias_mes_ant_js= _safe_json(d21_ant.get('dias_mes', []))
    dias_mes_js    = _safe_json(d21.get('dias_mes', []))
    label_ant      = _safe_json(d21_ant.get('periodo_label', ''))
    label_atual    = _safe_json(d21.get('periodo_label', ''))

    formas       = d17.get('formas', {})

    formas_nomes = _safe_json(list(formas.keys()))
    formas_vals  = _safe_json([v['valor'] for v in formas.values()])
    formas_qtde  = _safe_json([v['qtde'] for v in formas.values()])

    paleta = ['#1565C0','#FF8F00','#2E7D32','#6A1B9A','#00838F',
              '#D84315','#37474F','#558B2F','#AD1457','#0277BD']
    formas_cores = _safe_json(paleta[:len(formas)])

    # ── Dados 50012 (Latitude/Longitude + Analítico) ────────────────────────
    d50_total      = d50.get('total_registros', 0)
    d50_com_coord  = d50.get('com_coordenadas', 0)
    d50_sem_coord  = d50.get('sem_coordenadas', 0)
    d50_cobertura  = d50.get('cobertura_pct', 0.0)
    
    d50_equipes    = d50.get('equipes', [])
    d50_bairros    = d50.get('bairros_top15', [])
    d50_centro     = d50.get('centro_geografico', {'latitude': None, 'longitude': None})
    d50_bbox       = d50.get('bbox_geografico', {})

    # Analíticos novos
    kpis50         = d50.get('kpis', {})
    criticas50     = d50.get('criticas', [])
    faixas50       = d50.get('faixas_consumo', {})
    leituristas50  = d50.get('leituristas', [])
    grupos50       = d50.get('grupos', [])
    cobranca50     = d50.get('cobranca', [])
    bairros_fat50  = d50.get('bairros_fat', [])
    desvio50       = d50.get('desvio', [])
    leit_dia50     = d50.get('leit_dia', [])
    categorias50   = d50.get('categorias', [])
    hidrometros50  = d50.get('hidrometros', [])

    # KPIs analíticos
    k50_tot_lig  = kpis50.get('total_ligacoes', d50_total)
    k50_vol      = kpis50.get('total_vol', 0)
    k50_fat      = kpis50.get('total_fat', 0)
    k50_media_v  = kpis50.get('media_vol', 0)
    k50_pct_norm = kpis50.get('pct_normal', 0)
    k50_pct_crit = kpis50.get('pct_critico', 0)
    k50_zeros    = kpis50.get('zeros_vol', 0)
    k50_aumento  = kpis50.get('n_aumento', 0)
    k50_leit_ant = kpis50.get('n_leit_ant', 0)
    k50_queda    = kpis50.get('n_queda', 0)
    k50_nao_real = kpis50.get('n_nao_realizada', 0)
    k50_hd_inv   = kpis50.get('n_hd_invertido', 0)

    # JSON para gráficos geográficos (existentes)
    equipesL_50 = _safe_json([eq.get('equipe', '?') for eq in d50_equipes])
    equipesQ_50 = _safe_json([eq.get('total', 0) for eq in d50_equipes])
    equipesC_50 = _safe_json([eq.get('cobertura_pct', 0) for eq in d50_equipes])
    bairrosL_50 = _safe_json([b.get('bairro', '?') for b in d50_bairros])
    bairrosQ_50 = _safe_json([b.get('total', 0) for b in d50_bairros])
    centro_lat_50 = d50_centro.get('latitude', 0) if d50_centro.get('latitude') else -26.12
    centro_lng_50 = d50_centro.get('longitude', 0) if d50_centro.get('longitude') else -48.60
    bbox_50 = _safe_json(d50_bbox)

    # JSON analíticos novos
    crit50_L  = _safe_json([c['label'] for c in criticas50])
    crit50_Q  = _safe_json([c['qtd']   for c in criticas50])
    fxc50_L   = _safe_json(list(faixas50.keys()))
    fxc50_Q   = _safe_json(list(faixas50.values()))
    leit50_L  = _safe_json([l['label']    for l in leituristas50])
    leit50_Q  = _safe_json([l['qtd']      for l in leituristas50])
    leit50_CR = _safe_json([l['criticos'] for l in leituristas50])
    dist_func_L = _safe_json(dist_func_L_list)
    dist_func_Q = _safe_json(dist_func_Q_list)

    # ── Hidrômetro × Tipo OS e Leiturista OS ────────────────────────────
    # Modelos (labels) e totais para lado esquerdo
    hid_modelos_js   = _safe_json([h['modelo'] for h in _hid_por_tipo])
    hid_totais_js    = _safe_json([h['total']   for h in _hid_por_tipo])
    # Para cada modelo, lista dos top tipos com qtd (para lado direito)
    hid_tipos_raw_js = _safe_json([
        [{'tipo': t, 'qtd': q} for t, q in list(h['tipos'].items())[:6]]
        for h in _hid_por_tipo
    ])
    top_tipos_os_js  = _safe_json(_top_tipos_os)
    # Retroativo 12 meses  hidrômetros
    # Passa dados brutos com mes/modelo/tipo/qtd; toda a agregação por modelo
    # é feita no JS para que o slider de período funcione dinamicamente.
    _hid_ret_meses = list(dict.fromkeys(r['mes'] for r in _hid_retroativo))
    hid_ret_meses_js  = _safe_json(_hid_ret_meses)
    hid_ret_series_js = _safe_json(_hid_retroativo)  # lista de {mes, modelo, tipo, qtd}
    # Leiturista
    leit_os_labels_js = _safe_json([l['leiturista'] for l in _leit_os])
    leit_os_totais_js = _safe_json([l['total']      for l in _leit_os])
    
    # Dados para gráfico de pizza: Distribuição por Funcionário
    leit50_NO = _safe_json([l['normais']  for l in leituristas50])
    leit50_TC = _safe_json([l['taxa_crit'] for l in leituristas50])
    leit50_V  = _safe_json([l['vol']      for l in leituristas50])
    grp50_L   = _safe_json([g['label'] for g in grupos50])
    grp50_V   = _safe_json([g['vol']   for g in grupos50])
    grp50_F   = _safe_json([g['valor'] for g in grupos50])
    cob50_L   = _safe_json([c['label'] for c in cobranca50])
    cob50_Q   = _safe_json([c['qtd']   for c in cobranca50])
    brf50_L   = _safe_json([b['bairro'] for b in bairros_fat50])
    brf50_V   = _safe_json([b['valor']  for b in bairros_fat50])
    brf50_Q   = _safe_json([b['qtd']    for b in bairros_fat50])
    dev50_L   = _safe_json([d['label'] for d in desvio50])
    dev50_Q   = _safe_json([d['qtd']   for d in desvio50])
    dia50_L   = _safe_json([d['data'] for d in leit_dia50])
    dia50_Q   = _safe_json([d['qtd']  for d in leit_dia50])
    dia50_V   = _safe_json([d['vol']  for d in leit_dia50])
    cat50_L   = _safe_json([c['label'] for c in categorias50])
    cat50_Q   = _safe_json([c['qtd']   for c in categorias50])
    hid50_L   = _safe_json([h['modelo'] for h in hidrometros50])
    hid50_Q   = _safe_json([h['qtd']   for h in hidrometros50])

    # Monta tabs habilitadas
    tabs_js_list = []
    if tem_8117:  tabs_js_list.append('"8117"')
    if tem_8121:  tabs_js_list.append('"8121"')
    if tem_8091:  tabs_js_list.append('"8091"')
    if tem_50012: tabs_js_list.append('"50012"')
    if tem_combo: tabs_js_list.append('"mapeamento"')
    if tem_pn_aba: tabs_js_list.append('"pn"')
    tabs_js = '[' + ','.join(tabs_js_list) + ']'

    # Default tab: primeira disponível (prioridade: 8117 > 8121 > 8091 > 50012 > mapeamento)
    if tem_8117:
        default_tab = '8117'
    elif tem_8121:
        default_tab = '8121'
    elif tem_8091:
        default_tab = '8091'
    elif tem_50012:
        default_tab = '50012'
    elif tem_combo:
        default_tab = 'mapeamento'
    elif tem_pn_aba:
      default_tab = 'pn'
    else:
        default_tab = '8117'

    # Insights HTML 8091
    _ins91_html = ''
    for it in ins91:
        _ins91_html += (
            f'<div class="insight-item {it["cor"]}">'
            f'<span class="insight-emoji">{it["emoji"]}</span>'
            f'<div class="insight-text">{it["texto"]}</div></div>'
        )

    # Mini-insights por gráfico 8091
    m91_situacao    = mini91.get('chart-91-situacao',    '')
    m91_fase        = mini91.get('chart-91-fase',        '')
    m91_comp        = mini91.get('chart-91-comp',        '')
    m91_categoria   = mini91.get('chart-91-categoria',   '')
    m91_bairros     = mini91.get('chart-91-bairros',     '')
    m91_grupos      = mini91.get('chart-91-grupos',      '')
    m91_linha       = mini91.get('chart-8091-linha',     '')
    m91_faixas_vol  = mini91.get('chart-91-faixas-vol',  '')
    m91_leiturista  = mini91.get('chart-91-leiturista',  '')
    m91_czero       = mini91.get('chart-91-consumo-zero','')
    m91_leit_dia    = mini91.get('chart-91-leit-dia',    '')
    # Mini-insights 8117
    m17_arrec_diario = mini17.get('chart-arrec-diario', '')
    m17_formas_pizza = mini17.get('chart-formas-pizza', '')
    m17_formas_barra = mini17.get('chart-formas-barra', '')
    # Mini-insights 8121
    m21_arrec_fat    = mini21.get('chart-arrec-fat',    '')
    m21_acumulado    = mini21.get('chart-acumulado',    '')
    # Mini-insights 50012
    m50_criticas     = mini50.get('chart-50-criticas',    '')
    m50_faixas       = mini50.get('chart-50-faixas',      '')
    m50_cobranca     = mini50.get('chart-50-cobranca',    '')
    m50_leituristas  = mini50.get('chart-50-leituristas', '')
    m50_grupos       = mini50.get('chart-50-grupos',      '')
    m50_desvio       = mini50.get('chart-50-desvio',      '')
    m50_bairros_fat  = mini50.get('chart-50-bairros-fat', '')
    m50_categorias   = mini50.get('chart-50-categorias',  '')
    m50_dias         = mini50.get('chart-50-dias',        '')
    m50_equipes      = mini50.get('chart-50-equipes',     '')
    m50_bairros      = mini50.get('chart-50-bairros',     '')
    m50_hidrometros  = mini50.get('chart-50-hidrometros', '')

    # ── Mini-insights 50012 (gerados em Python, HTML estático) ──────────────
    def _mini50(emoji, cor, texto):
        cor_map = {
            'green':  ('rgba(46,125,50,0.08)',  '#2E7D32', 'rgba(46,125,50,0.35)'),
            'red':    ('rgba(198,40,40,0.08)',   '#C62828', 'rgba(198,40,40,0.35)'),
            'orange': ('rgba(230,115,0,0.08)',   '#E65100', 'rgba(230,115,0,0.35)'),
            'blue':   ('rgba(21,101,192,0.08)',  '#1565C0', 'rgba(21,101,192,0.35)'),
            '':       ('rgba(100,116,139,0.07)', '#475569', 'rgba(100,116,139,0.25)'),
        }
        bg, txt_cor, border = cor_map.get(cor, cor_map[''])
        return (
            f'<div style="display:flex;align-items:flex-start;gap:8px;padding:9px 13px;'
            f'background:{bg};border-left:3px solid {border};border-radius:0 7px 7px 0;'
            f'margin-top:6px;font-size:0.82rem;line-height:1.45;color:{txt_cor};">'
            f'<span style="font-size:1rem;flex-shrink:0;margin-top:1px">{emoji}</span>'
            f'<span>{texto}</span></div>'
        )
    def _mini50_bloco(items):
        if not items: return ''
        return ''.join(_mini50(**i) for i in items)

    # Criticas
    m50_criticas = ''
    if tem_50012 and criticas50:
        tot_crit = sum(c['qtd'] for c in criticas50)
        top_crit = max(criticas50, key=lambda x: x['qtd']) if criticas50 else None
        normais  = next((c['qtd'] for c in criticas50 if 'NORMAL' in c['label'].upper()), 0)
        pct_norm = normais / tot_crit * 100 if tot_crit else 0
        itens = []
        if top_crit and 'NORMAL' not in top_crit['label'].upper():
            itens.append({'emoji':'','cor':'red','texto':f'<strong>Crítica mais frequente:</strong> {top_crit["label"]}  {top_crit["qtd"]:,} ligações ({top_crit["qtd"]/tot_crit*100:.1f}% do total)'})
        itens.append({'emoji':'' if pct_norm >= 70 else '','cor':'green' if pct_norm >= 70 else 'red','texto':f'<strong>Leituras normais:</strong> {normais:,} ({pct_norm:.1f}%)  {"nível saudável" if pct_norm >= 70 else "abaixo do esperado, investigar"}'})
        m50_criticas = _mini50_bloco(itens)

    # Faixas consumo
    m50_faixas = ''
    if tem_50012 and faixas50:
        tot_fx = sum(faixas50.values())
        top_fx = max(faixas50.items(), key=lambda x: x[1]) if faixas50 else ('', 0)
        zeros  = faixas50.get('0-5', 0)
        itens = [
            {'emoji':'','cor':'green','texto':f'<strong>Faixa mais frequente:</strong> {top_fx[0]} m³ com {top_fx[1]:,} ligações ({top_fx[1]/tot_fx*100:.1f}% da base)'},
        ]
        if zeros / tot_fx * 100 > 5 if tot_fx else False:
            itens.append({'emoji':'','cor':'orange','texto':f'<strong>Consumo muito baixo (05 m³):</strong> {zeros:,} ligações ({zeros/tot_fx*100:.1f}%)  verificar hidrômetros ou imóveis fechados'})
        m50_faixas = _mini50_bloco(itens)

    # Leituristas
    m50_leituristas = ''
    if tem_50012 and leituristas50:
        top_l  = max(leituristas50, key=lambda x: x['qtd'])
        top_tc = max(leituristas50, key=lambda x: x['taxa_crit'])
        media_tc = sum(l['taxa_crit'] for l in leituristas50) / len(leituristas50)
        itens = [
            {'emoji':'','cor':'green','texto':f'<strong>{top_l["label"].split()[0]}</strong> tem mais leituras: {top_l["qtd"]:,} ({top_l["normais"]:,} normais, {top_l["criticos"]:,} críticas)'},
            {'emoji':'' if top_tc["taxa_crit"] > media_tc * 1.5 else '','cor':'orange' if top_tc["taxa_crit"] > 30 else '','texto':f'<strong>Maior taxa de críticas:</strong> {top_tc["label"].split()[0]} com {top_tc["taxa_crit"]:.1f}% (média da equipe: {media_tc:.1f}%)'},
        ]
        m50_leituristas = _mini50_bloco(itens)

    # Grupos
    m50_grupos = ''
    if tem_50012 and grupos50:
        top_g  = max(grupos50, key=lambda x: x['vol'])
        tot_v  = sum(g['vol'] for g in grupos50)
        tot_f  = sum(g['valor'] for g in grupos50)
        receita_m3_geral = tot_f / tot_v if tot_v > 0 else 0
        itens = [
            {'emoji':'','cor':'green','texto':f'<strong>{top_g["label"]}</strong> tem maior volume: {top_g["vol"]:,.0f} m³ ({top_g["vol"]/tot_v*100:.1f}% do total)  faturamento {_fmt_brl(top_g["valor"])}'},
            {'emoji':'','cor':'','texto':f'<strong>Eficiência geral:</strong> {_fmt_brl(receita_m3_geral)}/m³ arrecadado entre os {len(grupos50)} grupos de leitura'},
        ]
        m50_grupos = _mini50_bloco(itens)

    # Desvio
    m50_desvio = ''
    if tem_50012 and desvio50:
        tot_dev = sum(d['qtd'] for d in desvio50)
        queda   = next((d['qtd'] for d in desvio50 if '< -50%' in d['label']), 0)
        alta    = next((d['qtd'] for d in desvio50 if '> +50%' in d['label']), 0)
        normais_dev = next((d['qtd'] for d in desvio50 if '0%' in d['label'] and '+20%' in d['label']), 0)
        itens = []
        if queda / tot_dev * 100 > 15 if tot_dev else False:
            itens.append({'emoji':'','cor':'red','texto':f'<strong>Queda acentuada (< -50%):</strong> {queda:,} ligações ({queda/tot_dev*100:.1f}%)  possível submedição ou corte do serviço'})
        if alta / tot_dev * 100 > 10 if tot_dev else False:
            itens.append({'emoji':'','cor':'orange','texto':f'<strong>Aumento acentuado (> +50%):</strong> {alta:,} ligações ({alta/tot_dev*100:.1f}%)  verificar vazamentos ou ligações irregulares'})
        if not itens:
            itens.append({'emoji':'','cor':'green','texto':f'<strong>Consumo dentro do padrão histórico</strong> na maioria das ligações. Anomalias residuais: {queda+alta:,} registros'})
        m50_desvio = _mini50_bloco(itens)

    # Categorias
    m50_categorias = ''
    if tem_50012 and categorias50:
        tot_cat = sum(c['qtd'] for c in categorias50)
        top_cat = max(categorias50, key=lambda x: x['qtd'])
        pct_top = top_cat['qtd'] / tot_cat * 100 if tot_cat else 0
        itens = [
            {'emoji':'','cor':'blue' if top_cat['label']=='RESIDENCIAL' else '','texto':f'<strong>{top_cat["label"]}</strong> é a categoria dominante: {top_cat["qtd"]:,} ligações ({pct_top:.1f}% da base)'},
        ]
        outros = [c for c in categorias50 if c['label'] != top_cat['label']]
        if outros:
            seg = max(outros, key=lambda x: x['qtd'])
            itens.append({'emoji':'','cor':'','texto':f'<strong>Segunda categoria:</strong> {seg["label"]} com {seg["qtd"]:,} ligações ({seg["qtd"]/tot_cat*100:.1f}%)'})
        m50_categorias = _mini50_bloco(itens)

    # Bairros fat
    m50_bairros_fat = ''
    if tem_50012 and bairros_fat50:
        tot_fat_b = sum(b['valor'] for b in bairros_fat50)
        top_b = bairros_fat50[0] if bairros_fat50 else None
        if top_b:
            itens = [
                {'emoji':'','cor':'green','texto':f'<strong>{top_b["bairro"].title()}</strong> lidera: {_fmt_brl(top_b["valor"])} ({top_b["valor"]/tot_fat_b*100:.1f}% do faturamento, {top_b["qtd"]:,} ligações)'},
            ]
            if len(bairros_fat50) >= 3:
                top3 = sum(b["valor"] for b in bairros_fat50[:3])
                itens.append({'emoji':'','cor':'','texto':f'<strong>Top 3 bairros</strong> concentram {top3/tot_fat_b*100:.1f}% do faturamento total ({_fmt_brl(top3)})'})
            m50_bairros_fat = _mini50_bloco(itens)

    pct_txt    = f"{ins['pct_realizado']:.1f}%"
    tend_icon  = '' if ins['tendencia'] == 'alta' else ''
    tend_label = 'Tendência de Alta' if ins['tendencia'] == 'alta' else 'Tendência de Baixa'
    tend_color = '#2E7D32' if ins['tendencia'] == 'alta' else '#C62828'

    # ── Pre-gera conteúdo PN inline para embed direto na aba ────────────
    _pn_inline_css  = ''
    _pn_inline_body = ''
    _pn_inline_js   = ''
    _d17_pn = dados.get('macro_8117') or dados.get('macro_8117_pn')
    _d21_pn = dados.get('macro_8121') or dados.get('macro_8121_pn')
    _d91_pn = dados.get('macro_8091') or dados.get('macro_8091_pn')
    if tem_pn_aba and _d17_pn and _d21_pn:
        try:
            _dados_pn_inline = {
                'macro_8117': _d17_pn,
                'macro_8121': _d21_pn,
          'macro_8091': _d91_pn,
                'periodo': periodo,
                'cidade': cidade,
            }
            _pn_inline_css, _pn_inline_body, _pn_inline_js = gerar_pn_html(
                _dados_pn_inline, caminho_saida, _return_parts=True
            )
        except Exception:
            pass
    _pn_vars = (
        '--navy:#06203D;--teal:#0D5C63;--gold:#D98E04;--mint:#0B8A6F;'
        '--danger:#C03A2B;--card:#FFFFFF;--ink:#12263A;--muted:#5C6E82;--line:#D7E0EC'
    )
    _panel_pn_html = (
        '<div class="tab-panel" id="panel-pn">'
        '<style id="pn-scoped-style">' + _pn_inline_css + '</style>'
        '<div style="' + _pn_vars + '">' + _pn_inline_body + '</div>'
        '<script>' + _pn_inline_js + '</script>'
        '</div>'
    ) if tem_pn_aba else ''

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Dashboard Financeiro  {cidade} | Itapoá Saneamento</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/chartjs-plugin-datalabels/2.2.0/chartjs-plugin-datalabels.min.js"></script>
<style>
  :root {{
    --navy:    #011F3F;
    --navy2:   #0D3B6B;
    --orange:  #F27D16;
    --orange2: #FF8F00;
    --green:   #2E7D32;
    --red:     #C62828;
    --bg:      #F0F4F8;
    --card:    #FFFFFF;
    --border:  #DDE3EE;
    --text:    #1A2535;
    --muted:   #6B7A96;
    --shadow:  rgba(1,31,63,0.10);
  }}
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  body {{
    font-family: 'Segoe UI', system-ui, sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
  }}

  /* ── HEADER ── */
  .header {{
    background: linear-gradient(135deg, var(--navy) 0%, var(--navy2) 60%, #1A4D8F 100%);
    padding: 28px 40px 0;
    position: relative;
    overflow: hidden;
  }}
  .header::before {{
    content: '';
    position: absolute;
    top: -60px; right: -60px;
    width: 280px; height: 280px;
    background: radial-gradient(circle, rgba(242,125,22,0.18) 0%, transparent 70%);
    border-radius: 50%;
    pointer-events: none;
  }}
  .header-top {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    flex-wrap: wrap;
    gap: 16px;
    margin-bottom: 24px;
  }}
  .header-left h1 {{
    color: #fff;
    font-size: 1.55rem;
    font-weight: 800;
    letter-spacing: -0.3px;
  }}
  .header-left h1 span {{ color: var(--orange); }}
  .header-left p {{
    color: rgba(255,255,255,0.65);
    font-size: 0.85rem;
    margin-top: 4px;
  }}
  .header-badge {{
    background: rgba(255,255,255,0.1);
    border: 1px solid rgba(255,255,255,0.2);
    border-radius: 12px;
    padding: 10px 20px;
    text-align: right;
    backdrop-filter: blur(8px);
  }}
  .header-badge .periodo {{
    color: var(--orange);
    font-size: 0.75rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.8px;
  }}
  .header-badge .datas {{
    color: #fff;
    font-size: 0.95rem;
    font-weight: 700;
    margin-top: 2px;
  }}
  .header-badge .gerado {{
    color: rgba(255,255,255,0.5);
    font-size: 0.7rem;
    margin-top: 4px;
  }}

  /* ── TABS NAVIGATION ── */
  .tabs-nav {{
    display: flex;
    gap: 0;
    border-bottom: none;
    position: relative;
    z-index: 10;
  }}
  .tab-btn {{
    padding: 12px 28px;
    background: rgba(255,255,255,0.08);
    border: none;
    border-radius: 10px 10px 0 0;
    color: rgba(255,255,255,0.6);
    font-size: 0.88rem;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.2s ease;
    display: flex;
    align-items: center;
    gap: 8px;
    position: relative;
    font-family: inherit;
    letter-spacing: 0.3px;
  }}
  .tab-btn:hover {{
    background: rgba(255,255,255,0.14);
    color: rgba(255,255,255,0.9);
  }}
  .tab-btn.active {{
    background: var(--bg);
    color: var(--navy);
    font-weight: 700;
  }}
  .tab-btn.active::after {{
    content: '';
    position: absolute;
    bottom: -1px; left: 0; right: 0;
    height: 2px;
    background: var(--orange);
    border-radius: 2px 2px 0 0;
  }}
  .tab-dot {{
    width: 8px; height: 8px;
    border-radius: 50%;
    display: inline-block;
  }}
  .dot-blue   {{ background: #42A5F5; }}
  .dot-orange {{ background: #FF8F00; }}
  .dot-green  {{ background: #66BB6A; }}

  /* ── MAIN ── */
  .main {{ padding: 32px 40px; max-width: 1400px; margin: 0 auto; }}

  /* ── TAB PANELS ── */
  .tab-panel {{ display: none; }}
  .tab-panel.active {{ display: block; }}

  /* ── KPI CARDS ── */
  .kpi-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 16px;
    margin-bottom: 28px;
  }}
  .kpi-card {{
    background: var(--card);
    border-radius: 16px;
    padding: 20px 22px;
    border: 1px solid var(--border);
    box-shadow: 0 4px 20px var(--shadow);
    display: flex;
    flex-direction: column;
    gap: 8px;
    opacity: 0;
    transform: translateY(24px);
    transition: opacity 0.5s ease, transform 0.5s ease;
    position: relative;
    overflow: hidden;
  }}
  .kpi-card.visible {{ opacity: 1; transform: translateY(0); }}
  .kpi-card::before {{
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    border-radius: 16px 16px 0 0;
  }}
  .kpi-card.blue::before   {{ background: linear-gradient(90deg, #1565C0, #42A5F5); }}
  .kpi-card.orange::before {{ background: linear-gradient(90deg, #FF8F00, #FFCA28); }}
  .kpi-card.green::before  {{ background: linear-gradient(90deg, #2E7D32, #66BB6A); }}
  .kpi-card.purple::before {{ background: linear-gradient(90deg, #6A1B9A, #AB47BC); }}
  .kpi-card.teal::before   {{ background: linear-gradient(90deg, #00838F, #4DD0E1); }}
  .kpi-card.red::before    {{ background: linear-gradient(90deg, #C62828, #EF9A9A); }}

  .kpi-icon {{
    width: 40px; height: 40px;
    border-radius: 10px;
    display: flex; align-items: center; justify-content: center;
    font-size: 1.25rem;
  }}
  .kpi-card.blue   .kpi-icon {{ background: #E3F2FD; }}
  .kpi-card.orange .kpi-icon {{ background: #FFF8E1; }}
  .kpi-card.green  .kpi-icon {{ background: #E8F5E9; }}
  .kpi-card.purple .kpi-icon {{ background: #F3E5F5; }}
  .kpi-card.teal   .kpi-icon {{ background: #E0F7FA; }}
  .kpi-card.red    .kpi-icon {{ background: #FFEBEE; }}

  .kpi-label {{ font-size: 0.75rem; color: var(--muted); font-weight: 500; text-transform: uppercase; letter-spacing: 0.5px; }}
  .kpi-value {{ font-size: 1.45rem; font-weight: 800; color: var(--text); line-height: 1; }}
  .kpi-sub   {{ font-size: 0.75rem; color: var(--muted); }}

  /* ── SECTION TITLE ── */
  .section-title {{
    font-size: 0.95rem;
    font-weight: 700;
    color: var(--text);
    margin-bottom: 16px;
    display: flex;
    align-items: center;
    gap: 10px;
  }}
  .section-title::before {{
    content: '';
    display: block;
    width: 4px; height: 18px;
    background: linear-gradient(180deg, var(--orange), var(--navy));
    border-radius: 4px;
  }}

  /* ── CHART CARDS ── */
  .chart-card {{
    background: var(--card);
    border-radius: 16px;
    padding: 26px;
    border: 1px solid var(--border);
    box-shadow: 0 4px 20px var(--shadow);
    margin-bottom: 22px;
    opacity: 0;
    transform: translateY(36px);
    transition: opacity 0.55s ease, transform 0.55s ease;
  }}
  .chart-card.visible {{ opacity: 1; transform: translateY(0); }}
  .chart-header {{
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    margin-bottom: 20px;
    flex-wrap: wrap;
    gap: 10px;
  }}
  .chart-title {{ font-size: 0.95rem; font-weight: 700; color: var(--text); }}
  .chart-subtitle {{ font-size: 0.78rem; color: var(--muted); margin-top: 3px; }}
  .chart-badges {{ display: flex; gap: 8px; flex-wrap: wrap; }}
  .badge {{
    font-size: 0.7rem; font-weight: 600;
    padding: 3px 9px;
    border-radius: 20px;
    white-space: nowrap;
  }}
  .badge-blue   {{ background: #E3F2FD; color: #1565C0; }}
  .badge-orange {{ background: #FFF3E0; color: #E65100; }}
  .badge-green  {{ background: #E8F5E9; color: #2E7D32; }}
  .badge-navy   {{ background: #E8EDF5; color: #011F3F; }}
  .badge-purple {{ background: #F3E5F5; color: #6A1B9A; }}
  .badge-red    {{ background: #FFEBEE; color: #C62828; }}

  .btn-dias-toggle {{
    border: none;
    padding: 6px 14px;
    border-radius: 4px;
    font-size: 0.8rem;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.2s ease;
    font-family: 'Segoe UI', sans-serif;
  }}
  .btn-dias-toggle:hover {{
    opacity: 0.85;
    transform: translateY(-1px);
  }}
  .btn-dias-toggle:active {{
    transform: translateY(0);
  }}

  .chart-wrapper {{ position: relative; overflow: visible; }}
  .chart-wrapper canvas {{ max-height: 340px; }}
  #formas-barra-wrap {{ overflow-x: hidden; overflow-y: visible; }}
  #chart-formas-barra {{
    max-height: none !important;
    width: 100% !important;
    max-width: 100% !important;
    display: block;
  }}
  /* Tooltips nunca cortados */
  .chart-card {{ overflow: visible; }}

  /* ── GRID 2 COLUNAS ── */
  .grid-2 {{
    display: grid;
    grid-template-columns: 1fr 1fr;
    align-items: start;
    gap: 22px;
    margin-bottom: 22px;
  }}
  .grid-3 {{
    display: grid;
    grid-template-columns: 1fr 1fr 1fr;
    gap: 22px;
    margin-bottom: 22px;
  }}
  /* Bloco superior da Macro 50012: evita 3 cards apertados em telas médias */
  .grid-50012-top {{
    grid-template-columns: repeat(auto-fit, minmax(420px, 1fr));
  }}
  @media (max-width: 1100px) {{ .grid-3 {{ grid-template-columns: 1fr 1fr; }} }}
  @media (max-width: 900px) {{ .grid-2, .grid-3 {{ grid-template-columns: 1fr; }} }}
  @media (max-width: 980px) {{ .grid-50012-top {{ grid-template-columns: 1fr; }} }}

  /* ── INSIGHTS ── */
  .insights-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
    gap: 12px;
    margin-top: 20px;
  }}
  .insight-item {{
    background: var(--bg);
    border-radius: 10px;
    padding: 13px 15px;
    border-left: 3px solid var(--navy);
    display: flex;
    gap: 11px;
    align-items: flex-start;
    opacity: 0;
    transform: translateX(-16px);
    transition: opacity 0.45s ease, transform 0.45s ease;
  }}
  .insight-item.visible {{ opacity: 1; transform: translateX(0); }}
  .insight-item.orange {{ border-color: var(--orange); }}
  .insight-item.green  {{ border-color: var(--green); }}
  .insight-item.red    {{ border-color: var(--red); }}
  .insight-emoji {{ font-size: 1.2rem; line-height: 1; }}
  .insight-text  {{ font-size: 0.8rem; color: var(--text); line-height: 1.55; }}
  .insight-text strong {{ font-weight: 700; }}

  /* Mini-insights abaixo de cada gráfico */
  .mini-insights-block {{
    display: flex;
    flex-direction: column;
    gap: 4px;
    margin-top: 10px;
    padding-top: 10px;
    border-top: 1px solid rgba(0,0,0,0.06);
  }}
  .mini-insights-block:empty {{ display: none; }}

  /* ── KPI MINI GRID (8091) ── */
  .kpi-mini-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
    gap: 12px;
    margin-bottom: 22px;
  }}
  .kpi-mini {{
    background: #F0F4F8;
    border-radius: 12px;
    padding: 14px;
    border: 1px solid var(--border);
    text-align: center;
  }}
  .kpi-mini-label {{
    font-size: 0.68rem;
    color: var(--muted);
    text-transform: uppercase;
    font-weight: 600;
    letter-spacing: 0.5px;
  }}
  .kpi-mini-value {{
    font-size: 1.2rem;
    font-weight: 800;
    margin-top: 5px;
    line-height: 1;
  }}

  /* ── TABLE ── */
  .table-wrap {{ overflow-x: auto; margin-top: 8px; }}
  table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 0.83rem;
  }}
  th {{
    background: var(--navy);
    color: #fff;
    padding: 9px 13px;
    text-align: left;
    font-weight: 600;
    font-size: 0.75rem;
    text-transform: uppercase;
    letter-spacing: 0.4px;
  }}
  th:last-child {{ text-align: right; }}
  td {{
    padding: 9px 13px;
    border-bottom: 1px solid var(--border);
    color: var(--text);
  }}
  td:last-child {{ text-align: right; font-weight: 600; }}
  tr:hover td {{ background: #F5F8FF; }}
  tr:last-child td {{ border-bottom: none; font-weight: 700; background: #EEF2FF; }}


  /* ── RANGE COMPARE COMPONENT ── */
  .rc-wrap {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 12px 16px 10px;
    margin-bottom: 14px;
    user-select: none;
  }}
  .rc-header {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    flex-wrap: wrap;
    gap: 10px;
    margin-bottom: 10px;
  }}
  .rc-title {{
    font-size: 0.78rem; font-weight: 700;
    color: var(--muted); text-transform: uppercase;
    letter-spacing: 0.5px; display: flex; align-items: center; gap: 6px;
  }}
  .rc-range-labels {{ font-size: 0.78rem; color: var(--navy); font-weight: 600; }}
  .rc-slider-wrap {{ position: relative; height: 28px; margin: 0 8px; }}
  .rc-track {{
    position: absolute; top: 50%; left: 0; right: 0;
    height: 4px; background: var(--border);
    border-radius: 4px; transform: translateY(-50%);
  }}
  .rc-fill {{
    position: absolute; top: 50%; height: 4px;
    background: linear-gradient(90deg, var(--navy), var(--orange));
    border-radius: 4px; transform: translateY(-50%); pointer-events: none;
  }}
  .rc-input {{
    position: absolute; width: 100%; height: 100%; top: 0; left: 0;
    appearance: none; -webkit-appearance: none;
    background: transparent; pointer-events: none; margin: 0;
  }}
  .rc-input::-webkit-slider-thumb {{
    -webkit-appearance: none; width: 16px; height: 16px;
    border-radius: 50%; background: var(--navy);
    border: 2px solid #fff; box-shadow: 0 1px 4px rgba(0,0,0,0.25);
    pointer-events: all; cursor: grab; transition: transform 0.15s, background 0.15s;
  }}
  .rc-input::-webkit-slider-thumb:active {{ cursor: grabbing; transform: scale(1.2); background: var(--orange); }}
  .rc-input::-moz-range-thumb {{
    width: 16px; height: 16px; border-radius: 50%;
    background: var(--navy); border: 2px solid #fff;
    box-shadow: 0 1px 4px rgba(0,0,0,0.25); pointer-events: all; cursor: grab;
  }}
  .rc-checks {{ display: flex; gap: 8px; flex-wrap: wrap; margin-top: 8px; }}
  .rc-check-label {{
    display: flex; align-items: center; gap: 6px;
    font-size: 0.8rem; font-weight: 600; color: var(--text);
    cursor: pointer; padding: 5px 10px; border-radius: 7px;
    border: 1.5px solid var(--border); background: #fff; transition: all 0.15s;
  }}
  .rc-check-label:hover:not(.rc-disabled) {{ border-color: var(--navy); background: #EEF2FF; }}
  /* ano anterior → roxo */
  .rc-check-label.rc-active-year  {{ border-color: #7C3AED; background: #F5F3FF; color: #7C3AED; }}
  /* mesmo mês ano anterior → índigo */
  .rc-check-label.rc-active-ymonth {{ border-color: #4F46E5; background: #EEF2FF; color: #4F46E5; }}
  /* mês anterior → teal */
  .rc-check-label.rc-active-month {{ border-color: #0891B2; background: #ECFEFF; color: #0891B2; }}
  /* legado  mantém compatibilidade */
  .rc-check-label.rc-active {{ border-color: #7C3AED; background: #F5F3FF; color: #7C3AED; }}
  .rc-check-label.rc-disabled {{ opacity: 0.38; cursor: not-allowed; pointer-events: none; filter: grayscale(1); }}
  .rc-check-label input[type=checkbox] {{ width: 13px; height: 13px; cursor: pointer; }}
  /* badge por tipo */
  .rc-badge        {{ font-size: 0.68rem; padding: 1px 6px; border-radius: 4px; font-weight: 700; }}
  .rc-badge-year   {{ background: #EDE9FE; color: #7C3AED; }}
  .rc-badge-ymonth {{ background: #E0E7FF; color: #4F46E5; }}
  .rc-badge-month  {{ background: #CFFAFE; color: #0E7490; }}
  .rc-badge-gray   {{ background: #F1F5F9; color: #94A3B8; }}
  /* linha colorida à esquerda do label ativo */
  .rc-check-label.rc-active-year::before  {{ content:''; display:inline-block; width:3px; height:14px; border-radius:2px; background:#7C3AED; margin-right:2px; }}
  .rc-check-label.rc-active-ymonth::before {{ content:''; display:inline-block; width:3px; height:14px; border-radius:2px; background:#4F46E5; margin-right:2px; }}
  .rc-check-label.rc-active-month::before {{ content:''; display:inline-block; width:3px; height:14px; border-radius:2px; background:#0891B2; margin-right:2px; }}
  /* mês anterior ano passado → esmeralda */
  .rc-check-label.rc-active-pmy {{ border-color: #059669; background: #ECFDF5; color: #059669; }}
  .rc-badge-pmy  {{ background: #D1FAE5; color: #065F46; }}
  .rc-check-label.rc-active-pmy::before {{ content:''; display:inline-block; width:3px; height:14px; border-radius:2px; background:#059669; margin-right:2px; }}
  .rc-status {{
    display: none;
    margin-top: 8px;
    padding: 8px 10px;
    border-radius: 8px;
    font-size: 0.78rem;
    line-height: 1.35;
    border: 1px solid #FBD38D;
    background: #FFFAF0;
    color: #9C4221;
  }}
  .rc-status.show {{ display: block; }}

  /* ── FOOTER ── */
  .footer {{
    text-align: center;
    padding: 22px;
    color: var(--muted);
    font-size: 0.75rem;
    border-top: 1px solid var(--border);
    margin-top: 16px;
  }}
  .footer strong {{ color: var(--navy); }}

  /* ── DATATABLES CUSTOMIZAÇÃO ── */
  .dataTables_wrapper {{ font-family: 'Segoe UI', system-ui, sans-serif; }}
  .dataTables_header {{ padding: 12px 0; }}
  .dataTables_filter input {{ 
    padding: 8px 12px;
    border: 1px solid var(--border);
    border-radius: 6px;
    background: var(--bg);
    color: var(--text);
    font-size: 0.88rem;
  }}
  .dataTables_length select {{
    padding: 6px 8px;
    border: 1px solid var(--border);
    border-radius: 4px;
    background: var(--bg);
    color: var(--text);
  }}
  table.dataTable {{
    border-collapse: collapse;
    width: 100%;
  }}
  table.dataTable thead th {{
    background: var(--surface);
    border-bottom: 2px solid var(--border);
    padding: 10px 12px;
    text-align: left;
    font-weight: 600;
    color: var(--muted);
    font-size: 0.82em;
    position: relative;
    white-space: nowrap;
  }}
  table.dataTable thead th:after {{
    content: '';
    display: inline-block;
    width: 10px;
    height: 10px;
    background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 12 12'%3E%3Cpath d='M6 9L1 4h10z' fill='%23999'/%3E%3C/svg%3E");
    background-repeat: no-repeat;
    background-size: contain;
    margin-left: 6px;
    opacity: 0.3;
  }}
  table.dataTable thead .sorting {{ cursor: pointer; }}
  table.dataTable tbody td {{
    padding: 8px 12px;
    border-bottom: 1px solid var(--border);
    font-size: 0.83em;
  }}
  table.dataTable tbody tr:hover {{
    background: rgba(21,101,192,0.04);
  }}
  .dataTables_paginate {{ margin-top: 16px; padding-top: 12px; border-top: 1px solid var(--border); }}
  .paginate_button {{
    padding: 4px 8px;
    margin-right: 4px;
    border: 1px solid var(--border);
    border-radius: 4px;
    background: var(--card);
    color: var(--text);
    cursor: pointer;
    font-size: 0.85rem;
    transition: all 0.2s ease;
  }}
  .paginate_button:hover:not(.disabled) {{
    background: var(--orange);
    color: white;
    border-color: var(--orange);
  }}
  .paginate_button.current {{
    background: #1565C0;
    color: white;
    border-color: #1565C0;
  }}
  .dataTables_info {{
    padding-top: 12px;
    color: var(--muted);
    font-size: 0.8rem;
  }}
</style>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/datatables/1.10.21/css/dataTables.bootstrap4.min.css"/>
</head>
<body>

<div class="header">
  <div class="header-top">
    <div class="header-left">
      <h1>Dashboard <span>Financeiro</span></h1>
      <p>Itapoá Saneamento  {cidade} &nbsp;|&nbsp; {macros_badges}</p>
    </div>
    <div class="header-badge">
      <div class="periodo">Período Analisado</div>
      <div class="datas">{periodo['inicial']} → {periodo['final']}</div>
      <div class="gerado">Gerado em {gerado_em}</div>
    </div>
  </div>

  <!-- TABS -->
  <nav class="tabs-nav" id="tabs-nav">
    {'<button class="tab-btn" data-tab="8117" onclick="switchTab(\'8117\')"><span class="tab-dot dot-blue"></span>Macro 8117  Arrecadação</button>' if tem_8117 else ''}
    {'<button class="tab-btn" data-tab="8121" onclick="switchTab(\'8121\')"><span class="tab-dot dot-orange"></span>Macro 8121  Faturamento</button>' if tem_8121 else ''}
    {'<button class="tab-btn" data-tab="8091" onclick="switchTab(\'8091\')"><span class="tab-dot dot-green"></span>Macro 8091  Contas</button>' if tem_8091 else ''}
    {'<button class="tab-btn" data-tab="50012" onclick="switchTab(\'50012\')"><span class="tab-dot dot-purple"></span>Macro 50012  Analítico</button>' if tem_50012 else ''}
    {'<button class="tab-btn" data-tab="mapeamento" onclick="switchTab(\'mapeamento\')"><span class="tab-dot" style="background:#00695C"></span>Mapeamento das Equipes</button>' if tem_combo else ''}
    {'<button class="tab-btn" data-tab="pn" onclick="switchTab(\'pn\')"><span class="tab-dot" style="background:#00ACC1"></span>Painel de Negócios (PN)</button>' if tem_pn_aba else ''}
  </nav>
</div>

<div class="main">

  <!-- ═══════════════════════ TAB 8117 ═══════════════════════ -->
  <div class="tab-panel" id="panel-8117">

    <div class="kpi-grid">
      <div class="kpi-card blue" data-delay="0">
        <div class="kpi-icon"></div>
        <div class="kpi-label">Total Arrecadado</div>
        <div class="kpi-value" id="kpi-total-arr">R$ 0</div>
        <div class="kpi-sub">{periodo['inicial']} a {periodo['final']}</div>
      </div>
      <div class="kpi-card teal" data-delay="80">
        <div class="kpi-icon"></div>
        <div class="kpi-label">Média Diária</div>
        <div class="kpi-value" id="kpi-media">R$ 0</div>
        <div class="kpi-sub">Por dia de pagamento</div>
      </div>
      <div class="kpi-card purple" data-delay="160">
        <div class="kpi-icon"></div>
        <div class="kpi-label">Melhor Dia</div>
        <div class="kpi-value" id="kpi-melhor-val">R$ 0</div>
        <div class="kpi-sub" id="kpi-melhor-data">--</div>
      </div>
    </div>

    <div class="chart-card">
      <div class="chart-header">
        <div>
          <div class="chart-title">Arrecadação &amp; Acumulada</div>
          <div class="chart-subtitle" id="sub-arrec-diario">Evolução no período  Macro 8117</div>
        </div>
        <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
          <div style="display:flex;border:1px solid var(--border);border-radius:8px;overflow:hidden">
            <button id="btn8117-d" onclick="toggle8117('daily')"
              style="padding:6px 14px;border:none;background:#1565C0;color:#fff;font-size:0.8rem;font-weight:600;cursor:pointer;font-family:inherit">
               Diário
            </button>
            <button id="btn8117-m" onclick="toggle8117('monthly')"
              style="padding:6px 14px;border:none;background:transparent;color:#6B7A96;font-size:0.8rem;font-weight:600;cursor:pointer;font-family:inherit">
               Mensal
            </button>
            <button id="btn8117-y" onclick="toggle8117('yearly')"
              style="padding:6px 14px;border:none;background:transparent;color:#6B7A96;font-size:0.8rem;font-weight:600;cursor:pointer;font-family:inherit">
               Anual
            </button>
          </div>
          <span class="badge badge-orange">Linha = Acumulado</span>
        </div>
      </div>
      <div id="rc-8117-arrec"></div>
      <div class="chart-wrapper"><canvas id="chart-arrec-diario"></canvas></div>
      <div class="mini-insights-block">{m17_arrec_diario}</div>
    </div>

    <div class="grid-2">
      <div class="chart-card" style="margin-bottom:0">
        <div class="chart-header">
          <div>
            <div class="chart-title">Arrecadação por Forma</div>
            <div class="chart-subtitle">Distribuição percentual  Macro 8117</div>
          </div>
        </div>
        <div class="chart-wrapper"><canvas id="chart-formas-pizza"></canvas></div>
        <div class="mini-insights-block">{m17_formas_pizza}</div>
      </div>
      <div class="chart-card" style="margin-bottom:0">
        <div class="chart-header">
          <div>
            <div class="chart-title">Arrecadação por Forma</div>
            <div class="chart-subtitle">Valor arrecadado por forma de pagamento</div>
          </div>
        </div>
        <div class="chart-wrapper" id="formas-barra-wrap"><canvas id="chart-formas-barra"></canvas></div>
        <div class="mini-insights-block">{m17_formas_barra}</div>
      </div>
    </div>

    <div class="chart-card" style="margin-top:22px">
      <div class="section-title">Insights Automáticos  Arrecadação</div>
      <div class="insights-grid" id="insights-8117"></div>
    </div>

  </div>

  <!-- ═══════════════════════ TAB 8121 ═══════════════════════ -->
  <div class="tab-panel" id="panel-8121">

    <div class="kpi-grid">
      <div class="kpi-card blue" data-delay="0">
        <div class="kpi-icon"></div>
        <div class="kpi-label">Total Arrecadado</div>
        <div class="kpi-value" id="kpi-arr-21">R$ 0</div>
        <div class="kpi-sub">{periodo['inicial']} a {periodo['final']}</div>
      </div>
      <div class="kpi-card orange" data-delay="80">
        <div class="kpi-icon"></div>
        <div class="kpi-label">Total Faturado</div>
        <div class="kpi-value" id="kpi-fat-21">R$ 0</div>
        <div class="kpi-sub">{periodo['inicial']} a {periodo['final']}</div>
      </div>
      <div class="kpi-card green" data-delay="160">
        <div class="kpi-icon"></div>
        <div class="kpi-label">% Realizado</div>
        <div class="kpi-value" id="kpi-pct-21">0%</div>
        <div class="kpi-sub">Arrecadado / Faturado</div>
      </div>
      <div class="kpi-card purple" data-delay="240">
        <div class="kpi-icon"></div>
        <div class="kpi-label">Maior Faturamento</div>
        <div class="kpi-value" id="kpi-maior-fat">R$ 0</div>
        <div class="kpi-sub" id="kpi-maior-fat-data">--</div>
      </div>
    </div>

    <div class="chart-card">
      <div class="chart-header">
        <div>
          <div class="chart-title">Arrecadação vs Faturamento</div>
          <div class="chart-subtitle" id="sub-arrec-fat">Comparativo  Macro 8121</div>
        </div>
        <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
          <div style="display:flex;border:1px solid var(--border);border-radius:8px;overflow:hidden">
            <button id="btn8121-d" onclick="toggle8121('daily')"
              style="padding:6px 14px;border:none;background:#1565C0;color:#fff;font-size:0.8rem;font-weight:600;cursor:pointer;font-family:inherit">
               Diário
            </button>
            <button id="btn8121-m" onclick="toggle8121('monthly')"
              style="padding:6px 14px;border:none;background:transparent;color:#6B7A96;font-size:0.8rem;font-weight:600;cursor:pointer;font-family:inherit">
               Mensal
            </button>
            <button id="btn8121-y" onclick="toggle8121('yearly')"
              style="padding:6px 14px;border:none;background:transparent;color:#6B7A96;font-size:0.8rem;font-weight:600;cursor:pointer;font-family:inherit">
               Anual
            </button>
          </div>
          <span class="badge badge-blue">Arrecadação</span>
          <span class="badge badge-orange">Faturamento</span>
        </div>
      </div>
      <div id="rc-8121-arrec"></div>
      <div class="chart-wrapper"><canvas id="chart-arrec-fat"></canvas></div>
      <div class="mini-insights-block">{m21_arrec_fat}</div>
    </div>

    <div class="chart-card">
      <div class="chart-header">
        <div>
          <div class="chart-title">Curva Acumulada  Arrecadação vs Faturamento</div>
          <div class="chart-subtitle" id="sub-acumulado">Evolução acumulada no período</div>
        </div>
        <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
          <div style="display:flex;border:1px solid var(--border);border-radius:8px;overflow:hidden">
            <button id="btn-acum-d" onclick="toggleAcumulado('daily')"
              style="padding:6px 14px;border:none;background:#1565C0;color:#fff;font-size:0.8rem;font-weight:600;cursor:pointer;font-family:inherit">
               Diário
            </button>
            <button id="btn-acum-m" onclick="toggleAcumulado('monthly')"
              style="padding:6px 14px;border:none;background:transparent;color:#6B7A96;font-size:0.8rem;font-weight:600;cursor:pointer;font-family:inherit">
               Mensal
            </button>
          </div>
          <span class="badge badge-blue">Arrecadação Acum.</span>
          <span class="badge badge-orange">Faturamento Acum.</span>
        </div>
      </div>
      <div id="rc-8121-acum"></div>
      <div class="chart-wrapper"><canvas id="chart-acumulado"></canvas></div>
      <div class="mini-insights-block">{m21_acumulado}</div>
    </div>

    <div class="chart-card">
      <div class="section-title">Insights Automáticos  Faturamento x Arrecadação</div>
      <div class="insights-grid" id="insights-8121"></div>
    </div>

  </div>

  <!-- ═══════════════════════ TAB 8091 ═══════════════════════ -->
  <div class="tab-panel" id="panel-8091">

    <!-- KPIs 8091 -->
    <div class="kpi-mini-grid">
      <div class="kpi-mini">
        <div class="kpi-mini-label">Total Faturado</div>
        <div class="kpi-mini-value" id="k91-fat" style="color:#011F3F">R$ 0</div>
      </div>
      <div class="kpi-mini">
        <div class="kpi-mini-label">Água</div>
        <div class="kpi-mini-value" id="k91-agua" style="color:#1565C0">R$ 0</div>
      </div>
      <div class="kpi-mini">
        <div class="kpi-mini-label">Multas</div>
        <div class="kpi-mini-value" id="k91-mult" style="color:#C62828">R$ 0</div>
      </div>
      <div class="kpi-mini">
        <div class="kpi-mini-label">Ligações</div>
        <div class="kpi-mini-value" id="k91-lig" style="color:#2E7D32">0</div>
      </div>
      <div class="kpi-mini">
        <div class="kpi-mini-label">Volume m³</div>
        <div class="kpi-mini-value" id="k91-vol" style="color:#00838F">0</div>
      </div>
      <div class="kpi-mini">
        <div class="kpi-mini-label">Ticket Médio</div>
        <div class="kpi-mini-value" id="k91-media" style="color:#6A1B9A">R$ 0</div>
      </div>
    </div>

    <!-- Volume por Dia de Leitura -->
    <div class="chart-card">
      <div class="chart-header">
        <div>
          <div class="chart-title">Leituras Processadas no Período</div>
          <div class="chart-subtitle" id="sub-91-leit-dia">Visão objetiva com volume total, ligações lidas e linha de média geral por ligação</div>
        </div>
        <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
          <div style="display:flex;border:1px solid var(--border);border-radius:8px;overflow:hidden">
            <button id="btn91dia-d" onclick="toggle91LeitDia('daily')"
              style="padding:6px 14px;border:none;background:#1565C0;color:#fff;font-size:0.8rem;font-weight:600;cursor:pointer;font-family:inherit">
               Diário
            </button>
            <button id="btn91dia-m" onclick="toggle91LeitDia('monthly')"
              style="padding:6px 14px;border:none;background:transparent;color:#6B7A96;font-size:0.8rem;font-weight:600;cursor:pointer;font-family:inherit">
               Mensal
            </button>
          </div>
          <span class="badge badge-blue">m³</span>
          <span class="badge badge-navy">Ligações</span>
          <span class="badge badge-orange">Média geral m³/lig</span>
        </div>
      </div>
      <div id="rc-91-leit-dia"></div>
      <div class="chart-wrapper"><canvas id="chart-91-leit-dia"></canvas></div>
      <div class="mini-insights-block">{m91_leit_dia}</div>
    </div>

    <!-- Situação + Fase -->
    <div class="grid-2">
      <div class="chart-card" style="margin-bottom:0">
        <div class="chart-header">
          <div>
            <div class="chart-title">Situação das Faturas</div>
            <div class="chart-subtitle">Distribuição por status de pagamento</div>
          </div>
          <span class="badge badge-green">{meses_count} mês(es)</span>
        </div>
        <div class="chart-wrapper"><canvas id="chart-91-situacao"></canvas></div>
        <div class="mini-insights-block">{m91_situacao}</div>
      </div>
      <div class="chart-card" style="margin-bottom:0">
        <div class="chart-header">
          <div>
            <div class="chart-title">Fase de Cobrança</div>
            <div class="chart-subtitle">Ligações por fase</div>
          </div>
        </div>
        <div class="chart-wrapper"><canvas id="chart-91-fase"></canvas></div>
        <div class="mini-insights-block">{m91_fase}</div>
      </div>
    </div>

    <!-- Composição + Faixas -->
    <div class="grid-2" style="margin-top:22px">
      <div class="chart-card" style="margin-bottom:0">
        <div class="chart-header">
          <div>
            <div class="chart-title">Composição do Faturamento</div>
            <div class="chart-subtitle">Água vs Multas vs Outros</div>
          </div>
          <span class="badge badge-blue">Total: {total_91_fmt}</span>
        </div>
        <div class="chart-wrapper"><canvas id="chart-91-comp"></canvas></div>
        <div class="mini-insights-block">{m91_comp}</div>
      </div>
    </div>

    <!-- Categoria -->
    <div class="chart-card" style="margin-top:22px">
      <div class="chart-header">
        <div>
          <div class="chart-title">Faturamento por Categoria de Uso</div>
          <div class="chart-subtitle">Valor e quantidade de ligações por categoria</div>
        </div>
        <div class="chart-badges">
          <span class="badge badge-navy">Valor</span>
          <span class="badge badge-navy">Qtde</span>
        </div>
      </div>
      <div class="chart-wrapper"><canvas id="chart-91-categoria"></canvas></div>
      <div class="mini-insights-block">{m91_categoria}</div>
    </div>

    <!-- Top Bairros -->
    <div class="chart-card">
      <div class="chart-header">
        <div>
          <div class="chart-title">Top 10 Bairros por Faturamento</div>
          <div class="chart-subtitle">Ranking e detalhamento</div>
        </div>
      </div>
      <div class="grid-2" style="margin-bottom:0;gap:28px">
        <div class="chart-wrapper"><canvas id="chart-91-bairros"></canvas></div>
        <div class="mini-insights-block">{m91_bairros}</div>
        <div style="overflow-x:auto">
          <table>
            <tr>
              <th>#</th>
              <th>Bairro</th>
              <th>Lig.</th>
              <th style="text-align:right">Faturado</th>
            </tr>
            {_bairros_rows}
          </table>
        </div>
      </div>
    </div>

    <!-- Receita por m³ por Grupo -->
    <div class="chart-card">
      <div class="chart-header">
        <div>
          <div class="chart-title">Receita por m³ por Grupo de Leitura</div>
          <div class="chart-subtitle">Eficiência tarifária  R$ arrecadado por metro cúbico consumido</div>
        </div>
        <span class="badge badge-orange">R$/m³</span>
      </div>
      <div class="chart-wrapper"><canvas id="chart-91-grupos"></canvas></div>
      <div class="mini-insights-block">{m91_grupos}</div>
    </div>

    <!-- Indicadores de inadimplência -->
    <div class="chart-card">
      <div class="section-title" style="margin-bottom:12px">Indicadores de Inadimplência</div>
      <div style="display:flex;align-items:center;justify-content:center;gap:40px;flex-wrap:wrap;padding:10px 0" id="rings-91"></div>
    </div>

    <!-- Evolução mensal (multi-mês) -->
    <div id="chart-8091-linha-wrap" class="chart-card">
      <div class="chart-header">
        <div>
          <div class="chart-title">Evolução Mensal do Faturamento</div>
          <div class="chart-subtitle">Faturamento mês a mês  Macro 8091</div>
        </div>
        <span class="badge badge-green">{meses_count} meses</span>
      </div>
      <div id="rc-8091-linha"></div>
      <canvas id="chart-8091-linha"></canvas>
      <div class="mini-insights-block">{m91_linha}</div>
    </div>

    <!-- Faixas de Consumo (m³) -->
    <div class="chart-card" data-delay="200">
      <div class="chart-header">
        <div>
          <div class="chart-title">Distribuição de Consumo por Faixa (m³)</div>
          <div class="chart-subtitle">Quantidade de ligações por faixa de volume medido</div>
        </div>
        <span class="badge badge-blue">Volume</span>
      </div>
      <div class="chart-wrapper"><canvas id="chart-91-faixas-vol"></canvas></div>
      <div class="mini-insights-block">{m91_faixas_vol}</div>
    </div>

    <!-- Leituristas -->
    <div class="chart-card" data-delay="250">
      <div class="chart-header">
        <div>
          <div class="chart-title">Desempenho por Leiturista</div>
          <div class="chart-subtitle">Volume lido, ligações e volume médio por ligação</div>
        </div>
        <span class="badge badge-green">Operacional</span>
      </div>
      <div class="chart-wrapper"><canvas id="chart-91-leiturista"></canvas></div>
      <div class="mini-insights-block">{m91_leiturista}</div>
    </div>

    <!-- Consumo zero -->
    <div class="chart-card" data-delay="300">
      <div class="chart-header">
        <div>
          <div class="chart-title">Ligações com Consumo Zero (m³)</div>
          <div class="chart-subtitle">Possíveis irregularidades, hidrômetros parados ou imóveis fechados</div>
        </div>
        <span class="badge" style="background:#FFF3E0;color:#E65100"> Atenção</span>
      </div>
      <div class="chart-wrapper"><canvas id="chart-91-consumo-zero"></canvas></div>
      <div class="mini-insights-block">{m91_czero}</div>
    </div>

    <!-- Insights 8091  ao final da aba -->
    <div class="chart-card">
      <div class="section-title">Insights Automáticos  Macro 8091</div>
      <div class="insights-grid" id="insights-8091">
        {_ins91_html}
      </div>
    </div>

  </div><!-- /panel-8091 -->

  <!-- ═══════════════════════ TAB 50012 ══════════════════════════ -->
  <div class="tab-panel" id="panel-50012">

    <!-- KPIs analíticos -->
    <div class="kpi-grid">
      <div class="kpi-card blue" data-delay="0">
        <div class="kpi-label">Ligações</div>
        <div class="kpi-value" id="k50-lig">0</div>
        <div class="kpi-sub">100% ativas</div>
      </div>
      <div class="kpi-card teal" data-delay="80">
        <div class="kpi-label">Volume Total</div>
        <div class="kpi-value" id="k50-vol">0 m³</div>
        <div class="kpi-sub" id="k50-media-vol">média  m³/lig</div>
      </div>
      <div class="kpi-card orange" data-delay="160">
        <div class="kpi-label">Faturamento</div>
        <div class="kpi-value" id="k50-fat">R$ 0</div>
        <div class="kpi-sub">total do período</div>
      </div>
      <div class="kpi-card green" data-delay="240">
        <div class="kpi-label">Leituras Normais</div>
        <div class="kpi-value" id="k50-pct-norm">0%</div>
        <div class="kpi-sub">sem críticas</div>
      </div>
      <div class="kpi-card red" data-delay="320">
        <div class="kpi-label">Leituras Críticas</div>
        <div class="kpi-value" id="k50-pct-crit">0%</div>
        <div class="kpi-sub">requerem atenção</div>
      </div>
      <div class="kpi-card purple" data-delay="400">
        <div class="kpi-label">Volume Zero</div>
        <div class="kpi-value" id="k50-zeros">0</div>
        <div class="kpi-sub">possíveis irregularidades</div>
      </div>
    </div>

    <!-- Alertas operacionais -->
    <div style="display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:28px">
      <div style="background:var(--card);border:1px solid var(--border);border-left:4px solid #C62828;border-radius:12px;padding:14px 16px;box-shadow:0 2px 10px var(--shadow)">
        <div style="font-size:0.72rem;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px">Aumento ≥ 101%</div>
        <div class="kpi-value" id="k50-aumento" style="font-size:1.6rem;color:#C62828">0</div>
      </div>
      <div style="background:var(--card);border:1px solid var(--border);border-left:4px solid #FF8F00;border-radius:12px;padding:14px 16px;box-shadow:0 2px 10px var(--shadow)">
        <div style="font-size:0.72rem;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px">Leit = Anterior</div>
        <div class="kpi-value" id="k50-leit-ant" style="font-size:1.6rem;color:#FF8F00">0</div>
      </div>
      <div style="background:var(--card);border:1px solid var(--border);border-left:4px solid #FF8F00;border-radius:12px;padding:14px 16px;box-shadow:0 2px 10px var(--shadow)">
        <div style="font-size:0.72rem;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px">Queda ≥ 50%</div>
        <div class="kpi-value" id="k50-queda" style="font-size:1.6rem;color:#FF8F00">0</div>
      </div>
      <div style="background:var(--card);border:1px solid var(--border);border-left:4px solid #C62828;border-radius:12px;padding:14px 16px;box-shadow:0 2px 10px var(--shadow)">
        <div style="font-size:0.72rem;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px">Não Realizada</div>
        <div class="kpi-value" id="k50-nao-real" style="font-size:1.6rem;color:#C62828">0</div>
      </div>
      <div style="background:var(--card);border:1px solid var(--border);border-left:4px solid #1565C0;border-radius:12px;padding:14px 16px;box-shadow:0 2px 10px var(--shadow)">
        <div style="font-size:0.72rem;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px">HD Invertido</div>
        <div class="kpi-value" id="k50-hd-inv" style="font-size:1.6rem;color:#1565C0">0</div>
      </div>
    </div>

    <!-- Linha 1: Críticas + Faixas + Cobrança -->
    <div class="grid-3 grid-50012-top" style="margin-bottom:22px">
      <div class="chart-card" data-delay="0">
        <div class="chart-header">
          <div>
            <div class="chart-title">Críticas de Medição</div>
            <div class="chart-subtitle">Distribuição por tipo de crítica atual</div>
          </div>
          <span class="badge badge-orange">Qualidade</span>
        </div>
        <div class="chart-wrapper"><canvas id="chart-50-criticas"></canvas>
      <div class="mini-insights-block">{m50_criticas}</div></div>
      </div>
      <div class="chart-card" data-delay="80">
        <div class="chart-header">
          <div>
            <div class="chart-title">Faixas de Consumo</div>
            <div class="chart-subtitle">Ligações por volume medido (m³)</div>
          </div>
          <span class="badge badge-blue">Volume</span>
        </div>
        <div class="chart-wrapper"><canvas id="chart-50-faixas"></canvas>
      <div class="mini-insights-block">{m50_faixas}</div></div>
      </div>
      <div class="chart-card" data-delay="160">
        <div class="chart-header">
          <div>
            <div class="chart-title">Situação de Cobrança</div>
            <div class="chart-subtitle">Status comercial das ligações</div>
          </div>
          <span class="badge badge-green">Comercial</span>
        </div>
        <div class="chart-wrapper"><canvas id="chart-50-cobranca"></canvas></div>
        <div class="mini-insights-block">{m50_cobranca}</div>
      </div>
    </div>

    <!-- Leituristas -->
    <div class="chart-card" data-delay="100" style="margin-bottom:22px">
      <div class="chart-header">
        <div>
          <div class="chart-title">Desempenho por Leiturista</div>
          <div class="chart-subtitle">Volume total lido, leituras normais vs. críticas e taxa de criticidade</div>
        </div>
        <span class="badge badge-green">Operacional</span>
      </div>
      <div class="chart-wrapper"><canvas id="chart-50-leituristas" style="max-height:320px"></canvas>
      <div class="mini-insights-block">{m50_leituristas}</div></div>
    </div>

    <!-- Linha 2: Grupos + Desvio -->
    <div class="grid-2" style="margin-bottom:22px">
      <div class="chart-card" data-delay="0">
        <div class="chart-header">
          <div>
            <div class="chart-title">Volume + Faturamento por Grupo</div>
            <div class="chart-subtitle">m³ medido e R$ faturado por grupo de leitura</div>
          </div>
          <span class="badge badge-blue">Grupos</span>
        </div>
        <div class="chart-wrapper"><canvas id="chart-50-grupos"></canvas>
      <div class="mini-insights-block">{m50_grupos}</div></div>
      </div>
      <div class="chart-card" data-delay="80">
        <div class="chart-header">
          <div>
            <div class="chart-title">Desvio vs. Média Histórica</div>
            <div class="chart-subtitle">Comparação do consumo atual com a média cadastrada</div>
          </div>
          <span class="badge" style="background:#FFF3E0;color:#E65100"> Anomalias</span>
        </div>
        <div class="chart-wrapper"><canvas id="chart-50-desvio"></canvas>
      <div class="mini-insights-block">{m50_desvio}</div></div>
      </div>
    </div>

    <!-- Linha 3: Bairros fat + Categorias -->
    <div class="grid-2" style="margin-bottom:22px">
      <div class="chart-card" data-delay="0">
        <div class="chart-header">
          <div>
            <div class="chart-title">Faturamento por Bairro  Top 12</div>
            <div class="chart-subtitle">Receita total por bairro no período</div>
          </div>
          <span class="badge badge-orange">Receita</span>
        </div>
        <div class="chart-wrapper"><canvas id="chart-50-bairros-fat"></canvas>
      <div class="mini-insights-block">{m50_bairros_fat}</div></div>
      </div>
      <div class="chart-card" data-delay="80">
        <div class="chart-header">
          <div>
            <div class="chart-title">Categorias de Consumo</div>
            <div class="chart-subtitle">Distribuição por tipo de ligação</div>
          </div>
          <span class="badge badge-green">Categorias</span>
        </div>
        <div class="chart-wrapper"><canvas id="chart-50-categorias"></canvas>
      <div class="mini-insights-block">{m50_categorias}</div></div>
      </div>
    </div>

    <!-- Leituras por dia -->
    <div class="chart-card" data-delay="0" style="margin-bottom:22px">
      <div class="chart-header">
        <div>
          <div class="chart-title">Leituras por Dia  Volume e Quantidade</div>
          <div class="chart-subtitle">Produção diária de leituras e volume medido no período</div>
        </div>
        <div style="display:flex;gap:12px;align-items:center">
          <span class="badge badge-blue">Timeline</span>
          <div style="display:flex;background:rgba(0,0,0,0.06);border-radius:6px;padding:3px;border:1px solid #DDE3EE">
            <button id="btn-dias-diario" class="btn-dias-toggle" style="background:#1565C0;color:white">Diário</button>
            <button id="btn-dias-mensal" class="btn-dias-toggle" style="background:transparent;color:#6B7A96">Mensal</button>
          </div>
        </div>
      </div>
      <div id="rc-50-dias"></div>
      <div class="chart-wrapper"><canvas id="chart-50-dias" style="max-height:200px"></canvas>
      <div class="mini-insights-block">{m50_dias}</div></div>
    </div>

    <!-- Dados Geográficos -->
    <div class="grid-2" style="margin-bottom:22px">
      <div class="chart-card" data-delay="0">
        <div class="chart-header">
          <div>
            <div class="chart-title">OSs por Leiturista (Coord.)</div>
            <div class="chart-subtitle">Distribuição geográfica por equipe</div>
          </div>
        </div>
        <div class="chart-wrapper"><canvas id="chart-50-equipes"></canvas>
      <div class="mini-insights-block">{m50_equipes}</div></div>
      </div>
      <div class="chart-card" data-delay="80">
        <div class="chart-header">
          <div>
            <div class="chart-title">Top 15 Bairros  Ligações</div>
            <div class="chart-subtitle">Quantidade de OSs por bairro geográfico</div>
          </div>
        </div>
        <div class="chart-wrapper"><canvas id="chart-50-bairros"></canvas>
      <div class="mini-insights-block">{m50_bairros}</div></div>
      </div>
    </div>

    <!-- Dados Geográficos card -->
    <div class="chart-card" style="margin-bottom:22px">
      <div class="chart-header">
        <div>
          <div class="chart-title">Cobertura Geográfica</div>
          <div class="chart-subtitle">Centro e limites calculados de todas as ligações com coordenadas</div>
        </div>
      </div>
      <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:16px;padding:16px 0 0">
        <div style="background:rgba(21,101,192,0.07);border:1px solid #1565C0;border-radius:8px;padding:14px;text-align:center">
          <div style="font-size:0.8rem;color:var(--muted);margin-bottom:6px">Total</div>
          <div id="k50-total" style="font-weight:700;font-size:1.4rem;color:#1565C0">0</div>
        </div>
        <div style="background:rgba(0,131,143,0.07);border:1px solid #00838F;border-radius:8px;padding:14px;text-align:center">
          <div style="font-size:0.8rem;color:var(--muted);margin-bottom:6px">Com Coord.</div>
          <div id="k50-coord" style="font-weight:700;font-size:1.4rem;color:#00838F">0</div>
        </div>
        <div style="background:rgba(198,40,40,0.07);border:1px solid #C62828;border-radius:8px;padding:14px;text-align:center">
          <div style="font-size:0.8rem;color:var(--muted);margin-bottom:6px">Sem Coord.</div>
          <div id="k50-sem" style="font-weight:700;font-size:1.4rem;color:#C62828">0</div>
        </div>
        <div style="background:rgba(46,125,50,0.07);border:1px solid #2E7D32;border-radius:8px;padding:14px;text-align:center">
          <div style="font-size:0.8rem;color:var(--muted);margin-bottom:6px">Cobertura</div>
          <div id="k50-cobertura" style="font-weight:700;font-size:1.4rem;color:#2E7D32">0%</div>
        </div>
      </div>
      <div style="display:grid;grid-template-columns:repeat(2,1fr);gap:12px;margin-top:12px">
        <div style="background:rgba(123,31,162,0.07);border:1px solid #7B1FA2;border-radius:8px;padding:12px">
          <div style="font-size:0.85rem;color:var(--muted);margin-bottom:4px">Centro Geográfico (Centróide)</div>
          <div style="font-size:0.85rem;color:var(--muted);margin-bottom:4px">Centro Geográfico do Atendimento</div>
          <div id="geo-centro" style="font-weight:600;font-family:monospace">--,&nbsp;--</div>
        </div>
        <div style="background:rgba(0,131,143,0.07);border:1px solid #00838F;border-radius:8px;padding:12px">
          <div style="font-size:0.85rem;color:var(--muted);margin-bottom:4px">Bounding Box (Limites)</div>
          <div style="font-size:0.85rem;color:var(--muted);margin-bottom:4px">Limites do Mapa (Extremos)</div>
          <div id="geo-bbox" style="font-weight:600;font-family:monospace;font-size:0.82rem">N: --, S: --, E: --, W: --</div>
        </div>
      </div>
    </div>

  </div><!-- /panel-50012 -->

  {_panel_pn_html}

  {_html_panel_mapeamento}

</div><!-- /main -->

<div class="footer">
  Gerado automaticamente pelo <strong>Mathools 1.0  Itapoá Saneamento</strong>
  &nbsp;|&nbsp; Dados: Waterfy {' &amp; '.join(['Macro ' + m for m in macros])}
  &nbsp;|&nbsp; {gerado_em}
  &nbsp;|&nbsp; <em>Documento confidencial  uso interno</em>
</div>

<script>

// ══════════════════════════════════════════════════════════════════════
//  RANGE COMPARE  filtro de período + comparação histórica
// ══════════════════════════════════════════════════════════════════════
// ── SINCRONIZAÇÃO DE PERÍODO POR ABA ─────────────────────────────────
// Quando o usuário arrasta o slider de qualquer gráfico, todos os outros
// gráficos da MESMA ABA sincronizam para o mesmo intervalo de datas.
// Gráficos de abas diferentes não são afetados.
// ══════════════════════════════════════════════════════════════════════
const _RC_SYNC = {{
  _registry: [],   // {{ containerId, panelId, syncFn }}
  _syncing: false, // evita loops de sincronização

  // Descobre em qual panel-XXXX o container está
  _getPanelId(containerId) {{
    const el = document.getElementById(containerId);
    if (!el) return null;
    const panel = el.closest('[id^="panel-"]');
    return panel ? panel.id : null;
  }},

  register(containerId, syncFn) {{
    this._registry = this._registry.filter(r => r.containerId !== containerId);
    const panelId = this._getPanelId(containerId);
    this._registry.push({{ containerId, panelId, syncFn }});
  }},

  broadcast(sourceId, dateStart, dateEnd) {{
    if (this._syncing) return;
    this._syncing = true;
    const sourcePanelId = this._getPanelId(sourceId);
    this._registry.forEach(r => {{
      // Só sincroniza dentro da mesma aba
      if (r.containerId !== sourceId && r.panelId === sourcePanelId) {{
        try {{ r.syncFn(dateStart, dateEnd); }} catch(e) {{}}
      }}
    }});
    this._syncing = false;
  }}
}};

//  Uso: buildRangeCompare({{ containerId, labels, datasets, rebuildFn }})
// ══════════════════════════════════════════════════════════════════════
function buildRangeCompare(cfg) {{
  const el = document.getElementById(cfg.containerId);
  if (!el) return;
  const allLabels   = cfg.labels   || [];
  const allDatasets = cfg.datasets || [];
  const N = allLabels.length;
  if (N === 0) {{ el.innerHTML = ''; return; }}

  let idxS = 0, idxE = N - 1, cmpAno = false, cmpMes = false, cmpYearMonth = false, cmpPMY = false;

  function _parseLabel(s) {{
    if (!s) return null;
    s = String(s).trim();
    let m = s.match(/^(\\d{{2}})\\/(\\d{{2}})\\/(\\d{{4}})$/);
    if (m) return new Date(+m[3], +m[2]-1, +m[1]);
    m = s.match(/^(\\d{{1,2}})\\/(\\d{{4}})$/);
    if (m) return new Date(+m[2], +m[1]-1, 1);
    // Formatos mensais adicionais usados em algumas macros (ex.: 2026-03)
    m = s.match(/^(\\d{{4}})-(\\d{{1,2}})$/);
    if (m) return new Date(+m[1], +m[2]-1, 1);
    m = s.match(/^(\\d{{4}})\\/(\\d{{1,2}})$/);
    if (m) return new Date(+m[1], +m[2]-1, 1);
    m = s.match(/^(\\d{{1,2}})-(\\d{{4}})$/);
    if (m) return new Date(+m[2], +m[1]-1, 1);
    const mn = {{jan:0,fev:1,feb:1,mar:2,abr:3,apr:3,mai:4,may:4,jun:5,
                jul:6,ago:7,aug:7,set:8,sep:8,out:9,oct:9,nov:10,dez:11,dec:11}};
    m = s.toLowerCase().match(/^([a-z]{{3}})\\s*(\\d{{4}})?/);
    if (m && mn[m[1]] !== undefined) return new Date(+(m[2]||2000), mn[m[1]], 1);
    return null;
  }}

  function _temDoisMeses(i0, i1) {{
    if (i1 - i0 + 1 < 1) return false;
    const d0 = _parseLabel(allLabels[i0]), d1 = _parseLabel(allLabels[i1]);
    // Para labels diários: habilita com qualquer range que tenha pelo menos 2 dias
    if (!d0 || !d1) return (i1 - i0 + 1) >= 2;
    // Habilita com pelo menos 1 dia de diferença  permite selecionar 1 mês inteiro
    // e comparar com o mês anterior ou mesmo período do ano passado
    return (d1.getFullYear() - d0.getFullYear()) * 12 + (d1.getMonth() - d0.getMonth()) >= 1
        || (i1 - i0 + 1) >= 2;
  }}

  function _buildCmpDs(i0, i1, tipo) {{
    const idxs = [];
    const meses = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];

    const dFirst = _parseLabel(allLabels[i0]);
    const dLast  = _parseLabel(allLabels[i1]);
    let periodoLabel = '';

    // ── Helper: reconstrói acumulado somando os diários do período mapeado ────
    // O acumulado de comparação NUNCA parte do valor histórico absoluto 
    // sempre recomeça do zero, somando os valores diários mapeados em sequência.
    // Assim: primeiro ponto = valor do primeiro dia, segundo = soma dos dois, etc.
    // dsDaily é o dataset Diário (di=0), idxsMapped é o array de índices mapeados.
    function _recalcAcum(idxsMapped, dsDaily) {{
      let soma = 0;
      return idxsMapped.map(j => {{
        if (j < 0) return null;
        const v = dsDaily.data[j] ?? 0;
        soma += v;
        return Math.round(soma * 100) / 100;
      }});
    }}

    // ── Helper: monta os datasets de comparação dado os índices mapeados ──────
    function _makeCmpDatasets(idxsMapped, suffix, cores, estilo) {{
      // Acumulado de comparação:
      // só recalcula progressivamente quando existe uma série diária pareada.
      // Em gráficos que já são "acumulado puro" (sem diária), mantém os valores
      // históricos mapeados para não deformar a curva de comparação.
      const _norm = (s) => String(s || '').toLowerCase()
        .replaceAll('á','a').replaceAll('ã','a').replaceAll('â','a')
        .replaceAll('é','e').replaceAll('ê','e')
        .replaceAll('í','i')
        .replaceAll('ó','o').replaceAll('ô','o').replaceAll('õ','o')
        .replaceAll('ú','u')
        .replaceAll('ç','c');

      function _pickDailyBaseDataset(dsLabel) {{
        const dailyBases = (allDatasets || []).filter(b => /diari/.test(_norm(b.label)));
        if (!dailyBases.length) return null;

        const alvo = _norm(dsLabel);
        const chaves = ['arrecad', 'fatur', 'volume', 'leit', 'agua', 'esgoto'];
        for (const k of chaves) {{
          if (alvo.includes(k)) {{
            const match = dailyBases.find(b => _norm(b.label).includes(k));
            if (match) return match;
          }}
        }}
        return dailyBases[0];
      }}

      return allDatasets.map((ds, di) => {{
        const corSerie = cores[di % cores.length];
        let data;
        if (ds.label && _norm(ds.label).includes('acumulad')) {{
          const dsDaily = _pickDailyBaseDataset(ds.label);
          data = dsDaily ? _recalcAcum(idxsMapped, dsDaily)
                         : idxsMapped.map(j => j >= 0 ? (ds.data[j] ?? null) : null);
        }} else {{
          data = idxsMapped.map(j => j >= 0 ? (ds.data[j] ?? null) : null);
        }}
        return {{
          label: ds.label + suffix,
          data,
          borderColor: corSerie,
          backgroundColor: 'transparent',
          borderWidth: estilo.borderWidth ?? 1.8,
          borderDash: estilo.borderDash,
          pointRadius: estilo.pointRadius ?? 3,
          pointHoverRadius: estilo.pointHoverRadius ?? 6,
          pointStyle: estilo.pointStyle ?? 'circle',
          fill: false, tension: 0.3, _isCompare: true,
          _periodo: periodoLabel,
          _tipo: tipo,
        }};
      }});
    }}

    // ── Modo "mês anterior (cada mês)" ───────────────────────────────────────
    // Cada ponto do período visível é comparado com o mês imediatamente anterior.
    // Ex: Dez/2025 → Nov/2025 | Jan/2026 → Dez/2025 | Fev/2026 → Jan/2026
    // No modo diário: 19/Fev/2026 → 19/Jan/2026
    if (tipo === 'month') {{
      if (dFirst && dLast) {{
        const dAntFirst = new Date(dFirst.getFullYear(), dFirst.getMonth() - 1, dFirst.getDate());
        const dAntLast  = new Date(dLast.getFullYear(),  dLast.getMonth()  - 1, dLast.getDate());
        const mFirst = meses[dAntFirst.getMonth()] + '/' + dAntFirst.getFullYear();
        const mLast  = meses[dAntLast.getMonth()]  + '/' + dAntLast.getFullYear();
        periodoLabel  = mFirst === mLast ? mFirst : mFirst + ' → ' + mLast;
      }}
      const suffix = periodoLabel ? `  ${{periodoLabel}}` : ' (mês ant.)';
      const idxsMonth = [];
      for (let i = i0; i <= i1; i++) {{
        const d = _parseLabel(allLabels[i]);
        if (!d) {{ idxsMonth.push(-1); continue; }}
        // Alvo: mesmo dia, 1 mês antes
        const dRef = new Date(d.getFullYear(), d.getMonth() - 1, d.getDate());
        let best = -1, bestD = Infinity;
        for (let j = 0; j < N; j++) {{
          const dj = _parseLabel(allLabels[j]);
          if (!dj) continue;
          // "Mês anterior": só aceita pontos do MESMO mês/ano alvo.
          if (dj.getFullYear() !== dRef.getFullYear() || dj.getMonth() !== dRef.getMonth()) continue;
          const diff = Math.abs(dj - dRef);
          if (diff < bestD) {{ bestD = diff; best = j; }}
        }}
        // Tolerância curta: evita "colar" em pontos errados.
        idxsMonth.push(best >= 0 && bestD <= 3 * 86400000 ? best : -1);
      }}
      return _makeCmpDatasets(idxsMonth, suffix,
        ['#0891B2','#0D9488'], // teal: Diária / Acumulado
        {{ borderDash: [4,4], pointRadius: 4, pointHoverRadius: 7, pointStyle: 'circle' }}
      );
    }}

    // ── Modo "mesmo mês  ano anterior" ──────────────────────────────────────
    // Cada ponto mapeado para o mesmo dia/mês do ano anterior.
    // Ex: 19/Fev/2026 → 19/Fev/2025 | 05/Mar/2026 → 05/Mar/2025
    if (tipo === 'year_month') {{
      if (dFirst && dLast) {{
        const aFirst = dFirst.getFullYear() - 1;
        const aLast  = dLast.getFullYear()  - 1;
        const mFirst = meses[dFirst.getMonth()] + '/' + aFirst;
        const mLast  = meses[dLast.getMonth()]  + '/' + aLast;
        periodoLabel = mFirst === mLast ? mFirst : mFirst + ' → ' + mLast;
      }}
      const suffix = periodoLabel ? `  ${{periodoLabel}}` : ' (mesmo mês 1 ano)';
      const idxsYM = [];
      for (let i = i0; i <= i1; i++) {{
        const d = _parseLabel(allLabels[i]);
        if (!d) {{ idxsYM.push(-1); continue; }}
        const dRef = new Date(d.getFullYear() - 1, d.getMonth(), d.getDate());
        let best = -1, bestD = Infinity;
        for (let j = 0; j < N; j++) {{
          const dj = _parseLabel(allLabels[j]);
          if (!dj) continue;
          // "Mesmo mês do ano passado": só aceita mês/ano exatos.
          if (dj.getFullYear() !== dRef.getFullYear() || dj.getMonth() !== dRef.getMonth()) continue;
          const diff = Math.abs(dj - dRef);
          if (diff < bestD) {{ bestD = diff; best = j; }}
        }}
        idxsYM.push(best >= 0 && bestD <= 3 * 86400000 ? best : -1);
      }}
      return _makeCmpDatasets(idxsYM, suffix,
        ['#4F46E5','#7C3AED'], // índigo: Diária / Acumulado
        {{ borderDash: [5,3], pointRadius: 4, pointHoverRadius: 7, pointStyle: 'rectRot' }}
      );
    }}

    // ── Modo "mês anterior  ano passado" ────────────────────────────────────
    // Cada ponto mapeado para o mês anterior do ano passado.
    // Ex: Mar/2026 → Fev/2025 | Fev/2026 → Jan/2025 | Jan/2026 → Dez/2024
    if (tipo === 'prev_month_year') {{
      if (dFirst && dLast) {{
        const dRefFirst = new Date(dFirst.getFullYear()-1, dFirst.getMonth()-1, dFirst.getDate());
        const dRefLast  = new Date(dLast.getFullYear()-1,  dLast.getMonth()-1,  dLast.getDate());
        const mFirst = meses[dRefFirst.getMonth()] + '/' + dRefFirst.getFullYear();
        const mLast  = meses[dRefLast.getMonth()]  + '/' + dRefLast.getFullYear();
        periodoLabel = mFirst === mLast ? mFirst : mFirst + ' → ' + mLast;
      }}
      const suffix = periodoLabel ? `  ${{periodoLabel}}` : ' (mês ant. 1 ano)';
      const idxsPMY = [];
      for (let i = i0; i <= i1; i++) {{
        const d = _parseLabel(allLabels[i]);
        if (!d) {{ idxsPMY.push(-1); continue; }}
        // Mesmo dia, 1 mês antes, 1 ano antes
        const dRef = new Date(d.getFullYear()-1, d.getMonth()-1, d.getDate());
        let best = -1, bestD = Infinity;
        for (let j = 0; j < N; j++) {{
          const dj = _parseLabel(allLabels[j]);
          if (!dj) continue;
          // "Mês anterior do ano passado": mês/ano alvo exatos.
          if (dj.getFullYear() !== dRef.getFullYear() || dj.getMonth() !== dRef.getMonth()) continue;
          const diff = Math.abs(dj - dRef);
          if (diff < bestD) {{ bestD = diff; best = j; }}
        }}
        idxsPMY.push(best >= 0 && bestD <= 3 * 86400000 ? best : -1);
      }}
      return _makeCmpDatasets(idxsPMY, suffix,
        ['#059669','#047857'], // esmeralda: Diária / Acumulado
        {{ borderDash: [3,3], pointRadius: 3, pointHoverRadius: 6, pointStyle: 'rectRot' }}
      );
    }}

    // ── Modo "mesmo período  ano passado" ────────────────────────────────────
    // O range inteiro é espelhado 1 ano atrás.
    // Ex: 19/Fev19/Mar/2026 → 19/Fev19/Mar/2025
    if (dFirst && dLast) {{
      const shift = (d) => new Date(d.getFullYear()-1, d.getMonth(), d.getDate());
      const dRefFirst = shift(dFirst);
      const dRefLast  = shift(dLast);
      const aFirst = dRefFirst.getFullYear(), aLast = dRefLast.getFullYear();
      periodoLabel = aFirst === aLast ? String(aFirst) : `${{aFirst}}${{aLast}}`;
    }}

    for (let i = i0; i <= i1; i++) {{
      const d = _parseLabel(allLabels[i]);
      if (!d) {{ idxs.push(-1); continue; }}
      const dRef = new Date(d.getFullYear()-1, d.getMonth(), d.getDate());
      let best = -1, bestD = Infinity;
      for (let j = 0; j < N; j++) {{
        const dj = _parseLabel(allLabels[j]);
        if (!dj) continue;
        // "Mesmo período ano passado": trava no mesmo mês/ano alvo.
        if (dj.getFullYear() !== dRef.getFullYear() || dj.getMonth() !== dRef.getMonth()) continue;
        const diff = Math.abs(dj - dRef);
        if (diff < bestD) {{ bestD = diff; best = j; }}
      }}
      idxs.push(best >= 0 && bestD <= 3 * 86400000 ? best : -1);
    }}

    const suffix = periodoLabel ? `  ${{periodoLabel}}` : ' (ano ant.)';

    return _makeCmpDatasets(idxs, suffix,
      ['#7C3AED','#C026D3'], // roxo: Diária / Acumulado
      {{ borderDash: [6,3], pointRadius: 3, pointHoverRadius: 6, pointStyle: 'triangle' }}
    );
  }}

  function _apply() {{
    const i0 = Math.min(idxS, idxE), i1 = Math.max(idxS, idxE);
    const filtL  = allLabels.slice(i0, i1+1);
    const filtDs = allDatasets.map(ds => {{ return {{ ...ds, data: ds.data.slice(i0, i1+1) }}; }});
    const cmpDs  = [];
    const _hasCmpSelecionado = cmpAno || cmpMes || cmpYearMonth || cmpPMY;
    if (cmpAno) cmpDs.push(..._buildCmpDs(i0, i1, 'year'));
    if (cmpYearMonth) cmpDs.push(..._buildCmpDs(i0, i1, 'year_month'));
    if (cmpPMY) cmpDs.push(..._buildCmpDs(i0, i1, 'prev_month_year'));
    if (cmpMes) cmpDs.push(..._buildCmpDs(i0, i1, 'month'));

    const _temValorComparativo = (cmpDs || []).some(ds =>
      Array.isArray(ds?.data) && ds.data.some(v => v !== null && v !== undefined && Number.isFinite(Number(v)))
    );

    cfg.rebuildFn(filtL, filtDs, cmpDs, i0, i1);
    _updateUI(i0, i1);

    const st = el.querySelector('.rc-status');
    if (st) {{
      if (_hasCmpSelecionado && !_temValorComparativo) {{
        st.innerHTML = ' Sem base histórica compatível para este recorte. Ajuste o período no slider ou desative a comparação.';
        st.classList.add('show');
      }} else {{
        st.textContent = '';
        st.classList.remove('show');
      }}
    }}
  }}

  function _updateUI(i0, i1) {{
    const lbl = el.querySelector('.rc-range-labels');
    if (lbl) lbl.textContent = allLabels[i0] + '  →  ' + allLabels[i1];
    const hab = _temDoisMeses(i0, i1);
    el.querySelectorAll('.rc-check-label').forEach(cl => {{
      cl.classList.toggle('rc-disabled', !hab);
      if (!hab) {{
        const inp = cl.querySelector('input');
        if (inp && inp.checked) inp.checked = false;
        cl.classList.remove('rc-active','rc-active-year','rc-active-ymonth','rc-active-month','rc-active-pmy');
        cl.querySelectorAll('.rc-badge').forEach(b => b.classList.add('rc-badge-gray'));
      }} else {{
        cl.querySelectorAll('.rc-badge').forEach(b => b.classList.remove('rc-badge-gray'));
      }}
    }});
    _fill();
  }}

  function _fill() {{
    const f = el.querySelector('.rc-fill');
    if (!f || N <= 1) return;
    const pL = (Math.min(idxS, idxE) / (N-1)) * 100;
    const pR = (Math.max(idxS, idxE) / (N-1)) * 100;
    f.style.left = pL + '%'; f.style.width = (pR - pL) + '%';
  }}

  const hab0 = _temDoisMeses(0, N-1);
  el.innerHTML =
    '<div class="rc-wrap">' +
    '<div class="rc-header">' +
    '<span class="rc-title">&#128197; Recorte de período</span>' +
    '<span class="rc-range-labels">' + allLabels[0] + '  \u2192  ' + allLabels[N-1] + '</span>' +
    '</div>' +
    '<div class="rc-slider-wrap">' +
    '<div class="rc-track"></div><div class="rc-fill"></div>' +
    '<input class="rc-input" id="' + cfg.containerId + '-s" type="range" min="0" max="' + (N-1) + '" value="0" step="1">' +
    '<input class="rc-input" id="' + cfg.containerId + '-e" type="range" min="0" max="' + (N-1) + '" value="' + (N-1) + '" step="1">' +
    '</div>' +
    '<div class="rc-checks">' +
    '<label class="rc-check-label' + (hab0 ? '' : ' rc-disabled') + '" id="' + cfg.containerId + '-la">' +
    '<input type="checkbox" id="' + cfg.containerId + '-ca">' +
    '<span> Mesmo período  ano passado</span>' +
    '<span class="rc-badge rc-badge-year' + (hab0 ? '' : ' rc-badge-gray') + '">\u22121 ano</span>' +
    '</label>' +
    '<label class="rc-check-label' + (hab0 ? '' : ' rc-disabled') + '" id="' + cfg.containerId + '-lpmy">' +
    '<input type="checkbox" id="' + cfg.containerId + '-cpmy">' +
    '<span> Mês anterior  ano passado</span>' +
    '<span class="rc-badge rc-badge-pmy' + (hab0 ? '' : ' rc-badge-gray') + '" id="' + cfg.containerId + '-bpmy">\u22121m \u22121a</span>' +
    '</label>' +
    '<label class="rc-check-label' + (hab0 ? '' : ' rc-disabled') + '" id="' + cfg.containerId + '-lm">' +
    '<input type="checkbox" id="' + cfg.containerId + '-cm">' +
    '<span>↩ Mês anterior (cada mês)</span>' +
    '<span class="rc-badge rc-badge-month' + (hab0 ? '' : ' rc-badge-gray') + '" id="' + cfg.containerId + '-bm">\u22121 mês</span>' +
    '</label>' +
    '</div>' +
    '<div class="rc-status" id="' + cfg.containerId + '-status"></div>' +
    '</div>';

  const inpS = el.querySelector('#' + cfg.containerId + '-s');
  const inpE = el.querySelector('#' + cfg.containerId + '-e');
  const cbA   = el.querySelector('#' + cfg.containerId + '-ca');
  const cbPMY = el.querySelector('#' + cfg.containerId + '-cpmy');
  const cbM   = el.querySelector('#' + cfg.containerId + '-cm');
  const lbA   = el.querySelector('#' + cfg.containerId + '-la');
  const lbPMY = el.querySelector('#' + cfg.containerId + '-lpmy');
  const lbM   = el.querySelector('#' + cfg.containerId + '-lm');
  const bdPMY = el.querySelector('#' + cfg.containerId + '-bpmy');
  const bdM   = el.querySelector('#' + cfg.containerId + '-bm');

  // Aplica defaultStart se configurado (ex: últimos 90 dias)
  if (cfg.defaultStart != null && cfg.defaultStart >= 0 && cfg.defaultStart < N) {{
    idxS = cfg.defaultStart;
    inpS.value = idxS;
  }}

  // Atualiza badges dinâmicos conforme a seleção do slider
  const _mesesNomes = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  function _atualizarBadges() {{
    const i1 = Math.max(parseInt(inpS.value), parseInt(inpE.value));
    const dLast = _parseLabel(allLabels[i1]);
    if (dLast) {{
      if (bdM) {{
        const dAnt = new Date(dLast.getFullYear(), dLast.getMonth() - 1, 1);
        bdM.textContent = _mesesNomes[dAnt.getMonth()] + '/' + dAnt.getFullYear();
      }}
      if (bdPMY) {{
        // Mês anterior do ano passado: 1 mês 1 ano
        const dPMY = new Date(dLast.getFullYear() - 1, dLast.getMonth() - 1, 1);
        bdPMY.textContent = _mesesNomes[dPMY.getMonth()] + '/' + dPMY.getFullYear();
      }}
    }} else {{
      if (bdM)   bdM.textContent   = '\u22121 mês';
      if (bdPMY) bdPMY.textContent = '\u22121m \u22121a';
    }}
  }}

  function _slide() {{
    idxS = parseInt(inpS.value); idxE = parseInt(inpE.value);
    if (idxS > idxE) {{
      if (document.activeElement === inpS) idxS = idxE; else idxE = idxS;
      inpS.value = idxS; inpE.value = idxE;
    }}
    _atualizarBadges();
    _apply();
    // Broadcast: notifica outros RCs com as datas do intervalo atual
    const d0 = _parseLabel(allLabels[Math.min(idxS, idxE)]);
    const d1 = _parseLabel(allLabels[Math.max(idxS, idxE)]);
    if (d0 && d1) _RC_SYNC.broadcast(cfg.containerId, d0, d1);
  }}
  inpS.addEventListener('input', _slide);
  inpE.addEventListener('input', _slide);

  // Registra este RC no bus de sincronização
  // syncFn: recebe (dateStart, dateEnd) e ajusta o slider para o range mais próximo
  _RC_SYNC.register(cfg.containerId, (dateStart, dateEnd) => {{
    let bestS = 0, bestE = N - 1, bestDS = Infinity, bestDE = Infinity;
    for (let j = 0; j < N; j++) {{
      const dj = _parseLabel(allLabels[j]);
      if (!dj) continue;
      const diffS = Math.abs(dj - dateStart);
      const diffE = Math.abs(dj - dateEnd);
      if (diffS < bestDS) {{ bestDS = diffS; bestS = j; }}
      if (diffE < bestDE) {{ bestDE = diffE; bestE = j; }}
    }}
    // Só sincroniza se encontrou datas razoavelmente próximas (dentro de 60 dias)
    const tolerancia = 60 * 86400000;
    if (bestDS < tolerancia || bestDE < tolerancia) {{
      idxS = bestS; idxE = bestE;
      inpS.value = idxS; inpE.value = idxE;
      _atualizarBadges();
      _apply();
    }}
  }});

  // Helpers para alternar classe de cor ativa por tipo
  function _setActive(lb, cls, on) {{
    if (!lb) return;
    lb.classList.remove('rc-active','rc-active-year','rc-active-ymonth','rc-active-month','rc-active-pmy');
    if (on) lb.classList.add(cls);
  }}

  cbA.addEventListener('change', () => {{
    cmpAno = cbA.checked;
    _setActive(lbA, 'rc-active-year', cmpAno);
    _apply();
  }});
  if (cbPMY) cbPMY.addEventListener('change', () => {{
    cmpPMY = cbPMY.checked;
    _setActive(lbPMY, 'rc-active-pmy', cmpPMY);
    _apply();
  }});
  cbM.addEventListener('change', () => {{
    cmpMes = cbM.checked;
    _setActive(lbM, 'rc-active-month', cmpMes);
    _apply();
  }});
  _atualizarBadges();
  _fill();
  _apply(); // dispara imediatamente para popular mini-insights com o range inicial
}}

// ── DADOS ────────────────────────────────────────────────────────────
const DATAS_17  = {datas_js};
const ARR_D     = {arr_d_js};
const ARR_AC    = {arr_ac_js};
const DATAS_21  = {datas_21_js};
const ARR21_D   = {arr21_d_js};
const ARR21_AC  = {arr21_ac_js};
const DIAS_MES  = {dias_mes_js};
const FAT_D     = {fat_d_js};
const FAT_AC    = {fat_ac_js};
const FAT_AC_ANT   = {fat_ac_ant_js};
const ARR_AC_ANT   = {arr_ac_ant_js};
const DIAS_MES_ANT = {dias_mes_ant_js};
const LABEL_ATUAL  = {label_atual};
const LABEL_ANT    = {label_ant};
const FORMAS_N  = {formas_nomes};
const FORMAS_V  = {formas_vals};
const FORMAS_Q  = {formas_qtde};
const FORMAS_C  = {formas_cores};
const TOTAL_ARR = {_safe_num(ins['total_arr'])};
const TOTAL_FAT = {_safe_num(ins['total_fat'])};
const MEDIA_ARR = {_safe_num(ins['media_arr'])};
const PCT       = {_safe_num(ins['pct_realizado'])};
const ACIMA_MEDIA = {_safe_num(ins['acima_media_arr'])};
const TOTAL_DIAS  = {_safe_num(ins['total_dias'])};
const MAIOR_ARR_V = {_safe_num(ins['maior_arr'][1])};
const MAIOR_ARR_D = "{ins['maior_arr'][0]}";
const MAIOR_FAT_V = {_safe_num(ins['maior_fat'][1])};
const MAIOR_FAT_D = "{ins['maior_fat'][0]}";
const DIAS_SEM    = {_safe_num(ins['dias_sem_arr'])};
const TEM_8117 = {'true' if tem_8117 else 'false'};
const TEM_8121 = {'true' if tem_8121 else 'false'};
const TEM_8091 = {'true' if tem_8091 else 'false'};
const TEM_50012 = {'true' if tem_50012 else 'false'};
const TABS_AVAILABLE = {tabs_js};
const DEFAULT_TAB  = "{default_tab}";

// ── DADOS 8091 ────────────────────────────────────────────────────────
const K91_FAT   = {_safe_num(k91_fat)};
const K91_AGUA  = {_safe_num(k91_agua)};
const K91_MULT  = {_safe_num(k91_mult)};
const K91_LIG   = {_safe_num(k91_lig)};
const K91_VOL   = {_safe_num(k91_vol)};
const K91_MEDIA = {_safe_num(k91_media)};
const SIT91_L = {sit91_L}; const SIT91_Q = {sit91_Q}; const SIT91_V = {sit91_V};
const CAT91_L = {cat91_L}; const CAT91_Q = {cat91_Q}; const CAT91_V = {cat91_V};
const FASE91_L = {fase91_L}; const FASE91_Q = {fase91_Q}; const FASE91_V = {fase91_V};
const EMIS91_L = {emis91_L}; const EMIS91_V = {emis91_V};
const BAR91_L = {bar91_L}; const BAR91_V = {bar91_V}; const BAR91_Q = {bar91_Q};
const GRP91_L = {grp91_L}; const GRP91_V = {grp91_V}; const GRP91_VOL = {grp91_VOL};
const FAIXAS91_L = {faixas91_L}; const FAIXAS91_V = {faixas91_V};
const FVOL91_L = {fvol91_L}; const FVOL91_V = {fvol91_V};
const LEIT91_L = {leit91_L}; const LEIT91_Q = {leit91_Q}; const LEIT91_VOL = {leit91_VOL}; const LEIT91_VM = {leit91_VM};
const LEIT91_CRIT = [];
const CZERO91_L = {czero91_L}; const CZERO91_Q = {czero91_Q};
const LDIA91_L = {ldia91_L}; const LDIA91_Q = {ldia91_Q}; const LDIA91_V = {ldia91_V};
const GRP_R91_L = {grp_r91_L}; const GRP_R91_RM = {grp_r91_RM}; const GRP_R91_V = {grp_r91_V}; const GRP_R91_VOL = {grp_r91_VOL};
const COMP91_L = {comp91_L}; const COMP91_V = {comp91_V};
const SERIE_L = {serie_labels}; const SERIE_V = {serie_totais};
const MESES_COUNT = {_safe_num(meses_count, 1)};

// ── DADOS 50012 (Latitude/Longitude) ────────────────────────────────────
const D50_TOTAL = {_safe_num(d50_total)};
const D50_COM_COORD = {_safe_num(d50_com_coord)};
const D50_SEM_COORD = {_safe_num(d50_sem_coord)};
const D50_COBERTURA = {_safe_num(d50_cobertura)};
const D50_EQUIPES_L = {equipesL_50};
const D50_EQUIPES_Q = {equipesQ_50};
const D50_EQUIPES_C = {equipesC_50};
const D50_BAIRROS_L = {bairrosL_50};
const D50_BAIRROS_Q = {bairrosQ_50};
const D50_CENTRO_LAT = {_safe_num(centro_lat_50)};
const D50_CENTRO_LNG = {_safe_num(centro_lng_50)};
const D50_BBOX = {bbox_50};
const TOTAL_91 = {_safe_num(total_91)};
const CONTAS_N = {contas_nomes}; const CONTAS_V = {contas_vals};
// Analíticos 50012
const K50_TOT_LIG  = {_safe_num(k50_tot_lig)};
const K50_VOL      = {_safe_num(k50_vol)};
const K50_FAT      = {_safe_num(k50_fat)};
const K50_MEDIA_V  = {_safe_num(k50_media_v)};
const K50_PCT_NORM = {_safe_num(k50_pct_norm)};
const K50_PCT_CRIT = {_safe_num(k50_pct_crit)};
const K50_ZEROS    = {_safe_num(k50_zeros)};
const K50_AUMENTO  = {_safe_num(k50_aumento)};
const K50_LEIT_ANT = {_safe_num(k50_leit_ant)};
const K50_QUEDA    = {_safe_num(k50_queda)};
const K50_NAO_REAL = {_safe_num(k50_nao_real)};
const K50_HD_INV   = {_safe_num(k50_hd_inv)};
const CRIT50_L = {crit50_L}; const CRIT50_Q = {crit50_Q};
const M50_CRITICAS    = {_safe_json(m50_criticas if tem_50012 else '')};
const M50_FAIXAS      = {_safe_json(m50_faixas if tem_50012 else '')};
const M50_LEITURISTAS = {_safe_json(m50_leituristas if tem_50012 else '')};
const M50_GRUPOS      = {_safe_json(m50_grupos if tem_50012 else '')};
const M50_DESVIO      = {_safe_json(m50_desvio if tem_50012 else '')};
const M50_CATEGORIAS  = {_safe_json(m50_categorias if tem_50012 else '')};
const M50_BAIRROS_FAT = {_safe_json(m50_bairros_fat if tem_50012 else '')};
const DIST_FUNC_L = {dist_func_L}; const DIST_FUNC_Q = {dist_func_Q};
const HID_MODELOS   = {hid_modelos_js};
const HID_TOTAIS    = {hid_totais_js};
const HID_TIPOS_RAW = {hid_tipos_raw_js};
const TOP_TIPOS_OS  = {top_tipos_os_js};
const HID_RET_MESES  = {hid_ret_meses_js};
const HID_RET_SERIES = {hid_ret_series_js};
const LEIT_OS_L     = {leit_os_labels_js};
const LEIT_OS_Q     = {leit_os_totais_js};
const FXC50_L  = {fxc50_L};  const FXC50_Q  = {fxc50_Q};
const LEIT50_L = {leit50_L}; const LEIT50_Q = {leit50_Q};
const LEIT50_CR= {leit50_CR};const LEIT50_NO= {leit50_NO};
const LEIT50_TC= {leit50_TC};const LEIT50_V = {leit50_V};
const GRP50_L  = {grp50_L};  const GRP50_V  = {grp50_V};  const GRP50_F = {grp50_F};
const COB50_L  = {cob50_L};  const COB50_Q  = {cob50_Q};
const BRF50_L  = {brf50_L};  const BRF50_V  = {brf50_V};  const BRF50_Q = {brf50_Q};
const DEV50_L  = {dev50_L};  const DEV50_Q  = {dev50_Q};
const DIA50_L  = {dia50_L};  const DIA50_Q  = {dia50_Q};  const DIA50_V = {dia50_V};
const CAT50_L  = {cat50_L};  const CAT50_Q  = {cat50_Q};
const HID50_L  = {hid50_L};  const HID50_Q  = {hid50_Q};

// ── HELPERS ──────────────────────────────────────────────────────────
function fmtBRL(v) {{
  return 'R$ ' + v.toLocaleString('pt-BR', {{minimumFractionDigits:2, maximumFractionDigits:2}});
}}
function fmtBRLk(v) {{
  if (v >= 1e6) return 'R$ ' + (v/1e6).toFixed(1).replace('.',',') + 'M';
  if (v >= 1e3) return 'R$ ' + (v/1e3).toFixed(1).replace('.',',') + 'k';
  return fmtBRL(v);
}}
function animateCounter(el, target, formatter, duration=1200) {{
  if (!el) return;
  const start = performance.now();
  function step(now) {{
    const p = Math.min((now - start) / duration, 1);
    const ease = 1 - Math.pow(1 - p, 3);
    el.textContent = formatter(target * ease);
    if (p < 1) requestAnimationFrame(step);
    else el.textContent = formatter(target);
  }}
  requestAnimationFrame(step);
}}

// ── REGISTRAR PLUGIN DATALABELS ───────────────────────────────────────
if (typeof Chart !== 'undefined' && typeof ChartDataLabels !== 'undefined') {{
  Chart.register(ChartDataLabels);
}}

// ── CONFIGURAÇÃO ANTI-COLISÃO PARA DATALABELS ─────────────────────────
function getDatalabelsConfig(tipo, showLabels) {{
  if (showLabels === false || showLabels === 0 || tipo === 'noop') return {{ display: false }};

  if (tipo === 'linha') {{
    // Linhas: rótulo BEM acima do ponto, caixa branca com borda sutil
    return {{
      display: true,
      anchor: 'end', align: 'top', offset: 10,
      font: {{ size: 11, weight: 'bold' }},
      color: '#1A2535',
      backgroundColor: 'rgba(255,255,255,0.95)',
      borderRadius: 4,
      padding: {{ x: 6, y: 3 }},
      borderColor: 'rgba(1,31,63,0.20)',
      borderWidth: 1,
      clamp: true, clip: false,
      formatter: (v) => (!v || v === 0) ? '' : fmtBRLk(v)
    }};
  }}

  // Barras verticais: dentro, centralizado
  return {{
    display: true,
    anchor: 'center', align: 'center', offset: 0,
    font: {{ size: 11, weight: 'bold' }},
    color: '#011F3F',
    backgroundColor: 'rgba(255,255,255,0.88)',
    borderRadius: 3,
    padding: {{ x: 5, y: 2 }},
    borderWidth: 0,
    clamp: true, clip: false,
    formatter: (v, ctx) => {{
      if (!v || v === 0) return '';
      const ds  = ctx.chart.data.datasets[ctx.datasetIndex];
      const vals = ds.data.filter(x => x > 0);
      const max  = vals.length ? Math.max(...vals) : 1;
      // Oculta se barra < 10% do máximo (muito pequena para caber texto)
      if (v / max < 0.10) return '';
      return fmtBRLk(v);
    }}
  }};
}}

// Barras horizontais com formatter customizado
function getDatalabelsHBar(fmt) {{
  // Helper seguro: extrai valor parsed sem lançar TypeError se ctx.parsed for undefined
  function _parsedVal(ctx) {{
    if (!ctx || !ctx.parsed) return 0;
    return ctx.parsed.x != null ? ctx.parsed.x : (ctx.parsed.y != null ? ctx.parsed.y : 0);
  }}
  function _maxVal(ctx) {{
    try {{
      const ds = ctx.chart.data.datasets[ctx.datasetIndex];
      return Math.max(...ds.data.filter(x => x > 0).concat([1]));
    }} catch(e) {{ return 1; }}
  }}
  return {{
    display: (ctx) => {{
      try {{ return _parsedVal(ctx) > 0; }} catch(e) {{ return false; }}
    }},
    clamp: false, clip: false,
    anchor: (ctx) => {{
      return _parsedVal(ctx) / _maxVal(ctx) < 0.15 ? 'end' : 'end';
    }},
    align: (ctx) => {{
      return _parsedVal(ctx) / _maxVal(ctx) < 0.15 ? 'right' : 'start';
    }},
    offset: (ctx) => {{
      return _parsedVal(ctx) / _maxVal(ctx) < 0.15 ? 4 : 6;
    }},
    font: {{ size: 11, weight: 'bold' }},
    color: (ctx) => {{
      return _parsedVal(ctx) / _maxVal(ctx) < 0.15 ? '#011F3F' : '#fff';
    }},
    textShadowBlur: (ctx) => {{
      return _parsedVal(ctx) / _maxVal(ctx) < 0.15 ? 0 : 3;
    }},
    textShadowColor: 'rgba(0,0,0,0.5)',
    backgroundColor: (ctx) => {{
      return _parsedVal(ctx) / _maxVal(ctx) < 0.15 ? 'rgba(255,255,255,0.92)' : 'transparent';
    }},
    borderRadius: 3,
    borderWidth: 0,
    padding: {{ x: 4, y: 2 }},
    formatter: (v, ctx) => {{
      if (!v || v === 0) return '';
      return typeof fmt === 'function' ? fmt(v) : fmtBRLk(v);
    }}
  }};
}}

// Pizza/Doughnut: % dentro da fatia com sombra forte
function getDatalabelsPie(total, fmt) {{
  const _tot = total;
  return {{
    display: true,
    // Posição FIXA no centro de cada fatia  funciona em todas as versões do plugin
    anchor: 'center',
    align: 'center',
    offset: 0,
    // Font adaptável: maior para fatias grandes, menor para pequenas
    font: (ctx) => {{
      if (!ctx || ctx.parsed == null) return {{ size: 9, weight: 'bold' }};
      const tot = _tot || ctx.chart.data.datasets[0].data.reduce((a,b)=>a+(b||0),0);
      const pct = tot > 0 ? ctx.parsed / tot * 100 : 0;
      return {{ size: pct >= 15 ? 13 : pct >= 5 ? 11 : 9, weight: 'bold' }};
    }},
    color: '#fff',
    textShadowBlur: 8,
    textShadowColor: 'rgba(0,0,0,0.85)',
    backgroundColor: 'rgba(0,0,0,0.30)',
    borderRadius: 3,
    borderWidth: 0,
    padding: {{ x: 4, y: 2 }},
    clamp: true,
    clip: false,
    formatter: (value, ctx) => {{
      if (!value || value === 0) return '';
      const tot = _tot || ctx.chart.data.datasets[0].data.reduce((a,b)=>a+(b||0),0);
      const pct = tot > 0 ? (value / tot * 100).toFixed(1) : '0.0';
      // Mostra tudo acima de 0.5%
      if (parseFloat(pct) < 0.5) return '';
      return pct + '%';
    }}
  }};
}}

// Pizza/Doughnut (inteligente): % dentro se couber, fora com linha se for pequena
function getDatalabelsPieOutside(total) {{
  const _tot = total;
  return {{
    display: true,
    anchor: (ctx) => {{
      const tot = _tot || ctx.chart.data.datasets[0].data.reduce((a,b)=>a+(b||0),0);
      const pct = tot > 0 ? (ctx.parsed || 0) / tot * 100 : 0;
      // Fatias > 8% dentro da fatia
      return pct > 8 ? 'center' : 'end';
    }},
    align: (ctx) => {{
      const tot = _tot || ctx.chart.data.datasets[0].data.reduce((a,b)=>a+(b||0),0);
      const pct = tot > 0 ? (ctx.parsed || 0) / tot * 100 : 0;
      return pct > 8 ? 'center' : 'end';
    }},
    offset: (ctx) => {{
      const tot = _tot || ctx.chart.data.datasets[0].data.reduce((a,b)=>a+(b||0),0);
      const pct = tot > 0 ? (ctx.parsed || 0) / tot * 100 : 0;
      return pct > 8 ? 0 : 18; // Fatias pequenas saem mais para longe
    }},
    clamp: false,
    clip: false,
    font: (ctx) => {{
      const tot = _tot || ctx.chart.data.datasets[0].data.reduce((a,b)=>a+(b||0),0);
      const pct = tot > 0 ? (ctx.parsed || 0) / tot * 100 : 0;
      if (pct > 8) return {{ size: 12, weight: 'bold' }};
      return {{ size: 10, weight: '600' }};
    }},
    color: (ctx) => {{
      const tot = _tot || ctx.chart.data.datasets[0].data.reduce((a,b)=>a+(b||0),0);
      const pct = tot > 0 ? (ctx.parsed || 0) / tot * 100 : 0;
      if (pct > 8) return '#fff';
      const ds = ctx.chart.data.datasets[ctx.datasetIndex];
      const colors = ds.backgroundColor || [];
      return Array.isArray(colors) ? (colors[ctx.dataIndex] || '#011F3F') : (colors || '#011F3F');
    }},
    textStrokeColor: (ctx) => {{
      const tot = _tot || ctx.chart.data.datasets[0].data.reduce((a,b)=>a+(b||0),0);
      const pct = tot > 0 ? (ctx.parsed || 0) / tot * 100 : 0;
      return pct > 8 ? 'rgba(0,0,0,0.4)' : '#fff';
    }},
    textStrokeWidth: (ctx) => {{
      const tot = _tot || ctx.chart.data.datasets[0].data.reduce((a,b)=>a+(b||0),0);
      const pct = tot > 0 ? (ctx.parsed || 0) / tot * 100 : 0;
      return pct > 8 ? 2 : 2.5;
    }},
    backgroundColor: (ctx) => {{
      const tot = _tot || ctx.chart.data.datasets[0].data.reduce((a,b)=>a+(b||0),0);
      const pct = tot > 0 ? (ctx.parsed || 0) / tot * 100 : 0;
      if (pct > 8) return 'transparent';
      const ds = ctx.chart.data.datasets[ctx.datasetIndex];
      const colors = ds.backgroundColor || [];
      const color = Array.isArray(colors) ? (colors[ctx.dataIndex] || '#ccc') : (colors || '#ccc');
      return color;
    }},
    borderRadius: 3,
    padding: (ctx) => {{
      const tot = _tot || ctx.chart.data.datasets[0].data.reduce((a,b)=>a+(b||0),0);
      const pct = tot > 0 ? (ctx.parsed || 0) / tot * 100 : 0;
      return pct > 8 ? {{ x: 0, y: 0 }} : {{ x: 5, y: 3 }};
    }},
    formatter: (value) => {{
      if (!value || value <= 0 || !_tot || _tot <= 0) return '';
      return (value / _tot * 100).toFixed(1) + '%';
    }}
  }};
}}


// ── TABS ─────────────────────────────────────────────────────────────
let _builtTabs = {{}};

function switchTab(tab) {{
  console.log(`[switchTab] Abrindo aba: ${{tab}}`);
  // Oculta todos os painéis
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));

  const panel = document.getElementById('panel-' + tab);
  console.log(`[switchTab] Panel encontrado para ${{tab}}:`, panel ? 'SIM' : 'NÃO');
  if (panel) {{
    panel.classList.add('active');
    console.log(`[switchTab] Panel ativado, display agora é:`, window.getComputedStyle(panel).display);
  }}

  const btn = document.querySelector('[data-tab="' + tab + '"]');
  console.log(`[switchTab] Botão encontrado para ${{tab}}:`, btn ? 'SIM' : 'NÃO');
  if (btn) btn.classList.add('active');

  // Build charts na primeira vez que a tab é aberta
  if (!_builtTabs[tab]) {{
    _builtTabs[tab] = true;
    const _safe = (fn, name) => {{ try {{ fn(); }} catch(e) {{ console.error('[' + name + ']', e); }} }};
    // setTimeout garante que o painel já está visível (display:block) antes de renderizar os canvas
    setTimeout(() => {{
      if (tab === '8117') {{
        _safe(buildCharts8117,  'buildCharts8117');
        _safe(initKPIs8117,    'initKPIs8117');
        _safe(initInsights8117, 'initInsights8117');
      }}
      if (tab === '8121') {{
        _safe(buildCharts8121,  'buildCharts8121');
        _safe(initInsights8121, 'initInsights8121');
      }}
      if (tab === '8091') {{
        _safe(initKPIs8091,    'initKPIs8091');
        _safe(buildCharts8091, 'buildCharts8091');
      }}
      if (tab === '50012') {{
        _safe(initKPIs50012,    'initKPIs50012');
        _safe(buildCharts50012, 'buildCharts50012');
        _safe(fillMiniInsights50012, 'fillMiniInsights50012');
        // Timeout extra para garantir canvas visível antes de renderizar
        setTimeout(() => {{
          if (typeof DIA50_L !== 'undefined' && DIA50_L && DIA50_L.length > 0)
            _safe(() => rebuildChartDias(modoLeiturasDias || 'daily'), 'rebuildChartDias');
        }}, 150);
      }}
      if (tab === 'mapeamento') {{
        // Timeout maior para garantir que o painel está visível
        setTimeout(() => {{
          _safe(buildChartHidEspelho,      'buildChartHidEspelho');
          _safe(buildChartHidRetroativo,   'buildChartHidRetroativo');
          _safe(buildChartLeitOS,          'buildChartLeitOS');
          _safe(buildChartDistribuicao,    'buildChartDistribuicao');
        }}, 100);
      }}
    }}, 50);
  }}

  setTimeout(() => {{
    document.querySelectorAll('#panel-' + tab + ' .kpi-card, #panel-' + tab + ' .chart-card').forEach(el => {{
      el.classList.add('visible');
    }});
    document.querySelectorAll('#panel-' + tab + ' .insight-item').forEach((el, i) => {{
      setTimeout(() => el.classList.add('visible'), i * 60);
    }});
  }}, 60);
}}

// ── CHART.JS DEFAULTS ────────────────────────────────────────────────
if (typeof Chart !== 'undefined') {{
  Chart.defaults.font.family = "'Segoe UI', system-ui, sans-serif";
  Chart.defaults.color = '#6B7A96';
  Chart.defaults.plugins.legend.labels.boxWidth = 12;
  Chart.defaults.plugins.legend.labels.padding = 14;
  // Tooltip global: animação rápida + nunca cortado
  Chart.defaults.plugins.tooltip.animation = {{ duration: 120 }};
  Chart.defaults.plugins.tooltip.position  = 'nearest';
}}

// ── TOOLTIP EXTERNO  nunca cortado por overflow ──────────────────────
// Cria um div flutuante no body que recebe os dados do Chart.js
(function() {{
  const tip = document.createElement('div');
  tip.id = 'chart-tooltip-ext';
  tip.style.cssText = [
    'position:fixed', 'z-index:99999', 'pointer-events:none',
    'background:rgba(1,31,63,0.96)', 'color:#fff',
    'border-radius:10px', 'padding:12px 16px',
    'font-size:13px', 'font-family:Segoe UI,system-ui,sans-serif',
    'box-shadow:0 8px 32px rgba(0,0,0,0.35)',
    'border:1px solid rgba(255,255,255,0.15)',
    'max-width:320px', 'line-height:1.55',
    'transition:opacity 0.12s ease',
    'opacity:0', 'display:block'
  ].join(';');
  document.body.appendChild(tip);
  window._chartTooltip = tip;
}})();

function externalTooltip(context) {{
  const tip   = window._chartTooltip;
  const model = context.tooltip;

  if (model.opacity === 0) {{
    tip.style.opacity = '0';
    return;
  }}

  // Monta HTML do tooltip
  let html = '';
  if (model.title && model.title.length) {{
    // _fmtTitleComparacao pode retornar array (quando há comparativos) ou string simples
    html += '<div style="font-weight:700;margin-bottom:6px;border-bottom:1px solid rgba(255,255,255,0.2);padding-bottom:5px;line-height:1.65">';
    model.title.forEach((line, i) => {{
      if (i === 0) {{
        html += '<div style="font-size:13px">' + line + '</div>';
      }} else {{
        html += '<div style="font-size:11px;opacity:0.78;margin-top:1px">' + line + '</div>';
      }}
    }});
    html += '</div>';
  }}
  if (model.body) {{
    // Separador antes de CADA grupo  atuais, month, year
    // lastTipo: null = não iniciou, 'current' = séries atuais, 'year'/'month' = comparação
    let lastTipo = null;
    model.body.forEach((b, i) => {{
      const dpIdx = model.dataPoints && model.dataPoints[i] ? model.dataPoints[i].datasetIndex : i;
      const ds    = model.chart ? model.chart.data.datasets[dpIdx] : null;
      const isCompare = ds && ds._isCompare;
      const tipo = isCompare ? (ds._tipo || 'unknown') : 'current';

      // Separador sempre que o grupo muda
      if (tipo !== lastTipo) {{
        html += '<div style="border-top:1px solid rgba(255,255,255,0.18);margin:5px 0 4px"></div>';
        lastTipo = tipo;
      }}

      const lc = model.labelColors[i] || {{}};
      const color = (lc.borderColor && lc.borderColor !== 'transparent')
        ? lc.borderColor : (lc.backgroundColor || '#fff');
      const isLine = ds && ds.type === 'line';
      const radius = isLine ? '50%' : '2px';
      const colorBox = color && color !== 'transparent'
        ? `<span style="display:inline-block;width:9px;height:9px;border-radius:${{radius}};background:${{color}};margin-right:7px;vertical-align:middle;flex-shrink:0"></span>`
        : '';

      b.lines.forEach(line => {{
        const opacity  = isCompare ? '0.85' : '1';
        const fontSize = isCompare ? '12px' : '13px';
        html += `<div style="display:flex;align-items:center;opacity:${{opacity}};font-size:${{fontSize}};margin-bottom:2px">${{colorBox}}${{line}}</div>`;
      }});
    }});
  }}
  tip.innerHTML = html;

  // Posiciona: calcula onde colocar em relação à janela
  const canvas = context.chart.canvas;
  const rect   = canvas.getBoundingClientRect();
  const ex     = rect.left + model.caretX;
  const ey     = rect.top  + model.caretY;
  const tw     = tip.offsetWidth  || 220;
  const th     = tip.offsetHeight || 80;
  const pad    = 14;
  const vw     = window.innerWidth;
  const vh     = window.innerHeight;

  let x = ex - tw / 2;            // centralizado no ponto
  let y = ey - th - pad;          // acima do ponto

  // Não sair pela direita
  if (x + tw + pad > vw) x = vw - tw - pad;
  // Não sair pela esquerda
  if (x < pad) x = pad;
  // Não sair pelo topo  coloca abaixo
  if (y < pad) y = ey + pad;
  // Não sair pela base
  if (y + th + pad > vh) y = ey - th - pad;

  tip.style.left    = x + 'px';
  tip.style.top     = y + 'px';
  tip.style.opacity = '1';
}}

const TT = {{
  enabled: false,          // desativa tooltip nativo do canvas
  external: externalTooltip,
  // Mantém callbacks funcionando normalmente
}};

// ── HELPER: título do tooltip com comparação ─────────────────────────────
// Formata o título do tooltip mostrando o período atual e todos os períodos
// de comparação ativos, um por linha.
//
// Exemplos:
//   sem cmp  → " 03/2026"
//   year     → " 03/2026    03/2025  (mesmo período 1 ano)"
//   year_month → " 03/2026    Mar/2025  (mesmo mês 1 ano)"
//   month    → " 03/2026   ↩ 02/2026  (mês anterior)"
//   múltiplos → linhas separadas para cada comparativo
function _fmtTitleComparacao(items, datasets, mensal) {{
  if (!items || !items.length) return '';

  const mNames = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  const lbl = items[0].label || '';

  // Grupos de comparação únicos (evita duplicar ano/mês quando há pares Diária+Acumulado)
  const cmpDatasets = (datasets || []).filter(ds => ds._isCompare);
  const tiposVistos = new Set();
  const cmpUnicos = cmpDatasets.filter(ds => {{
    const k = ds._tipo + '|' + (ds._periodo || '');
    if (tiposVistos.has(k)) return false;
    tiposVistos.add(k); return true;
  }});
  const temCmp = cmpUnicos.length > 0;

  // Formata a data atual conforme o modo
  function _fmtData(s, shift) {{
    shift = shift || 0; // shift em meses para calcular a referência
    const mDia = s.match(/^(\\d{{2}})\\/(\\d{{2}})\\/(\\d{{4}})$/);
    if (mDia) {{
      let d = new Date(+mDia[3], +mDia[2]-1, +mDia[1]);
      if (shift !== 0) d = new Date(d.getFullYear() + Math.trunc(shift/12), d.getMonth() + (shift%12), d.getDate());
      if (mensal === 'yearly') return String(d.getFullYear());
      if (mensal === true)     return String(d.getMonth()+1).padStart(2,'0') + '/' + d.getFullYear();
      return d.getDate().toString().padStart(2,'0') + ' ' + mNames[d.getMonth()] + ' ' + d.getFullYear();
    }}
    const mMes = s.match(/^(\\d{{2}})\\/(\\d{{4}})$/);
    if (mMes) {{
      let d = new Date(+mMes[2], +mMes[1]-1, 1);
      if (shift !== 0) d = new Date(d.getFullYear() + Math.trunc(shift/12), d.getMonth() + (shift%12), 1);
      if (mensal === 'yearly') return String(d.getFullYear());
      return String(d.getMonth()+1).padStart(2,'0') + '/' + d.getFullYear();
    }}
    if (/^\\d{{4}}$/.test(s)) {{
      return String(parseInt(s) + Math.trunc(shift/12));
    }}
    return s;
  }}

  const dataAtual = _fmtData(lbl);
  if (!temCmp) return ' ' + dataAtual;

  // Linha base: período atual
  const linhas = [' ' + dataAtual];

  // Uma linha por tipo de comparação ativo
  const iconeMap = {{ year: '', year_month: '', prev_month_year: '', month: '↩' }};
  const labelMap = {{ year: 'mesmo período 1 ano', year_month: 'mesmo mês 1 ano', prev_month_year: 'mês ant. 1 ano', month: 'mês anterior' }};
  const shiftMap = {{ year: -12, year_month: -12, prev_month_year: -13, month: -1 }};

  cmpUnicos.forEach(ds => {{
    const tipo   = ds._tipo || '';
    const icone  = iconeMap[tipo] || '↔';
    const shift  = shiftMap[tipo] || 0;
    const dataRef = _fmtData(lbl, shift);
    // Só mostra se conseguiu calcular uma data de referência válida
    if (dataRef && dataRef !== lbl) linhas.push(`${{icone}} ${{dataRef}}`);
    else linhas.push(`${{icone}} ${{ds._periodo || ''}}`);
  }});

  return linhas;
}}


// ── ZOOM CONFIG ───────────────────────────────────────────────────────
// Aplica zoom a um gráfico: scroll = zoom, arrastar = pan, duplo clique = reset
// Helper: cria gráfico
function makeChart(ctx, config) {{
  if (!ctx) return null;
  if (typeof Chart === 'undefined') {{
    console.error('Chart.js não foi carregado');
    return null;
  }}
  // Hover responsivo: highlight instantâneo
  config.options = config.options || {{}};
  config.options.hover = config.options.hover || {{}};
  config.options.hover.animationDuration = 0;
  // Tooltip: posicionador inteligente em todos os gráficos
  config.options.plugins = config.options.plugins || {{}};
  config.options.plugins.tooltip = config.options.plugins.tooltip || {{}};
  return new Chart(ctx, config);
}}

// ── CHARTS 8117 ───────────────────────────────────────────────────────
// ── Sort cronológico para labels mm/yyyy ou yyyy ───────────────────────
function _sortCrono(a, b) {{
  // Converte mm/yyyy → número yyyymm para comparação correta
  const toNum = s => {{
    const mMes = s.match(/^(\\d{{2}})\\/(\\d{{4}})$/);
    if (mMes) return parseInt(mMes[2]) * 100 + parseInt(mMes[1]);
    const mAno = s.match(/^(\\d{{4}})$/);
    if (mAno) return parseInt(mAno[1]) * 100;
    // dd/mm/yyyy
    const mDia = s.match(/^(\\d{{2}})\\/(\\d{{2}})\\/(\\d{{4}})$/);
    if (mDia) return parseInt(mDia[3]) * 10000 + parseInt(mDia[2]) * 100 + parseInt(mDia[1]);
    return 0;
  }};
  return toNum(a) - toNum(b);
}}

// ── Agrega dados diários em mensais ────────────────────────────────────
// acValues: se fornecido, usa o último valor acumulado de cada mês (mais preciso)
function _agruparMensal(datas, valores, acValues) {{
  const mapaVal  = {{}};
  const mapaAcum = {{}};  // último acumulado de cada mês
  const mapaOrdem = {{}};  // ordem de inserção
  let ordem = 0;

  datas.forEach((d, i) => {{
    const partes = d.split('/');
    const chave  = partes.length >= 3 ? partes[1] + '/' + partes[2] : d;
    if (!(chave in mapaVal)) {{ mapaVal[chave] = 0; mapaOrdem[chave] = ordem++; }}
    mapaVal[chave] += (valores[i] || 0);
    if (acValues) mapaAcum[chave] = acValues[i]; // sobrescreve com o último do mês
  }});

  // Ordena pela ordem de aparecimento
  const labels = Object.keys(mapaVal).sort((a, b) => mapaOrdem[a] - mapaOrdem[b]);
  const vals   = labels.map(k => Math.round(mapaVal[k] * 100) / 100); // evita float imprecision

  // Acumulado: usa o último ARR_AC do mês se disponível, senão recalcula
  let acum;
  if (acValues && labels.every(k => mapaAcum[k] !== undefined)) {{
    acum = labels.map(k => Math.round(mapaAcum[k] * 100) / 100);
  }} else {{
    let soma = 0;
    acum = vals.map(v => {{ soma += v; return Math.round(soma * 100) / 100; }});
  }}

  return {{ labels, vals, acum }};
}}

function _agruparAnual(datas, valores, acValues) {{
  // Agrega por ano: soma os valores e pega o último acumulado de cada ano
  const mapaVal  = {{}};
  const mapaAcum = {{}};
  const mapaOrdem = {{}};
  let ordem = 0;

  datas.forEach((d, i) => {{
    const partes = d.split('/');
    // Suporta dd/mm/yyyy ou mm/yyyy
    const chave = partes.length >= 3 ? partes[2] : (partes.length === 2 ? partes[1] : d);
    if (!(chave in mapaVal)) {{ mapaVal[chave] = 0; mapaOrdem[chave] = ordem++; }}
    mapaVal[chave] += (valores[i] || 0);
    if (acValues) mapaAcum[chave] = acValues[i];
  }});

  const labels = Object.keys(mapaVal).sort((a, b) => mapaOrdem[a] - mapaOrdem[b]);
  const vals   = labels.map(k => Math.round(mapaVal[k] * 100) / 100);

  let acum;
  if (acValues && labels.every(k => mapaAcum[k] !== undefined)) {{
    acum = labels.map(k => Math.round(mapaAcum[k] * 100) / 100);
  }} else {{
    let soma = 0;
    acum = vals.map(v => {{ soma += v; return Math.round(soma * 100) / 100; }});
  }}

  return {{ labels, vals, acum }};
}}

let _chart8117 = null;
let _modo8117  = 'daily';

function toggle8117(modo) {{
  _modo8117 = modo;
  const bd = document.getElementById('btn8117-d');
  const bm = document.getElementById('btn8117-m');
  const by = document.getElementById('btn8117-y');
  [['daily',bd],['monthly',bm],['yearly',by]].forEach(([m,b]) => {{
    if (!b) return;
    b.style.background = modo===m ? '#1565C0' : 'transparent';
    b.style.color      = modo===m ? '#fff'    : '#6B7A96';
  }});
  const sub = document.getElementById('sub-arrec-diario');
  if (sub) sub.textContent = modo==='daily'   ? 'Evolução diária  Macro 8117'
                           : modo==='monthly' ? 'Evolução mensal (soma por mês)  Macro 8117'
                           :                   'Evolução anual (soma por ano)  Macro 8117';
  _rebuild8117();
}}

// ── Mini-insight de tendência  atualiza ao trocar Diário ↔ Mensal ──────────
function _atualizarMiniInsight8117() {{
  const bloco = document.querySelector('#panel-8117 .mini-insights-block');
  if (!bloco) return;

  const mensal = _modo8117 === 'monthly';
  const yearly = _modo8117 === 'yearly';

  // Usa dados do range atual do slider; fallback para dataset completo
  const rDatas = (_rc8117Range.datas && _rc8117Range.datas.length) ? _rc8117Range.datas : DATAS_17;
  const rDataD = (_rc8117Range.dataD && _rc8117Range.dataD.length) ? _rc8117Range.dataD : ARR_D;
  const rDataAC = (_rc8117Range.dataAC && _rc8117Range.dataAC.length) ? _rc8117Range.dataAC : ARR_AC;

  // ── VIEW DIÁRIA ─────────────────────────────────────────────────────────
  if (!mensal && !yearly) {{
    const validos = rDataD.map((v, i) => [rDatas[i], v]).filter(x => x[1] > 0);
    if (!validos.length) {{ bloco.innerHTML = ''; return; }}

    const total = rDataAC.length ? rDataAC[rDataAC.length - 1] : validos.reduce((s, x) => s + x[1], 0);
    const media = validos.reduce((s, x) => s + x[1], 0) / validos.length;
    const melhor = validos.reduce((a, b) => b[1] > a[1] ? b : a);
    const pior   = validos.reduce((a, b) => b[1] < a[1] ? b : a);
    const acima  = validos.filter(x => x[1] > media).length;
    const semArr = rDataD.filter(v => v === 0).length;

    let html = _miniCard('','green',`<strong>Melhor dia:</strong> ${{melhor[0]}} com ${{fmtBRL(melhor[1])}} (${{total?(melhor[1]/total*100).toFixed(1):0}}% do total)`)
             + _miniCard('','',`<strong>Média diária:</strong> ${{fmtBRL(media)}}  ${{acima}} de ${{validos.length}} dias acima da média (${{(acima/validos.length*100).toFixed(0)}}%)`)
             + _miniCard('','red',`<strong>Menor dia:</strong> ${{pior[0]}} com ${{fmtBRL(pior[1])}} (${{total?(pior[1]/total*100).toFixed(1):0}}% do total)`);
    if (semArr > 0)
      html += _miniCard('ℹ','',`<strong>${{semArr}} dia(s) sem arrecadação</strong> no período selecionado`);
    bloco.innerHTML = html;
    return;
  }}

  // ── VIEW MENSAL ou ANUAL ────────────────────────────────────────────────
  const agg = yearly ? _agruparAnual(rDatas, rDataD, rDataAC)
                     : _agruparMensal(rDatas, rDataD, rDataAC);
  const labels = agg.labels; const vals = agg.vals;
  if (!vals.length) {{ bloco.innerHTML = ''; return; }}

  const total = vals.reduce((s, v) => s + v, 0);
  const media = total / vals.length;
  const melhor = labels[vals.indexOf(Math.max(...vals))];
  const pior   = labels[vals.indexOf(Math.min(...vals))];
  const melhorV = Math.max(...vals);
  const piorV   = Math.min(...vals);
  const unid = yearly ? 'ano' : 'mês';

  let html = _miniCard('','green',`<strong>Melhor ${{unid}}:</strong> ${{melhor}} com ${{fmtBRL(melhorV)}} (${{(melhorV/total*100).toFixed(1)}}% do total)`)
           + _miniCard('','',`<strong>Média ${{unid}}al:</strong> ${{fmtBRL(media)}}  ${{vals.length}} ${{unid}}(s) no período`);

  if (vals.length >= 2) {{
    const ult  = vals[vals.length - 1];
    const ante = vals[vals.length - 2];
    const var_pct = ante > 0 ? (ult - ante) / ante * 100 : 0;
    const sinal = var_pct >= 0 ? '' : '';
    const cor   = var_pct >= 5 ? 'green' : var_pct <= -5 ? 'red' : 'orange';
    const desc  = var_pct >= 0 ?
       `${{var_pct.toFixed(1)}}% <strong>acima</strong> de ${{labels[labels.length-2]}}`
      : `${{Math.abs(var_pct).toFixed(1)}}% <strong>abaixo</strong> de ${{labels[labels.length-2]}}`;
    html += _miniCard(sinal, cor,
      `<strong>Variação:</strong> ${{labels[labels.length-1]}} (${{fmtBRL(ult)}}) está ${{desc}} (${{fmtBRL(ante)}})`);
  }}

  bloco.innerHTML = html;
}}

// Helper: gera HTML de um mini-card (espelha _mini_insight do Python)
function _miniCard(emoji, cor, texto) {{
  const corMap = {{
    green:  ['rgba(46,125,50,0.08)',  '#2E7D32', 'rgba(46,125,50,0.35)'],
    red:    ['rgba(198,40,40,0.08)',  '#C62828', 'rgba(198,40,40,0.35)'],
    orange: ['rgba(230,115,0,0.08)',  '#E65100', 'rgba(230,115,0,0.35)'],
    blue:   ['rgba(21,101,192,0.08)', '#1565C0', 'rgba(21,101,192,0.35)'],
  }};
  const [bg, txtCor, border] = corMap[cor] || ['rgba(100,116,139,0.07)', '#475569', 'rgba(100,116,139,0.25)'];
  return `<div style="display:flex;align-items:flex-start;gap:8px;padding:9px 13px;`
       + `background:${{bg}};border-left:3px solid ${{border}};border-radius:0 7px 7px 0;`
       + `margin-top:6px;font-size:0.82rem;line-height:1.45;color:${{txtCor}};">`
       + `<span style="font-size:1rem;flex-shrink:0;margin-top:1px">${{emoji}}</span>`
       + `<span>${{texto}}</span></div>`;
}}

// Range atual do slider 8117 (i0, i1 no array de labels/dados)
let _rc8117Range = {{ i0: 0, i1: -1 }};

// Inicializa RangeCompare para 8117 (chamado 1x no buildCharts8117)
function _initRC8117(labels, dataD, dataAC, labelD) {{
  // No modo diário, inicia mostrando os últimos ~90 dias por padrão
  const defaultStart = (_modo8117 === 'daily' && labels.length > 90)
     ? labels.length - 90 : 0;
  buildRangeCompare({{
    containerId: 'rc-8117-arrec',
    labels:   labels,
    datasets: [
      {{ label: labelD, data: dataD }},
      {{ label: 'Acumulado', data: dataAC }},
    ],
    defaultStart,
    rebuildFn: (filtL, filtDs, cmpDs, i0, i1) => {{
      // Guarda o range atual para uso nos mini-insights
      _rc8117Range = {{ i0: i0 < 0 ? 0 : i0, i1: i1 < 0 ? (filtDs[0].data.length - 1) : i1,
                        datas: filtL,
                        dataD: filtDs[0].data,
                        dataAC: filtDs[1].data }};
      _atualizarMiniInsight8117();
      _drawChart8117(filtL, filtDs[0].data, filtDs[1].data, labelD, cmpDs);
    }},
  }});
}}

function _rebuild8117() {{
  const ctx = document.getElementById('chart-arrec-diario');
  if (!ctx) return;
  if (_chart8117) {{ _chart8117.destroy(); _chart8117 = null; }}

  const mensal = _modo8117 === 'monthly' ? true : (_modo8117 === 'yearly' ? 'yearly' : false);
  let labels, dataD, dataAC, labelD;

  if (_modo8117 === 'yearly') {{
    const agg = _agruparAnual(DATAS_17, ARR_D, ARR_AC);
    labels = agg.labels; dataD = agg.vals; dataAC = agg.acum;
    labelD = 'Arrecadação Anual';
  }} else if (_modo8117 === 'monthly') {{
    const agg = _agruparMensal(DATAS_17, ARR_D, ARR_AC);
    labels = agg.labels; dataD = agg.vals; dataAC = agg.acum;
    labelD = 'Arrecadação Mensal';
  }} else {{
    labels = DATAS_17; dataD = ARR_D; dataAC = ARR_AC;
    labelD = 'Arrecadação Diária';
  }}
  // Reinicializa o RangeCompare com os dados do modo atual
  _initRC8117(labels, dataD, dataAC, labelD);
}}

function _drawChart8117(labels, dataD, dataAC, labelD, cmpDs) {{
  const ctx = document.getElementById('chart-arrec-diario');
  if (!ctx) return;
  if (_chart8117) {{ _chart8117.destroy(); _chart8117 = null; }}

  const mensal = _modo8117 === 'monthly' ? true : (_modo8117 === 'yearly' ? 'yearly' : false);
  const modoAgregado = mensal !== false;

  // Recalcula acumulado relativo ao período selecionado (começa do zero)
  // independente do acumulado histórico passado pelo slice
  let soma = 0;
  const dataACRel = dataD.map(v => {{ soma += (v || 0); return Math.round(soma * 100) / 100; }});

  const totalAcum = dataACRel.length ? dataACRel[dataACRel.length - 1] : 0;
  const maxD = dataD.length ? Math.max(...dataD) : 1;

  const c2d = ctx.getContext('2d');
  const gradBar  = c2d.createLinearGradient(0, 0, 0, 300);
  gradBar.addColorStop(0, 'rgba(21,101,192,0.85)');
  gradBar.addColorStop(1, 'rgba(21,101,192,0.3)');
  const gradAcum = c2d.createLinearGradient(0, 0, 0, 300);
  gradAcum.addColorStop(0, 'rgba(255,143,0,0.18)');
  gradAcum.addColorStop(1, 'rgba(255,143,0,0.0)');

  // Agrupa cmpDs por _tipo para garantir ordem: todos do year juntos, todos do month juntos
  // Dentro de cada tipo: Acumulado (label contém 'Acumulado') → Diária
  const hexToRgba8117 = (hex, a) => {{
    if (!hex || !hex.startsWith('#')) return `rgba(124,58,237,${{a}})`;
    const r = parseInt(hex.slice(1,3),16), g = parseInt(hex.slice(3,5),16), b = parseInt(hex.slice(5,7),16);
    return `rgba(${{r}},${{g}},${{b}},${{a}})`;
  }};
  const cmpPorTipo = {{}};
  (cmpDs || []).forEach(d => {{
    const t = d._tipo || 'unknown';
    if (!cmpPorTipo[t]) cmpPorTipo[t] = {{ acum: [], diaria: [] }};
    if (d.label && d.label.includes('Acumulado')) cmpPorTipo[t].acum.push(d);
    else cmpPorTipo[t].diaria.push(d);
  }});

  // Ordem dos tipos na tooltip: month primeiro, year depois (conforme _apply)
  const tiposOrdem = ['month', 'year', 'year_month', 'prev_month_year', 'unknown'];
  const cmpOrdenados = tiposOrdem.flatMap(t => cmpPorTipo[t] ?
     [...cmpPorTipo[t].acum, ...cmpPorTipo[t].diaria]
    : []
  );

  const datasetsOrdenados = [
    // 1. Acumulado atual (linha)  sempre primeiro na legenda
    {{ label: 'Acumulado', data: dataACRel, type: 'line', borderColor: '#FF8F00',
       backgroundColor: gradAcum, borderWidth: 2.5, pointRadius: modoAgregado ? 5 : 3,
       pointHoverRadius: 7, pointBackgroundColor: '#FF8F00', fill: true,
       tension: 0.35, yAxisID: 'y2', order: 1 }},
    // 2. Arrecadação Diária atual (barra)  segundo
    {{ label: labelD, data: dataD, backgroundColor: gradBar,
       borderColor: '#1565C0', borderWidth: 1, borderRadius: 5, yAxisID: 'y', order: 2 }},
    // 3. Comparações ordenadas por grupo (Acumulado+Diária de cada tipo juntos)
    ...cmpOrdenados.map((d, i) => {{
      const isAcum = d.label && d.label.includes('Acumulado');
      if (isAcum) return {{
        ...d, type: 'line', yAxisID: 'y2', order: 3 + i, _isCompare: true,
        borderWidth: 2, pointRadius: modoAgregado ? 4 : 2, fill: false, tension: 0.35,
      }};
      return {{
        ...d, type: 'bar', yAxisID: 'y', order: 3 + i, _isCompare: true,
        borderWidth: 1.5,
        backgroundColor: hexToRgba8117(d.borderColor, 0.30),
        borderRadius: 3,
      }};
    }}),
  ];

  _chart8117 = makeChart(ctx, {{
    type: 'bar',
    data: {{
      labels,
      datasets: datasetsOrdenados,
    }},
    options: {{
      responsive: true,
      interaction: {{ mode: 'index', intersect: false }},
      animation: {{ duration: 900, easing: 'easeOutQuart' }},
      plugins: {{
        legend: {{ position: 'top' }},
        datalabels: modoAgregado ? {{
          clamp: true, clip: false,
          display: (ctx) => {{
            const v = ctx.dataset.data[ctx.dataIndex];
            if (!v || v === 0) return false;
            // Barras (atuais e anteriores): mostrar se >= 10% do máximo da própria série
            if (ctx.dataset.type !== 'line') {{
              const vals = ctx.dataset.data.filter(x => x > 0);
              const max  = vals.length ? Math.max(...vals) : 1;
              return v / max >= 0.10;
            }}
            // Linhas: sempre mostrar último ponto; demais só se espaço suficiente
            return ctx.dataIndex === ctx.dataset.data.length - 1 || true;
          }},
          // Posição: barras atuais → centro | barras ant → topo (acima) | linhas atuais → topo | linhas ant → baixo
          anchor: (ctx) => {{
            if (ctx.dataset.type !== 'line') return ctx.dataset._isCompare ? 'end' : 'center';
            return 'end';
          }},
          align:  (ctx) => {{
            if (ctx.dataset.type !== 'line') return ctx.dataset._isCompare ? 'top' : 'center';
            return ctx.dataset._isCompare ? 'bottom' : 'top';
          }},
          offset: (ctx) => {{
            if (ctx.dataset.type !== 'line') return ctx.dataset._isCompare ? 4 : 0;
            return 8;
          }},
          font: (ctx) => ({{ size: ctx.dataset._isCompare ? 10 : 11, weight: 'bold' }}),
          color: (ctx) => {{
            if (ctx.dataset._isCompare) return ctx.dataset.borderColor || '#0891B2';
            return ctx.dataset.type === 'line' ? '#1A2535' : '#011F3F';
          }},
          backgroundColor: (ctx) => {{
            if (ctx.dataset._isCompare) return 'rgba(255,255,255,0.92)';
            return ctx.dataset.type === 'line' ? 'rgba(255,255,255,0.96)' : 'rgba(255,255,255,0.88)';
          }},
          borderRadius: (ctx) => ctx.dataset.type === 'line' ? 4 : 3,
          borderColor: (ctx) => {{
            if (ctx.dataset._isCompare) return (ctx.dataset.borderColor || '#0891B2') + '55';
            return ctx.dataset.type === 'line' ? 'rgba(1,31,63,0.20)' : 'transparent';
          }},
          borderWidth: (ctx) => ctx.dataset.type === 'line' ? 1 : 0,
          padding: {{ x: 5, y: 2 }},
          formatter: (v, ctx) => {{
            if (!v || v === 0) return '';
            if (!ctx.dataset._isCompare) return fmtBRLk(v);

            // Calcula o mês de referência dinamicamente para aquele ponto
            const xLabel = ctx.chart.data.labels[ctx.dataIndex] || '';
            const tipo   = ctx.dataset._tipo || '';
            const meses  = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
            const icone  = tipo === 'year' ? '' : '↩';

            // Tenta calcular a data de referência a partir do label do ponto
            let refLabel = '';
            const mMes = xLabel.match(/^(\\d{{2}})\\/(\\d{{4}})$/);
            if (mMes) {{
              const mo = parseInt(mMes[1]) - 1, y = parseInt(mMes[2]);
              if (tipo === 'year') {{
                // Mesmo mês, ano anterior: Nov/2025 → Nov/2024
                refLabel = meses[mo] + '/' + (y - 1);
              }} else if (tipo === 'prev_month_year') {{
                // Mês anterior do ano passado: Mar/2026 → Fev/2025
                const ref = new Date(y - 1, mo - 1, 1);
                refLabel = meses[ref.getMonth()] + '/' + ref.getFullYear();
              }} else {{
                // Mês anterior: Nov/2025 → Out/2025
                const ref = new Date(y, mo - 1, 1);
                refLabel = meses[ref.getMonth()] + '/' + ref.getFullYear();
              }}
            }} else {{
              // Fallback: usa o _periodo estático do dataset
              const periodoMatch = (ctx.dataset.label || '').match(/  (.+)$/);
              refLabel = periodoMatch ? periodoMatch[1] : '';
            }}

            return refLabel
              ? fmtBRLk(v) + '\\n' + icone + ' ' + refLabel
              : fmtBRLk(v);
          }}
        }} : {{ display: false }},
        tooltip: {{ ...TT, callbacks: {{
          title: items => _fmtTitleComparacao(items, datasetsOrdenados, mensal),
          label: c => {{
            const lbl  = c.dataset.label || '';
            const val  = fmtBRL(c.parsed.y);
            const tipo = c.dataset._tipo || '';
            const idx  = c.dataIndex;

            // ── Séries de comparação ──────────────────────────────────────────
            if (c.dataset._isCompare) {{
              const isAcum = lbl.toLowerCase().includes('acumulado');
              const partes    = lbl.split('  ');
              const nomeSerie = partes[0].trim();

              const meta = {{
                year:  {{ icone: '', tag: 'ano passado' }},
                prev_month_year: {{ icone: '', tag: 'mês ant. 1 ano' }},
                month: {{ icone: '↩',  tag: 'mês anterior' }},
              }}[tipo] || {{ icone: '↔', tag: tipo }};

              const tipoSerie = isAcum ? 'Acumulado' : (modoAgregado ? 'Mensal' : 'Diário');

              // Busca o valor atual correspondente (mesmo idx, mesmo tipo de série)
              const dsAtual = datasetsOrdenados.find(d =>
                !d._isCompare &&
                (isAcum ? d.label === 'Acumulado' : d.label === labelD)
              );
              const valAtual = dsAtual ? (dsAtual.data[idx] ?? null) : null;
              const valComp  = c.parsed.y;

              // Calcula variação %
              let deltaStr = '';
              if (valAtual != null && valComp != null && valComp > 0) {{
                const delta = ((valAtual - valComp) / valComp) * 100;
                const sinal = delta >= 0 ? '+' : '';
                const cor   = delta >= 0 ? '▲' : '▼';
                deltaStr = `  ${{cor}} ${{sinal}}${{delta.toFixed(1)}}%`;
              }}

              return ` ${{meta.icone}} ${{tipoSerie}} (${{meta.tag}}): ${{val}}${{deltaStr}}`;
            }}

            // ── Acumulado atual (linha laranja) ───────────────────────────────
            if (lbl === 'Acumulado') {{
              const pct = totalAcum ? (c.parsed.y / totalAcum * 100).toFixed(1) : 0;
              return `  Acumulado: ${{val}}  ·  ${{pct}}% do total`;
            }}

            // ── Arrecadação atual (barras azuis) ──────────────────────────────
            const modLabel = mensal === 'yearly' ? 'Arrecadação do ano'
                           : modoAgregado ? 'Arrecadação do mês'
                           : 'Arrecadação do dia';
            return `  ${{modLabel}}: ${{val}}`;
          }}
        }} }}
      }},
      scales: {{
        x: {{ grid: {{ color: 'rgba(0,0,0,0.04)' }}, ticks: {{ maxRotation: modoAgregado ? 0 : 45 }} }},
        y: {{
          position: 'left',
          display: (cmpDs || []).some(d => !( (cmpDs.indexOf(d) % 2 === 1) )),
          min: 0, max: maxD * 4,
          ticks: {{ callback: v => fmtBRLk(v), color: '#1565C0' }},
          grid: {{ color: 'rgba(0,0,0,0.04)' }},
          title: {{ display: (cmpDs||[]).length > 0, text: 'Diário', font: {{ size: 10 }}, color: '#1565C0' }}
        }},
        y2: {{ position: 'right', ticks: {{ callback: v => fmtBRLk(v) }},
              grid: {{ color: 'rgba(0,0,0,0.05)' }},
              title: {{ display: true, text: 'Acumulado', font: {{ size: 10 }}, color: '#FF8F00' }} }}
      }}
    }}
  }});
}}

function buildCharts8117() {{
  if (!TEM_8117) return;
  _rebuild8117();

  // Pizza formas
  (function() {{
    const ctx = document.getElementById('chart-formas-pizza');
    if (!ctx) return;
    makeChart(ctx, {{
      type: 'doughnut',
      data: {{ labels: FORMAS_N, datasets: [{{ data: FORMAS_V, backgroundColor: FORMAS_C,
               borderColor: '#fff', borderWidth: 3, hoverOffset: 12 }}] }},
      options: {{
        responsive: true, cutout: '55%',
        layout: {{ padding: 34 }},
        animation: {{ duration: 1400, easing: 'easeOutBack' }},
        plugins: {{ legend: {{ position:'right', labels:{{ padding:12, font:{{size:11}}, usePointStyle:true, pointStyleWidth:10, generateLabels: (chart) => {{ 
          const ds = chart.data.datasets[0];
          const total = ds.data.reduce((a,b)=>a+(b||0),0);
          return chart.data.labels.map((label, i) => {{
            const val = ds.data[i] || 0;
            const pct = total > 0 ? (val / total * 100).toFixed(1) : '0.0';
            const displayText = label.length > 18 ? label.slice(0,16)+'' : label;
            return {{ text: `${{displayText}} (${{pct}}%)`, fillStyle: ds.backgroundColor[i], strokeStyle: ds.backgroundColor[i], lineWidth: 0, hidden: false, index: i }};
          }});
        }} }}}},
          datalabels: {{
            display: true,
            anchor: 'center',
            align: 'center',
            offset: 0,
            clamp: true,
            clip: false,
            font: (ctx) => {{
              const total = ctx.chart.data.datasets[0].data.reduce((a,b)=>a+(b||0),0);
              const pct = total > 0 ? (ctx.parsed / total * 100) : 0;
              return {{ size: pct >= 12 ? 13 : 11, weight: 'bold' }};
            }},
            color: '#fff',
            textShadowBlur: 6,
            textShadowColor: 'rgba(0,0,0,0.6)',
            formatter: (value, ctx) => {{
              const total = ctx.chart.data.datasets[0].data.reduce((a,b)=>a+(b||0),0);
              const pct = total > 0 ? (value / total * 100).toFixed(1) : '0.0';
              if (pct < 5) return '';
              return pct + '%';
            }}
          }},
          tooltip: {{ ...TT, callbacks: {{ label: c => ` ${{c.label}}: ${{fmtBRL(c.parsed)}} (${{(c.parsed / FORMAS_V.reduce((a,b)=>a+b,0)*100).toFixed(1)}}%)` }} }} }}
      }}
    }});
  }})();

  // Barras formas horizontal  setTimeout garante que a pizza renderizou primeiro
  setTimeout(function() {{
    const ctx = document.getElementById('chart-formas-barra');
    if (!ctx) return;
    const sorted = FORMAS_N.map((n,i) => ({{n, v:FORMAS_V[i], c:FORMAS_C[i]}})).sort((a,b) => b.v - a.v);

    // Quebra labels longos em múltiplas linhas para liberar espaço do gráfico.
    const WRAP_CHARS_FORMAS = 36; // Ajuste aqui o "X caracteres"
    const MAX_WRAP_LINES_FORMAS = 2;
    const wrapLabel = (text, maxChars = WRAP_CHARS_FORMAS) => {{
      const s = String(text || '').trim();
      if (!s) return [''];
      if (s.length <= maxChars) return [s];
      // Normaliza separadores para facilitar quebra em nomes longos.
      const normalized = s.replace(/\\//g, '/ ').replace(/-/g, '- ');
      const words = normalized.split(/\\s+/);
      const lines = [];
      let cur = '';
      for (const w of words) {{
        const test = cur ? `${{cur}} ${{w}}` : w;
        if (test.length > maxChars && cur) {{
          lines.push(cur);
          cur = w;
        }} else {{
          cur = test;
        }}
      }}
      if (cur) lines.push(cur);
      if (lines.length > MAX_WRAP_LINES_FORMAS) {{
        const trimmed = lines.slice(0, MAX_WRAP_LINES_FORMAS);
        const last = trimmed[MAX_WRAP_LINES_FORMAS - 1] || '';
        trimmed[MAX_WRAP_LINES_FORMAS - 1] = last.length > maxChars - 1
          ?
           last.slice(0, maxChars - 2) + '...'
          : last + '...';
        return trimmed;
      }}
      return lines;
    }};

    // Controle único de densidade visual do gráfico: compacto | normal | espacado.
    const FORMAS_SPACING_MODE = 'normal';
    const FORMAS_SPACING_PRESETS = {{
      compacto: {{ lineHeight: 18, rowGap: 12, rowMin: 48, barThickness: 20, categoryPercentage: 0.80, yTickPadding: 6 }},
      normal:   {{ lineHeight: 19, rowGap: 16, rowMin: 56, barThickness: 18, categoryPercentage: 0.76, yTickPadding: 8 }},
      espacado: {{ lineHeight: 20, rowGap: 18, rowMin: 58, barThickness: 16, categoryPercentage: 0.72, yTickPadding: 10 }}
    }};
    const spacingCfg = FORMAS_SPACING_PRESETS[FORMAS_SPACING_MODE] || FORMAS_SPACING_PRESETS.espacado;

    // Anti-sobreposição: espaço vertical por categoria + gap fixo entre barras.
    const lineHeight = spacingCfg.lineHeight;
    const rowGap     = spacingCfg.rowGap;
    const rowMin     = spacingCfg.rowMin;
    const countLines = (s) => wrapLabel(s).length;

    // "Borda" fixa interna do gráfico (respiro constante acima/abaixo).
    const chartFrameGapY = 20;
    const minDynamicHeight = 290;
    const totalHeight = sorted.reduce((acc, item) => {{
      return acc + Math.max(rowMin, countLines(item.n) * lineHeight + rowGap);
    }}, 0);
    ctx.parentElement.style.height = Math.max(minDynamicHeight, totalHeight + (chartFrameGapY * 2)) + 'px';
    ctx.style.maxHeight = 'none';
    ctx.style.height = '100%';
    ctx.style.width = '100%';

    // Largura do eixo Y proporcional ao maior label (evita espaço vazio excessivo à esquerda).
    const maxLineChars = sorted.reduce((acc, item) => {{
      const lens = wrapLabel(item.n).map(line => line.length);
      const localMax = lens.length ? Math.max(...lens) : 0;
      return Math.max(acc, localMax);
    }}, 0);
    const yAxisWidth = Math.min(240, Math.max(150, Math.round(maxLineChars * 6.2) + 20));

    makeChart(ctx, {{
      type: 'bar',
      data: {{
        labels: sorted.map(x => x.n),
        datasets: [{{ label: 'Valor Arrecadado', data: sorted.map(x => x.v),
          backgroundColor: sorted.map(x => x.c), borderRadius: 5,
          barThickness: spacingCfg.barThickness,
          categoryPercentage: spacingCfg.categoryPercentage,
          barPercentage: 0.9 }}]
      }},
      options: {{
        indexAxis: 'y', responsive: true,
        maintainAspectRatio: false,
        animation: {{ duration: 1200, easing: 'easeOutBack' }},
        layout: {{ padding: {{ left: 4, right: 90, top: chartFrameGapY, bottom: chartFrameGapY }} }},
        plugins: {{ 
          legend: {{ display: false }},
          datalabels: {{
            display: true,
            anchor: 'end',
            align: 'right',
            offset: 6,
            clamp: false,
            clip: false,
            font: {{ size: 11, weight: '600' }},
            color: '#011F3F',
            formatter: (v) => (!v || v === 0) ? '' : fmtBRLk(v)
          }},
          tooltip: {{ ...TT, callbacks: {{ label: c => ' ' + fmtBRL(c.parsed.x) }} }} 
        }},
        scales: {{
          x: {{
            ticks: {{ callback: v => fmtBRLk(v), maxTicksLimit: 6 }},
            grid: {{ color: 'rgba(0,0,0,0.05)' }}
          }},
          y: {{
            grid: {{ display: false }},
            afterFit: (scale) => {{ scale.width = yAxisWidth; }},
            ticks: {{
              font: {{ size: 11, lineHeight: 1.4 }},
              padding: spacingCfg.yTickPadding,
              autoSkip: false,
              maxRotation: 0,
              callback: function(value) {{
                return wrapLabel(this.getLabelForValue(value));
              }}
            }}
          }}
        }}
      }}
    }});
  }}, 150);

  // KPIs são atualizados via initKPIs8117 (chamada separada no switchTab)
}}

function initKPIs8117() {{
  animateCounter(document.getElementById('kpi-total-arr'), TOTAL_ARR, fmtBRL);
  animateCounter(document.getElementById('kpi-media'), MEDIA_ARR, fmtBRL);
  animateCounter(document.getElementById('kpi-melhor-val'), MAIOR_ARR_V, fmtBRL);
  const kpiMelhorData = document.getElementById('kpi-melhor-data');
  if (kpiMelhorData) kpiMelhorData.textContent = MAIOR_ARR_D || '--';
}}

function initInsights8117() {{
  const container = document.getElementById('insights-8117');
  if (!container) return;
  const items = [
    {{ emoji:'', cor:'green', texto:`<strong>Total arrecadado:</strong> ${{fmtBRL(TOTAL_ARR)}} no período.` }},
    {{ emoji:'', cor:'', texto:`<strong>Média diária:</strong> ${{fmtBRL(MEDIA_ARR)}} por dia de pagamento.` }},
    {{ emoji:'', cor:'green', texto:`<strong>Melhor dia:</strong> ${{MAIOR_ARR_D}} com ${{fmtBRL(MAIOR_ARR_V)}}.` }},
    {{ emoji:'', cor:'green', texto:`<strong>Dias acima da média:</strong> ${{ACIMA_MEDIA}} de ${{TOTAL_DIAS}} dias (${{TOTAL_DIAS > 0 ? (ACIMA_MEDIA/TOTAL_DIAS*100).toFixed(0) : 0}}%).` }},
    DIAS_SEM > 0
      ?
       {{ emoji:'', cor:'red',   texto:`<strong>Atenção:</strong> ${{DIAS_SEM}} dia(s) sem arrecadação registrada.` }}
      : {{ emoji:'', cor:'green', texto:'<strong>Sem falhas:</strong> Arrecadação em todos os dias do período.' }},
  ];
  items.forEach((it, i) => {{
    const div = document.createElement('div');
    div.className = `insight-item ${{it.cor}}`;
    div.style.transitionDelay = (i * 70) + 'ms';
    div.innerHTML = `<span class="insight-emoji">${{it.emoji}}</span><div class="insight-text">${{it.texto}}</div>`;
    container.appendChild(div);
  }});
}}

// ── CHARTS 8121 ───────────────────────────────────────────────────────
let _chart8121Comp = null;
let _chart8121Acum = null;
let _modo8121 = 'daily';

function toggle8121(modo) {{
  _modo8121 = modo;
  const bd = document.getElementById('btn8121-d');
  const bm = document.getElementById('btn8121-m');
  const by = document.getElementById('btn8121-y');
  [['daily',bd],['monthly',bm],['yearly',by]].forEach(([m,b]) => {{
    if (!b) return;
    b.style.background = modo===m ? '#1565C0' : 'transparent';
    b.style.color      = modo===m ? '#fff'    : '#6B7A96';
  }});
  const sub = document.getElementById('sub-arrec-fat');
  if (sub) sub.textContent = modo==='daily'   ? 'Comparativo diário  Macro 8121'
                           : modo==='monthly' ? 'Comparativo mensal (soma por mês)  Macro 8121'
                           :                   'Comparativo anual (soma por ano)  Macro 8121';
  _rebuild8121();
}}

function _initRC8121Comp(labels, dataArr, dataFat, labelArr, labelFat) {{
  buildRangeCompare({{
    containerId: 'rc-8121-arrec',
    labels:   labels,
    datasets: [
      {{ label: labelArr, data: dataArr }},
      {{ label: labelFat, data: dataFat }},
    ],
    rebuildFn: (filtL, filtDs, cmpDs) => {{
      _drawChart8121Comp(filtL, filtDs[0].data, filtDs[1].data, labelArr, labelFat, cmpDs);
    }},
  }});
}}

function _drawChart8121Comp(labels, dataArr, dataFat, labelArr, labelFat, cmpDs) {{
  const ctx = document.getElementById('chart-arrec-fat');
  if (!ctx) return;
  if (_chart8121Comp) {{ _chart8121Comp.destroy(); _chart8121Comp = null; }}

  const mensal = _modo8121 === 'monthly' ? true : (_modo8121 === 'yearly' ? 'yearly' : false);
  const modoAgregado = mensal !== false;

  const dlCfg = modoAgregado ? {{
    display: (ctx) => {{
      const v = ctx.dataset.data[ctx.dataIndex];
      if (!v || v === 0) return false;
      const vals = ctx.dataset.data.filter(x => x > 0);
      const max  = vals.length ? Math.max(...vals) : 1;
      return v / max >= 0.10;
    }},
    anchor: 'center',
    align: (ctx) => ctx.dataset._isCompare ? 'top' : 'center',
    offset: (ctx) => ctx.dataset._isCompare ? 4 : 0,
    font: (ctx) => ({{ size: ctx.dataset._isCompare ? 10 : 11, weight: 'bold' }}),
    color: (ctx) => {{
      if (ctx.dataset._isCompare) {{
        // Combina com a cor da série pai: Arrecadação=azul, Faturamento=laranja
        const nome = (ctx.dataset.label || '').toLowerCase();
        return nome.includes('arrecad') ? '#1565C0' : '#E65100';
      }}
      const actualIdx = ctx.chart.data.datasets.slice(0, ctx.datasetIndex).filter(d => !d._isCompare).length;
      return actualIdx === 0 ? '#0D3270' : '#7A2E00';
    }},
    backgroundColor: (ctx) => {{
      if (ctx.dataset._isCompare) {{
        const nome = (ctx.dataset.label || '').toLowerCase();
        return nome.includes('arrecad') ? 'rgba(219,234,254,0.94)' : 'rgba(255,237,213,0.94)';
      }}
      const actualIdx = ctx.chart.data.datasets.slice(0, ctx.datasetIndex).filter(d => !d._isCompare).length;
      return actualIdx === 0 ? 'rgba(219,234,254,0.94)' : 'rgba(255,237,213,0.94)';
    }},
    borderRadius: 3,
    padding: {{ x: 4, y: 2 }},
    borderWidth: (ctx) => ctx.dataset._isCompare ? 1 : 0,
    borderColor: (ctx) => ctx.dataset._isCompare
      ? ((ctx.dataset.borderColor || '#0891B2') + '55') : 'transparent',
    formatter: (v) => (!v || v === 0) ? '' : fmtBRLk(v)
  }} : {{ display: false }};

  // Separar comparativos em: Arrecadação ant (di=0) e Faturamento ant (di=1)
  // Ordem final: Arrecadação atual (barra) → Faturamento atual (barra) → Arrecadação ant (linha) → Faturamento ant (linha)
  // Agrupa cmpDs por _tipo (igual ao 8117) para evitar mistura por paridade
  const cmp8121PorTipo = {{}};
  (cmpDs || []).forEach(d => {{
    const t = d._tipo || 'unknown';
    if (!cmp8121PorTipo[t]) cmp8121PorTipo[t] = {{ arr: [], fat: [] }};
    if (d.label && d.label.toLowerCase().includes('arrecad')) cmp8121PorTipo[t].arr.push(d);
    else cmp8121PorTipo[t].fat.push(d);
  }});
  const tiposOrdem8121 = ['month', 'year', 'year_month', 'prev_month_year', 'unknown'];
  const cmpArr = tiposOrdem8121.flatMap(t => cmp8121PorTipo[t] ? cmp8121PorTipo[t].arr : []);
  const cmpFat = tiposOrdem8121.flatMap(t => cmp8121PorTipo[t] ? cmp8121PorTipo[t].fat : []);

  // Cores semitransparentes por tipo para as barras de comparação
  const hexToRgba8121 = (hex, a) => {{
    if (!hex || !hex.startsWith('#')) return `rgba(0,0,0,${{a}})`;
    const r = parseInt(hex.slice(1,3),16), g = parseInt(hex.slice(3,5),16), b = parseInt(hex.slice(5,7),16);
    return `rgba(${{r}},${{g}},${{b}},${{a}})`;
  }};

  const makeCmpBar = (d, cor) => ({{
    ...d, _isCompare: true,
    type: 'bar', yAxisID: 'y', order: 3,
    borderWidth: 1, borderRadius: 3,
    borderColor: cor,
    backgroundColor: hexToRgba8121(cor, 0.40),
  }});

  // Ordem: Arr atual (azul escuro) → Arr ant (azul claro) → Fat atual (laranja) → Fat ant (laranja claro)
  // Agrupa por tipo de dado para comparação visual lado a lado
  const datasets8121 = [
    // Arrecadação atual  azul escuro
    {{ label: labelArr, data: dataArr,
       backgroundColor: 'rgba(21,101,192,0.85)', borderColor: '#1565C0',
       borderWidth: 1, borderRadius: 4, yAxisID: 'y', order: 1 }},
    // Arrecadação anterior  azul claro transparente
    ...cmpArr.map(d => makeCmpBar(d, '#1565C0')),
    // Faturamento atual  laranja escuro
    {{ label: labelFat, data: dataFat,
       backgroundColor: 'rgba(255,143,0,0.85)', borderColor: '#FF8F00',
       borderWidth: 1, borderRadius: 4, yAxisID: 'y', order: 1 }},
    // Faturamento anterior  laranja claro transparente
    ...cmpFat.map(d => makeCmpBar(d, '#FF8F00')),
  ];

  _chart8121Comp = makeChart(ctx, {{
    type: 'bar',
    data: {{ labels, datasets: datasets8121 }},
    options: {{
      responsive: true, interaction: {{ mode: 'index', intersect: false }},
      animation: {{ duration: 900, easing: 'easeOutQuart' }},
      plugins: {{
        legend: {{ position: 'top' }},
        datalabels: dlCfg,
        tooltip: {{ ...TT, callbacks: {{
          title: items => _fmtTitleComparacao(items, datasets8121, mensal),
          label: c => {{
            const lbl  = c.dataset.label || '';
            const tipo = c.dataset._tipo || '';
            const val  = fmtBRL(c.parsed.y);
            const nome = lbl.split('  ')[0].trim();
            const idx  = c.dataIndex;
            if (c.dataset._isCompare) {{
              const isArr = nome.toLowerCase().includes('arrecad');
              const meta = {{
                year:  {{ icone: '', tag: 'ano passado' }},
                prev_month_year: {{ icone: '', tag: 'mês ant. 1 ano' }},
                month: {{ icone: '↩',  tag: 'mês anterior' }},
              }}[tipo] || {{ icone: '↔', tag: tipo }};
              // Delta % em relação ao valor atual correspondente
              const dsAtual = datasets8121.find(d =>
                !d._isCompare &&
                (isArr ? d.label === labelArr : d.label === labelFat)
              );
              const valAtual = dsAtual ? (dsAtual.data[idx] ?? null) : null;
              const valComp  = c.parsed.y;
              let deltaStr = '';
              if (valAtual != null && valComp != null && valComp > 0) {{
                const delta = ((valAtual - valComp) / valComp) * 100;
                const sinal = delta >= 0 ? '+' : '';
                deltaStr = `  ${{delta >= 0 ? '▲' : '▼'}} ${{sinal}}${{delta.toFixed(1)}}%`;
              }}
              return ` ${{meta.icone}} ${{nome}} (${{meta.tag}}): ${{val}}${{deltaStr}}`;
            }}
            const icone = nome.toLowerCase().includes('arrecad') ? '' : '';
            return ` ${{icone}} ${{nome}}: ${{val}}`;
          }}
        }} }}
      }},
      scales: {{
        x: {{ grid: {{ color: 'rgba(0,0,0,0.04)' }}, ticks: {{ maxRotation: modoAgregado ? 0 : 45 }} }},
        y: {{
          position: 'left',
          ticks: {{ callback: v => fmtBRLk(v) }},
          grid: {{ color: 'rgba(0,0,0,0.05)' }},
        }}
      }}
    }}
  }});
}}

function _rebuild8121() {{
  const mensal = _modo8121 === 'monthly' ? true : (_modo8121 === 'yearly' ? 'yearly' : false);

  // ── Gráfico 1: Comparativo Arr vs Fat ─────────────────────────────
  (function() {{
    let labels, dataArr, dataFat, labelArr, labelFat;
    if (_modo8121 === 'yearly') {{
      const aggArr = _agruparAnual(DATAS_21, ARR21_D, ARR21_AC);
      const aggFat = _agruparAnual(DATAS_21, FAT_D, FAT_AC);
      const allLabels = [...new Set([...aggArr.labels, ...aggFat.labels])].sort(_sortCrono);
      labels  = allLabels;
      dataArr = allLabels.map(l => {{ const i = aggArr.labels.indexOf(l); return i>=0 ? aggArr.vals[i] : 0; }});
      dataFat = allLabels.map(l => {{ const i = aggFat.labels.indexOf(l); return i>=0 ? aggFat.vals[i] : 0; }});
      labelArr = 'Arrecadação Anual'; labelFat = 'Faturamento Anual';
    }} else if (_modo8121 === 'monthly') {{
      const aggArr = _agruparMensal(DATAS_21, ARR21_D, ARR21_AC);
      const aggFat = _agruparMensal(DATAS_21, FAT_D, FAT_AC);
      const allLabels = [...new Set([...aggArr.labels, ...aggFat.labels])].sort(_sortCrono);
      labels  = allLabels;
      dataArr = allLabels.map(l => {{ const i = aggArr.labels.indexOf(l); return i>=0 ? aggArr.vals[i] : 0; }});
      dataFat = allLabels.map(l => {{ const i = aggFat.labels.indexOf(l); return i>=0 ? aggFat.vals[i] : 0; }});
      labelArr = 'Arrecadação Mensal'; labelFat = 'Faturamento Mensal';
    }} else {{
      // Modo diário: usa arrecadação própria da 8121  independente da 8117
      labels  = DATAS_21;
      dataArr = ARR21_D;
      dataFat = FAT_D;
      labelArr = 'Arrecadação Diária'; labelFat = 'Faturamento Diário';
    }}
    _initRC8121Comp(labels, dataArr, dataFat, labelArr, labelFat);
  }})();

  // ── Gráfico 2: Acumulado  gerenciado por _rebuildAcumulado() ──────
  _rebuildAcumulado();
}}

let _chartAcumulado = null;
let _modoAcumulado  = 'daily';

function toggleAcumulado(modo) {{
  _modoAcumulado = modo;
  const bd = document.getElementById('btn-acum-d');
  const bm = document.getElementById('btn-acum-m');
  if (bd) {{ bd.style.background = modo==='daily' ? '#1565C0' : 'transparent'; bd.style.color = modo==='daily' ? '#fff' : '#6B7A96'; }}
  if (bm) {{ bm.style.background = modo==='monthly' ? '#1565C0' : 'transparent'; bm.style.color = modo==='monthly' ? '#fff' : '#6B7A96'; }}
  const sub = document.getElementById('sub-acumulado');
  if (sub) sub.textContent = modo==='daily'
     ?
     'Evolução acumulada diária no período'
    : 'Evolução acumulada mensal no período';
  _rebuildAcumulado();
}}

function _initRC8121Acum(labels, datasets) {{
  buildRangeCompare({{
    containerId: 'rc-8121-acum',
    labels:   labels,
    datasets: datasets.map(d => ({{ label: d.label, data: d.data }})),
    rebuildFn: (filtL, filtDs, cmpDs) => {{
      _drawChart8121Acum(filtL, filtDs, cmpDs);
    }},
  }});
}}

function _drawChart8121Acum(labels, filtDs, cmpDs) {{
  const ctx = document.getElementById('chart-acumulado');
  if (!ctx) return;
  if (_chartAcumulado) {{ _chartAcumulado.destroy(); _chartAcumulado = null; }}

  const mensal = _modoAcumulado === 'monthly';

  const c2d = ctx.getContext('2d');
  const gradArr = c2d.createLinearGradient(0,0,0,300);
  gradArr.addColorStop(0,'rgba(21,101,192,0.25)'); gradArr.addColorStop(1,'rgba(21,101,192,0.0)');
  const gradFat = c2d.createLinearGradient(0,0,0,300);
  gradFat.addColorStop(0,'rgba(255,143,0,0.2)'); gradFat.addColorStop(1,'rgba(255,143,0,0.0)');

  const dataArr = filtDs[0] ? filtDs[0].data : [];
  const dataFat = filtDs[1] ? filtDs[1].data : [];
  // Séries anteriores fixas (quando passadas como filtDs[2] e [3] via _rebuildAcumulado)
  const dataArrAnt = filtDs[2] ? filtDs[2].data : [];
  const dataFatAnt = filtDs[3] ? filtDs[3].data : [];
  const labelArrAnt = filtDs[2] ? filtDs[2].label : null;
  const labelFatAnt = filtDs[3] ? filtDs[3].label : null;

  // Recalcula acumulado relativo ao início do período selecionado.
  // O RC passa um slice do acumulado histórico  subtrai o valor do primeiro ponto
  // para que o gráfico parta do zero no início do recorte.
  function _relativo(arr) {{
    if (!arr || !arr.length) return arr;
    const primeiro = arr.find(v => v != null && v > 0);
    if (!primeiro) return arr; // todos zero  retorna como está
    return arr.map(v => v != null ? Math.max(0, Math.round((v - primeiro) * 100) / 100) : null);
  }}

  const dataArrRel = _relativo(dataArr);
  const dataFatRel = _relativo(dataFat);

  // Garante _tipo nos cmpDs vindos do RC (podem não ter _tipo se eram séries fixas)
  const cmpDsTyped = (cmpDs || []).map(d => ({{ ...d, _isCompare: true, _tipo: d._tipo || 'year' }}));

  // Agrupa por tipo para ordem correta na tooltip: month → year
  const acumPorTipo = {{}};
  cmpDsTyped.forEach(d => {{
    const t = d._tipo || 'unknown';
    if (!acumPorTipo[t]) acumPorTipo[t] = [];
    acumPorTipo[t].push(d);
  }});
  const tiposOrdemAcum = ['month', 'year', 'year_month', 'prev_month_year', 'unknown'];
  const cmpDsOrdenados = tiposOrdemAcum.flatMap(t => acumPorTipo[t] || []);

  const datasets = [
    {{ label: 'Arrecadação Acumulada',
       data: dataArrRel, borderColor:'#1565C0', backgroundColor:gradArr,
       borderWidth:2.5, pointRadius: mensal ? 6 : 3, pointHoverRadius:7,
       fill:true, tension:0.4, order: 1 }},
    {{ label: 'Faturamento Acumulado',
       data: dataFatRel, borderColor:'#FF8F00', backgroundColor:gradFat,
       borderWidth:2.5, pointRadius: mensal ? 6 : 3, pointHoverRadius:7,
       pointStyle:'rectRot', fill:true, tension:0.4, order: 1 }},
    // Séries anteriores fixas (DIAS_MES_ANT)  mantidas com _tipo marcado
    ...(labelArrAnt ? [{{
       label: labelArrAnt, data: dataArrAnt,
       borderColor:'#0891B2', backgroundColor:'transparent',
       borderWidth:1.8, borderDash:[5,4], pointRadius: mensal ? 4 : 2, pointHoverRadius:6,
       fill:false, tension:0.4, order: 2, _isCompare: true, _tipo: 'year'
    }}] : []),
    ...(labelFatAnt ? [{{
       label: labelFatAnt, data: dataFatAnt,
       borderColor:'#0D9488', backgroundColor:'transparent',
       borderWidth:1.8, borderDash:[5,4], pointRadius: mensal ? 4 : 2, pointHoverRadius:6,
       pointStyle:'rectRot', fill:false, tension:0.4, order: 2, _isCompare: true, _tipo: 'year'
    }}] : []),
    // cmpDs do RC ordenados por tipo
    ...cmpDsOrdenados.map(d => ({{ ...d, order: 3 }}))
  ];

  const dlCfg = mensal ? {{
    clamp: true, clip: false,
    display: (ctx) => {{
      const v = ctx.dataset.data[ctx.dataIndex];
      if (!v || v === 0) return false;
      if (ctx.dataset._isCompare) {{
        // Só mostrar no último ponto das séries anteriores
        const data = ctx.dataset.data.filter(x => x != null && x > 0);
        return v === data[data.length - 1];
      }}
      return true;
    }},
    anchor: 'end',
    align: (ctx) => {{
      if (ctx.dataset._isCompare) return 'bottom';
      const actualIdx = ctx.chart.data.datasets.slice(0, ctx.datasetIndex).filter(d => !d._isCompare).length;
      return actualIdx === 0 ? 'top' : 'bottom';
    }},
    offset: 8,
    font: (ctx) => ({{ size: ctx.dataset._isCompare ? 10 : 11, weight: 'bold' }}),
    color: (ctx) => {{
      if (ctx.dataset._isCompare) return ctx.dataset.borderColor || '#0891B2';
      const actualIdx = ctx.chart.data.datasets.slice(0, ctx.datasetIndex).filter(d => !d._isCompare).length;
      return actualIdx === 0 ? '#1565C0' : '#E65100';
    }},
    backgroundColor: 'rgba(255,255,255,0.96)',
    borderRadius: 4,
    padding: {{ x: 6, y: 3 }},
    borderColor: (ctx) => {{
      if (ctx.dataset._isCompare) return (ctx.dataset.borderColor || '#0891B2') + '55';
      const actualIdx = ctx.chart.data.datasets.slice(0, ctx.datasetIndex).filter(d => !d._isCompare).length;
      return actualIdx === 0 ? 'rgba(21,101,192,0.30)' : 'rgba(230,81,0,0.30)';
    }},
    borderWidth: 1,
    formatter: (v, ctx) => {{
      if (!v || v === 0) return '';
      const lbl = ctx.dataset.label || '';
      const periodoMatch = lbl.match(/  (.+)$/);
      const periodo = periodoMatch ? periodoMatch[1] : null;
      if (ctx.dataset._isCompare && periodo) {{
        return fmtBRLk(v) + '\\n↩ ' + periodo;
      }}
      return fmtBRLk(v);
    }}
  }} : {{ display: false }};

  _chartAcumulado = makeChart(ctx, {{
    type: 'line',
    data: {{ labels, datasets }},
    options: {{
      responsive:true, interaction:{{mode:'index',intersect:false}},
      animation:{{duration:900,easing:'easeOutQuart'}},
      plugins: {{
        legend: {{position:'top'}},
        datalabels: dlCfg,
        tooltip: {{...TT, callbacks:{{
          title: items => _fmtTitleComparacao(items, datasets, mensal),
          label: c => {{
            const lbl  = c.dataset.label || '';
            const tipo = c.dataset._tipo || '';
            const val  = fmtBRL(c.parsed.y);
            const idx  = c.dataIndex;
            const nome = lbl.split('  ')[0].trim();

            if (c.dataset._isCompare) {{
              const isArr = nome.toLowerCase().includes('arrecad');
              const meta = {{
                year:  {{ icone: '', tag: 'ano passado' }},
                prev_month_year: {{ icone: '', tag: 'mês ant. 1 ano' }},
                month: {{ icone: '↩',  tag: 'mês anterior' }},
              }}[tipo] || {{ icone: '↔', tag: tipo }};

              // Delta % vs série atual correspondente
              const dsAtual = datasets.find(d =>
                !d._isCompare &&
                (isArr ? d.label.includes('Arrecadação') : d.label.includes('Faturamento'))
              );
              const valAtual = dsAtual ? (dsAtual.data[idx] ?? null) : null;
              const valComp  = c.parsed.y;
              let deltaStr = '';
              if (valAtual != null && valComp != null && valComp > 0) {{
                const delta = ((valAtual - valComp) / valComp) * 100;
                deltaStr = `  ${{delta >= 0 ? '▲' : '▼'}} ${{delta >= 0 ? '+' : ''}}${{delta.toFixed(1)}}%`;
              }}
              return ` ${{meta.icone}} ${{nome}} (${{meta.tag}}): ${{val}}${{deltaStr}}`;
            }}

            // Séries atuais
            const base = nome.includes('Arrecadação')
              ? (dataFatRel.length ? dataFatRel[dataFatRel.length-1] : 0)
              : (dataArrRel.length ? dataArrRel[dataArrRel.length-1] : 0);
            const pct = base > 0 ? `  · ${{(c.parsed.y/base*100).toFixed(1)}}%` : '';
            const icone = nome.includes('Arrecadação') ? '' : '';
            return ` ${{icone}} ${{nome}}: ${{val}}${{pct}}`;
          }}
        }} }}
      }},
      scales: {{
        x: {{grid:{{color:'rgba(0,0,0,0.04)'}}, ticks:{{maxRotation: mensal ? 0 : 45}}}},
        y: {{ticks:{{callback:v=>fmtBRLk(v)}}, grid:{{color:'rgba(0,0,0,0.05)'}}}}
      }}
    }}
  }});
}}

function _rebuildAcumulado() {{
  const mensal = _modoAcumulado === 'monthly';
  const temAnt = FAT_AC_ANT.length > 0 || ARR_AC_ANT.length > 0;

  let labels, dataArr, dataFat;
  if (mensal) {{
    const aggArr = _agruparMensal(DATAS_21, ARR21_D, ARR21_AC);
    const aggFat = _agruparMensal(DATAS_21, FAT_D, FAT_AC);
    const allLabels = [...new Set([...aggArr.labels, ...aggFat.labels])].sort(_sortCrono);
    labels  = allLabels;
    dataArr = aggArr.acum.slice(0, allLabels.length);
    dataFat = aggFat.acum.slice(0, allLabels.length);
  }} else {{
    labels  = temAnt ? DIAS_MES.map(d => 'Dia ' + d) : DATAS_21;
    dataArr = ARR21_AC.slice(0, DATAS_21.length);
    dataFat = FAT_AC;
  }}

  // Datasets base para RC (sem séries anteriores  RC gera os comparativos)
  const baseDsRc = [
    {{ label: 'Arrecadação Acumulada', data: dataArr }},
    {{ label: 'Faturamento Acumulado', data: dataFat }},
  ];

  // Séries anteriores ficam fora do RC (são fixas, não filtráveis)
  if (!mensal && temAnt) {{
    const antArr = DIAS_MES.map(d => {{ const i=DIAS_MES_ANT.indexOf(d); return i>=0?ARR_AC_ANT[i]:null; }});
    const antFat = DIAS_MES.map(d => {{ const i=DIAS_MES_ANT.indexOf(d); return i>=0?FAT_AC_ANT[i]:null; }});
    baseDsRc.push({{ label: LABEL_ANT?'Arrecadação  '+LABEL_ANT:'Arrec. Anterior', data:antArr }});
    baseDsRc.push({{ label: LABEL_ANT?'Faturamento  '+LABEL_ANT:'Fat. Anterior',   data:antFat }});
  }}

  _initRC8121Acum(labels, baseDsRc);
}}


function buildCharts8121() {{
  if (!TEM_8121) return;
  _rebuild8121();
  animateCounter(document.getElementById('kpi-arr-21'), TOTAL_ARR, fmtBRL);
  animateCounter(document.getElementById('kpi-fat-21'), TOTAL_FAT, fmtBRL);
  animateCounter(document.getElementById('kpi-pct-21'), PCT, v => v.toFixed(1)+'%');
  animateCounter(document.getElementById('kpi-maior-fat'), MAIOR_FAT_V, fmtBRL);
  const kpiMaiorFatData = document.getElementById('kpi-maior-fat-data');
  if (kpiMaiorFatData) kpiMaiorFatData.textContent = MAIOR_FAT_D || '--';
}}

function initInsights8121() {{
  const container = document.getElementById('insights-8121');
  if (!container) return;
  const items = [
    {{ emoji:'', cor:'', texto:`<strong>Total arrecadado:</strong> ${{fmtBRL(TOTAL_ARR)}} no período.` }},
    {{ emoji:'', cor:'orange', texto:`<strong>Total faturado:</strong> ${{fmtBRL(TOTAL_FAT)}} no período.` }},
    {{ emoji:'', cor: PCT >= 90 ? 'green' : PCT >= 70 ? '' : 'red',
       texto:`<strong>Realização:</strong> ${{PCT.toFixed(1)}}% do faturamento foi arrecadado.` }},
    MAIOR_FAT_V > 0
      ?
       {{ emoji:'', cor:'orange', texto:`<strong>Maior faturamento:</strong> ${{MAIOR_FAT_D}} com ${{fmtBRL(MAIOR_FAT_V)}}.` }}
      : null,
  ].filter(Boolean);
  items.forEach((it, i) => {{
    const div = document.createElement('div');
    div.className = `insight-item ${{it.cor}}`;
    div.style.transitionDelay = (i * 70) + 'ms';
    div.innerHTML = `<span class="insight-emoji">${{it.emoji}}</span><div class="insight-text">${{it.texto}}</div>`;
    container.appendChild(div);
  }});
}}

// ── KPIs 8091 ─────────────────────────────────────────────────────────
function initKPIs8091() {{
  animateCounter(document.getElementById('k91-fat'),   K91_FAT,   fmtBRL);
  animateCounter(document.getElementById('k91-agua'),  K91_AGUA,  fmtBRL);
  animateCounter(document.getElementById('k91-mult'),  K91_MULT,  fmtBRL);
  animateCounter(document.getElementById('k91-lig'),   K91_LIG,   v => Math.round(v).toLocaleString('pt-BR'));
  animateCounter(document.getElementById('k91-vol'),   K91_VOL,   v => Math.round(v).toLocaleString('pt-BR') + ' m³');
  animateCounter(document.getElementById('k91-media'), K91_MEDIA, fmtBRL);
}}

// ── CHARTS 8091 ───────────────────────────────────────────────────────
let _chart8091Linha = null;

function _initRC8091Linha(labels, dataV) {{
  buildRangeCompare({{
    containerId: 'rc-8091-linha',
    labels:   labels,
    datasets: [{{ label: 'Faturamento Mensal', data: dataV }}],
    rebuildFn: (filtL, filtDs, cmpDs) => {{
      _drawChart8091Linha(filtL, filtDs[0].data, cmpDs);
    }},
  }});
}}

function _drawChart8091Linha(labels, dataV, cmpDs) {{
  const ctx = document.getElementById('chart-8091-linha');
  if (!ctx) return;
  if (_chart8091Linha) {{ _chart8091Linha.destroy(); _chart8091Linha = null; }}

  // Gradiente para a barra atual
  const c2d = ctx.getContext('2d');
  const gradBar = c2d.createLinearGradient(0, 0, 0, 300);
  gradBar.addColorStop(0, 'rgba(46,125,50,0.85)');
  gradBar.addColorStop(1, 'rgba(46,125,50,0.45)');

  // Agrupa cmpDs por tipo  padrão dos outros gráficos
  const cmp91PorTipo = {{}};
  (cmpDs || []).forEach(d => {{
    const t = d._tipo || 'unknown';
    if (!cmp91PorTipo[t]) cmp91PorTipo[t] = [];
    cmp91PorTipo[t].push({{ ...d, _isCompare: true }});
  }});
  const tiposOrdem91 = ['month', 'year', 'prev_month_year', 'unknown'];
  const cmpOrdenados91 = tiposOrdem91.flatMap(t => cmp91PorTipo[t] || []);

  // Helper cor semitransparente
  const hexToRgba91 = (hex, a) => {{
    if (!hex || !hex.startsWith('#')) return `rgba(46,125,50,${{a}})`;
    const r = parseInt(hex.slice(1,3),16), g = parseInt(hex.slice(3,5),16), b = parseInt(hex.slice(5,7),16);
    return `rgba(${{r}},${{g}},${{b}},${{a}})`;
  }};

  const datasets91 = [
    // Faturamento atual  verde
    {{ label: 'Faturamento Mensal', data: dataV,
       backgroundColor: gradBar, borderColor: '#2E7D32',
       borderWidth: 1, borderRadius: 5, yAxisID: 'y', order: 1 }},
    // Comparativos  barras semitransparentes na mesma cor verde
    ...cmpOrdenados91.map(d => ({{
      ...d, type: 'bar', yAxisID: 'y', order: 2,
      borderWidth: 1, borderRadius: 3,
      borderColor: d.borderColor || '#2E7D32',
      backgroundColor: hexToRgba91(d.borderColor || '#2E7D32', 0.35),
    }}))
  ];

  _chart8091Linha = makeChart(ctx, {{
    type: 'bar',
    data: {{ labels, datasets: datasets91 }},
    options: {{
      responsive: true, interaction: {{ mode: 'index', intersect: false }},
      animation: {{ duration: 1200, easing: 'easeOutQuart' }},
      plugins: {{
        legend: {{ position: 'top' }},
        datalabels: {{
          display: (ctx) => {{
            const v = ctx.dataset.data[ctx.dataIndex];
            if (!v || v === 0) return false;
            const vals = ctx.dataset.data.filter(x => x > 0);
            const max  = vals.length ? Math.max(...vals) : 1;
            return v / max >= 0.10;
          }},
          anchor: 'center',
          align: (ctx) => ctx.dataset._isCompare ? 'top' : 'center',
          offset: (ctx) => ctx.dataset._isCompare ? 4 : 0,
          font: (ctx) => ({{ size: ctx.dataset._isCompare ? 10 : 11, weight: 'bold' }}),
          color: (ctx) => ctx.dataset._isCompare
            ? (ctx.dataset.borderColor || '#2E7D32')
            : '#1B5E20',
          backgroundColor: (ctx) => ctx.dataset._isCompare
            ? 'rgba(255,255,255,0.92)'
            : 'rgba(232,245,233,0.94)',
          borderRadius: 3,
          padding: {{ x: 4, y: 2 }},
          borderWidth: (ctx) => ctx.dataset._isCompare ? 1 : 0,
          borderColor: (ctx) => ctx.dataset._isCompare
            ? ((ctx.dataset.borderColor || '#2E7D32') + '55') : 'transparent',
          formatter: (v) => (!v || v === 0) ? '' : fmtBRLk(v)
        }},
        tooltip: {{ ...TT, callbacks: {{
          title: items => _fmtTitleComparacao(items, datasets91, true),
          label: c => {{
            const tipo = c.dataset._tipo || '';
            const val  = fmtBRL(c.parsed.y);
            const idx  = c.dataIndex;
            if (c.dataset._isCompare) {{
              const meta = {{
                year:  {{ icone: '', tag: 'ano passado' }},
                prev_month_year: {{ icone: '', tag: 'mês ant. 1 ano' }},
                month: {{ icone: '↩',  tag: 'mês anterior' }},
              }}[tipo] || {{ icone: '↔', tag: tipo }};
              const valAtual = dataV[idx] ?? null;
              const valComp  = c.parsed.y;
              let deltaStr = '';
              if (valAtual != null && valComp != null && valComp > 0) {{
                const delta = ((valAtual - valComp) / valComp) * 100;
                deltaStr = `  ${{delta >= 0 ? '▲' : '▼'}} ${{delta >= 0 ? '+' : ''}}${{delta.toFixed(1)}}%`;
              }}
              return ` ${{meta.icone}} Faturamento (${{meta.tag}}): ${{val}}${{deltaStr}}`;
            }}
            return `  Faturamento: ${{val}}`;
          }}
        }} }}
      }},
      scales: {{
        x: {{ grid: {{ color: 'rgba(0,0,0,0.04)' }}, ticks: {{ maxRotation: 45 }} }},
        y: {{ ticks: {{ callback: v => fmtBRLk(v) }}, grid: {{ color: 'rgba(0,0,0,0.05)' }} }}
      }}
    }}
  }});
}}

function buildCharts8091() {{
  if (!TEM_8091) return;

  // Situação (doughnut)
  (function() {{
    const ctx = document.getElementById('chart-91-situacao');
    if (!ctx || !SIT91_L.length) return;
    const cores = ['#2E7D32','#1565C0','#C62828','#FF8F00','#6A1B9A','#37474F','#00838F','#D84315'];
    const totalSit = SIT91_Q.reduce((a,b)=>a+b,0);
    const totalSitV = SIT91_V.reduce((a,b)=>a+b,0);
    makeChart(ctx, {{
      type:'doughnut',
      data:{{ labels:SIT91_L, datasets:[{{ data:SIT91_Q, backgroundColor:cores,
               borderColor:'#fff', borderWidth:3, hoverOffset:10 }}] }},
      options:{{
        responsive:true, cutout:'55%',
        animation:{{duration:1400,easing:'easeOutBack'}},
        plugins:{{ legend:{{ position:'right', labels:{{ padding:12, font:{{size:10}}, usePointStyle:true, pointStyleWidth:8 }}}},
          datalabels: getDatalabelsPie(totalSit, v => v.toLocaleString('pt-BR') + ' lig'),
          tooltip:{{...TT, callbacks:{{label: c =>
            ` ${{c.label}}: ${{c.parsed.toLocaleString('pt-BR')}} fat. (${{(c.parsed/totalSit*100).toFixed(1)}}%)  ${{fmtBRL(SIT91_V[c.dataIndex])}} (${{totalSitV ? (SIT91_V[c.dataIndex]/totalSitV*100).toFixed(1) : 0}}% do valor)`
          }}}} }}
      }}
    }});
  }})();

  // Fase cobrança (barra horiz)
  (function() {{
    const ctx = document.getElementById('chart-91-fase');
    if (!ctx || !FASE91_L.length) return;
    const cores = ['#2E7D32','#FF8F00','#1565C0','#C62828','#6A1B9A','#37474F','#00838F'];
    const totalFase = FASE91_Q.reduce((a,b)=>a+b,0);
    const totalFaseV = FASE91_V.reduce((a,b)=>a+b,0);
    makeChart(ctx, {{
      type:'bar',
      data:{{ labels:FASE91_L, datasets:[{{ label:'Ligações', data:FASE91_Q,
               backgroundColor:cores, borderRadius:5 }}] }},
      options:{{
        indexAxis:'y', responsive:true,
        animation:{{duration:1200,easing:'easeOutBack'}},
        plugins:{{ 
          legend:{{ position:'right', labels:{{ padding:12, font:{{size:10}}, usePointStyle:true, pointStyleWidth:8 }}}},
          datalabels: getDatalabelsHBar(v => fmtBRLk(v)),
          tooltip:{{...TT, callbacks:{{label: c =>
            ` ${{c.parsed.x.toLocaleString('pt-BR')}} lig (${{(c.parsed.x/totalFase*100).toFixed(1)}}%)  ${{fmtBRL(FASE91_V[c.dataIndex])}} (${{totalFaseV ? (FASE91_V[c.dataIndex]/totalFaseV*100).toFixed(1) : 0}}% do valor)`
          }}}} }},
        scales:{{
          x:{{ticks:{{callback:v=>v.toLocaleString('pt-BR')}},grid:{{color:'rgba(0,0,0,0.05)'}}}},
          y:{{grid:{{display:false}}}}
        }}
      }}
    }});
  }})();

  // Composição (doughnut)
  (function() {{
    const ctx = document.getElementById('chart-91-comp');
    if (!ctx) return;
    makeChart(ctx, {{
      type:'doughnut',
      data:{{ labels:COMP91_L, datasets:[{{ data:COMP91_V,
               backgroundColor:['#1565C0','#C62828','#37474F'],
               borderColor:'#fff', borderWidth:3, hoverOffset:12 }}] }},
      options:{{
        responsive:true, cutout:'60%',
        animation:{{duration:1400,easing:'easeOutBack'}},
        plugins:{{ legend:{{ position:'right', labels:{{ padding:12, font:{{size:10}}, usePointStyle:true, pointStyleWidth:8 }}}},
          datalabels: getDatalabelsPie(COMP91_V.reduce((a,b)=>a+b,0), fmtBRL),
          tooltip:{{...TT, callbacks:{{label: c =>
            ` ${{c.label}}: ${{fmtBRL(c.parsed)}} (${{K91_FAT ? (c.parsed/K91_FAT*100).toFixed(1) : 0}}%)`
          }}}} }}
      }}
    }});
  }})();

  // Categoria (combo)
  (function() {{
    const ctx = document.getElementById('chart-91-categoria');
    if (!ctx || !CAT91_L.length) return;
    const totalCatV = CAT91_V.reduce((a,b)=>a+b,0);
    const totalCatQ = CAT91_Q.reduce((a,b)=>a+b,0);
    makeChart(ctx, {{
      data:{{
        labels:CAT91_L,
        datasets:[
          {{ type:'bar', label:'Valor Faturado', data:CAT91_V,
             backgroundColor:['#1565C0','#FF8F00','#2E7D32','#6A1B9A','#C62828'],
             borderRadius:6, yAxisID:'y' }},
          {{ type:'line', label:'Qtde Ligações', data:CAT91_Q,
             borderColor:'#011F3F', backgroundColor:'transparent', borderWidth:2.5,
             pointRadius:6, pointBackgroundColor:'#011F3F', tension:0.3, yAxisID:'y2' }}
        ]
      }},
      options:{{
        responsive:true, interaction:{{mode:'index',intersect:false}},
        animation:{{duration:1300,easing:'easeOutQuart'}},
        plugins:{{ 
          legend:{{position:'top'}},
          datalabels: getDatalabelsConfig('barra', true),
          tooltip:{{...TT, callbacks:{{label: c =>
            c.dataset.label === 'Valor Faturado'
              ?
               ` Valor: ${{fmtBRL(c.parsed.y)}} (${{totalCatV ? (c.parsed.y/totalCatV*100).toFixed(1) : 0}}%)`
              : ` Ligações: ${{c.parsed.y.toLocaleString('pt-BR')}} (${{totalCatQ ? (c.parsed.y/totalCatQ*100).toFixed(1) : 0}}%)`
          }}}} }},
        scales:{{
          x:{{grid:{{color:'rgba(0,0,0,0.04)'}}}},
          y:{{position:'left', ticks:{{callback:v=>fmtBRLk(v)}},grid:{{color:'rgba(0,0,0,0.05)'}}}},
          y2:{{position:'right',ticks:{{callback:v=>v.toLocaleString('pt-BR')}},grid:{{display:false}}}}
        }}
      }}
    }});
  }})();

  // Bairros (barra horiz)
  (function() {{
    const ctx = document.getElementById('chart-91-bairros');
    if (!ctx || !BAR91_L.length) return;
    const pal = ['#011F3F','#1565C0','#1976D2','#1E88E5','#2196F3','#42A5F5','#64B5F6','#90CAF9','#BBDEFB','#E3F2FD'];
    const totalBarV = BAR91_V.reduce((a,b)=>a+b,0);
    makeChart(ctx, {{
      type:'bar',
      data:{{ labels:BAR91_L, datasets:[{{ label:'Faturado', data:BAR91_V,
               backgroundColor:pal, borderRadius:4 }}] }},
      options:{{
        indexAxis:'y', responsive:true,
        animation:{{duration:1200,easing:'easeOutBack'}},
        plugins:{{ 
          legend:{{ position:'right', labels:{{ padding:12, font:{{size:10}}, usePointStyle:true, pointStyleWidth:8 }}}},
          datalabels: getDatalabelsHBar(v => fmtBRLk(v)),
          tooltip:{{...TT, callbacks:{{label: c =>
            ` ${{fmtBRL(c.parsed.x)}} (${{totalBarV ? (c.parsed.x/totalBarV*100).toFixed(1) : 0}}%)  ${{BAR91_Q[c.dataIndex].toLocaleString('pt-BR')}} lig`
          }}}} }},
        scales:{{
          x:{{ticks:{{callback:v=>fmtBRLk(v)}},grid:{{color:'rgba(0,0,0,0.05)'}}}},
          y:{{grid:{{display:false}},ticks:{{font:{{size:10}}}}}}
        }}
      }}
    }});
  }})();

  // Receita por m³ por grupo
  (function() {{
    const ctx = document.getElementById('chart-91-grupos');
    if (!ctx || !GRP_R91_L.length) return;
    const cores = GRP_R91_RM.map(v => v < 1 ? 'rgba(198,40,40,0.75)' : v < 4 ? 'rgba(242,125,22,0.75)' : 'rgba(1,31,63,0.75)');
    const totalGrpV = GRP_R91_VOL.reduce((a,b)=>a+b,0);
    const totalGrpFat = GRP_R91_V.reduce((a,b)=>a+b,0);
    makeChart(ctx, {{
      data: {{
        labels: GRP_R91_L,
        datasets: [
          {{ type:'bar', label:'Receita/m³ (R$)', data:GRP_R91_RM,
             backgroundColor:cores, borderRadius:5, yAxisID:'y',
             datalabels: {{
               display: true,
               anchor: 'center', align: 'center',
               color: '#fff', font: {{ size: 10, weight: '700' }},
               formatter: v => 'R$' + v.toFixed(2)
             }} }},
          {{ type:'line', label:'Volume (m³)', data:GRP_R91_VOL,
             borderColor:'#F27D16', backgroundColor:'transparent',
             borderWidth:2, pointRadius:4, pointBackgroundColor:'#F27D16',
             tension:0.3, yAxisID:'y2',
             datalabels: {{
               display: true,
               anchor: 'end', align: 'top', offset: 4,
               color: '#E65100', font: {{ size: 9, weight: '600' }},
               formatter: v => v >= 1000
                 ?
                  'R$ ' + (v/1000).toLocaleString('pt-BR', {{minimumFractionDigits:1}}) + 'k'
                 : v.toLocaleString('pt-BR')
             }} }}
        ]
      }},
      options:{{
        responsive:true, interaction:{{mode:'index',intersect:false}},
        animation:{{duration:1300,easing:'easeOutQuart'}},
        plugins:{{
          legend:{{position:'top'}},
          datalabels: {{ display: false }},
          tooltip:{{...TT, callbacks:{{label: c =>
            c.dataset.label === 'Receita/m³ (R$)'
              ?
               ` R$${{c.parsed.y.toFixed(2)}}/m³  ${{fmtBRL(GRP_R91_V[c.dataIndex])}} (${{totalGrpFat ? (GRP_R91_V[c.dataIndex]/totalGrpFat*100).toFixed(1) : 0}}% do fat.)`
              : ` ${{c.parsed.y.toLocaleString('pt-BR')}} m³ (${{totalGrpV ? (c.parsed.y/totalGrpV*100).toFixed(1) : 0}}% do vol.)`
          }}}}
        }},
        scales:{{
          x:{{grid:{{color:'rgba(0,0,0,0.04)'}},ticks:{{maxRotation:45}}}},
          y:{{position:'left', title:{{display:true,text:'R$/m³'}}, ticks:{{callback:v=>'R$'+v.toFixed(2)}},grid:{{color:'rgba(0,0,0,0.05)'}}}},
          y2:{{position:'right',title:{{display:true,text:'Volume m³'}}, ticks:{{callback:v=>v.toLocaleString('pt-BR')+' m³'}},grid:{{display:false}}}}
        }}
      }}
    }});
  }})();

  // Rings inadimplência
  (function() {{
    const ring91 = document.getElementById('rings-91');
    if (!ring91 || !SIT91_L.length) return;
    let totPago=0, totDebito=0, totIsento=0, valPago=0, valDebito=0;
    SIT91_L.forEach((s,i) => {{
      if (s.includes('PAGO'))   {{ totPago   += SIT91_Q[i]; valPago   += SIT91_V[i]; }}
      if (s.includes('DEBITO')) {{ totDebito += SIT91_Q[i]; valDebito += SIT91_V[i]; }}
      if (s.includes('ISENTO')) {{ totIsento += SIT91_Q[i]; }}
    }});
    const tot = totPago + totDebito + totIsento || 1;
    function mkRing(label, pct, color, value, sub) {{
      const r=44, sw=10, circ=2*Math.PI*r, dash=(pct/100*circ).toFixed(2);
      return `<div style="text-align:center">
        <svg width="110" height="110" viewBox="0 0 110 110" style="transform:rotate(-90deg)">
          <circle cx="55" cy="55" r="${{r}}" fill="none" stroke="#EEF2FF" stroke-width="${{sw}}"/>
          <circle cx="55" cy="55" r="${{r}}" fill="none" stroke="${{color}}" stroke-width="${{sw}}"
            stroke-dasharray="${{dash}} ${{circ}}" stroke-linecap="round"/>
          <text x="55" y="52" text-anchor="middle" font-size="16" font-weight="800"
            fill="${{color}}" style="transform:rotate(90deg);transform-origin:55px 55px">${{pct.toFixed(1)}}%</text>
          <text x="55" y="68" text-anchor="middle" font-size="9" fill="#6B7A96"
            style="transform:rotate(90deg);transform-origin:55px 55px">${{sub}}</text>
        </svg>
        <div style="font-size:0.8rem;color:#6B7A96;margin-top:8px;text-transform:uppercase;font-weight:600">${{label}}</div>
        <div style="font-size:1.2rem;font-weight:800;color:${{color}};margin-top:2px">${{value > 0 ? fmtBRL(value) : sub}}</div>
      </div>`;
    }}
    ring91.innerHTML =
      mkRing('Pago',      totPago/tot*100,   '#2E7D32', valPago,   totPago.toLocaleString('pt-BR')+' fat.') +
      mkRing('Em Débito', totDebito/tot*100,  '#C62828', valDebito, totDebito.toLocaleString('pt-BR')+' fat.') +
      mkRing('Isento',    totIsento/tot*100,  '#37474F', 0,         totIsento.toLocaleString('pt-BR')+' fat.');
  }})();

  // Evolução mensal linha
  (function() {{
    if (SERIE_L.length <= 1) {{
      const wrap = document.getElementById('chart-8091-linha-wrap');
      if (wrap) wrap.style.display = 'none';
      return;
    }}
    _initRC8091Linha(SERIE_L, SERIE_V);
  }})();

  // Faixas de consumo m³
  (function() {{
    const ctx = document.getElementById('chart-91-faixas-vol');
    if (!ctx || !FVOL91_L.length) return;
    const cores = ['#C62828','#EF6C00','#F9A825','#2E7D32','#1565C0','#0277BD','#6A1B9A','#37474F'];
    const total = FVOL91_V.reduce((a,b)=>a+b,0);
    makeChart(ctx, {{
      type:'bar',
      data:{{ labels:FVOL91_L, datasets:[{{
        label:'Ligações', data:FVOL91_V,
        backgroundColor:cores, borderRadius:6, borderSkipped:false
      }}] }},
      options:{{
        responsive:true, animation:{{duration:1200,easing:'easeOutQuart'}},
        plugins:{{ legend:{{ position:'right', labels:{{ padding:12, font:{{size:10}}, usePointStyle:true, pointStyleWidth:8 }}}},
          tooltip:{{...TT, callbacks:{{label: c =>
            ` ${{c.parsed.y.toLocaleString('pt-BR')}} ligações (${{(c.parsed.y/total*100).toFixed(1)}}%)`
          }}}}
        }},
        scales:{{
          x:{{grid:{{display:false}}}},
          y:{{ticks:{{callback:v=>v.toLocaleString('pt-BR')}},grid:{{color:'rgba(0,0,0,0.05)'}}}}
        }}
      }}
    }});
  }})();

  // Leituristas  Volume, Ligações e Média por Ligação
  (function() {{
    const ctx = document.getElementById('chart-91-leiturista');
    if (!ctx || !LEIT91_L.length) return;

    // Ordena por volume total decrescente
    const idx = LEIT91_VOL.map((v,i) => i).sort((a,b) => LEIT91_VOL[b] - LEIT91_VOL[a]);
    const labels  = idx.map(i => LEIT91_L[i]);
    const vols    = idx.map(i => LEIT91_VOL[i]);
    const qtds    = idx.map(i => LEIT91_Q[i]);
    const medias  = idx.map(i => LEIT91_VM[i]);

    const totalVol = vols.reduce((a,b)=>a+b,0);
    const totalQtd = qtds.reduce((a,b)=>a+b,0);
    const mediaGeral = totalQtd > 0 ? totalVol / totalQtd : 0;

    makeChart(ctx, {{
      data: {{
        labels,
        datasets: [
          {{ type:'bar', label:'Volume Lido (m³)', data:vols,
             backgroundColor:'rgba(21,101,192,0.75)', borderColor:'#1565C0',
             borderWidth:1, borderRadius:5, yAxisID:'y',
             datalabels: {{ display: false }} }},
          {{ type:'bar', label:'Ligações Lidas', data:qtds,
             backgroundColor:'rgba(46,125,50,0.70)', borderColor:'#2E7D32',
             borderWidth:1, borderRadius:5, yAxisID:'y2',
             datalabels: {{ display: false }} }},
          {{ type:'line', label:'Vol. Médio/Lig. (m³)', data:medias,
             borderColor:'#F27D16', backgroundColor:'rgba(242,125,22,0.12)',
             borderWidth:2.5, pointRadius:6, pointBackgroundColor:'#F27D16',
             tension:0.3, yAxisID:'y3',
             datalabels: {{
               display: true,
               anchor: 'end', align: 'top', offset: 4,
               font: {{ size: 10, weight: '600' }},
               color: (ctx) => {{
                 if (!ctx || ctx.parsed == null) return '#F27D16';
                 return ctx.parsed.y > mediaGeral * 1.3 ? '#C62828'
                      : ctx.parsed.y < mediaGeral * 0.7 ? '#1565C0'
                      : '#F27D16';
               }},
               formatter: v => v != null ? v.toLocaleString('pt-BR', {{minimumFractionDigits:1, maximumFractionDigits:1}}) + ' m³' : ''
             }} }}
        ]
      }},
      options: {{
        responsive: true,
        interaction: {{ mode:'index', intersect:false }},
        animation: {{ duration:1300, easing:'easeOutQuart' }},
        plugins: {{
          legend: {{ position:'top' }},
          datalabels: {{ display: false }},
          tooltip: {{ ...TT, callbacks: {{ label: c => {{
            if (c.dataset.label === 'Volume Lido (m³)')
              return ` Volume: ${{c.parsed.y.toLocaleString('pt-BR')}} m³ (${{totalVol ? (c.parsed.y/totalVol*100).toFixed(1) : 0}}%)`;
            if (c.dataset.label === 'Ligações Lidas')
              return ` Ligações: ${{c.parsed.y.toLocaleString('pt-BR')}} (${{totalQtd ? (c.parsed.y/totalQtd*100).toFixed(1) : 0}}%)`;
            const flag = c.parsed.y > mediaGeral * 1.3 ? '  alto' : c.parsed.y < mediaGeral * 0.7 ? '  baixo' : '';
            return ` Vol. Médio: ${{c.parsed.y.toLocaleString('pt-BR', {{minimumFractionDigits:1}})}} m³/lig${{flag}}`;
          }} }} }}
        }},
        scales: {{
          x: {{ grid:{{color:'rgba(0,0,0,0.04)'}}, ticks:{{maxRotation:30, font:{{size:10}}}} }},
          y:  {{ position:'left',  title:{{display:true,text:'Volume (m³)'}},
                 ticks:{{callback:v=>v.toLocaleString('pt-BR')+' m³'}}, grid:{{color:'rgba(0,0,0,0.05)'}} }},
          y2: {{ position:'right', title:{{display:true,text:'Ligações'}},
                 ticks:{{callback:v=>v.toLocaleString('pt-BR')}}, grid:{{display:false}} }},
          y3: {{ position:'right', title:{{display:true,text:'Média m³/lig'}},
                 ticks:{{callback:v=>v.toFixed(1)+' m³'}}, grid:{{display:false}}, offset:true }}
        }}
      }}
    }});
  }})();

  // Consumo zero por categoria
  (function() {{
    const ctx = document.getElementById('chart-91-consumo-zero');
    if (!ctx || !CZERO91_L.length) return;
    const totalZero = CZERO91_Q.reduce((a,b)=>a+b,0);
    makeChart(ctx, {{
      type:'bar',
      data:{{ labels:CZERO91_L, datasets:[{{
        label:'Ligações com consumo zero', data:CZERO91_Q,
        backgroundColor:['#C62828','#EF6C00','#F9A825','#1565C0','#2E7D32'],
        borderRadius:6,
      }}] }},
      options:{{
        indexAxis:'y', responsive:true,
        animation:{{duration:1200,easing:'easeOutBack'}},
        plugins:{{ legend:{{ position:'right', labels:{{ padding:12, font:{{size:10}}, usePointStyle:true, pointStyleWidth:8 }}}},
          datalabels: getDatalabelsHBar(v => v.toLocaleString('pt-BR')),
          tooltip:{{...TT, callbacks:{{label: c =>
            ` ${{c.parsed.x.toLocaleString('pt-BR')}} lig (${{totalZero ? (c.parsed.x/totalZero*100).toFixed(1) : 0}}% do total zero)`
          }}}}
        }},
        scales:{{
          x:{{ticks:{{callback:v=>v.toLocaleString('pt-BR')}},grid:{{color:'rgba(0,0,0,0.05)'}}}},
          y:{{grid:{{display:false}}}}
        }}
      }}
    }});
  }})();

  // Volume por dia/mês de leitura  inicia ao abrir a aba
  _rebuild91LeitDia();
}}

// ── toggle/rebuild gráfico Leitura por Dia (escopo global  acessível via onclick) ──
let _chart91LeitDia = null;
let _modo91LeitDia  = (Array.isArray(LDIA91_L) && LDIA91_L.length > 45) ? 'monthly' : 'daily';

function toggle91LeitDia(modo) {{
  _modo91LeitDia = modo;
  const bd = document.getElementById('btn91dia-d');
  const bm = document.getElementById('btn91dia-m');
  if (bd) {{ bd.style.background = modo==='daily' ? '#1565C0' : 'transparent'; bd.style.color = modo==='daily' ? '#fff' : '#6B7A96'; }}
  if (bm) {{ bm.style.background = modo==='monthly' ? '#1565C0' : 'transparent'; bm.style.color = modo==='monthly' ? '#fff' : '#6B7A96'; }}
  const sub = document.getElementById('sub-91-leit-dia');
  if (sub) sub.textContent = modo==='daily'
     ?
     'Visão diária com volume, ligações lidas e linha de média geral por ligação'
    : 'Consolidado mensal com volume total, ligações lidas e linha de média geral por ligação';
  _rebuild91LeitDia();
}}

function _rebuild91LeitDia() {{
  const mensal = _modo91LeitDia === 'monthly';
  let labels, dataV, dataQ;

  if (mensal) {{
    const aggV = _agruparMensal(LDIA91_L, LDIA91_V);
    const aggQ = _agruparMensal(LDIA91_L, LDIA91_Q);
    labels = aggV.labels; dataV = aggV.vals; dataQ = aggQ.vals;
  }} else {{
    labels = LDIA91_L; dataV = LDIA91_V; dataQ = LDIA91_Q;
  }}

  const bd = document.getElementById('btn91dia-d');
  const bm = document.getElementById('btn91dia-m');
  if (bd) {{ bd.style.background = mensal ? 'transparent' : '#1565C0'; bd.style.color = mensal ? '#6B7A96' : '#fff'; }}
  if (bm) {{ bm.style.background = mensal ? '#1565C0' : 'transparent'; bm.style.color = mensal ? '#fff' : '#6B7A96'; }}

  const defaultStart = (!mensal && labels.length > 45) ? labels.length - 45 : 0;

  buildRangeCompare({{
    containerId: 'rc-91-leit-dia',
    labels,
    datasets: [
      {{ label: 'Volume Total (m³)', data: dataV }},
      {{ label: 'Ligações Lidas',    data: dataQ }},
    ],
    defaultStart,
    rebuildFn: (filtL, filtDs) => {{
      _drawChart91LeitDia(filtL, filtDs[0].data, filtDs[1].data, mensal);
    }},
  }});
}}

function _drawChart91LeitDia(labels, dataV, dataQ, mensal) {{
  const ctx = document.getElementById('chart-91-leit-dia');
  if (!ctx) return;
  if (_chart91LeitDia) {{ _chart91LeitDia.destroy(); _chart91LeitDia = null; }}

  const manyPoints  = !mensal && labels.length > 36;
  const totalDiaVol = dataV.reduce((a,b)=>a+b,0);
  const totalDiaQ   = dataQ.reduce((a,b)=>a+b,0);
  const mediaGeralLig = totalDiaQ > 0 ? parseFloat((totalDiaVol / totalDiaQ).toFixed(2)) : 0;
  const mediaGeralSerie = labels.map(() => mediaGeralLig);

  const datasets91LD = [
    {{ type:'bar', label:'Volume Total (m³)', data:dataV,
       backgroundColor:'rgba(21,101,192,0.65)', borderColor:'#1565C0',
       borderWidth:1, borderRadius:4, yAxisID:'y',
       categoryPercentage: manyPoints ? 0.9 : 0.72,
       barPercentage: manyPoints ? 0.95 : 0.8 }},
    {{ type:'line', label:'Ligações Lidas', data:dataQ,
       borderColor:'#2E7D32', backgroundColor:'transparent',
       borderWidth:2.2, pointRadius: mensal ? 4 : (manyPoints ? 0 : 2), pointHoverRadius: 4,
       pointBackgroundColor:'#2E7D32', pointBorderWidth:0,
       fill:false, tension:0.28, yAxisID:'y2' }},
     {{
       type:'line', label:'Média Geral por Ligação (m³)', data:mediaGeralSerie,
       borderColor:'rgba(242,125,22,0.45)', backgroundColor:'transparent',
       borderWidth:1.6, pointRadius:0, pointHoverRadius:0,
       pointBackgroundColor:'transparent', pointBorderWidth:0,
       borderDash:[7,6], fill:false, tension:0, yAxisID:'y3'
     }}
  ];

  const dlCfg = getDatalabelsConfig(mensal ? 'barra' : 'noop', mensal && labels.length <= 18);

  _chart91LeitDia = makeChart(ctx, {{
    data: {{ labels, datasets: datasets91LD }},
    options: {{
      responsive:true, interaction:{{mode:'index',intersect:false}},
      animation:{{duration:900,easing:'easeOutQuart'}},
      plugins: {{
        legend: {{
          position:'top',
          labels: {{
            usePointStyle:true,
            pointStyleWidth:12,
            boxWidth:12,
            padding:12,
            font:{{size:11, weight:'600'}}
          }}
        }},
        datalabels: dlCfg,
        tooltip: {{...TT, callbacks: {{
          title: items => items?.[0]?.label || '',
          label: c => {{
            const lbl  = c.dataset.label || '';
            if (lbl === 'Volume Total (m³)')
              return `  Volume: ${{c.parsed.y.toLocaleString('pt-BR')}} m³`;
            if (lbl === 'Ligações Lidas')
              return `  Ligações: ${{c.parsed.y.toLocaleString('pt-BR')}}`;
            if (lbl === 'Média Geral por Ligação (m³)')
              return `  Média geral: ${{c.parsed.y.toLocaleString('pt-BR',{{minimumFractionDigits:1, maximumFractionDigits:1}})}} m³/lig`;
            return '';
          }}
        }}}}
      }},
      scales: {{
        x: {{
          grid:{{display:false}},
          ticks:{{
            autoSkip:true,
            maxTicksLimit: mensal ? 12 : (manyPoints ? 10 : 18),
            maxRotation: mensal ? 0 : (manyPoints ? 0 : 45),
            minRotation:0,
            font:{{size:10}}
          }}
        }},
        y: {{position:'left', title:{{display:true,text:'Volume m³',font:{{size:10}}}},
             ticks:{{callback:v=>v.toLocaleString('pt-BR')+' m³',font:{{size:10}}}}, grid:{{color:'rgba(0,0,0,0.05)'}}}},
        y2:{{position:'right', title:{{display:true,text:'Ligações',font:{{size:10}}}},
             ticks:{{callback:v=>v.toLocaleString('pt-BR'),font:{{size:10}}}}, grid:{{display:false}}}},
           y3:{{display:false, position:'right', title:{{display:false,text:'Média m³/lig',font:{{size:10}}}},
             ticks:{{display:false, callback:v=>v.toLocaleString('pt-BR')+' m³',font:{{size:10}}}},
             grid:{{display:false}}, offset:true}}
      }}
    }}
  }});
}}

// ── INIT ─────────────────────────────────────────────────────────
// Macro 50012 - KPIs
function initKPIs50012() {{
  // KPIs analíticos
  animateCounter(document.getElementById('k50-lig'),      K50_TOT_LIG, v => Math.round(v).toLocaleString('pt-BR'));
  animateCounter(document.getElementById('k50-vol'),      K50_VOL,     v => Math.round(v).toLocaleString('pt-BR') + ' m³');
  animateCounter(document.getElementById('k50-fat'),      K50_FAT,     fmtBRL);
  animateCounter(document.getElementById('k50-pct-norm'), K50_PCT_NORM, v => v.toFixed(1) + '%');
  animateCounter(document.getElementById('k50-pct-crit'), K50_PCT_CRIT, v => v.toFixed(1) + '%');
  animateCounter(document.getElementById('k50-zeros'),    K50_ZEROS,   v => Math.round(v).toLocaleString('pt-BR'));
  // Subtítulo média
  const subVol = document.getElementById('k50-media-vol');
  if (subVol) subVol.textContent = 'média ' + K50_MEDIA_V.toLocaleString('pt-BR', {{minimumFractionDigits:1}}) + ' m³/lig';
  // Alertas
  animateCounter(document.getElementById('k50-aumento'),  K50_AUMENTO,  v => Math.round(v).toLocaleString('pt-BR'));
  animateCounter(document.getElementById('k50-leit-ant'), K50_LEIT_ANT, v => Math.round(v).toLocaleString('pt-BR'));
  animateCounter(document.getElementById('k50-queda'),    K50_QUEDA,    v => Math.round(v).toLocaleString('pt-BR'));
  animateCounter(document.getElementById('k50-nao-real'), K50_NAO_REAL, v => Math.round(v).toLocaleString('pt-BR'));
  animateCounter(document.getElementById('k50-hd-inv'),   K50_HD_INV,   v => Math.round(v).toLocaleString('pt-BR'));
  // Geográficos
  animateCounter(document.getElementById('k50-total'),    D50_TOTAL,    v => Math.round(v).toLocaleString('pt-BR'));
  animateCounter(document.getElementById('k50-coord'),    D50_COM_COORD, v => Math.round(v).toLocaleString('pt-BR'));
  animateCounter(document.getElementById('k50-sem'),      D50_SEM_COORD, v => Math.round(v).toLocaleString('pt-BR'));
  animateCounter(document.getElementById('k50-cobertura'), D50_COBERTURA, v => v.toFixed(1) + '%');
  if (D50_CENTRO_LAT && D50_CENTRO_LNG) {{
    document.getElementById('geo-centro').textContent = D50_CENTRO_LAT.toFixed(4) + ', ' + D50_CENTRO_LNG.toFixed(4);
  }}
  const bbox = D50_BBOX;
  if (bbox && bbox.north) {{
    document.getElementById('geo-bbox').textContent =
      `N: ${{bbox.north.toFixed(4)}}, S: ${{bbox.south.toFixed(4)}}, E: ${{bbox.east.toFixed(4)}}, W: ${{bbox.west.toFixed(4)}}`;
  }}
}}

// Macro 50012 - Gráficos
function buildCharts50012() {{
  const palBlue = ['#011F3F','#1565C0','#1976D2','#1E88E5','#2196F3','#42A5F5','#64B5F6','#90CAF9','#BBDEFB','#E3F2FD','#0D47A1','#0277BD'];
  const palMix  = ['#2E7D32','#1565C0','#C62828','#FF8F00','#6A1B9A','#37474F','#00838F','#D84315','#558B2F','#AD1457','#0277BD','#F57F17'];

  // ── 1. Críticas (doughnut) ─────────────────────────────────────────────
  (function() {{
    const ctx = document.getElementById('chart-50-criticas');
    if (!ctx || !CRIT50_L.length) return;
    const totalCrit = CRIT50_Q.reduce((a,b)=>a+b,0);
    makeChart(ctx, {{
      type: 'doughnut',
      data: {{ labels: CRIT50_L, datasets: [{{ data: CRIT50_Q,
        backgroundColor: palMix, borderColor: '#fff', borderWidth: 3, hoverOffset: 10 }}] }},
      options: {{
        responsive: true, cutout: '55%',
        layout: {{ padding: 20 }},
        animation: {{ duration: 1400, easing: 'easeOutBack' }},
        plugins: {{
          legend: {{ position:'right', labels:{{ padding:12, font:{{size:11}}, usePointStyle:true, pointStyleWidth:10, generateLabels: (chart) => {{ const ds = chart.data.datasets[0]; return chart.data.labels.map((label, i) => ({{ text: label.length > 22 ? label.slice(0,20)+'' : label, fillStyle: ds.backgroundColor[i], strokeStyle: ds.backgroundColor[i], lineWidth: 0, hidden: false, index: i }})); }} }}}},
          datalabels: getDatalabelsPie(totalCrit, v => v.toLocaleString('pt-BR')),
          tooltip: {{ ...TT, callbacks: {{ label: c =>
            ` ${{c.label}}: ${{c.parsed.toLocaleString('pt-BR')}} (${{(c.parsed/totalCrit*100).toFixed(1)}}%)`
          }} }}
        }}
      }}
    }});
  }})();

  // ── 2. Faixas de consumo (bar horizontal) ─────────────────────────────
  (function() {{
    const ctx = document.getElementById('chart-50-faixas');
    if (!ctx || !FXC50_L.length) return;
    const totalFx = FXC50_Q.reduce((a,b)=>a+b,0);
    const cores = ['#2E7D32','#1565C0','#4CAF50','#1976D2','#FF8F00','#E65100','#C62828','#880E4F','#4A148C'];
    makeChart(ctx, {{
      type: 'bar',
      data: {{ labels: FXC50_L, datasets: [{{ label: 'Ligações', data: FXC50_Q,
        backgroundColor: cores.slice(0, FXC50_L.length), borderRadius: 5 }}] }},
      options: {{
        indexAxis: 'y', responsive: true,
        animation: {{ duration: 1200, easing: 'easeOutBack' }},
        plugins: {{
          legend: {{ display: false }},
          datalabels: getDatalabelsHBar(v => v.toLocaleString('pt-BR')),
          tooltip: {{ ...TT, callbacks: {{ label: c =>
            ` ${{c.parsed.x.toLocaleString('pt-BR')}} lig (${{totalFx ? (c.parsed.x/totalFx*100).toFixed(1) : 0}}%)`
          }} }}
        }},
        scales: {{
          x: {{ ticks: {{ callback: v => v.toLocaleString('pt-BR') }}, grid: {{ color: 'rgba(0,0,0,0.05)' }} }},
          y: {{ grid: {{ display: false }} }}
        }}
      }}
    }});
  }})();

  // ── 3. Cobrança (doughnut) ─────────────────────────────────────────────
  (function() {{
    const ctx = document.getElementById('chart-50-cobranca');
    if (!ctx || !COB50_L.length) return;
    const totalCob = COB50_Q.reduce((a,b)=>a+b,0);
    const coresCob = ['#2E7D32','#FF8F00','#E65100','#C62828','#880E4F','#4A148C','#37474F'];
    makeChart(ctx, {{
      type: 'doughnut',
      data: {{ labels: COB50_L, datasets: [{{ data: COB50_Q,
        backgroundColor: coresCob, borderColor: '#fff', borderWidth: 3, hoverOffset: 10 }}] }},
      options: {{
        responsive: true, cutout: '55%',
        layout: {{ padding: 20 }},
        animation: {{ duration: 1400, easing: 'easeOutBack' }},
        plugins: {{
          legend: {{ position:'right', labels:{{ padding:12, font:{{size:11}}, usePointStyle:true, pointStyleWidth:10, generateLabels: (chart) => {{ const ds = chart.data.datasets[0]; return chart.data.labels.map((label, i) => ({{ text: label.length > 22 ? label.slice(0,20)+'' : label, fillStyle: ds.backgroundColor[i], strokeStyle: ds.backgroundColor[i], lineWidth: 0, hidden: false, index: i }})); }} }}}},
          tooltip: {{ ...TT, callbacks: {{ label: c =>
            ` ${{c.label}}: ${{c.parsed.toLocaleString('pt-BR')}} (${{(c.parsed/totalCob*100).toFixed(1)}}%)`
          }} }}
        }}
      }}
    }});
  }})();

  // ── 4. Leituristas (combo bar + linha taxa) ───────────────────────────
  (function() {{
    const ctx = document.getElementById('chart-50-leituristas');
    if (!ctx || !LEIT50_L.length) return;

    // Ordena por total (normais + críticas) decrescente
    const idx = LEIT50_NO.map((v,i) => i)
      .sort((a,b) => (LEIT50_NO[b]+LEIT50_CR[b]) - (LEIT50_NO[a]+LEIT50_CR[a]));

    const labels  = idx.map(i => LEIT50_L[i]);
    const normais = idx.map(i => LEIT50_NO[i]);
    const criticos= idx.map(i => LEIT50_CR[i]);
    const taxas   = idx.map(i => LEIT50_TC[i]);
    const totais  = idx.map((_,i) => normais[i] + criticos[i]);
    const totalGeral = totais.reduce((a,b)=>a+b,0) || 1;

    makeChart(ctx, {{
      data: {{
        labels,
        datasets: [
          {{
            type: 'bar', label: 'Normais', data: normais,
            backgroundColor: 'rgba(46,125,50,0.80)', borderRadius: 4,
            yAxisID: 'y', stack: 'leit',
            datalabels: {{ display: false }}
          }},
          {{
            type: 'bar', label: 'Críticas', data: criticos,
            backgroundColor: 'rgba(198,40,40,0.80)', borderRadius: 4,
            yAxisID: 'y', stack: 'leit',
            datalabels: {{
              display: (ctx) => {{
                // Só mostra rótulo no topo da barra empilhada (dataset Críticas)
                // e apenas se a barra total for grande o suficiente
                const total = totais[ctx.dataIndex] || 0;
                const maxTotal = Math.max(...totais);
                return total / maxTotal > 0.05;
              }},
              anchor: 'end', align: 'top', offset: 4,
              clamp: true, clip: false,
              font: {{ size: 11, weight: '700' }},
              color: '#1A2535',
              backgroundColor: 'rgba(255,255,255,0.92)',
              borderRadius: 4,
              padding: {{ x: 5, y: 2 }},
              formatter: (v, ctx) => {{
                const total = totais[ctx.dataIndex];
                if (!total) return '';
                const pct = (total / totalGeral * 100).toFixed(0);
                return total.toLocaleString('pt-BR') + ' (' + pct + '%)';
              }}
            }}
          }},
          {{
            type: 'line', label: 'Taxa Crítica (%)', data: taxas,
            borderColor: '#FF8F00', backgroundColor: 'transparent',
            borderWidth: 2, pointRadius: 5, pointBackgroundColor: '#FF8F00',
            pointBorderColor: '#fff', pointBorderWidth: 2,
            tension: 0.3, yAxisID: 'y2',
            datalabels: {{
              display: true,
              anchor: 'end', align: 'top', offset: 6,
              clamp: true, clip: false,
              font: {{ size: 10, weight: '600' }},
              color: '#B45309',
              backgroundColor: 'rgba(255,255,255,0.92)',
              borderRadius: 3,
              padding: {{ x: 4, y: 2 }},
              formatter: v => v != null ? v.toFixed(1) + '%' : ''
            }}
          }},
        ]
      }},
      options: {{
        responsive: true,
        interaction: {{ mode: 'index', intersect: false }},
        animation: {{ duration: 1300, easing: 'easeOutQuart' }},
        plugins: {{
          legend: {{ position: 'top' }},
          datalabels: {{ display: false }},  // desativa global, cada dataset controla o seu
          tooltip: {{ ...TT, callbacks: {{ label: c => {{
            if (c.dataset.label === 'Taxa Crítica (%)')
              return ` Taxa: ${{c.parsed.y.toFixed(1)}}%${{c.parsed.y > 35 ? ' ' : ''}}`;
            const total = totais[c.dataIndex] || 0;
            const pct = total ? (c.parsed.y / total * 100).toFixed(1) : '0';
            return ` ${{c.dataset.label}}: ${{c.parsed.y.toLocaleString('pt-BR')}} (${{pct}}% do total)`;
          }} }} }}
        }},
        scales: {{
          x: {{
            stacked: true,
            grid: {{ display: false }},
            ticks: {{ maxRotation: 30, font: {{ size: 10 }} }}
          }},
          y: {{
            stacked: true,
            grid: {{ color: 'rgba(0,0,0,0.05)' }},
            ticks: {{ callback: v => v.toLocaleString('pt-BR') }},
            title: {{ display: true, text: 'Leituras' }}
          }},
          y2: {{
            position: 'right',
            grid: {{ display: false }},
            ticks: {{ callback: v => v + '%', font: {{ size: 10 }} }},
            title: {{ display: true, text: 'Taxa crítica' }},
            min: 0, max: 100
          }}
        }}
      }}
    }});
  }})();

  // ── 5. Grupos (bar + linha fatura) ────────────────────────────────────
  (function() {{
    const ctx = document.getElementById('chart-50-grupos');
    if (!ctx || !GRP50_L.length) return;
    makeChart(ctx, {{
      data: {{
        labels: GRP50_L,
        datasets: [
          {{ type: 'bar', label: 'Volume m³', data: GRP50_V,
             backgroundColor: 'rgba(21,101,192,0.65)', borderColor: '#1565C0',
             borderWidth: 1, borderRadius: 4, yAxisID: 'y',
             datalabels: {{ display: false }} }},
          {{ type: 'line', label: 'Faturamento R$', data: GRP50_F,
             borderColor: '#FF8F00', backgroundColor: 'transparent',
             borderWidth: 2.5, pointRadius: 4, pointBackgroundColor: '#FF8F00',
             tension: 0.3, yAxisID: 'y2',
             datalabels: {{
               display: true,
               align: (ctx) => ctx.dataIndex % 2 === 0 ? 'top' : 'bottom',
               anchor: 'center',
               offset: 10,
               clamp: true, clip: false,
               font: {{ size: 9, weight: '600' }},
               color: '#B45309',
               backgroundColor: 'rgba(255,255,255,0.88)',
               borderRadius: 3,
               padding: {{ x: 3, y: 1 }},
               formatter: (v, ctx) => {{
                 if (!v || v === 0) return '';
                 const data = ctx.chart.data.datasets[ctx.datasetIndex].data;
                 const max = Math.max(...data.filter(x => x > 0));
                 if (ctx.dataIndex === 0 || ctx.dataIndex === data.length - 1) return fmtBRLk(v);
                 const prev = data[ctx.dataIndex - 1] || v;
                 return Math.abs(v - prev) / max > 0.015 ? fmtBRLk(v) : '';
               }}
             }} }},
        ]
      }},
      options: {{
        responsive: true, interaction: {{ mode: 'index', intersect: false }},
        animation: {{ duration: 1300, easing: 'easeOutQuart' }},
        plugins: {{
          legend: {{ position: 'top' }},
          datalabels: {{ display: false }},
          tooltip: {{ ...TT, callbacks: {{ label: c =>
            c.dataset.label === 'Volume m³'
              ?
               ` Volume: ${{c.parsed.y.toLocaleString('pt-BR')}} m³`
              : ` Fatura: ${{fmtBRL(c.parsed.y)}}`
          }} }}
        }},
        scales: {{
          x: {{ grid: {{ display: false }}, ticks: {{ maxRotation: 45, font: {{ size: 10 }} }} }},
          y:  {{ grid: {{ color: 'rgba(0,0,0,0.05)' }},
                ticks: {{ callback: v => v.toLocaleString('pt-BR') + ' m³' }},
                title: {{ display: true, text: 'm³' }} }},
          y2: {{ position: 'right', grid: {{ display: false }},
                ticks: {{ callback: v => fmtBRLk(v) }},
                title: {{ display: true, text: 'R$' }} }}
        }}
      }}
    }});
  }})();

  // ── 6. Desvio vs média histórica ──────────────────────────────────────
  (function() {{
    const ctx = document.getElementById('chart-50-desvio');
    if (!ctx || !DEV50_L.length) return;
    const totalDev = DEV50_Q.reduce((a,b)=>a+b,0);
    const coresDev = ['#C62828','#E65100','#FF8F00','#2E7D32','#FF8F00','#C62828'];
    makeChart(ctx, {{
      type: 'bar',
      data: {{ labels: DEV50_L, datasets: [{{ label: 'Ligações', data: DEV50_Q,
        backgroundColor: coresDev.slice(0, DEV50_L.length), borderRadius: 5,
        datalabels: {{
          display: true,
          anchor: 'center', align: 'center',
          font: {{ size: 12, weight: 'bold' }},
          color: '#ffffff',
          textShadowBlur: 6,
          textShadowColor: 'rgba(0,0,0,0.6)',
          backgroundColor: 'rgba(0,0,0,0.28)',
          borderRadius: 4,
          padding: {{ x: 7, y: 4 }},
          formatter: (v, ctx) => {{
            if (!v) return '';
            const total = ctx.chart.data.datasets[0].data.reduce((a,b)=>a+(b||0),0);
            return v.toLocaleString('pt-BR');
          }}
        }}
    }}] }},
      options: {{
        responsive: true,
        animation: {{ duration: 1200, easing: 'easeOutBack' }},
        plugins: {{
          legend: {{ display: false }},
          tooltip: {{ ...TT, callbacks: {{ label: c =>
            ` ${{c.parsed.y.toLocaleString('pt-BR')}} lig (${{totalDev ? (c.parsed.y/totalDev*100).toFixed(1) : 0}}%)`
          }} }}
        }},
        scales: {{
          x: {{ grid: {{ display: false }} }},
          y: {{ ticks: {{ callback: v => v.toLocaleString('pt-BR') }},
               grid: {{ color: 'rgba(0,0,0,0.05)' }} }}
        }}
      }}
    }});
  }})();

  // ── 7. Bairros por faturamento (bar horizontal) ───────────────────────
  (function() {{
    const ctx = document.getElementById('chart-50-bairros-fat');
    if (!ctx || !BRF50_L.length) return;
    const totalBrf = BRF50_V.reduce((a,b)=>a+b,0);
    makeChart(ctx, {{
      type: 'bar',
      data: {{ labels: BRF50_L, datasets: [{{ label: 'Faturado R$', data: BRF50_V,
        backgroundColor: palBlue.slice(0, BRF50_L.length), borderRadius: 4 }}] }},
      options: {{
        indexAxis: 'y', responsive: true,
        animation: {{ duration: 1200, easing: 'easeOutBack' }},
        plugins: {{
          legend: {{ display: false }},
          datalabels: getDatalabelsHBar(fmtBRLk),
          tooltip: {{ ...TT, callbacks: {{ label: c =>
            ` ${{fmtBRL(c.parsed.x)}} (${{totalBrf ? (c.parsed.x/totalBrf*100).toFixed(1) : 0}}%)  ${{BRF50_Q[c.dataIndex].toLocaleString('pt-BR')}} lig`
          }} }}
        }},
        scales: {{
          x: {{ ticks: {{ callback: v => fmtBRLk(v) }}, grid: {{ color: 'rgba(0,0,0,0.05)' }} }},
          y: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 10 }} }} }}
        }}
      }}
    }});
  }})();

  // ── 8. Categorias (polarArea) ─────────────────────────────────────────
  (function() {{
    const ctx = document.getElementById('chart-50-categorias');
    if (!ctx || !CAT50_L.length) return;
    const totalCat = CAT50_Q.reduce((a,b)=>a+b,0);
    makeChart(ctx, {{
      type: 'doughnut',
      data: {{ labels: CAT50_L, datasets: [{{ data: CAT50_Q,
        backgroundColor: ['#1565C0','#FF8F00','#2E7D32','#6A1B9A','#C62828','#37474F'],
        borderColor: '#fff', borderWidth: 3, hoverOffset: 10 }}] }},
      options: {{
        responsive: true, cutout: '55%',
        layout: {{ padding: 20 }},
        animation: {{ duration: 1400, easing: 'easeOutBack' }},
        plugins: {{
          legend: {{ position:'right', labels:{{ padding:12, font:{{size:11}}, usePointStyle:true, pointStyleWidth:10, generateLabels: (chart) => {{ const ds = chart.data.datasets[0]; return chart.data.labels.map((label, i) => ({{ text: label.length > 22 ? label.slice(0,20)+'' : label, fillStyle: ds.backgroundColor[i], strokeStyle: ds.backgroundColor[i], lineWidth: 0, hidden: false, index: i }})); }} }}}},
          // Datalabels desativados: com ~95% em RESIDENCIAL as fatias menores
          // empilham labels no topo. A legenda lateral já informa tudo.
          datalabels: {{
            display: (ctx) => {{
              // Só mostra label dentro da fatia se ela tiver >= 4% do total
              const pct = totalCat > 0 ? (ctx.parsed / totalCat * 100) : 0;
              return pct >= 4;
            }},
            anchor: 'center', align: 'center',
            font: {{ size: 12, weight: 'bold' }},
            color: '#fff',
            textShadowBlur: 6, textShadowColor: 'rgba(0,0,0,0.7)',
            formatter: (v) => {{
              const pct = totalCat > 0 ? (v / totalCat * 100).toFixed(1) : '0';
              return pct + '%';
            }}
          }},
          tooltip: {{ ...TT, callbacks: {{ label: c =>
            ` ${{c.label}}: ${{c.parsed.toLocaleString('pt-BR')}} (${{(c.parsed/totalCat*100).toFixed(1)}}%)`
          }} }}
        }}
      }}
    }});
  }})();

  // ── 9. Leituras por dia será inicializado com rebuildChartDias('daily')

  // ── 10. Equipes (pie geográfico) ──────────────────────────────────────
  if (D50_EQUIPES_L.length > 0) {{
    const ctx = document.getElementById('chart-50-equipes');
    if (ctx) {{
      const _totalEq = D50_EQUIPES_Q.reduce((a,b)=>a+b,0) || 1;
      makeChart(ctx, {{
        type: 'pie',
        data: {{ labels: D50_EQUIPES_L, datasets: [{{ data: D50_EQUIPES_Q,
          backgroundColor: ['#1565C0','#FF8F00','#2E7D32','#6A1B9A','#00838F','#D84315','#37474F','#558B2F'],
          borderColor: '#fff', borderWidth: 3, hoverOffset: 10 }}] }},
        options: {{
          responsive: true,
          layout: {{ padding: 30 }},
          animation: {{ duration: 1300, easing: 'easeOutBack' }},
          plugins: {{
            legend: {{ position: 'right', labels: {{ padding: 12, font: {{ size: 11 }}, usePointStyle: true }} }},
            datalabels: {{
              display: (ctx) => (ctx.dataset.data[ctx.dataIndex] / _totalEq * 100) >= 3,
              anchor: (ctx) => (ctx.dataset.data[ctx.dataIndex] / _totalEq * 100) >= 10 ? 'center' : 'end',
              align: (ctx) => (ctx.dataset.data[ctx.dataIndex] / _totalEq * 100) >= 10 ? 'center' : 'end',
              offset: (ctx) => (ctx.dataset.data[ctx.dataIndex] / _totalEq * 100) >= 10 ? 0 : 14,
              clamp: false, clip: false,
              font: {{ size: 11, weight: 'bold' }},
              color: (ctx) => (ctx.dataset.data[ctx.dataIndex] / _totalEq * 100) >= 10 ? '#fff' : '#1A2535',
              textShadowBlur: (ctx) => (ctx.dataset.data[ctx.dataIndex] / _totalEq * 100) >= 10 ? 4 : 0,
              textShadowColor: 'rgba(0,0,0,0.5)',
              backgroundColor: (ctx) => {{
                const pct = ctx.dataset.data[ctx.dataIndex] / _totalEq * 100;
                if (pct >= 10) return 'rgba(255,255,255,0.20)';
                const cores = ['#1565C0','#FF8F00','#2E7D32','#6A1B9A','#00838F','#D84315','#37474F','#558B2F'];
                return cores[ctx.dataIndex % cores.length];
              }},
              borderRadius: 4,
              padding: {{ x: 5, y: 3 }},
              formatter: (v, ctx) => {{
                const pct = (v / _totalEq * 100).toFixed(1);
                return v.toLocaleString('pt-BR') + ' (' + pct + '%)';
              }}
            }},
            tooltip: {{ ...TT, callbacks: {{ label: c =>
              ` ${{c.label}}: ${{c.parsed.toLocaleString('pt-BR')}} (${{(c.parsed/_totalEq*100).toFixed(1)}}%)`
            }} }}
          }}
        }}
      }});
    }}
  }}

  // ── 11. Bairros geográficos (bar horizontal) ─────────────────────────
  if (D50_BAIRROS_L.length > 0) {{
    const ctx = document.getElementById('chart-50-bairros');
    if (ctx) {{
      const totalBar = D50_BAIRROS_Q.reduce((a,b)=>a+b,0);
      makeChart(ctx, {{
        type: 'bar',
        data: {{ labels: D50_BAIRROS_L, datasets: [{{ label: 'Ligações', data: D50_BAIRROS_Q,
          backgroundColor: '#00838F', borderRadius: 4 }}] }},
        options: {{
          indexAxis: 'y', responsive: true,
          animation: {{ duration: 1200, easing: 'easeOutBack' }},
          plugins: {{
            legend: {{ display: false }},
            datalabels: getDatalabelsHBar(v => v.toLocaleString('pt-BR')),
            tooltip: {{ ...TT, callbacks: {{ label: c =>
              ` ${{c.parsed.x.toLocaleString('pt-BR')}} lig (${{totalBar ? (c.parsed.x/totalBar*100).toFixed(1) : 0}}%)`
            }} }}
          }},
          scales: {{
            x: {{ ticks: {{ callback: v => v.toLocaleString('pt-BR') }}, grid: {{ color: 'rgba(0,0,0,0.05)' }} }},
            y: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 10 }} }} }}
          }}
        }}
      }});
    }}
  }}

}}

// ── AGREGAÇÃO DIÁRIA/MENSAL ──────────────────────────────────────────
let chartDias = null;
let modoLeiturasDias = 'daily';  // 'daily' ou 'monthly'

function aggregateDailyToMonthly(labels, quantities, volumes) {{
  const monthlyQ = Array(12).fill(0);
  const monthlyV = Array(12).fill(0);
  const monthNames = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
  let currentMonth = new Date().getMonth();
  
  for (let i = 0; i < labels.length; i++) {{
    const label = String(labels[i]);
    let month = currentMonth;
    
    if (label.includes('-')) {{
      const parts = label.split('-');
      if (parts[0].length === 4) month = parseInt(parts[1]) - 1;
      else month = parseInt(parts[1]) - 1;
    }} else if (label.includes('/')) {{
      const parts = label.split('/');
      if (parts.length >= 2) month = parseInt(parts[1]) - 1;
    }}
    
    if (month >= 0 && month < 12) {{
      monthlyQ[month] += quantities[i] || 0;
      monthlyV[month] += volumes[i] || 0;
    }}
  }}
  
  const monthsWithData = [];
  const dataQ = [], dataV = [];
  for (let i = 0; i < 12; i++) {{
    if (monthlyQ[i] > 0 || monthlyV[i] > 0) {{
      monthsWithData.push(monthNames[i]);
      dataQ.push(monthlyQ[i]);
      dataV.push(monthlyV[i]);
    }}
  }}
  return {{ labels: monthsWithData, quantities: dataQ, volumes: dataV }};
}}

function _initRC50Dias(labels, dataQ, dataV) {{
  buildRangeCompare({{
    containerId: 'rc-50-dias',
    labels:   labels,
    datasets: [
      {{ label: 'Leituras', data: dataQ }},
      {{ label: 'Volume m³', data: dataV }},
    ],
    rebuildFn: (filtL, filtDs, cmpDs) => {{
      _drawChart50Dias(filtL, filtDs[0].data, filtDs[1].data, cmpDs);
    }},
  }});
}}

function _drawChart50Dias(labels, dataQ, dataV, cmpDs) {{
  const ctx = document.getElementById('chart-50-dias');
  if (!ctx) return;
  if (chartDias) {{ chartDias.destroy(); chartDias = null; }}

  const mode = modoLeiturasDias;
  const showLabels = mode === 'monthly';
  const totalDiaQ = dataQ.reduce((a,b)=>a+b,0);
  const mensal50 = mode === 'monthly';

  // Agrupa cmpDs por tipo para ordem correta
  const cmp50PorTipo = {{}};
  (cmpDs || []).forEach(d => {{
    const t = d._tipo || 'unknown';
    if (!cmp50PorTipo[t]) cmp50PorTipo[t] = [];
    cmp50PorTipo[t].push({{ ...d, _isCompare: true }});
  }});
  const tiposOrdem50 = ['month', 'year', 'year_month', 'prev_month_year', 'unknown'];
  const cmpOrdenados50 = tiposOrdem50.flatMap(t => cmp50PorTipo[t] || []);

  const baseDatasets = [
    {{ type: 'bar', label: 'Leituras', data: dataQ,
      backgroundColor: 'rgba(21,101,192,0.6)', borderColor: '#1565C0',
      borderWidth: 1, borderRadius: 3, yAxisID: 'y' }},
    {{ type: 'line', label: 'Volume m³', data: dataV,
      borderColor: '#FF8F00', backgroundColor: 'transparent',
      borderWidth: 2, pointRadius: 4, pointBackgroundColor: '#FF8F00',
      tension: 0.3, yAxisID: 'y2', borderDash: [4,3] }},
    ...cmpOrdenados50.map(d => ({{ ...d, yAxisID: 'y', order: 3 }}))
  ];

  chartDias = makeChart(ctx, {{
    type: 'bar',
    data: {{ labels: labels, datasets: baseDatasets }},
    options: {{
      responsive: true, interaction: {{ mode: 'index', intersect: false }},
      animation: {{ duration: 1400, easing: 'easeOutQuart' }},
      plugins: {{
        legend: {{ position: 'top' }},
        datalabels: getDatalabelsConfig('barra', showLabels),
        tooltip: {{ ...TT, callbacks: {{
          title: items => _fmtTitleComparacao(items, baseDatasets, mensal50),
          label: c => {{
            const lbl  = c.dataset.label || '';
            const tipo = c.dataset._tipo || '';
            const val  = c.parsed.y;
            const idx  = c.dataIndex;

            if (c.dataset._isCompare) {{
              const meta = {{
                year:  {{ icone: '', tag: 'ano passado' }},
                prev_month_year: {{ icone: '', tag: 'mês ant. 1 ano' }},
                month: {{ icone: '↩',  tag: 'mês anterior' }},
              }}[tipo] || {{ icone: '↔', tag: tipo }};
              // Delta % vs série atual de Leituras
              const valAtual = dataQ[idx] ?? null;
              const valComp  = val;
              let deltaStr = '';
              if (valAtual != null && valComp != null && valComp > 0) {{
                const delta = ((valAtual - valComp) / valComp) * 100;
                deltaStr = `  ${{delta >= 0 ? '▲' : '▼'}} ${{delta >= 0 ? '+' : ''}}${{delta.toFixed(1)}}%`;
              }}
              return ` ${{meta.icone}} Leituras (${{meta.tag}}): ${{val.toLocaleString('pt-BR')}}${{deltaStr}}`;
            }}

            if (lbl === 'Leituras')
              return `  Leituras: ${{val.toLocaleString('pt-BR')}} (${{totalDiaQ ? (val/totalDiaQ*100).toFixed(1) : 0}}%)`;
            return `  Volume: ${{val.toLocaleString('pt-BR')}} m³`;
          }}
        }} }},
      }},
      scales: {{
        x: {{ grid: {{ display: false }}, title: {{ display: true, text: mode === 'monthly' ? 'Mês' : 'Dia do mês' }} }},
        y:  {{ grid: {{ color: 'rgba(0,0,0,0.05)' }},
              ticks: {{ callback: v => v.toLocaleString('pt-BR') }},
              title: {{ display: true, text: 'Leituras' }} }},
        y2: {{ position: 'right', grid: {{ display: false }},
              ticks: {{ callback: v => v.toLocaleString('pt-BR') + ' m³' }},
              title: {{ display: true, text: 'Volume m³' }} }}
      }}
    }}
  }});
}}

function rebuildChartDias(mode) {{
  if (!DIA50_L.length) return;

  let labels = DIA50_L;
  let dataQ = DIA50_Q;
  let dataV = DIA50_V;

  if (mode === 'monthly') {{
    const agg = aggregateDailyToMonthly(DIA50_L, DIA50_Q, DIA50_V);
    labels = agg.labels;
    dataQ = agg.quantities;
    dataV = agg.volumes;
  }}

  _initRC50Dias(labels, dataQ, dataV);
}}


// ── INICIALIZAÇÃO  compatível com script no final do <body> ──────────
// Quando o script fica no fim do body, DOMContentLoaded já disparou;
// usamos readyState para disparar imediatamente nesses casos.

// ── GRÁFICOS DO PAINEL MAPEAMENTO ────────────────────────────────────────────

function buildChartDistribuicao() {{
  const labels = DIST_FUNC_L;
  const data   = DIST_FUNC_Q;
  if (!labels || !labels.length) return;
  const total = data.reduce((a,b)=>a+b,0) || 1;
  const cores  = ['#1565C0','#FF8F00','#6A1B9A','#1976D2','#00838F',
                  '#AD1457','#F57C00','#5E35B1','#0097A7','#7B1FA2'];

  // Pizza
  (function() {{
    const c = Chart.getChart('chart-distribuicao-pie');
    if (c) c.destroy();
    const ctx = document.getElementById('chart-distribuicao-pie');
    if (!ctx) return;
    makeChart(ctx, {{
      type: 'pie',
      data: {{ labels,
        datasets: [{{ data, backgroundColor: cores.slice(0,labels.length),
          borderColor:'#fff', borderWidth:2.5, hoverOffset:8 }}] }},
      options: {{
        responsive:true, maintainAspectRatio:true,
        animation:{{ duration:1300, easing:'easeOutBack' }},
        plugins:{{
          datalabels: getDatalabelsPie(total),
          legend:{{ position:'bottom', labels:{{ padding:8, font:{{size:10,weight:'500'}} }} }},
          tooltip:{{...TT, callbacks:{{label: c =>
            ` ${{c.label}}: ${{c.parsed.toLocaleString('pt-BR')}} (${{(c.parsed/total*100).toFixed(1)}}%)`
          }}}}
        }}
      }}
    }});
  }})();

  // Barras horizontais
  (function() {{
    const c = Chart.getChart('chart-funcionarios-barras');
    if (c) c.destroy();
    const ctx = document.getElementById('chart-funcionarios-barras');
    if (!ctx) return;
    makeChart(ctx, {{
      type: 'bar',
      data: {{ labels,
        datasets: [{{ label:'Total de OSs', data,
          backgroundColor: cores.slice(0,labels.length),
          borderRadius:6 }}] }},
      options: {{
        indexAxis:'y', responsive:true, maintainAspectRatio:true,
        animation:{{ duration:1200, easing:'easeOutBack' }},
        layout:{{ padding:{{ right:80 }} }},
        plugins:{{
          legend:{{ display:false }},
          datalabels:{{
            display:true, anchor:'end', align:'right', offset:6,
            clamp:false, clip:false,
            font:{{size:11,weight:'700'}}, color:'#1A2535',
            formatter:(v)=> v ? v.toLocaleString('pt-BR')+' ('+( v/total*100).toFixed(1)+'%)' : ''
          }},
          tooltip:{{...TT, callbacks:{{label: c=>
            ` ${{c.parsed.x.toLocaleString('pt-BR')}} OSs (${{(c.parsed.x/total*100).toFixed(1)}}%)`
          }}}}
        }},
        scales:{{
          x:{{ ticks:{{callback:v=>v.toLocaleString('pt-BR')}}, grid:{{color:'rgba(0,0,0,0.05)'}} }},
          y:{{ grid:{{display:false}} }}
        }}
      }}
    }});
  }})();
}}

function buildChartHidEspelho() {{
  (function() {{
    const ctxEl = document.getElementById('chart-hid-espelho');
    if (!ctxEl || !HID_MODELOS.length) return;

    const labels   = HID_MODELOS;
    const totais   = HID_TOTAIS;
    const tiposRaw = HID_TIPOS_RAW;
    const maxTotal = Math.max(...totais, 1);

    const CORES_TIPOS = [
      '#C62828','#2E7D32','#1565C0','#FF8F00','#6A1B9A',
      '#00838F','#AD1457','#37474F','#558B2F','#D84315'
    ];
    const tipoColorMap = {{}};
    TOP_TIPOS_OS.forEach((t, i) => {{ tipoColorMap[t] = CORES_TIPOS[i % CORES_TIPOS.length]; }});
    function corTipo(t) {{ return tipoColorMap[t] || '#90A4AE'; }}

    const dpr    = window.devicePixelRatio || 1;
    const W      = ctxEl.parentElement.offsetWidth || 900;
    const rowH   = 36;
    const PAD_T  = 40;
    const PAD_B  = 40;
    const H      = PAD_T + labels.length * rowH + PAD_B;

    ctxEl.width  = W * dpr;
    ctxEl.height = H * dpr;
    ctxEl.style.width  = W + 'px';
    ctxEl.style.height = H + 'px';

    const ctx = ctxEl.getContext('2d');
    ctx.scale(dpr, dpr);

    // Layout 3 colunas: [barra total] | [modelo central] | [top 3 tipos]
    const COL_MDL_W = 60;
    const PAD_L     = 10;
    const PAD_R     = 10;
    const GAP       = 6;
    const totalAreaW = W - PAD_L - PAD_R;
    const xMdlL     = PAD_L + (totalAreaW - COL_MDL_W) / 2;
    const xMdlR     = xMdlL + COL_MDL_W;
    const colEsqW   = xMdlL - GAP - PAD_L;
    const colDirW   = W - PAD_R - xMdlR - GAP;
    const xDirStart = xMdlR + GAP;

    // ── Tooltip ──────────────────────────────────────────────────────────────
    let _espTip = document.getElementById('hid-esp-tip');
    if (!_espTip) {{
      _espTip = document.createElement('div');
      _espTip.id = 'hid-esp-tip';
      _espTip.style.cssText = [
        'position:fixed','z-index:9999','pointer-events:none','display:none',
        'background:rgba(1,31,63,0.95)','color:#fff','font-size:12px','line-height:1.5',
        'padding:8px 12px','border-radius:8px','max-width:320px','white-space:pre-line',
        'box-shadow:0 4px 16px rgba(0,0,0,0.25)','border:1px solid rgba(255,255,255,0.12)'
      ].join(';');
      document.body.appendChild(_espTip);
    }}

    // Metadados por linha para hit-test
    const _espRows = labels.map((modelo, i) => {{
      const tipos = (tiposRaw[i] || []).slice(0, 3);
      const soma  = tipos.reduce((s, t) => s + t.qtd, 0) || 1;
      return {{ modelo, total: totais[i], tipos, soma }};
    }});

    ctxEl.addEventListener('mousemove', function(e) {{
      const rect = ctxEl.getBoundingClientRect();
      const mx   = e.clientX - rect.left;
      const my   = e.clientY - rect.top;
      const rowIdx = Math.floor((my - PAD_T) / rowH);
      if (rowIdx < 0 || rowIdx >= _espRows.length) {{ _espTip.style.display = 'none'; return; }}
      const row = _espRows[rowIdx];

      // Hover na coluna esquerda (total)
      const xBarEnd   = PAD_L + colEsqW;
      const totalW    = Math.max(4, (row.total / maxTotal) * (colEsqW - 8));
      const xBarStart = xBarEnd - totalW;
      if (mx >= xBarStart && mx <= xBarEnd) {{
        _espTip.textContent = row.modelo + '  \u25b8  Total\\n' + row.total + ' OSs no per\u00edodo';
        _showEspTip(e); return;
      }}

      // Hover na coluna direita (top 3 tipos)
      if (mx >= xDirStart && mx <= xDirStart + colDirW) {{
        let xCur = xDirStart;
        for (const t of row.tipos) {{
          const segW = Math.max(1, (t.qtd / row.soma) * colDirW);
          if (mx >= xCur && mx <= xCur + segW) {{
            const pct = ((t.qtd / row.soma) * 100).toFixed(1);
            _espTip.textContent = row.modelo + '  \u25b8  ' + t.tipo + '\\n' + t.qtd + ' OSs  (' + pct + '% dos top 3)';
            _showEspTip(e); return;
          }}
          xCur += segW;
        }}
      }}

      _espTip.style.display = 'none';
    }});

    ctxEl.addEventListener('mouseleave', () => {{ _espTip.style.display = 'none'; }});
    ctxEl.style.cursor = 'default';

    function _showEspTip(e) {{
      _espTip.style.display = 'block';
      const tw = _espTip.offsetWidth, th = _espTip.offsetHeight;
      let tx = e.clientX + 14, ty = e.clientY - th / 2;
      if (tx + tw > window.innerWidth - 8) tx = e.clientX - tw - 14;
      if (ty < 8) ty = 8;
      if (ty + th > window.innerHeight - 8) ty = window.innerHeight - th - 8;
      _espTip.style.left = tx + 'px';
      _espTip.style.top  = ty + 'px';
    }}

    // Abreviaturas (mesmo padrão do retroativo)
    function abreviarTipo(nome, pixelsDisp) {{
      const ABREVS = [
        [/SUBSTITUIÇÃO DE HIDRÔMETRO\\s*-?\\s*/i, 'Sub.Hidr. '],
        [/VERIFICAÇÃO DE HIDRÔMETRO A PEDIDO DO CLIENTE/i, 'Verif.(cliente)'],
        [/VERIFICAÇÃO DE HIDRÔMETRO/i, 'Verif.Hidr.'],
        [/VERIFICAÇÃO DE EXCESSO DE CONSUMO\\s*-?\\s*/i, 'Verif.Exc.Cons. '],
        [/VERIFICAÇÃO DE CONSUMO ZERO/i, 'Verif.Cons.Zero'],
        [/CORREÇÃO DE HIDRÔMETRO INVERTIDO/i, 'Corr.Invertido'],
        [/IDENTIFICAR PAGAMENTO DE FATURA/i, 'Ident.Pgto.Fat.'],
        [/SUSPENSÃO ABASTECIMENTO/i, 'Susp.Abast.'],
        [/SOLICITAÇÃO\\s*-?\\s*RECALC/i, 'Solic.Recalc.'],
        [/RECOMPOSIÇÃO DE LAJOTA/i, 'Recomp.Lajota'],
        [/MANUTENÇÃO DE ABRIGO/i, 'Man.Abrigo'],
        [/PREVENTIVA/i, 'Prev.'],
        [/CORRETIVA/i, 'Corret.'],
        [/DANIFICADO/i, 'Danif.'],
        [/INSTALAÇÃO/i, 'Inst.'],
      ];
      let resultado = nome;
      for (const [regex, abrev] of ABREVS) {{
        resultado = resultado.replace(regex, abrev).trim();
        if (ctx.measureText(resultado).width <= pixelsDisp) break;
      }}
      return resultado;
    }}

    // Cabecalhos
    ctx.fillStyle = '#6B7A96';
    ctx.font = 'bold 10px Segoe UI, system-ui, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('Total', PAD_L + colEsqW / 2, 18);
    ctx.fillText('Modelo', xMdlL + COL_MDL_W / 2, 18);
    ctx.fillText('Top 3 Tipos de OS  (qtd.)', xMdlR + GAP + colDirW / 2, 18);

    ctx.strokeStyle = '#DDE3EE';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(PAD_L, 24); ctx.lineTo(W - PAD_R, 24);
    ctx.stroke();

    labels.forEach((modelo, i) => {{
      const y    = PAD_T + i * rowH;
      const bH   = rowH - 6;
      const yBar = y + 3;
      const yMid = yBar + bH / 2;

      if (i % 2 === 0) {{
        ctx.fillStyle = 'rgba(0,0,0,0.022)';
        ctx.fillRect(0, y, W, rowH);
      }}

      // Barra esquerda (total de OSs) -- cresce da coluna Modelo para a esquerda
      const totalW = Math.max(4, (totais[i] / maxTotal) * (colEsqW - 8));
      const xBarEnd  = PAD_L + colEsqW;
      const xBarStart = xBarEnd - totalW;
      const rL     = Math.min(5, bH / 2);

      ctx.shadowColor = 'rgba(21,101,192,0.18)';
      ctx.shadowBlur  = 4;
      ctx.fillStyle   = '#1565C0';
      ctx.beginPath();
      ctx.moveTo(xBarStart + rL, yBar);
      ctx.lineTo(xBarEnd, yBar);
      ctx.lineTo(xBarEnd, yBar + bH);
      ctx.lineTo(xBarStart + rL, yBar + bH);
      ctx.arcTo(xBarStart, yBar + bH, xBarStart, yBar, rL);
      ctx.arcTo(xBarStart, yBar, xBarStart + rL, yBar, rL);
      ctx.closePath();
      ctx.fill();
      ctx.shadowBlur = 0;

      ctx.font = 'bold 11px Segoe UI, system-ui, sans-serif';
      if (totalW > 22) {{
        ctx.fillStyle = '#fff';
        ctx.textAlign = 'center';
        ctx.fillText(totais[i], xBarStart + totalW / 2, yMid + 4);
      }} else {{
        ctx.fillStyle = '#1565C0';
        ctx.textAlign = 'right';
        ctx.fillText(totais[i], xBarStart - 4, yMid + 4);
      }}

      // Coluna central: modelo (pill escuro)
      ctx.fillStyle = '#011F3F';
      const mdlR = 5;
      if (ctx.roundRect) {{
        ctx.beginPath(); ctx.roundRect(xMdlL, yBar + 1, COL_MDL_W, bH - 2, mdlR); ctx.fill();
      }} else {{
        ctx.fillRect(xMdlL, yBar + 1, COL_MDL_W, bH - 2);
      }}
      ctx.fillStyle = '#FFFFFF';
      ctx.font = 'bold 11px Segoe UI, system-ui, sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText(modelo, xMdlL + COL_MDL_W / 2, yMid + 4);

      // Coluna direita: top 3 tipos de OS
      const tipos = (tiposRaw[i] || []).slice(0, 3);
      if (!tipos.length) return;

      const somaTop3 = tipos.reduce((s, t) => s + t.qtd, 0) || 1;
      let xCur = xDirStart;

      tipos.forEach((t, ti) => {{
        const segW  = Math.max(1, (t.qtd / somaTop3) * colDirW);
        const isLast = ti === tipos.length - 1;
        const r      = isLast ? Math.min(5, bH / 2) : 0;
        const cor    = corTipo(t.tipo);

        ctx.shadowColor = 'rgba(0,0,0,0.08)';
        ctx.shadowBlur  = 2;
        ctx.fillStyle   = cor;

        if (r > 0) {{
          ctx.beginPath();
          ctx.moveTo(xCur, yBar);
          ctx.lineTo(xCur + segW - r, yBar);
          ctx.arcTo(xCur + segW, yBar, xCur + segW, yBar + r, r);
          ctx.arcTo(xCur + segW, yBar + bH, xCur + segW - r, yBar + bH, r);
          ctx.lineTo(xCur, yBar + bH);
          ctx.closePath();
          ctx.fill();
        }} else {{
          ctx.fillRect(xCur, yBar, segW, bH);
        }}
        ctx.shadowBlur = 0;

        // Separador entre segmentos
        if (ti < tipos.length - 1 && segW > 2) {{
          ctx.strokeStyle = 'rgba(255,255,255,0.55)';
          ctx.lineWidth   = 1.5;
          ctx.beginPath();
          ctx.moveTo(xCur + segW, yBar + 2);
          ctx.lineTo(xCur + segW, yBar + bH - 2);
          ctx.stroke();
        }}

        // Texto dentro da barra
        const PAD_INNER = 5;
        const pixDisp   = segW - PAD_INNER * 2;
        ctx.font = 'bold 9px Segoe UI, system-ui, sans-serif';
        if (pixDisp > 18) {{
          ctx.fillStyle = '#fff';
          ctx.textAlign = 'left';
          const qtdStr   = ' (' + t.qtd + ')';
          const qtdW     = ctx.measureText(qtdStr).width;
          const nomeDisp = pixDisp - qtdW;
          let nomeExib   = t.tipo;
          if (ctx.measureText(nomeExib).width > nomeDisp) {{
            nomeExib = abreviarTipo(nomeExib, nomeDisp);
          }}
          if (ctx.measureText(nomeExib).width > nomeDisp) {{
            while (nomeExib.length > 2 && ctx.measureText(nomeExib + '').width > nomeDisp) {{
              nomeExib = nomeExib.slice(0, -1);
            }}
            nomeExib = nomeExib.trimEnd() + '';
          }}
          // Se nomeDisp for minúsculo, exibe só abreviação sem qtd
          if (nomeDisp < 10) {{
            const soloDisp = pixDisp;
            nomeExib = abreviarTipo(t.tipo, soloDisp);
            if (ctx.measureText(nomeExib).width > soloDisp) {{
              while (nomeExib.length > 2 && ctx.measureText(nomeExib + '').width > soloDisp) {{
                nomeExib = nomeExib.slice(0, -1);
              }}
              nomeExib = nomeExib.trimEnd() + '';
            }}
            ctx.fillText(nomeExib, xCur + PAD_INNER, yMid + 3);
          }} else {{
            ctx.fillText(nomeExib + qtdStr, xCur + PAD_INNER, yMid + 3);
          }}
        }} else if (pixDisp > 8) {{
          ctx.fillStyle = '#fff';
          ctx.textAlign = 'center';
          ctx.fillText(t.qtd, xCur + segW / 2, yMid + 3);
        }}

        xCur += segW;
      }});
    }});

    // Linhas separadoras verticais
    ctx.strokeStyle = '#DDE3EE';
    ctx.lineWidth = 1.5;
    [xMdlL - GAP/2, xMdlR + GAP/2].forEach(x => {{
      ctx.beginPath();
      ctx.moveTo(x, 26);
      ctx.lineTo(x, PAD_T + labels.length * rowH + 4);
      ctx.stroke();
    }});

    // Legenda na base
    const legY = PAD_T + labels.length * rowH + 14;
    let legX   = PAD_L;
    ctx.font = '9px Segoe UI, system-ui, sans-serif';
    TOP_TIPOS_OS.slice(0, 8).forEach((tipo, i) => {{
      const cor = corTipo(tipo);
      ctx.fillStyle = cor;
      ctx.fillRect(legX, legY, 9, 9);
      ctx.fillStyle = '#475569';
      ctx.textAlign = 'left';
      const lbl = tipo.length > 24 ? tipo.slice(0, 22) + '...' : tipo;
      ctx.fillText(lbl, legX + 12, legY + 8);
      legX += ctx.measureText(lbl).width + 24;
      if (legX > W - 160) legX = PAD_L;
    }});

  }})();
}}

function buildChartHidRetroativo() {{
  const ctxEl = document.getElementById('chart-hid-retroativo');
  if (!ctxEl || typeof HID_RET_SERIES === 'undefined' || !HID_RET_SERIES.length) return;

  // Função de parse de data reutilizada
  const parseDate = s => {{
    const str = String(s).trim();
    if (str.includes('/')) {{ const [m, y] = str.split('/'); return parseInt(y) * 100 + parseInt(m); }}
    if (str.includes('-')) {{ const [y, m] = str.split('-'); return parseInt(y) * 100 + parseInt(m); }}
    return 0;
  }};

  // 1. USA HID_RET_MESES COMO BASE DO SLIDER (período completo do payload)
  //    Inclui TODOS os meses, mesmo aqueles sem modelo identificado (modelo "")
  const todosOsMeses = (typeof HID_RET_MESES !== 'undefined' && HID_RET_MESES.length)
    ? [...HID_RET_MESES]
    : [...new Set(HID_RET_SERIES.map(r => r.mes))];

  // 2. ORDENAÇÃO CRONOLÓGICA BLINDADA (Mais antigo na esquerda, mais novo na direita)
  const mesesCrono = todosOsMeses.sort((a, b) => parseDate(a) - parseDate(b));

  const N = mesesCrono.length;
  if (N === 0) return;

  // 3. DADOS VÁLIDOS PARA RENDERIZAÇÃO DAS BARRAS
  //    Modelo válido = pelo menos 1 letra E pelo menos 1 número (ex: Y14, A25, Z25)
  //    Exclui: "", "adm", "200", "400", vazios, etc.
  const dadosValidos = HID_RET_SERIES.filter(r => {{
    if (!r.modelo || r.qtd <= 0) return false;
    const m = String(r.modelo).trim();
    return /[a-zA-Z]/.test(m) && /[0-9]/.test(m);
  }});

  // ── Configuração de Cores ──
  const CORES_TIPOS = ['#C62828','#2E7D32','#1565C0','#FF8F00','#6A1B9A','#00838F','#AD1457','#37474F','#558B2F','#D84315'];
  const tipoTotalGlobal = {{}};
  dadosValidos.forEach(r => {{ tipoTotalGlobal[r.tipo] = (tipoTotalGlobal[r.tipo] || 0) + r.qtd; }});
  const tiposOrdenados = Object.keys(tipoTotalGlobal).sort((a,b) => tipoTotalGlobal[b]-tipoTotalGlobal[a]);
  const tipoColorMap = {{}};
  tiposOrdenados.forEach((t, i) => {{ tipoColorMap[t] = CORES_TIPOS[i % CORES_TIPOS.length]; }});
  function corTipo(t) {{ return tipoColorMap[t] || '#90A4AE'; }}

  // ── Lógica de Agregação ──
  function agregaPorModelo(from, to) {{
    const mesesSlice = new Set(mesesCrono.slice(from, to + 1));
    const acc = {{}};
    dadosValidos.forEach(r => {{
      if (!mesesSlice.has(r.mes)) return;
      const mod = String(r.modelo).trim();
      if (!acc[mod]) acc[mod] = {{ total: 0, tipos: {{}} }};
      acc[mod].total += r.qtd;
      acc[mod].tipos[r.tipo] = (acc[mod].tipos[r.tipo] || 0) + r.qtd;
    }});
    return Object.entries(acc).map(([modelo, v]) => ({{
        modelo, total: v.total,
        tipos: Object.entries(v.tipos).map(([tipo, qtd]) => ({{ tipo, qtd }})).sort((a,b) => b.qtd - a.qtd)
    }})).sort((a,b) => b.total - a.total);
  }}

  // ── Slider e Interface ──
  const wrapper = ctxEl.parentElement;
  if (!wrapper.querySelector('.hid-ret-slider-wrap')) {{
    const sliderWrap = document.createElement('div');
    sliderWrap.className = 'hid-ret-slider-wrap';
    sliderWrap.style.cssText = 'padding:10px 16px 4px;user-select:none';
    sliderWrap.innerHTML = `
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:6px">
        <span style="font-size:11px;color:#6B7A96;white-space:nowrap">Período Selecionado:</span>
        <span id="hid-ret-label" style="font-size:12px;font-weight:600;color:#1565C0;flex:1;text-align:center"></span>
      </div>
      <div class="hid-ret-track-wrap" style="position:relative;height:28px;display:flex;align-items:center">
        <div class="hid-ret-track" style="position:absolute;left:0;right:0;height:4px;background:#DDE3EE;border-radius:2px">
          <div id="hid-ret-fill" style="position:absolute;height:100%;background:#1565C0;border-radius:2px"></div>
        </div>
        <input id="hid-ret-from" type="range" min="0" max="${{N-1}}" value="0" step="1" style="position:absolute;width:100%;appearance:none;background:transparent;pointer-events:none;height:28px">
        <input id="hid-ret-to" type="range" min="0" max="${{N-1}}" value="${{N-1}}" step="1" style="position:absolute;width:100%;appearance:none;background:transparent;pointer-events:none;height:28px">
      </div>
      <style>
        #hid-ret-from::-webkit-slider-thumb,#hid-ret-to::-webkit-slider-thumb{{
          appearance:none;width:16px;height:16px;border-radius:50%;background:#1565C0;border:2px solid #fff;box-shadow:0 1px 4px rgba(0,0,0,.25);pointer-events:all;cursor:pointer;position:relative;z-index:2;
        }}
      </style>`;
    wrapper.insertBefore(sliderWrap, ctxEl);
  }}

  // ── Tooltip overlay ──
  let _hidTip = document.getElementById('hid-ret-tip');
  if (!_hidTip) {{
    _hidTip = document.createElement('div');
    _hidTip.id = 'hid-ret-tip';
    _hidTip.style.cssText = [
      'position:fixed','z-index:9999','pointer-events:none','display:none',
      'background:rgba(1,31,63,0.95)','color:#fff','font-size:12px','line-height:1.5',
      'padding:8px 12px','border-radius:8px','max-width:320px','white-space:pre-line',
      'box-shadow:0 4px 16px rgba(0,0,0,0.25)','border:1px solid rgba(255,255,255,0.12)'
    ].join(';');
    document.body.appendChild(_hidTip);
  }}

  // Guarda metadados das linhas para hit-test no mousemove
  let _hidRows = [], _hidMeta = {{}};

  function _updateHidMeta(rows, W, rowH, PAD_T, COL_MDL_W, PAD_L, PAD_R, GAP) {{
    const totalAreaW = W - PAD_L - PAD_R;
    const xMdlL = PAD_L + (totalAreaW - COL_MDL_W) / 2;
    const xMdlR = xMdlL + COL_MDL_W;
    const colDirW = W - PAD_R - xMdlR - GAP;
    const xDirStart = xMdlR + GAP;
    _hidRows = rows;
    _hidMeta = {{ rowH, PAD_T, xDirStart, colDirW }};
  }}

  ctxEl.addEventListener('mousemove', function(e) {{
    if (!_hidRows.length) return;
    const rect  = ctxEl.getBoundingClientRect();
    const mx    = e.clientX - rect.left;
    const my    = e.clientY - rect.top;
    const {{ rowH, PAD_T, xDirStart, colDirW }} = _hidMeta;

    const rowIdx = Math.floor((my - PAD_T) / rowH);
    if (rowIdx < 0 || rowIdx >= _hidRows.length) {{ _hidTip.style.display = 'none'; return; }}
    const row = _hidRows[rowIdx];
    if (mx < xDirStart || mx > xDirStart + colDirW) {{ _hidTip.style.display = 'none'; return; }}

    // Descobrir em qual segmento está
    const tipos3   = row.tipos.slice(0, 3);
    const somaTop3 = tipos3.reduce((s, t) => s + t.qtd, 0) || 1;
    let xCur = xDirStart;
    let tipContent = null;
    for (const t of tipos3) {{
      const segW = Math.max(2, (t.qtd / somaTop3) * colDirW);
      if (mx >= xCur && mx <= xCur + segW) {{
        const pct = ((t.qtd / somaTop3) * 100).toFixed(1);
        tipContent = `${{row.modelo}}  ▸  ${{t.tipo}}\n${{t.qtd}} OSs  (${{pct}}% dos top 3)`;
        break;
      }}
      xCur += segW;
    }}

    if (tipContent) {{
      _hidTip.textContent = tipContent;
      _hidTip.style.display = 'block';
      const tw = _hidTip.offsetWidth, th = _hidTip.offsetHeight;
      let tx = e.clientX + 14, ty = e.clientY - th / 2;
      if (tx + tw > window.innerWidth - 8) tx = e.clientX - tw - 14;
      if (ty < 8) ty = 8;
      if (ty + th > window.innerHeight - 8) ty = window.innerHeight - th - 8;
      _hidTip.style.left = tx + 'px';
      _hidTip.style.top  = ty + 'px';
    }} else {{
      _hidTip.style.display = 'none';
    }}
  }});
  ctxEl.addEventListener('mouseleave', () => {{ _hidTip.style.display = 'none'; }});
  ctxEl.style.cursor = 'default';

  const inpFrom = document.getElementById('hid-ret-from');
  const inpTo   = document.getElementById('hid-ret-to');
  const fillEl  = document.getElementById('hid-ret-fill');
  const labelEl = document.getElementById('hid-ret-label');

  function updateSliderUI(from, to) {{
    const pct = 100 / Math.max(N - 1, 1);
    fillEl.style.left  = (from * pct) + '%';
    fillEl.style.width = ((to - from) * pct) + '%';
    labelEl.textContent = mesesCrono[from] + ' → ' + mesesCrono[to];
  }}

  function drawCanvas(from, to) {{
    const rows = agregaPorModelo(from, to);
    const maxTotal = Math.max(...rows.map(r => r.total), 1);
    const dpr = window.devicePixelRatio || 1;
    const W = ctxEl.parentElement.offsetWidth || 900;
    const rowH = 42; const PAD_T = 44; const PAD_B = 48;
    const H = PAD_T + Math.max(rows.length, 1) * rowH + PAD_B;

    ctxEl.width = W * dpr; ctxEl.height = H * dpr;
    ctxEl.style.width = W + 'px'; ctxEl.style.height = H + 'px';
    const ctx = ctxEl.getContext('2d');
    ctx.scale(dpr, dpr); ctx.clearRect(0, 0, W, H);
    ctx.clearRect(0, 0, W, H);

    const COL_MDL_W  = 60;
    const PAD_L      = 10;
    const PAD_R      = 10;
    const GAP        = 6;
    const totalAreaW = W - PAD_L - PAD_R;
    const xMdlL      = PAD_L + (totalAreaW - COL_MDL_W) / 2;
    const xMdlR      = xMdlL + COL_MDL_W;
    const colEsqW    = xMdlL - GAP - PAD_L;
    const colDirW    = W - PAD_R - xMdlR - GAP;

    // Atualiza metadados para o tooltip
    _updateHidMeta(rows, W, rowH, PAD_T, COL_MDL_W, PAD_L, PAD_R, GAP);

    ctx.fillStyle = '#6B7A96';
    ctx.font = 'bold 10px Segoe UI, system-ui, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('Total', PAD_L + colEsqW / 2, 18);
    ctx.fillText('Modelo', xMdlL + COL_MDL_W / 2, 18);
    ctx.fillText('Top 3 Tipos de OS  (qtd.)', xMdlR + GAP + colDirW / 2, 18);

    ctx.strokeStyle = '#DDE3EE';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(PAD_L, 24); ctx.lineTo(W - PAD_R, 24);
    ctx.stroke();

    if (!rows.length) {{
      ctx.fillStyle = '#6B7A96';
      ctx.font = '13px Segoe UI, system-ui, sans-serif';
      ctx.fillText('Nenhum dado no período selecionado', W / 2, PAD_T + 20);
      return;
    }}

    rows.forEach((row, i) => {{
      const y    = PAD_T + i * rowH;
      const bH   = rowH - 6;
      const yBar = y + 3;
      const yMid = yBar + bH / 2;

      if (i % 2 === 0) {{
        ctx.fillStyle = 'rgba(0,0,0,0.022)';
        ctx.fillRect(0, y, W, rowH);
      }}

      const totalW    = Math.max(4, (row.total / maxTotal) * (colEsqW - 8));
      const xBarEnd   = PAD_L + colEsqW;
      const xBarStart = xBarEnd - totalW;
      const rL        = Math.min(5, bH / 2);

      ctx.shadowColor = 'rgba(21,101,192,0.18)';
      ctx.shadowBlur  = 4;
      ctx.fillStyle   = '#1565C0';
      ctx.beginPath();
      ctx.moveTo(xBarStart + rL, yBar);
      ctx.lineTo(xBarEnd, yBar);
      ctx.lineTo(xBarEnd, yBar + bH);
      ctx.lineTo(xBarStart + rL, yBar + bH);
      ctx.arcTo(xBarStart, yBar + bH, xBarStart, yBar, rL);
      ctx.arcTo(xBarStart, yBar, xBarStart + rL, yBar, rL);
      ctx.closePath();
      ctx.fill();
      ctx.shadowBlur = 0;

      ctx.font = 'bold 11px Segoe UI, system-ui, sans-serif';
      if (totalW > 22) {{
        ctx.fillStyle = '#fff';
        ctx.textAlign = 'center';
        ctx.fillText(row.total, xBarStart + totalW / 2, yMid + 4);
      }} else {{
        ctx.fillStyle = '#1565C0';
        ctx.textAlign = 'right';
        ctx.fillText(row.total, xBarStart - 4, yMid + 4);
      }}

      ctx.fillStyle = '#011F3F';
      const mdlR = 5;
      if (ctx.roundRect) {{
        ctx.beginPath(); ctx.roundRect(xMdlL, yBar + 1, COL_MDL_W, bH - 2, mdlR); ctx.fill();
      }} else {{
        ctx.fillRect(xMdlL, yBar + 1, COL_MDL_W, bH - 2);
      }}
      ctx.fillStyle = '#FFFFFF';
      ctx.font = 'bold 11px Segoe UI, system-ui, sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText(row.modelo, xMdlL + COL_MDL_W / 2, yMid + 4);

      const tipos = row.tipos.slice(0, 3);
      if (!tipos.length) return;

      // Proporção baseada na soma dos top3 → barra ocupa 100% da coluna direita
      const somaTop3  = tipos.reduce((s, t) => s + t.qtd, 0) || 1;
      const xDirStart = xMdlR + GAP;
      let xCur = xDirStart;

      // Abreviaturas inteligentes para nomes longos de tipo OS
      function abreviarTipo(nome, pixelsDisp) {{
        const ABREVS = [
          [/SUBSTITUIÇÃO DE HIDRÔMETRO\\s*-?\\s*/i, 'Sub.Hidr. '],
          [/VERIFICAÇÃO DE HIDRÔMETRO A PEDIDO DO CLIENTE/i, 'Verif.(cliente)'],
          [/VERIFICAÇÃO DE HIDRÔMETRO/i, 'Verif.Hidr.'],
          [/CORREÇÃO DE HIDRÔMETRO INVERTIDO/i, 'Corr.Invertido'],
          [/INSTALAÇÃO DE HIDRÔMETRO DE TESTE/i, 'Inst.Teste'],
          [/AFERIÇÃO DE HIDRÔMETRO/i, 'Aferição'],
          [/RETIRADA HIDRÔMETRO/i, 'Retirada'],
          [/CADASTRO DE HIDRÔMETRO/i, 'Cadastro'],
          [/ALTERAÇÃO CADASTRAL.*HIDRÔMETRO/i, 'Alt.Cadastral'],
          [/PREVENTIVA/i, 'Prev.'],
          [/CORRETIVA/i, 'Corret.'],
          [/DANIFICADO/i, 'Danif.'],
        ];
        // Tenta abreviar progressivamente até caber
        let resultado = nome;
        for (const [regex, abrev] of ABREVS) {{
          resultado = resultado.replace(regex, abrev).trim();
          if (ctx.measureText(resultado).width <= pixelsDisp) break;
        }}
        return resultado;
      }}

      tipos.forEach((t, ti) => {{
        const segW   = Math.max(2, (t.qtd / somaTop3) * colDirW);
        const isLast = ti === tipos.length - 1;
        const rSeg   = isLast ? Math.min(5, bH / 2) : 0;
        const cor    = corTipo(t.tipo);

        ctx.shadowColor = 'rgba(0,0,0,0.10)';
        ctx.shadowBlur  = 3;
        ctx.fillStyle   = cor;

        if (rSeg > 0) {{
          ctx.beginPath();
          ctx.moveTo(xCur, yBar);
          ctx.lineTo(xCur + segW - rSeg, yBar);
          ctx.arcTo(xCur + segW, yBar, xCur + segW, yBar + rSeg, rSeg);
          ctx.arcTo(xCur + segW, yBar + bH, xCur + segW - rSeg, yBar + bH, rSeg);
          ctx.lineTo(xCur, yBar + bH);
          ctx.closePath();
          ctx.fill();
        }} else {{
          ctx.fillRect(xCur, yBar, segW, bH);
        }}
        ctx.shadowBlur = 0;

        // Linha separadora entre segmentos
        if (ti < tipos.length - 1 && segW > 2) {{
          ctx.strokeStyle = 'rgba(255,255,255,0.55)';
          ctx.lineWidth   = 1.5;
          ctx.beginPath();
          ctx.moveTo(xCur + segW, yBar + 2);
          ctx.lineTo(xCur + segW, yBar + bH - 2);
          ctx.stroke();
        }}

        // Texto dentro da barra
        const PAD_INNER = 5;
        const pixDisp   = segW - PAD_INNER * 2;
        const fontSize  = bH >= 22 ? 10 : 9;
        ctx.font = `bold ${{fontSize}}px Segoe UI, system-ui, sans-serif`;
        if (pixDisp > 18) {{
          ctx.fillStyle  = '#fff';
          ctx.textAlign  = 'left';
          const qtdStr   = ' (' + t.qtd + ')';
          const qtdW     = ctx.measureText(qtdStr).width;
          const nomeDisp = pixDisp - qtdW;

          let nomeExib = t.tipo;
          // Tenta abreviar para caber com a quantidade
          if (ctx.measureText(nomeExib).width > nomeDisp) {{
            nomeExib = abreviarTipo(nomeExib, nomeDisp);
          }}
          if (ctx.measureText(nomeExib).width > nomeDisp) {{
            while (nomeExib.length > 2 && ctx.measureText(nomeExib + '').width > nomeDisp) {{
              nomeExib = nomeExib.slice(0, -1);
            }}
            nomeExib = nomeExib.trimEnd() + '';
          }}
          // Se não couber nem abreviação + qtd, tenta só abreviação
          if (nomeDisp < 10) {{
            const soloDisp = pixDisp;
            nomeExib = abreviarTipo(t.tipo, soloDisp);
            if (ctx.measureText(nomeExib).width > soloDisp) {{
              while (nomeExib.length > 2 && ctx.measureText(nomeExib + '').width > soloDisp) {{
                nomeExib = nomeExib.slice(0, -1);
              }}
              nomeExib = nomeExib.trimEnd() + '';
            }}
            ctx.fillText(nomeExib, xCur + PAD_INNER, yMid + 4);
          }} else {{
            ctx.fillText(nomeExib + qtdStr, xCur + PAD_INNER, yMid + 4);
          }}
        }} else if (pixDisp > 8) {{
          // Barra pequena: mostra só a quantidade centralizada
          ctx.fillStyle = '#fff';
          ctx.textAlign = 'center';
          ctx.font = `bold 9px Segoe UI, system-ui, sans-serif`;
          ctx.fillText(t.qtd, xCur + segW / 2, yMid + 4);
        }}
        xCur += segW;
      }});

      // top3 agora ocupa 100% da coluna direita  sem resto cinza
    }});

    ctx.strokeStyle = '#DDE3EE';
    ctx.lineWidth = 1.5;
    [xMdlL - GAP/2, xMdlR + GAP/2].forEach(x => {{
      ctx.beginPath();
      ctx.moveTo(x, 26);
      ctx.lineTo(x, PAD_T + rows.length * rowH + 4);
      ctx.stroke();
    }});

    const legY = PAD_T + rows.length * rowH + 14;
    let legX   = PAD_L;
    ctx.font = '9px Segoe UI, system-ui, sans-serif';
    tiposOrdenados.slice(0, 8).forEach(tipo => {{
      const cor = corTipo(tipo);
      ctx.fillStyle = cor;
      ctx.fillRect(legX, legY, 9, 9);
      ctx.fillStyle = '#475569';
      ctx.textAlign = 'left';
      const lbl = tipo.length > 24 ? tipo.slice(0, 22) + '...' : tipo;
      ctx.fillText(lbl, legX + 12, legY + 8);
      legX += ctx.measureText(lbl).width + 24;
      if (legX > W - 160) {{ legX = PAD_L; }}
    }});
  }}

  let curFrom = 0, curTo = N - 1;
  updateSliderUI(curFrom, curTo);
  drawCanvas(curFrom, curTo);

  function onSlide() {{
    let f = parseInt(inpFrom.value), t = parseInt(inpTo.value);
    if (f >= t) {{
      if (this === inpFrom) {{ f = Math.max(0, t - 1); inpFrom.value = f; }}
      else                  {{ t = Math.min(N - 1, f + 1); inpTo.value = t; }}
    }}
    curFrom = f; curTo = t;
    updateSliderUI(f, t);
    drawCanvas(f, t);
  }}

  inpFrom.removeEventListener('input', onSlide);
  inpTo.removeEventListener('input', onSlide);
  inpFrom.addEventListener('input', onSlide);
  inpTo.addEventListener('input', onSlide);
}}
function buildChartLeitOS() {{
  const _chart_leit_inst = Chart.getChart('chart-leit-os');
  if (_chart_leit_inst) _chart_leit_inst.destroy();

  // ── 2. Leiturista × Total de OSs ─────────────────────────────────────────
  (function() {{
    const ctx2 = document.getElementById('chart-leit-os');
    if (!ctx2 || !LEIT_OS_L.length) return;

    const total = LEIT_OS_Q.reduce((a, b) => a + b, 0) || 1;
    const CORES = ['#1565C0','#2E7D32','#C62828','#FF8F00','#6A1B9A','#00838F','#AD1457','#37474F'];

    makeChart(ctx2, {{
      type: 'bar',
      data: {{
        labels: LEIT_OS_L,
        datasets: [{{
          label: 'Total de OSs',
          data: LEIT_OS_Q,
          backgroundColor: LEIT_OS_L.map((_, i) => CORES[i % CORES.length]),
          borderRadius: 6,
          borderSkipped: false,
        }}]
      }},
      options: {{
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        animation: {{ duration: 1200, easing: 'easeOutBack' }},
        layout: {{ padding: {{ right: 80 }} }},
        plugins: {{
          legend: {{ display: false }},
          datalabels: {{
            display: true,
            anchor: 'end',
            align: 'right',
            offset: 6,
            clamp: false,
            clip: false,
            font: {{ size: 11, weight: '700' }},
            color: '#1A2535',
            formatter: (v) => {{
              if (!v) return '';
              const pct = (v / total * 100).toFixed(1);
              return v.toLocaleString('pt-BR') + ' (' + pct + '%)';
            }}
          }},
          tooltip: {{ ...TT, callbacks: {{ label: c =>
            ` ${{c.parsed.x.toLocaleString('pt-BR')}} OSs (${{(c.parsed.x/total*100).toFixed(1)}}%)`
          }} }}
        }},
        scales: {{
          x: {{
            ticks: {{ callback: v => v.toLocaleString('pt-BR') }},
            grid: {{ color: 'rgba(0,0,0,0.05)' }}
          }},
          y: {{
            grid: {{ display: false }},
            ticks: {{ font: {{ size: 11, weight: '600' }} }}
          }}
        }}
      }}
    }});
  }})();
}}



// ── GRÁFICO ESPELHO: Hidrômetro × Tipo de OS ─────────────────────────────────

// ── GRÁFICO: Leiturista × Total de OSs ───────────────────────────────────────


// ── CHARTS MAPEAMENTO ─────────────────────────────────────────────────────


// ── GRÁFICOS DO PAINEL MAPEAMENTO ────────────────────────────────────────────

// ── MINI-INSIGHTS DA ABA 50012 ───────────────────────────────────────────────
function fillMiniInsights50012() {{
  const map50 = {{
    'mini-chart-50-criticas':    M50_CRITICAS,
    'mini-chart-50-faixas':      M50_FAIXAS,
    'mini-chart-50-leituristas': M50_LEITURISTAS,
    'mini-chart-50-grupos':      M50_GRUPOS,
    'mini-chart-50-desvio':      M50_DESVIO,
    'mini-chart-50-bairros-fat': M50_BAIRROS_FAT,
    'mini-chart-50-categorias':  M50_CATEGORIAS,
  }};
  Object.entries(map50).forEach(([id, html]) => {{
    if (!html) return;
    const el = document.getElementById(id);
    if (el) el.innerHTML = html;
  }});
}}

function _initDashboard() {{
  try {{
    console.log('%c[Dashboard] Iniciando...', 'color:green;font-weight:bold');
    console.log('%cDEFAULT_TAB:', 'color:blue', DEFAULT_TAB);
    switchTab(DEFAULT_TAB);
  }} catch(e) {{
    console.error('[Dashboard] Erro em switchTab:', e);
    // Fallback: força o primeiro painel visível mesmo sem charts
    const panels = document.querySelectorAll('.tab-panel');
    if (panels.length) panels[0].classList.add('active');
    const btns = document.querySelectorAll('.tab-btn');
    if (btns.length) btns[0].classList.add('active');
  }}
  // chart-50-dias é inicializado via switchTab('50012') para garantir que o canvas está visível

  const btnDiario = document.getElementById('btn-dias-diario');
  const btnMensal = document.getElementById('btn-dias-mensal');
  if (btnDiario && btnMensal) {{
    btnDiario.addEventListener('click', () => {{
      if (modoLeiturasDias === 'daily') return;
      modoLeiturasDias = 'daily';
      btnDiario.style.background = '#1565C0'; btnDiario.style.color = 'white';
      btnMensal.style.background = 'transparent'; btnMensal.style.color = '#6B7A96';
      rebuildChartDias('daily');
    }});
    btnMensal.addEventListener('click', () => {{
      if (modoLeiturasDias === 'monthly') return;
      modoLeiturasDias = 'monthly';
      btnDiario.style.background = 'transparent'; btnDiario.style.color = '#6B7A96';
      btnMensal.style.background = '#1565C0'; btnMensal.style.color = 'white';
      rebuildChartDias('monthly');
    }});
  }}
  const obs = new IntersectionObserver((entries) => {{
    entries.forEach(entry => {{
      if (entry.isIntersecting) {{
        const el = entry.target;
        const delay = parseInt(el.dataset.delay || 0);
        setTimeout(() => el.classList.add('visible'), delay);
        obs.unobserve(el);
      }}
    }});
  }}, {{ threshold: 0.10 }});
  document.querySelectorAll('.kpi-card, .chart-card').forEach(el => obs.observe(el));
}}

if (document.readyState === 'loading') {{
  document.addEventListener('DOMContentLoaded', _initDashboard);
}} else {{
  _initDashboard();
}}
</script>
<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/datatables/1.10.21/js/jquery.dataTables.min.js"></script>
<script>
  // Inicializar DataTables para tabela de OSs quando disponível
  document.addEventListener('DOMContentLoaded', function() {{
    if ($.fn.dataTable) {{
      const _dtOpts = {{
        paging: true,
        pageLength: 25,
        searching: true,
        ordering: true,
        order: [[0, 'asc']],
        lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'Todos']],
        language: {{
          sEmptyTable:     "Nenhum dado disponível",
          sInfo:           "Mostrando _START_ a _END_ de _TOTAL_ registros",
          sInfoEmpty:      "Mostrando 0 a 0 de 0 registros",
          sInfoFiltered:   "(filtrado de _MAX_ registros)",
          sLengthMenu:     "Mostrar _MENU_ registros",
          sLoadingRecords: "Carregando...",
          sProcessing:     "Processando...",
          sSearch:         "Buscar:",
          sZeroRecords:    "Nenhum registro encontrado",
          oPaginate: {{
            sFirst:    "Primeiro",
            sLast:     "Último",
            sNext:     "Próximo",
            sPrevious: "Anterior"
          }}
        }},
        columnDefs: [{{ targets: '_all', searchable: true, orderable: true }}],
        dom: '<"dataTables_header"lf>t<"dataTables_footer"ip>'
      }};
      if (document.querySelector('table#tabela-os-cruzadas'))
        $('#tabela-os-cruzadas').DataTable(_dtOpts);
      if (document.querySelector('table#tabela-os-sem-coord-com'))
        $('#tabela-os-sem-coord-com').DataTable(Object.assign({{}}, _dtOpts, {{ order: [[0, 'asc']] }}));
      if (document.querySelector('table#tabela-os-sem-coord-outras'))
        $('#tabela-os-sem-coord-outras').DataTable(Object.assign({{}}, _dtOpts, {{ order: [[0, 'asc']] }}));
    }}
  }});
</script>
</body>
</html>"""

    write_text_utf8(caminho_saida, html)

    return caminho_saida


def gerar_html_mapeamento_automatico(usuario: str, senha: str, caminho_saida: str, 
                                      periodo: str = None, cidade: str = None, 
                                      log_fn=None) -> str:
    """
    Gera dashboard HTML com mapeamento de OSs automaticamente integrado.
    
    Fluxo completo:
      1. Login no Waterfy
      2. Baixa payload das Ordens de Serviço (OSs ativas)
      3. Cruza com macro 50012 (lat/long por matrícula)
      4. Organiza por funcionário
      5. Gera HTML com todas as abas + mapa interativo
    
    Args:
        usuario        : usuário Waterfy
        senha          : senha Waterfy
        caminho_saida  : onde salvar o HTML (ex: 'Dashboard_Waterfy/dashboard.html')
        periodo        : período (default: data atual, ex: '03-2026')
        cidade         : cidade (default: 'ITAPOÁ')
        log_fn         : função de callback para logs
    
    Returns:
        string com o path do arquivo gerado
    
    Raises:
        Exception se waterfy_engine não estiver disponível ou login falhar
    """
    if not HAS_WATERFY:
        raise Exception("waterfy_engine não disponível. Certifique-se de estar na mesma pasta.")
    
    from datetime import datetime as dt
    
    if log_fn:
        log_fn(" [MAPEAMENTO AUTOMÁTICO] Iniciando fluxo de integração completo...")
    
    # Define período e cidade padrão
    if not periodo:
        periodo = dt.now().strftime('%m-%Y')
    if not cidade:
        cidade = 'ITAPOÁ'
    
    # Passo 1: Integra mapeamento (login → download OSs → 50012 → cruzar)
    if log_fn:
        log_fn("▶ PASSO 1: Iniciando autenticação e download de dados...")
    
    combo_resultado = integrar_combo_mapa(usuario, senha, log_fn)
    
    if combo_resultado.get('status') == 'error':
        raise Exception(f"Falha ao integrar mapa: {combo_resultado.get('msg')}")
    
    if log_fn:
        log_fn(f" Mapeamento integrado com sucesso")
    
    # Passo 2: Prepara dados para gerar_html (formato esperado)
    dados_dashboard = {
        'periodo': periodo,
        'cidade': cidade,
        'combo_mapa': combo_resultado,  # ← Integra resultado do mapeamento automático
    }
    
    if log_fn:
        log_fn("▶ PASSO 2: Gerando HTML do dashboard...")
    
    # Passo 3: Gera HTML
    resultado_caminho = gerar_html(dados_dashboard, caminho_saida)
    
    if log_fn:
        log_fn(f" Dashboard gerado com sucesso: {resultado_caminho}")
    
    return resultado_caminho
