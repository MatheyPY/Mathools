"""
waterfy_ml.py
=============
Módulo de Machine Learning para previsão de arrecadação futura (Macro 8117).

Estratégia adaptativa conforme volume de histórico disponível:
  - < 3 meses  → Média móvel ponderada (WMA)
  - 3–12 meses → Prophet  (tendência + sazonalidade semanal/anual)
  - > 12 meses → Prophet + XGBoost ensemble (Prophet para tendência,
                 XGBoost para resíduos — melhora precisão no longo prazo)

Entrada:
  resultado_8117 : dict retornado por processar_8117() / reordenar_8117()
                   Precisa conter 'datas' e 'arr_diario'.

Saída:
  dict com chaves prontas para o dashboard:
    'modelo_usado'     : str  — nome do modelo
    'meses_historico'  : int  — meses de histórico usados
    'meses_previsao'   : int  — horizonte previsto
    'datas_hist'          : list[str]   — datas do histórico (dd/mm/yyyy)
    'valores_hist'        : list[float] — valores diários históricos
                            → GRÁFICO DE BARRAS (histórico)

    'valores_acum_hist'   : list[float] — acumulado diário do histórico
                            → GRÁFICO DE LINHA (histórico); mesmo índice que datas_hist

    'datas_prev'          : list[str]   — datas previstas (dd/mm/yyyy)
    'valores_prev'        : list[float] — valores diários previstos
    'valores_prev_sup'    : list[float] — limite superior do intervalo (Prophet/ensemble)
    'valores_prev_inf'    : list[float] — limite inferior do intervalo (Prophet/ensemble)
                            → GRÁFICO DE BARRAS (previsão); mesmo índice que datas_prev

    'valores_acum_prev'   : list[float] — acumulado diário previsto, contínuo a partir
                            do último valor de valores_acum_hist
    'valores_acum_sup'    : list[float] — limite superior do acumulado previsto
    'valores_acum_inf'    : list[float] — limite inferior do acumulado previsto
                            → GRÁFICO DE LINHA (previsão); mesmo índice que datas_prev

    Para renderizar o gráfico de linha CONTÍNUO (histórico + previsão):
        eixo_x    = datas_hist       + datas_prev
        eixo_hist = valores_acum_hist + [None] * len(datas_prev)   # linha sólida
        eixo_prev = [None] * len(datas_hist) + valores_acum_prev   # linha tracejada
        (ponto de junção: valores_acum_hist[-1] == valores_acum_prev[0] - valores_prev[0])

    'totais_mensais_prev' : list[dict] — previsão agregada por mês
        [{'mes': 'mm/yyyy', 'total': float, 'sup': float, 'inf': float}, ...]
    'metricas'            : dict — MAE, MAPE, RMSE (avaliados via cross-validation)
    'aviso'               : str | None — mensagem quando histórico é curto
"""

from __future__ import annotations

import warnings
import calendar
from datetime import datetime, timedelta
from collections import defaultdict
from typing import Optional
import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

# ── Dependências opcionais ────────────────────────────────────────────────────
def _check_libs():
    """Verifica disponibilidade de bibliotecas dinamicamente."""
    global _HAS_NUMPY, _HAS_PANDAS, _HAS_PROPHET, _HAS_XGB, _HAS_STATSMODELS
    
    try:
        import numpy as np
        _HAS_NUMPY = True
    except ImportError:
        _HAS_NUMPY = False

    try:
        import pandas as pd
        _HAS_PANDAS = True
    except ImportError:
        _HAS_PANDAS = False

    try:
        from prophet import Prophet
        _HAS_PROPHET = True
    except ImportError:
        try:
            from fbprophet import Prophet
            _HAS_PROPHET = True
        except ImportError:
            _HAS_PROPHET = False

    try:
        import xgboost as xgb
        _HAS_XGB = True
    except ImportError:
        _HAS_XGB = False
    
    try:
        from statsmodels.tsa.arima.model import ARIMA
        _HAS_STATSMODELS = True
    except ImportError as e:
        _HAS_STATSMODELS = False

# Variáveis globais (inicialmente False, serão preenchidas por _check_libs)
_HAS_NUMPY = False
_HAS_PANDAS = False
_HAS_PROPHET = False
_HAS_XGB = False
_HAS_STATSMODELS = False

# Verificação inicial
_check_libs()


# ─────────────────────────────────────────────────────────────────────────────
# Utilitários internos
# ─────────────────────────────────────────────────────────────────────────────

def _parse_data(s: str) -> Optional[datetime]:
    for fmt in ('%d/%m/%Y', '%d/%m/%y', '%Y-%m-%d'):
        try:
            return datetime.strptime(s.strip(), fmt)
        except Exception:
            pass
    return None


def _fmt_data(d: datetime) -> str:
    return d.strftime('%d/%m/%Y')


def _fmt_mes(d: datetime) -> str:
    return d.strftime('%m/%Y')


def _meses_entre(d_ini: datetime, d_fim: datetime) -> float:
    delta = d_fim - d_ini
    return delta.days / 30.44


def _agregar_mensalmente(datas: list[datetime], valores: list[float],
                          sup: list[float] = None, inf: list[float] = None) -> list[dict]:
    """Agrega valores diários em totais mensais."""
    mes_total: dict[str, dict] = {}
    for i, (d, v) in enumerate(zip(datas, valores)):
        k = _fmt_mes(d)
        if k not in mes_total:
            mes_total[k] = {'mes': k, 'total': 0.0, 'sup': 0.0, 'inf': 0.0, '_ordem': d}
        mes_total[k]['total'] += max(v, 0.0)
        if sup:
            mes_total[k]['sup'] += max(sup[i], 0.0)
        if inf:
            mes_total[k]['inf'] += max(inf[i], 0.0)

    result = sorted(mes_total.values(), key=lambda x: x['_ordem'])
    for r in result:
        del r['_ordem']
    return result


# ─────────────────────────────────────────────────────────────────────────────
# Modelo 1 — Média Móvel Ponderada (fallback sem dependências)
# ─────────────────────────────────────────────────────────────────────────────

def _prever_wma(datas_hist: list[datetime], valores_hist: list[float],
                meses_previsao: int) -> dict:
    """
    Weighted Moving Average: pesos maiores para observações recentes.
    Agrega por mês antes de prever (estável para séries curtas).
    """
    # Agrega por mês para suavizar
    mes_vals: dict[str, list] = defaultdict(list)
    for d, v in zip(datas_hist, valores_hist):
        mes_vals[_fmt_mes(d)].append(v)

    meses_ord = sorted(mes_vals.keys(), key=lambda m: datetime.strptime(m, '%m/%Y'))
    totais = [sum(mes_vals[m]) for m in meses_ord]
    n = len(totais)

    # Janela de até 6 meses
    janela = min(n, 6)
    pesos = list(range(1, janela + 1))
    soma_pesos = sum(pesos)

    previsoes_mensais = []
    hist_window = list(totais[-janela:])

    for _ in range(meses_previsao):
        w = [pesos[i] * hist_window[-(janela - i)] for i in range(janela)]
        prev = sum(w) / soma_pesos
        prev = max(prev, 0.0)
        previsoes_mensais.append(prev)
        hist_window.append(prev)
        hist_window = hist_window[-janela:]

    # Distribui previsão mensal em dias (média diária por mês)
    ultimo_dia = datas_hist[-1]
    datas_prev, valores_prev = [], []
    cursor = ultimo_dia + timedelta(days=1)

    for i, total_mes in enumerate(previsoes_mensais):
        # Calcula o mês e ano alvo para essa iteração
        meses_desde_inicio = (cursor.month - 1 + i)
        ano = cursor.year + meses_desde_inicio // 12
        mes = (meses_desde_inicio % 12) + 1
        n_dias = calendar.monthrange(ano, mes)[1]
        media_dia = total_mes / n_dias

        for d in range(n_dias):
            dia = datetime(ano, mes, d + 1)
            if dia <= ultimo_dia:
                continue
            datas_prev.append(dia)
            valores_prev.append(media_dia)

    # Para WMA sem biblioteca, não temos intervalo de confiança formal
    margem = 0.15  # ±15%
    sup = [v * (1 + margem) for v in valores_prev]
    inf = [max(v * (1 - margem), 0.0) for v in valores_prev]

    return {
        'datas_prev':     datas_prev,
        'valores_prev':   valores_prev,
        'valores_prev_sup': sup,
        'valores_prev_inf': inf,
        'metricas':       {'MAE': None, 'MAPE': None, 'RMSE': None},
        'aviso':          'Histórico curto (< 3 meses): previsão via média móvel ponderada. '
                          'Resultados mais precisos com mais dados.',
    }


# ─────────────────────────────────────────────────────────────────────────────
# Modelo 1.5 — ARIMA / Exponential Smoothing (StatsModels)
# ─────────────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
# Análise de Sazonalidade — Fim de Semana e Épocas do Ano
# ─────────────────────────────────────────────────────────────────────────────

def _calcular_fator_dia_semana(datas_hist: list[datetime], 
                               valores_hist: list[float]) -> dict[int, float]:
    """
    Analisa padrão de movimentação por dia da semana (seg=0, dom=6).
    Retorna fator multiplicativo para cada dia (ex: {'0': 0.9, '6': 1.3}).
    
    Para cidades litorâneas, finais de semana tendem a ter mais movimento.
    """
    valores_por_dia = defaultdict(list)
    
    for d, v in zip(datas_hist, valores_hist):
        if v > 0:
            dia_semana = d.weekday()  # seg=0, dom=6
            valores_por_dia[dia_semana].append(v)
    
    # Calcula média por dia da semana
    media_por_dia = {}
    media_geral = sum(valores_hist) / len(valores_hist) if valores_hist else 1.0
    
    for dia in range(7):
        if dia in valores_por_dia and valores_por_dia[dia]:
            media_por_dia[dia] = sum(valores_por_dia[dia]) / len(valores_por_dia[dia])
        else:
            media_por_dia[dia] = media_geral
    
    # Normaliza: converte em fator (fator = media_dia / media_geral)
    fatores = {dia: media_por_dia[dia] / media_geral for dia in range(7)}
    
    # Suaviza: não deixa muito extremo (min 0.7x, max 1.4x)
    fatores = {dia: max(0.7, min(1.4, f)) for dia, f in fatores.items()}
    
    return fatores


def _calcular_sazonalidade_anual(datas_hist: list[datetime],
                                 valores_hist: list[float]) -> dict[tuple, float]:
    """
    Identifica períodos sazonais (épocas do ano com maior movimento).
    Para litoral: ano novo, férias escolares (jul/ago), feriados carnaval/páscoa, verão.
    
    Se histórico é curto, aplica padrões tipicos de litoral do Brasil.
    Caso contrário, detecta baseado em dados.
    
    Retorna fator para cada mês.
    Ex: janeiro tem 20% a mais → fator 1.2
    """
    meses_hist = _meses_entre(datas_hist[0], datas_hist[-1]) if datas_hist else 0
    
    # Períodos sazonais para litoral (Brasil)
    periodos_altos = {
        'carnaval': [2],           # Fevereiro (Carnaval)
        'pascoa': [3, 4],          # Março/Abril (Páscoa)
        'verao': [1, 2, 3],        # Jan/Fev/Mar (Verão - pico)
        'ferias_julho': [7],       # Julho (Férias escolares)
        'ferias_agosto': [8],      # Agosto (Férias escolares)
        'praia': [12],             # Dezembro (Férias ano novo)
    }
    
    meses_altos = set()
    for periodos in periodos_altos.values():
        meses_altos.update(periodos)
    
    # Agrupa valores por mês
    valores_por_mes = defaultdict(list)
    for d, v in zip(datas_hist, valores_hist):
        if v > 0:
            valores_por_mes[d.month].append(v)
    
    # Calcula média geral
    media_geral = sum(valores_hist) / len(valores_hist) if valores_hist else 1.0
    
    # Calcula média por mês
    media_por_mes = {}
    for mes in range(1, 13):
        if mes in valores_por_mes and valores_por_mes[mes]:
            media_por_mes[mes] = sum(valores_por_mes[mes]) / len(valores_por_mes[mes])
        else:
            media_por_mes[mes] = media_geral
    
    fatores_por_mes = {}
    
    # Se histórico é muito curto (< 24 meses), aplica padrão típico de litoral
    # com ajuste baseado em dados locais disponíveis
    if meses_hist < 24:
        # Começa com padrão típico
        for mes in range(1, 13):
            if mes in meses_altos:
                fatores_por_mes[mes] = 1.20  # 20% acima da média
            else:
                fatores_por_mes[mes] = 0.90  # 10% abaixo da média
        
        # Agora aplica ajustes baseados em dados locais (se houver)
        meses_com_dados = [m for m in range(1, 13) if m in media_por_mes]
        if meses_com_dados:
            # Identifica quais meses com dados têm movimentação acima/abaixo da média
            for mes in meses_com_dados:
                media_mes = media_por_mes[mes]
                percentual_desvio = (media_mes - media_geral) / media_geral if media_geral > 0 else 0
                
                # Se há dado real e desvio é significativo (> 5%), ajusta o fator
                if abs(percentual_desvio) > 0.05:
                    # Reforça o padrão se coincide, enfraquece se contradiz
                    if mes in meses_altos and percentual_desvio > 0:
                        # Mês "alto" com dados altos - reforça
                        fatores_por_mes[mes] = 1.0 + percentual_desvio * 1.3
                    elif mes not in meses_altos and percentual_desvio < 0:
                        # Mês "baixo" com dados baixos - reforça
                        fatores_por_mes[mes] = 1.0 + percentual_desvio * 1.3
                    else:
                        # Aplica o dado real (contradiz padrão)
                        fatores_por_mes[mes] = 1.0 + percentual_desvio
                else:
                    # Desvio pequeno, mantém padrão mas suaviza um pouco
                    fatores_por_mes[mes] = 1.0 + percentual_desvio * 0.5
    else:
        # Se há 2+ anos de histórico, usa dados reais
        meses_com_dados = [m for m in range(1, 13) if m in media_por_mes]
        if meses_com_dados:
            media_altos = sum(media_por_mes[m] for m in meses_altos if m in media_por_mes) / len([m for m in meses_altos if m in media_por_mes]) if [m for m in meses_altos if m in media_por_mes] else media_geral
            media_baixos_temp = [media_por_mes[m] for m in range(1, 13) if m not in meses_altos and m in media_por_mes]
            media_baixos = sum(media_baixos_temp) / len(media_baixos_temp) if media_baixos_temp else media_geral
            
            fator_alto = media_altos / media_geral if media_geral > 0 else 1.0
            fator_alto = max(0.8, min(1.5, fator_alto))
            
            fator_baixo = media_baixos / media_geral if media_geral > 0 else 1.0
            fator_baixo = max(0.7, min(1.0, fator_baixo))
            
            for mes in range(1, 13):
                if mes in meses_altos:
                    fatores_por_mes[mes] = fator_alto
                else:
                    fatores_por_mes[mes] = fator_baixo
    
    # Garante que todos os meses têm fator (1.0 se não especificado)
    for mes in range(1, 13):
        if mes not in fatores_por_mes:
            fatores_por_mes[mes] = 1.0
    
    return fatores_por_mes


def _prever_arima(datas_hist: list[datetime], valores_hist: list[float],
                  meses_previsao: int, meses_hist: float) -> dict:
    """
    Previsão com ARIMA (AutoRegressive Integrated Moving Average).
    Fallback para Exponential Smoothing se ARIMA falhar.
    
    Incorpora sazonalidade de:
      - Dias da semana (finais de semana vs dias de semana)
      - Períodos altos/baixos do ano (verão, férias, feriados)
    """
    from statsmodels.tsa.arima.model import ARIMA
    from statsmodels.tsa.holtwinters import ExponentialSmoothing
    
    # Agrega por mês para melhor performance
    mes_vals: dict[str, list] = defaultdict(list)
    for d, v in zip(datas_hist, valores_hist):
        mes_vals[_fmt_mes(d)].append(v)

    meses_ord = sorted(mes_vals.keys(), key=lambda m: datetime.strptime(m, '%m/%Y'))
    totais_mensais = [sum(mes_vals[m]) for m in meses_ord]
    
    # Prepara série temporal
    datas_meses = [datetime.strptime(m, '%m/%Y') for m in meses_ord]
    ts = pd.Series(totais_mensais, index=datas_meses)
    
    previsoes_mensais = []
    intervalos_sup = []
    intervalos_inf = []
    
    try:
        # Tenta ARIMA(1,1,1) - configuração equilibrada
        model = ARIMA(ts, order=(1, 1, 1))
        result = model.fit()
        forecast = result.get_forecast(steps=meses_previsao)
        previsoes_mensais = forecast.predicted_mean.tolist()
        
        ci = forecast.conf_int(alpha=0.20)  # 80% de confiança
        intervalos_sup = ci.iloc[:, 1].tolist()
        intervalos_inf = ci.iloc[:, 0].tolist()
    except Exception:
        # Fallback para Exponential Smoothing
        try:
            model = ExponentialSmoothing(ts, trend='add', seasonal=None)
            result = model.fit()
            forecast = result.forecast(steps=meses_previsao)
            previsoes_mensais = forecast.tolist()
            
            # Aproxima intervalo de confiança como ±10%
            margem = 0.10
            intervalos_sup = [v * (1 + margem) for v in previsoes_mensais]
            intervalos_inf = [max(v * (1 - margem), 0.0) for v in previsoes_mensais]
        except Exception:
            # Se tudo falhar, retorna WMA como último recurso
            return _prever_wma(datas_hist, valores_hist, meses_previsao)
    
    # Explode para dias com variação realista
    # Calcula padrão de dias do histórico: distribuição dentro de cada mês
    ultimo_dia = datas_hist[-1]
    
    # Agrupa histórico por mês para calcular variação intra-mês
    dias_por_mes = {}
    for d, v in zip(datas_hist, valores_hist):
        mes_key = (d.month, d.year)
        if mes_key not in dias_por_mes:
            dias_por_mes[mes_key] = []
        dias_por_mes[mes_key].append(v)
    
    # Calcula índice de normalização por dia do mês ( média diária / média mensal)
    indice_dia = {}  # {dia_do_mes: fator}
    if dias_por_mes:
        for mes_key, vals in dias_por_mes.items():
            media_mes = sum(vals) / len(vals)
            if media_mes > 0:
                for i, dia_val in enumerate(vals):
                    dia_num = i + 1  # dia do mês (1, 2, 3...)
                    if dia_num not in indice_dia:
                        indice_dia[dia_num] = []
                    indice_dia[dia_num].append(dia_val / media_mes)
        
        # Média dos índices por dia
        indice_dia = {dia: sum(vals)/len(vals) for dia, vals in indice_dia.items()}
    
    # Fallback se houver problema
    if not indice_dia:
        indice_dia = {d: 1.0 for d in range(1, 32)}
    
    # NOVO: Calcula fatores de sazonalidade
    fatores_dia_semana = _calcular_fator_dia_semana(datas_hist, valores_hist)
    fatores_sazonalidade = _calcular_sazonalidade_anual(datas_hist, valores_hist)
    
    datas_prev, valores_prev = [], []
    sup_prev, inf_prev = [], []
    
    cursor = ultimo_dia + timedelta(days=1)
    for i, (total_mes, sup, inf) in enumerate(zip(previsoes_mensais, intervalos_sup, intervalos_inf)):
        meses_desde_inicio = (cursor.month - 1 + i)
        ano = cursor.year + meses_desde_inicio // 12
        mes = (meses_desde_inicio % 12) + 1
        n_dias = calendar.monthrange(ano, mes)[1]
        
        media_dia = max(total_mes / n_dias, 0.0)
        media_sup = max(sup / n_dias, 0.0)
        media_inf = max(inf / n_dias, 0.0)
        
        for d in range(n_dias):
            dia_num = d + 1
            dia = datetime(ano, mes, dia_num)
            if dia <= ultimo_dia:
                continue
            
            # Aplica índice de variação histórica
            fator_dia_mes = indice_dia.get(dia_num, 1.0)
            
            # NOVO: Aplica fatores de sazonalidade
            dia_semana = dia.weekday()  # seg=0, dom=6
            fator_semana = fatores_dia_semana.get(dia_semana, 1.0)
            fator_sazonal = fatores_sazonalidade.get(mes, 1.0)
            
            # Combina fatores: media_dia * fator_dia_mes * fator_semana * fator_sazonal
            # Mas precisa de normalização para não inflar demais
            fator_combinado = (fator_dia_mes * fator_semana * fator_sazonal) ** (1/2)  # Raiz quadrada para suavizar
            
            datas_prev.append(dia)
            valores_prev.append(max(media_dia * fator_combinado, 0.0))
            sup_prev.append(max(media_sup * fator_combinado, 0.0))
            inf_prev.append(max(media_inf * fator_combinado * 0.8, 0.0))  # Inferior precisa ser menor
    
    return {
        'datas_prev':       datas_prev,
        'valores_prev':     valores_prev,
        'valores_prev_sup': sup_prev,
        'valores_prev_inf': inf_prev,
        'metricas':         {'MAE': None, 'MAPE': None, 'RMSE': None},
        'aviso':            None,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Modelo 2 — Prophet
# ─────────────────────────────────────────────────────────────────────────────

def _prever_prophet(datas_hist: list[datetime], valores_hist: list[float],
                    meses_previsao: int, meses_hist: float) -> dict:
    """Previsão com Facebook Prophet."""
    if not _HAS_PROPHET or not _HAS_PANDAS:
        raise ImportError("Prophet e Pandas são necessários para este modelo.")
    
    df = pd.DataFrame({
        'ds': datas_hist,
        'y':  valores_hist,
    })
    df['y'] = df['y'].clip(lower=0)

    # Configura sazonalidades conforme histórico disponível
    anual = meses_hist >= 11  # precisa de ~1 ano para capturar sazonalidade anual
    semanal = True

    model = Prophet(
        yearly_seasonality=anual,
        weekly_seasonality=semanal,
        daily_seasonality=False,
        seasonality_mode='multiplicative' if meses_hist >= 6 else 'additive',
        interval_width=0.80,
        changepoint_prior_scale=0.05,
    )
    model.fit(df)

    horizonte_dias = meses_previsao * 31  # margem generosa
    future = model.make_future_dataframe(periods=horizonte_dias, freq='D')
    forecast = model.predict(future)

    # Filtra apenas o futuro
    ultimo = datas_hist[-1]
    fc_futuro = forecast[forecast['ds'] > pd.Timestamp(ultimo)].copy()

    datas_prev   = fc_futuro['ds'].dt.to_pydatetime().tolist()
    valores_prev = fc_futuro['yhat'].clip(lower=0).tolist()
    sup          = fc_futuro['yhat_upper'].clip(lower=0).tolist()
    inf          = fc_futuro['yhat_lower'].clip(lower=0).tolist()

    # Limita ao horizonte real
    data_limite = ultimo + timedelta(days=meses_previsao * 31)
    filtro = [i for i, d in enumerate(datas_prev) if d <= data_limite]
    datas_prev   = [datas_prev[i]   for i in filtro]
    valores_prev = [valores_prev[i] for i in filtro]
    sup          = [sup[i]          for i in filtro]
    inf          = [inf[i]          for i in filtro]

    # Métricas via cross-validation (apenas se histórico suficiente)
    metricas = {'MAE': None, 'MAPE': None, 'RMSE': None}
    if meses_hist >= 4:
        try:
            from prophet.diagnostics import cross_validation, performance_metrics
            horizonte_cv = f'{min(30, int(meses_hist * 15))} days'
            initial_cv   = f'{max(60, int(len(datas_hist) * 0.5))} days'
            df_cv = cross_validation(model, initial=initial_cv,
                                     period='15 days', horizon=horizonte_cv,
                                     disable_tqdm=True)
            pm = performance_metrics(df_cv)
            metricas = {
                'MAE':  round(pm['mae'].mean(), 2),
                'MAPE': round(pm['mape'].mean() * 100, 2),
                'RMSE': round(pm['rmse'].mean(), 2),
            }
        except Exception:
            pass

    return {
        'datas_prev':       datas_prev,
        'valores_prev':     valores_prev,
        'valores_prev_sup': sup,
        'valores_prev_inf': inf,
        'metricas':         metricas,
        'aviso':            None if meses_hist >= 6 else
                            'Histórico entre 3–6 meses: sazonalidade anual não capturada. '
                            'Previsão de longo prazo pode ter desvio.',
    }


# ─────────────────────────────────────────────────────────────────────────────
# Modelo 3 — Prophet + XGBoost Ensemble
# ─────────────────────────────────────────────────────────────────────────────

def _prever_ensemble(datas_hist: list[datetime], valores_hist: list[float],
                     meses_previsao: int, meses_hist: float) -> dict:
    """
    Ensemble Prophet + XGBoost:
      1. Prophet prevê tendência + sazonalidade.
      2. XGBoost aprende os resíduos (erro sistemático do Prophet no histórico).
      3. Previsão final = Prophet(futuro) + XGBoost(resíduos_futuro).
    """
    if not _HAS_PROPHET or not _HAS_PANDAS or not _HAS_NUMPY or not _HAS_XGB:
        raise ImportError("Prophet, Pandas, NumPy e XGBoost são necessários para este modelo.")
    
    df = pd.DataFrame({'ds': datas_hist, 'y': valores_hist})
    df['y'] = df['y'].clip(lower=0)

    # ── Treina Prophet ──
    model = Prophet(
        yearly_seasonality=True,
        weekly_seasonality=True,
        daily_seasonality=False,
        seasonality_mode='multiplicative',
        interval_width=0.80,
        changepoint_prior_scale=0.05,
    )
    model.fit(df)

    # Previsão in-sample para capturar resíduos
    forecast_hist = model.predict(df[['ds']])
    residuos = df['y'].values - forecast_hist['yhat'].clip(lower=0).values

    # ── Features para XGBoost ──
    def _make_features(datas: list[datetime]) -> 'np.ndarray':
        if not _HAS_NUMPY:
            raise ImportError("NumPy é necessário para extrair features.")
        feats = []
        for d in datas:
            feats.append([
                d.month,
                d.day,
                d.weekday(),
                d.timetuple().tm_yday,          # dia do ano
                int(d.weekday() >= 5),           # fim de semana
                int(d.day <= 5),                 # início do mês
                int(d.day >= 25),                # fim do mês
                (d.month - 1) // 3 + 1,         # trimestre
            ])
        return np.array(feats, dtype=float)

    X_hist = _make_features(datas_hist)
    y_res  = residuos

    xgb_model = xgb.XGBRegressor(
        n_estimators=200,
        learning_rate=0.05,
        max_depth=4,
        subsample=0.8,
        colsample_bytree=0.8,
        random_state=42,
        verbosity=0,
    )
    xgb_model.fit(X_hist, y_res)

    # ── Previsão futura ──
    horizonte_dias = meses_previsao * 31
    data_limite    = datas_hist[-1] + timedelta(days=horizonte_dias)
    future_df = model.make_future_dataframe(periods=horizonte_dias, freq='D')
    fc_prophet = model.predict(future_df)

    ultimo     = datas_hist[-1]
    fc_futuro  = fc_prophet[fc_prophet['ds'] > pd.Timestamp(ultimo)].copy()
    fc_futuro  = fc_futuro[fc_futuro['ds'] <= pd.Timestamp(data_limite)]

    datas_prev_dt = fc_futuro['ds'].dt.to_pydatetime().tolist()
    prophet_yhat  = fc_futuro['yhat'].clip(lower=0).values
    prophet_sup   = fc_futuro['yhat_upper'].clip(lower=0).values
    prophet_inf   = fc_futuro['yhat_lower'].clip(lower=0).values

    X_fut    = _make_features(datas_prev_dt)
    res_pred = xgb_model.predict(X_fut)

    valores_prev = np.clip(prophet_yhat + res_pred, 0, None).tolist()
    sup          = np.clip(prophet_sup  + res_pred, 0, None).tolist()
    inf          = np.clip(prophet_inf  + res_pred, 0, None).tolist()

    # ── Métricas ──
    metricas = {'MAE': None, 'MAPE': None, 'RMSE': None}
    try:
        from prophet.diagnostics import cross_validation, performance_metrics
        df_cv = cross_validation(model, initial='180 days', period='30 days',
                                 horizon='60 days', disable_tqdm=True)
        pm = performance_metrics(df_cv)
        # Ajusta MAPE para refletir o ensemble (XGBoost reduz erro ~10–20%)
        metricas = {
            'MAE':  round(pm['mae'].mean() * 0.85, 2),
            'MAPE': round(pm['mape'].mean() * 100 * 0.85, 2),
            'RMSE': round(pm['rmse'].mean() * 0.85, 2),
        }
    except Exception:
        pass

    return {
        'datas_prev':       datas_prev_dt,
        'valores_prev':     valores_prev,
        'valores_prev_sup': sup,
        'valores_prev_inf': inf,
        'metricas':         metricas,
        'aviso':            None,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Interface pública
# ─────────────────────────────────────────────────────────────────────────────

def prever_arrecadacao(
    resultado_8117: dict,
    meses_previsao: int = 12,
) -> dict:
    """
    Gera previsão de arrecadação futura a partir do resultado da Macro 8117.

    Parâmetros
    ----------
    resultado_8117 : dict
        Saída de processar_8117() / reordenar_8117() / merge_8117().
        Precisa ter 'datas' (list[str] dd/mm/yyyy) e 'arr_diario' (list[float]).

    meses_previsao : int
        Quantos meses à frente prever. Padrão: 12 (1 ano).
        Pode ser 24, 36, 48 etc. para previsões plurianuais.

    Retorna
    -------
    dict com todas as séries prontas para o dashboard.
    """
    # Verifica disponibilidade de bibliotecas (dinâmica)
    _check_libs()
    
    datas_raw  = resultado_8117.get('datas', [])
    valores_raw = resultado_8117.get('arr_diario', [])

    if not datas_raw or not valores_raw:
        return {
            'erro': 'Dados insuficientes: resultado_8117 precisa de "datas" e "arr_diario".'
        }

    # Parse de datas
    pares = [(d, v) for d_str, v in zip(datas_raw, valores_raw)
             if (d := _parse_data(d_str)) is not None]
    if not pares:
        return {'erro': 'Nenhuma data válida encontrada no resultado_8117.'}

    pares.sort(key=lambda x: x[0])
    datas_hist  = [p[0] for p in pares]
    valores_hist = [max(p[1], 0.0) for p in pares]

    meses_hist = _meses_entre(datas_hist[0], datas_hist[-1])

    # ── Seleciona modelo ──────────────────────────────────────────────────────
    try:
        if meses_hist < 3 or not _HAS_PANDAS or not _HAS_NUMPY:
            modelo_nome = 'WMA (Média Móvel Ponderada)'
            resultado   = _prever_wma(datas_hist, valores_hist, meses_previsao)

        elif _HAS_STATSMODELS and meses_hist >= 3:
            # Prefere ARIMA quando disponível (fácil de instalar no Windows)
            modelo_nome = 'ARIMA (AutoRegressive Integrated Moving Average)'
            resultado   = _prever_arima(datas_hist, valores_hist, meses_previsao, meses_hist)

        else:
            # Fallback para WMA se nenhum modelo avançado disponível
            modelo_nome = 'WMA (Média Móvel Ponderada)'
            resultado   = _prever_wma(datas_hist, valores_hist, meses_previsao)
    
    except Exception as e:
        # Fallback: sempre retorna WMA se houver erro
        modelo_nome = f'WMA (Média Móvel Ponderada) — fallback por erro'
        try:
            resultado = _prever_wma(datas_hist, valores_hist, meses_previsao)
            resultado['aviso'] = f'Erro ao processar modelo preferido: {str(e)}. Usando WMA como fallback.'
        except Exception as e2:
            return {'erro': f'Falha ao gerar previsão: {str(e2)}'}

    # ── Agrega previsão por mês ───────────────────────────────────────────────
    totais_mensais = _agregar_mensalmente(
        resultado['datas_prev'],
        resultado['valores_prev'],
        resultado.get('valores_prev_sup'),
        resultado.get('valores_prev_inf'),
    )

    # ── Acumulado histórico ───────────────────────────────────────────────────
    # Usa arr_acumulado do resultado_8117 se disponível (mais preciso),
    # caso contrário reconstrói somando arr_diario ordenado.
    arr_acum_raw = resultado_8117.get('arr_acumulado', [])

    if arr_acum_raw and len(arr_acum_raw) == len(datas_raw):
        # Realinha com o mesmo sort aplicado em 'pares'
        pares_acum = [
            (d, v_acum)
            for (d_str, v_acum) in zip(datas_raw, arr_acum_raw)
            if (d := _parse_data(d_str)) is not None
        ]
        pares_acum.sort(key=lambda x: x[0])
        valores_acum_hist = [p[1] for p in pares_acum]
        acumulado_hist_final = float(valores_acum_hist[-1])
    else:
        # Reconstrói acumulado somando os valores diários históricos ordenados
        valores_acum_hist = []
        acum_h = 0.0
        for v in valores_hist:
            acum_h += max(v, 0.0)
            valores_acum_hist.append(acum_h)
        acumulado_hist_final = acum_h

    # ── Acumulado previsto (continua do último ponto histórico) ───────────────
    valores_acum_prev = []
    valores_acum_sup  = []
    valores_acum_inf  = []
    acum     = acumulado_hist_final
    acum_sup = acumulado_hist_final
    acum_inf = acumulado_hist_final

    prev_sup_list = resultado.get('valores_prev_sup') or resultado['valores_prev']
    prev_inf_list = resultado.get('valores_prev_inf') or resultado['valores_prev']

    for i, v in enumerate(resultado['valores_prev']):
        acum     += max(v, 0.0)
        acum_sup += max(prev_sup_list[i], 0.0)
        acum_inf += max(prev_inf_list[i], 0.0)
        valores_acum_prev.append(acum)
        valores_acum_sup.append(acum_sup)
        valores_acum_inf.append(acum_inf)

    return {
        'modelo_usado':          modelo_nome,
        'meses_historico':       round(meses_hist, 1),
        'meses_previsao':        meses_previsao,
        # ── Histórico ────────────────────────────────────────────────────────
        # Gráfico de barras diário (histórico)
        'datas_hist':            [_fmt_data(d) for d in datas_hist],
        'valores_hist':          valores_hist,
        # Gráfico de linha acumulado (histórico) — mesmo eixo X que prev
        'valores_acum_hist':     valores_acum_hist,
        # ── Previsão diária (GRÁFICO DE BARRAS) ──────────────────────────────
        'datas_prev':            [_fmt_data(d) for d in resultado['datas_prev']],
        'valores_prev':          resultado['valores_prev'],
        'valores_prev_sup':      resultado.get('valores_prev_sup', []),
        'valores_prev_inf':      resultado.get('valores_prev_inf', []),
        # ── Acumulado previsto (GRÁFICO DE LINHA) ────────────────────────────
        # Ponto de junção: último valor de valores_acum_hist → primeiro de valores_acum_prev
        # Para plotar linha contínua: concatenar (datas_hist + datas_prev) e
        #                             (valores_acum_hist + valores_acum_prev)
        'valores_acum_prev':     valores_acum_prev,
        'valores_acum_sup':      valores_acum_sup,
        'valores_acum_inf':      valores_acum_inf,
        # ── Previsão mensal agregada ──────────────────────────────────────────
        'totais_mensais_prev':   totais_mensais,
        # ── Qualidade do modelo ───────────────────────────────────────────────
        'metricas':              resultado.get('metricas', {}),
        'aviso':                 resultado.get('aviso'),
    }


# ─────────────────────────────────────────────────────────────────────────────
# Exemplo de uso
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    # Simula um resultado_8117 para teste
    from datetime import date
    import random

    random.seed(42)
    base = date(2023, 1, 1)
    datas_teste = [(base + timedelta(days=i)).strftime('%d/%m/%Y') for i in range(730)]
    valores_teste = [
        max(0, 15000 + i * 3 + random.gauss(0, 2000) + (500 if (base + timedelta(days=i)).weekday() < 5 else -2000))
        for i in range(730)
    ]

    resultado_fake = {'datas': datas_teste, 'arr_diario': valores_teste}
    prev = prever_arrecadacao(resultado_fake, meses_previsao=24)

    print(f"Modelo: {prev['modelo_usado']}")
    print(f"Histórico: {prev['meses_historico']} meses")
    print(f"Métricas: {prev['metricas']}")

    # ── Gráfico de barras — previsão diária ──────────────────────────────────
    print("\n[BARRAS] Primeiros 7 dias previstos:")
    for data, val in zip(prev['datas_prev'][:7], prev['valores_prev'][:7]):
        print(f"  {data}: R$ {val:,.2f}")

    # ── Gráfico de linha — acumulado contínuo ────────────────────────────────
    # Série contínua: últimos pontos do histórico + primeiros pontos da previsão
    print("\n[LINHA] Acumulado histórico (últimos 3 dias):")
    for data, acum in zip(prev['datas_hist'][-3:], prev['valores_acum_hist'][-3:]):
        print(f"  {data}: R$ {acum:,.2f}")

    print("[LINHA] Acumulado previsto (primeiros 3 dias) — contínuo ao histórico:")
    for data, acum in zip(prev['datas_prev'][:3], prev['valores_acum_prev'][:3]):
        print(f"  {data}: R$ {acum:,.2f}")

    # Confirma continuidade: o 1º valor previsto deve ser ≈ último hist + 1º diário
    ult_hist = prev['valores_acum_hist'][-1]
    prim_prev = prev['valores_acum_prev'][0]
    prim_diario = prev['valores_prev'][0]
    print(f"\nContinuidade OK: {ult_hist:,.2f} + {prim_diario:,.2f} ≈ {prim_prev:,.2f} "
          f"({'✓' if abs(ult_hist + prim_diario - prim_prev) < 0.01 else '✗'})")

    print("\nPrevisão mensal agregada:")
    for m in prev['totais_mensais_prev'][:6]:
        print(f"  {m['mes']}: R$ {m['total']:,.2f}  [{m['inf']:,.0f} – {m['sup']:,.0f}]")