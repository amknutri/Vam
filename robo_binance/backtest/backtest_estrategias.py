"""
Backtest das estratégias do robô de Futuros (v8) — comparação honesta.

O que faz:
  1. Baixa velas de 1h e taxas de funding dos últimos 2 anos (dados PÚBLICOS
     da Binance, sem chave de API, sem login).
  2. Simula o robô hora a hora, com taxas, slippage, funding e as mesmas
     regras de risco da v8 (stop, alvo, freio, limite diário etc.).
  3. Compara variantes e grava o resultado em resultado_backtest.txt e
     resultado_backtest.json (mais um CSV de operações por variante).

NÃO envia ordem nenhuma. Não usa chave. Só lê dados públicos.

Uso (PowerShell):
    cd C:\\RoboBacktest
    python backtest_estrategias.py

Os indicadores e a decisão de entrada da confluência vêm do PRÓPRIO arquivo
do robô (bot_textual_v8_FUTUROS.py), para o teste ser fiel ao que roda.
"""
import bisect
import csv
import importlib.util
import io
import json
import math
import os
import sys
import time
import urllib.error
import urllib.request
import zipfile
from datetime import datetime, timedelta, timezone

# ----------------------------------------------------------------------------
# Configuração
# ----------------------------------------------------------------------------
INICIO = "2024-10-01"          # início do período avaliado (UTC)
FIM = "2026-10-01"             # fim (exclusive)
AQUECIMENTO_DIAS = 25          # dados antes do início, só para aquecer indicadores
BANCA_INICIAL = 5000.0
PASTA_DADOS = "dados"
PASTA_SAIDA = "."

CAMINHOS_BOT = [
    "bot_textual_v8_FUTUROS.py",
    os.path.join("..", "RoboBinanceFuturos", "bot_textual_v8_FUTUROS.py"),
    r"C:\RoboBinanceFuturos\bot_textual_v8_FUTUROS.py",
]

# As 10 mais líquidas com histórico completo no período
TOP10 = ["BTC", "ETH", "SOL", "BNB", "XRP", "DOGE", "ADA", "LINK", "AVAX", "LTC"]

# Moedas da lista do robô que NÃO existem em Futuros com esse nome: o robô
# nunca consegue operá-las (em Futuros a SHIB é "1000SHIB"), então ficam fora.
NUNCA_OPERADAS = {"SHIB"}
SIMBOLO_FUTUROS = {}

UMA_HORA = 3600 * 1000
QUATRO_HORAS = 4 * UMA_HORA
OITO_HORAS = 8 * UMA_HORA
FUSO_BRASILIA = timedelta(hours=-3)


def log(msg):
    print(msg, flush=True)


# ----------------------------------------------------------------------------
# Carrega o robô (só para usar as funções de indicador e de decisão)
# ----------------------------------------------------------------------------
def carregar_robo(caminho_extra=None):
    candidatos = ([caminho_extra] if caminho_extra else []) + CAMINHOS_BOT
    for caminho in candidatos:
        if caminho and os.path.exists(caminho):
            spec = importlib.util.spec_from_file_location("robo_v8", caminho)
            modulo = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(modulo)
            return modulo, os.path.abspath(caminho)
    log("ERRO: não encontrei o arquivo bot_textual_v8_FUTUROS.py.")
    log("Coloque-o nesta pasta ou em C:\\RoboBinanceFuturos e rode de novo.")
    sys.exit(1)


# ----------------------------------------------------------------------------
# Download de dados públicos (com cache em disco)
# ----------------------------------------------------------------------------
def _get_json(url, tentativas=4):
    ultimo_erro = None
    for i in range(tentativas):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 backtest-robo"})
            with urllib.request.urlopen(req, timeout=30) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code in (400, 404):
                raise
            ultimo_erro = e
        except Exception as e:  # rede instável: tenta de novo
            ultimo_erro = e
        time.sleep(2 * (i + 1))
    raise ultimo_erro


def _get_zip_csv(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 backtest-robo"})
    with urllib.request.urlopen(req, timeout=60) as r:
        conteudo = r.read()
    with zipfile.ZipFile(io.BytesIO(conteudo)) as z:
        nome = z.namelist()[0]
        texto = z.read(nome).decode("utf-8")
    linhas = list(csv.reader(io.StringIO(texto)))
    if linhas and not linhas[0][0].lstrip("-").isdigit():
        linhas = linhas[1:]  # arquivos novos têm cabeçalho
    return linhas


def _ms(data_txt):
    return int(datetime.strptime(data_txt, "%Y-%m-%d").replace(tzinfo=timezone.utc).timestamp() * 1000)


def _meses(inicio_ms, fim_ms):
    d = datetime.fromtimestamp(inicio_ms / 1000, tz=timezone.utc).replace(day=1)
    fim = datetime.fromtimestamp(fim_ms / 1000, tz=timezone.utc)
    while d < fim:
        yield d.strftime("%Y-%m")
        d = (d.replace(day=28) + timedelta(days=4)).replace(day=1)


def baixar_velas(simbolo, inicio_ms, fim_ms):
    """Velas de 1h: [ts, open, high, low, close, volume, volume_em_usdt]."""
    os.makedirs(PASTA_DADOS, exist_ok=True)
    arquivo = os.path.join(PASTA_DADOS, f"{simbolo}_1h.csv")
    if os.path.exists(arquivo):
        with open(arquivo, encoding="utf-8") as f:
            velas = [[int(r[0])] + [float(x) for x in r[1:7]] for r in csv.reader(f)]
        if velas and velas[-1][0] >= fim_ms - 2 * UMA_HORA:
            return velas, "cache"

    velas, fonte = [], None
    # Fonte 1: API pública de Futuros
    try:
        t = inicio_ms
        while t < fim_ms:
            url = (f"https://fapi.binance.com/fapi/v1/klines?symbol={simbolo}&interval=1h"
                   f"&startTime={t}&endTime={fim_ms - 1}&limit=1500")
            lote = _get_json(url)
            if not lote:
                break
            velas += [[int(k[0]), float(k[1]), float(k[2]), float(k[3]), float(k[4]), float(k[5]), float(k[7])] for k in lote]
            t = int(lote[-1][0]) + UMA_HORA
            time.sleep(0.15)
        fonte = "fapi.binance.com"
    except Exception as e:
        log(f"   {simbolo}: API de Futuros indisponível ({type(e).__name__}); tentando arquivos públicos...")
        velas = []
    # Fonte 2: arquivos mensais públicos (data.binance.vision)
    if not velas:
        for mes in _meses(inicio_ms, fim_ms):
            url = f"https://data.binance.vision/data/futures/um/monthly/klines/{simbolo}/1h/{simbolo}-1h-{mes}.zip"
            try:
                for k in _get_zip_csv(url):
                    velas.append([int(k[0]), float(k[1]), float(k[2]), float(k[3]), float(k[4]), float(k[5]), float(k[7])])
            except Exception:
                continue  # mês sem arquivo (moeda ainda não listada, ou mês corrente)
        fonte = "data.binance.vision" if velas else None
    if not velas:
        return [], None
    vistos, limpas = set(), []
    for v in sorted(velas):
        if v[0] not in vistos and inicio_ms <= v[0] < fim_ms:
            vistos.add(v[0])
            limpas.append(v)
    with open(arquivo, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(limpas)
    return limpas, fonte


def baixar_funding(simbolo, inicio_ms, fim_ms):
    """Taxas de funding: lista de (ts, taxa). Taxa positiva: LONG paga."""
    arquivo = os.path.join(PASTA_DADOS, f"{simbolo}_funding.csv")
    if os.path.exists(arquivo):
        with open(arquivo, encoding="utf-8") as f:
            dados = [(int(r[0]), float(r[1])) for r in csv.reader(f)]
        if dados:
            return dados, "cache"
    dados = []
    try:
        t = inicio_ms
        while t < fim_ms:
            url = (f"https://fapi.binance.com/fapi/v1/fundingRate?symbol={simbolo}"
                   f"&startTime={t}&endTime={fim_ms - 1}&limit=1000")
            lote = _get_json(url)
            if not lote:
                break
            dados += [(int(x["fundingTime"]), float(x["fundingRate"])) for x in lote]
            t = int(lote[-1]["fundingTime"]) + 1
            time.sleep(0.15)
    except Exception:
        dados = []
    if not dados:
        for mes in _meses(inicio_ms, fim_ms):
            url = f"https://data.binance.vision/data/futures/um/monthly/fundingRate/{simbolo}/{simbolo}-fundingRate-{mes}.zip"
            try:
                for r in _get_zip_csv(url):
                    dados.append((int(r[0]), float(r[2])))
            except Exception:
                continue
    if not dados:
        return [], None
    dados = sorted(set(dados))
    with open(arquivo, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(dados)
    return dados, "baixado"


# ----------------------------------------------------------------------------
# Indicadores pré-calculados (uma vez por moeda; todas as variantes reaproveitam)
# ----------------------------------------------------------------------------
class _EstadoFalso:
    """Imita o pedaço do robô que a função avaliar_confluencia lê."""

    def __init__(self, robo):
        self.rsi_cache, self.macd_cache, self.bollinger_cache = {}, {}, {}
        self.estocastico_cache, self.candle_cache = {}, {}
        self.variacao_curta_cache, self.regime_cache = {}, {}
        self.confluencia_minima_padrao = robo.CONFLUENCIA_MINIMA_PADRAO
        self.confluencia_minima_banca_pequena = robo.CONFLUENCIA_MINIMA_BANCA_PEQUENA
        self.banca_total = BANCA_INICIAL


def agregar_4h(velas):
    """Junta velas de 1h em velas de 4h (alinhadas em 00/04/08/12/16/20 UTC)."""
    grupos = {}
    for v in velas:
        chave = v[0] - (v[0] % QUATRO_HORAS)
        g = grupos.get(chave)
        if g is None:
            grupos[chave] = [chave, v[1], v[2], v[3], v[4], v[5], v[6], 1]
        else:
            g[2] = max(g[2], v[2]); g[3] = min(g[3], v[3]); g[4] = v[4]
            g[5] += v[5]; g[6] += v[6]; g[7] += 1
    return [g[:7] for g in sorted(grupos.values()) if g[7] == 4]


def calcular_sinais_moeda(args):
    """Para cada vela de 1h fechada: decisão da confluência (exatamente a função
    do robô), ATR, volume 24h, volume relativo, tendência 4h e canais Donchian/ADX."""
    caminho_bot, moeda, velas = args
    robo, _ = carregar_robo(caminho_bot)
    estado = _EstadoFalso(robo)
    simbolo = f"{moeda}/USDT"
    n = len(velas)
    janela = max(robo.RSI_PERIODO, robo.ATR_PERIODO, robo.MACD_LENTA + robo.MACD_SINAL, robo.BB_PERIODO,
                 robo.ESTOCASTICO_PERIODO_K + robo.ESTOCASTICO_SUAVIZACAO_K + robo.ESTOCASTICO_PERIODO_D) + 20

    # --- 4h: tendência (EMA50 sobre 60 velas fechadas) e Donchian/ADX (50 velas) ---
    v4 = agregar_4h(velas)
    tend4, dc4 = {}, {}
    for j in range(len(v4)):
        fim_vela = v4[j][0] + QUATRO_HORAS
        if j + 1 >= robo.TENDENCIA_EMA_PERIODO + 10:
            closes = [c[4] for c in v4[j - (robo.TENDENCIA_EMA_PERIODO + 10) + 1: j + 1]]
            ema = robo.calcular_ema(closes, robo.TENDENCIA_EMA_PERIODO)
            if ema is not None:
                tend4[fim_vela] = "ALTA" if closes[-1] >= ema else "BAIXA"
        if j + 1 >= 50:
            bloco = v4[j - 49: j + 1]
            ent = bloco[-robo.DONCHIAN_ADX_ENTRADA:]
            sai = bloco[-robo.DONCHIAN_ADX_SAIDA:]
            dc4[fim_vela] = (max(c[2] for c in ent), min(c[3] for c in ent),
                             max(c[2] for c in sai), min(c[3] for c in sai),
                             robo.calcular_adx(bloco, robo.DONCHIAN_ADX_PERIODO))

    sinais = [None] * n
    ultimo_tend, ultimo_dc = None, None
    soma_qv24 = 0.0
    for i in range(n):
        soma_qv24 += velas[i][6]
        if i >= 24:
            soma_qv24 -= velas[i - 24][6]
        fecha = velas[i][0] + UMA_HORA
        if fecha in tend4:
            ultimo_tend = tend4[fecha]
        if fecha in dc4:
            ultimo_dc = dc4[fecha]
        if i + 1 < janela:
            continue
        ohlcv = [v[:6] for v in velas[i + 1 - janela: i + 1]]
        closes = [c[4] for c in ohlcv]
        estado.rsi_cache[simbolo] = robo.calcular_rsi(closes)
        atr = robo.calcular_atr(ohlcv)
        estado.macd_cache[simbolo] = robo.calcular_macd(closes)
        estado.bollinger_cache[simbolo] = robo.calcular_bollinger(closes)
        estado.estocastico_cache[simbolo] = robo.calcular_estocastico(ohlcv)
        estado.candle_cache[simbolo] = robo.detectar_padrao_candle(ohlcv)
        estado.regime_cache[simbolo] = robo.calcular_regime(estado.bollinger_cache[simbolo])
        n_var = getattr(robo, "VARIACAO_JANELA_VELAS", 1)
        estado.variacao_curta_cache[simbolo] = ((closes[-1] - closes[-1 - n_var]) / closes[-1 - n_var] * 100
                                                if closes[-1 - n_var] > 0 else None)
        vols = [c[5] for c in ohlcv]
        relvol = vols[-1] / (sum(vols[-21:-1]) / 20) if sum(vols[-21:-1]) > 0 else None
        vol24 = soma_qv24 if i >= 23 else 0.0
        acao, _txt, _reg = robo.BotTraderApp.avaliar_confluencia(estado, simbolo, closes[-1], vol24)
        sinais[i] = (acao, atr, vol24, relvol, ultimo_tend, ultimo_dc)
    return moeda, sinais


# ----------------------------------------------------------------------------
# Simulador de carteira
# ----------------------------------------------------------------------------
class Variante:
    def __init__(self, codigo, nome, sistema="confluencia", moedas=None, tendencia=False,
                 filtro_volume=False, trava_ganho=True, freio=True, custo_mult=1.0,
                 dc_tamanho_fixo=False, dc_alavancagem=1, max_posicoes=None):
        self.codigo, self.nome, self.sistema = codigo, nome, sistema
        self.moedas, self.tendencia, self.filtro_volume = moedas, tendencia, filtro_volume
        self.trava_ganho, self.freio, self.custo_mult = trava_ganho, freio, custo_mult
        self.dc_tamanho_fixo, self.dc_alavancagem = dc_tamanho_fixo, dc_alavancagem
        self.max_posicoes = max_posicoes


def _dia_brasilia(ms):
    return (datetime.fromtimestamp(ms / 1000, tz=timezone.utc) + FUSO_BRASILIA).date()


def simular(var, robo, dados, sinais, funding, ordem_moedas):
    taxa = robo.TAKER_FEE_PCT / 100 * var.custo_mult
    slip = robo.SLIPPAGE_ESTIMADO_PCT / 100 * var.custo_mult
    alav = min(3, robo.ALAVANCAGEM_MAXIMA_PERMITIDA)
    max_pos = var.max_posicoes or robo.MAX_POSICOES_SIMULTANEAS
    moedas = [m for m in ordem_moedas if m in dados and (var.moedas is None or m in var.moedas)]
    inicio_ms, fim_ms = _ms(INICIO), _ms(FIM)
    linha_tempo = sorted({v[0] for m in moedas for v in dados[m]["velas"] if inicio_ms <= v[0] < fim_ms})

    banca = BANCA_INICIAL
    posicoes = {}          # moeda -> dict
    operacoes = []
    curva = []             # (ts, patrimônio marcado a mercado)
    dia_atual, lucro_dia, pico_dia, ganho_sistema_dia = None, 0.0, 0.0, 0.0
    pausado_dia, dias_limite = False, []
    cooldown, stops_recentes, freio_ate = {}, [], 0
    ambiguas = 0
    ultimo_preco = {}

    def fechar(m, p, preco_bruto, motivo, ts_fecha):
        nonlocal banca, lucro_dia, ganho_sistema_dia, stops_recentes, freio_ate
        preco = preco_bruto * (1 - slip * p["lado"])
        bruto = p["lado"] * (preco - p["entrada"]) * p["qtd"]
        taxa_saida = preco * p["qtd"] * taxa
        liquido = bruto - p["taxa_entrada"] - taxa_saida + p["funding"]
        banca += liquido
        lucro_dia += liquido
        ganho_sistema_dia += liquido
        operacoes.append({
            "moeda": m, "lado": "LONG" if p["lado"] > 0 else "SHORT",
            "abertura": p["ts_abre"], "fechamento": ts_fecha, "motivo": motivo,
            "entrada": p["entrada"], "saida": preco, "qtd": p["qtd"], "risco": p["risco"],
            "liquido": liquido, "taxas": p["taxa_entrada"] + taxa_saida, "funding": p["funding"],
            "r": liquido / p["risco"] if p["risco"] > 0 else 0.0,
        })
        if var.freio:
            if liquido >= 0:
                stops_recentes = []
            else:
                cooldown[m] = ts_fecha + robo.COOLDOWN_MOEDA_HORAS * UMA_HORA
                stops_recentes = [t for t in stops_recentes if ts_fecha - t < robo.FREIO_JANELA_HORAS * UMA_HORA] + [ts_fecha]
                if len(stops_recentes) >= robo.FREIO_STOPS_SEGUIDOS:
                    freio_ate = ts_fecha + robo.FREIO_PAUSA_HORAS * UMA_HORA
                    stops_recentes = []

    for ts in linha_tempo:
        ts_fecha = ts + UMA_HORA
        velas_agora = {}
        for m in moedas:
            i = dados[m]["indice"].get(ts)
            if i is not None:
                velas_agora[m] = i

        # 1) funding cobrado no início da hora (00/08/16 UTC ou conforme histórico)
        for m, p in posicoes.items():
            if m in velas_agora:
                taxa_f = funding[m].get(ts)
                if taxa_f is not None:
                    preco_ref = dados[m]["velas"][velas_agora[m]][1]
                    p["funding"] -= p["lado"] * p["qtd"] * preco_ref * taxa_f

        # 2) stops e alvos dentro da vela (premissa conservadora: se a vela tocou
        #    stop e alvo, considera que o STOP veio primeiro)
        for m in list(posicoes):
            if m not in velas_agora:
                continue
            p = posicoes[m]
            _, o, h, l, c = dados[m]["velas"][velas_agora[m]][:5]
            lado = p["lado"]
            favoravel, contra = (h, l) if lado > 0 else (l, h)
            def passou(preco, nivel, a_favor):
                return (preco >= nivel) if (lado > 0) == a_favor else (preco <= nivel)
            saida = None
            if passou(o, p["sl"], False):
                saida = (o, "STOP (abriu além do stop)")
            elif p["tp"] is not None and passou(o, p["tp"], True):
                saida = (o, "ALVO")
            elif passou(contra, p["sl"], False):
                if p["tp"] is not None and passou(favoravel, p["tp"], True):
                    ambiguas += 1
                saida = (p["sl"], "STOP" if not p.get("be") else "STOP NO ZERO A ZERO")
            elif p["tp"] is not None and passou(favoravel, p["tp"], True):
                saida = (p["tp"], "ALVO")
            elif (var.sistema == "confluencia" and not p.get("be")
                  and passou(favoravel, p["entrada"] + lado * p["dist_r"], True)):
                folga = p["entrada"] * (2 * robo.TAKER_FEE_PCT + robo.SLIPPAGE_ESTIMADO_PCT) / 100
                p["sl"] = p["entrada"] + lado * folga
                p["be"] = True
                if passou(c, p["sl"], False):
                    saida = (p["sl"], "STOP NO ZERO A ZERO")
            if saida:
                fechar(m, p, saida[0], saida[1], ts_fecha)
                del posicoes[m]

        for m, i in velas_agora.items():
            ultimo_preco[m] = dados[m]["velas"][i][4]

        # 3) Donchian: saída pelo canal largo (preço de fechamento da hora)
        if var.sistema == "donchian":
            for m in list(posicoes):
                if m not in velas_agora:
                    continue
                s = sinais[m][velas_agora[m]]
                if not s or not s[5]:
                    continue
                c = dados[m]["velas"][velas_agora[m]][4]
                _maxE, _minE, maxS, minS, _adx = s[5]
                p = posicoes[m]
                if (p["lado"] > 0 and c < minS) or (p["lado"] < 0 and c > maxS):
                    fechar(m, p, c, "SAÍDA PELO CANAL", ts_fecha)
                    del posicoes[m]

        # 4) virada do dia (horário de Brasília) e limites diários
        dia = _dia_brasilia(ts_fecha)
        if dia != dia_atual:
            dia_atual, lucro_dia, pico_dia, ganho_sistema_dia, pausado_dia = dia, 0.0, 0.0, 0.0, False
        aberto = sum(p["lado"] * (ultimo_preco.get(m, p["entrada"]) - p["entrada"]) * p["qtd"] for m, p in posicoes.items())
        if not pausado_dia:
            if lucro_dia + aberto <= -abs(banca * robo.LIMITE_PERDA_DIARIA_PCT / 100):
                pausado_dia = True
                dias_limite.append(dia)
            pico_dia = max(pico_dia, lucro_dia)
            if not pausado_dia and pico_dia - lucro_dia >= banca * robo.DRAWDOWN_MAXIMO_PCT / 100:
                pausado_dia = True
                dias_limite.append(dia)
        curva.append((ts_fecha, banca + aberto))

        # 5) entradas novas (na ordem da lista de moedas do robô)
        if pausado_dia:
            continue
        if var.sistema == "confluencia" and var.trava_ganho and ganho_sistema_dia >= robo.STOP_GANHO_DIARIO_CONFLUENCIA_USD:
            continue
        longs = sum(1 for p in posicoes.values() if p["lado"] > 0)
        shorts = len(posicoes) - longs
        for m in moedas:
            if len(posicoes) >= max_pos:
                break
            if m in posicoes or m not in velas_agora:
                continue
            s = sinais[m][velas_agora[m]]
            if not s:
                continue
            acao, atr, vol24, relvol, tend, dc = s
            c = dados[m]["velas"][velas_agora[m]][4]
            if var.sistema == "confluencia":
                if acao not in ("COMPRA", "VENDA") or not atr or atr <= 0:
                    continue
            else:
                if not dc or dc[4] is None or dc[4] < robo.DONCHIAN_ADX_MINIMO:
                    continue
                acao = "COMPRA" if c > dc[0] else ("VENDA" if c < dc[1] else None)
                if acao is None:
                    continue
            lado = 1 if acao == "COMPRA" else -1
            if lado > 0 and longs >= robo.MAX_LONGS_SIMULTANEOS:
                continue
            if lado < 0 and shorts >= robo.MAX_SHORTS_SIMULTANEOS:
                continue
            if var.tendencia and tend != ("ALTA" if lado > 0 else "BAIXA"):
                continue
            if var.freio and (ts_fecha < freio_ate or ts_fecha < cooldown.get(m, 0)):
                continue
            if var.filtro_volume and (relvol is None or relvol < robo.FILTRO_VOLUME_MULTIPLO):
                continue
            # funding contra a posição (última taxa conhecida)
            taxa_f = funding[m].get("ultima_ate", lambda t: None)(ts_fecha)
            if taxa_f is not None and lado * taxa_f * 100 > robo.FUNDING_MAX_CONTRA_PCT:
                continue

            entrada = c * (1 + slip * lado)
            if var.sistema == "confluencia":
                dist_sl = atr * robo.SL_ATR_MULTIPLO
                dist_tp = dist_sl * robo.RELACAO_RISCO_RETORNO
                custo_pct = 2 * (robo.TAKER_FEE_PCT + robo.SLIPPAGE_ESTIMADO_PCT)
                if dist_tp / c * 100 < robo.CUSTO_MULTIPLO_MINIMO * custo_pct:
                    continue
            else:
                dist_sl = entrada * robo.DONCHIAN_ADX_SL_PCT / 100
                dist_tp = None

            risco_alvo = banca * robo.RISCO_POR_TRADE_PCT / 100
            if sum(p["risco"] for p in posicoes.values()) + risco_alvo > banca * robo.MAX_RISCO_TOTAL_PCT / 100:
                continue
            if var.sistema == "donchian" and var.dc_tamanho_fixo:
                livre = banca - sum(p["margem"] for p in posicoes.values())
                nocional = livre * robo.DONCHIAN_ADX_FRACAO_BANCA_PCT / 100
                qtd = nocional / entrada
                alav_usada = var.dc_alavancagem
            else:
                qtd = risco_alvo / dist_sl
                teto = banca * alav * robo.MAX_EXPOSICAO_POR_POSICAO_PCT / 100
                if qtd * entrada > teto:
                    qtd = teto / entrada
                alav_usada = alav
            if qtd <= 0:
                continue
            nocional = qtd * entrada
            margem = nocional / alav_usada
            livre = banca - sum(p["margem"] for p in posicoes.values())
            if margem > livre * 0.95:
                continue
            posicoes[m] = {
                "lado": lado, "entrada": entrada, "qtd": qtd, "risco": dist_sl * qtd,
                "dist_r": dist_sl, "sl": entrada - lado * dist_sl,
                "tp": (entrada + lado * dist_tp) if dist_tp else None,
                "taxa_entrada": nocional * taxa, "funding": 0.0, "margem": margem,
                "ts_abre": ts_fecha,
            }
            if lado > 0:
                longs += 1
            else:
                shorts += 1

    # fecha o que sobrou no último preço
    for m, p in list(posicoes.items()):
        fechar(m, p, ultimo_preco.get(m, p["entrada"]), "FIM DO TESTE", linha_tempo[-1] + UMA_HORA)
    return operacoes, curva, dias_limite, ambiguas


# ----------------------------------------------------------------------------
# Métricas
# ----------------------------------------------------------------------------
def metricas(operacoes, curva, dias_limite, ambiguas):
    r = {"operacoes": len(operacoes), "velas_ambiguas": ambiguas}
    if not operacoes:
        return r
    ganhos = [o["liquido"] for o in operacoes if o["liquido"] > 0]
    perdas = [o["liquido"] for o in operacoes if o["liquido"] <= 0]
    soma_g, soma_p = sum(ganhos), -sum(perdas)
    r["fator_lucro"] = round(soma_g / soma_p, 3) if soma_p > 0 else None
    r["taxa_acerto_pct"] = round(100 * len(ganhos) / len(operacoes), 1)
    r["resultado_usdt"] = round(sum(o["liquido"] for o in operacoes), 2)
    r["retorno_pct"] = round(100 * r["resultado_usdt"] / BANCA_INICIAL, 2)
    r["taxas_pagas_usdt"] = round(sum(o["taxas"] for o in operacoes), 2)
    r["funding_usdt"] = round(sum(o["funding"] for o in operacoes), 2)
    r["media_r"] = round(sum(o["r"] for o in operacoes) / len(operacoes), 3)
    r["ganho_medio"] = round(soma_g / len(ganhos), 2) if ganhos else 0
    r["perda_media"] = round(-soma_p / len(perdas), 2) if perdas else 0
    pico, dd_max = -1e18, 0.0
    for _, v in curva:
        pico = max(pico, v)
        dd_max = max(dd_max, (pico - v) / pico * 100 if pico > 0 else 0)
    r["queda_maxima_pct"] = round(dd_max, 2)
    meses = max(1e-9, (curva[-1][0] - curva[0][0]) / (1000 * 3600 * 24 * 30.44))
    r["operacoes_por_mes"] = round(len(operacoes) / meses, 1)
    r["retorno_medio_mes_pct"] = round(r["retorno_pct"] / meses, 2)
    seq, pior_seq = 0, 0
    for o in sorted(operacoes, key=lambda x: x["fechamento"]):
        seq = seq + 1 if o["liquido"] <= 0 else 0
        pior_seq = max(pior_seq, seq)
    r["pior_sequencia_perdas"] = pior_seq
    semanas = {}
    for d in dias_limite:
        k = d.isocalendar()[:2]
        semanas[k] = semanas.get(k, 0) + 1
    r["dias_limite_batido"] = len(dias_limite)
    r["semanas_com_mais_de_2_dias_limite"] = sum(1 for v in semanas.values() if v > 2)
    meio = curva[0][0] + (curva[-1][0] - curva[0][0]) / 2
    for nome, filtro in (("1a_metade", lambda o: o["fechamento"] < meio), ("2a_metade", lambda o: o["fechamento"] >= meio)):
        sub = [o for o in operacoes if filtro(o)]
        g = sum(o["liquido"] for o in sub if o["liquido"] > 0)
        p = -sum(o["liquido"] for o in sub if o["liquido"] <= 0)
        r[f"fator_lucro_{nome}"] = round(g / p, 3) if p > 0 else None
        r[f"resultado_{nome}_usdt"] = round(g - p, 2)
    for lado in ("LONG", "SHORT"):
        sub = [o for o in operacoes if o["lado"] == lado]
        g = sum(o["liquido"] for o in sub if o["liquido"] > 0)
        p = -sum(o["liquido"] for o in sub if o["liquido"] <= 0)
        r[f"{lado.lower()}_operacoes"] = len(sub)
        r[f"{lado.lower()}_fator_lucro"] = round(g / p, 3) if p > 0 else None
    por_moeda = {}
    for o in operacoes:
        por_moeda[o["moeda"]] = por_moeda.get(o["moeda"], 0.0) + o["liquido"]
    r["resultado_por_moeda"] = {k: round(v, 2) for k, v in sorted(por_moeda.items(), key=lambda x: x[1])}
    motivos = {}
    for o in operacoes:
        motivos[o["motivo"]] = motivos.get(o["motivo"], 0) + 1
    r["motivos_saida"] = motivos
    por_mes = {}
    for o in operacoes:
        k = datetime.fromtimestamp(o["fechamento"] / 1000, tz=timezone.utc).strftime("%Y-%m")
        por_mes[k] = por_mes.get(k, 0.0) + o["liquido"]
    r["resultado_por_mes"] = {k: round(v, 2) for k, v in sorted(por_mes.items())}
    r["meses_positivos"] = sum(1 for v in por_mes.values() if v > 0)
    r["meses_total"] = len(por_mes)
    return r


def veredito(m):
    """Mesma régua definida para a Demo (antes de ver os resultados)."""
    if not m.get("operacoes"):
        return "sem operações"
    fl = m.get("fator_lucro") or 0
    ok_fl = fl > 1.2
    ok_semana = m.get("semanas_com_mais_de_2_dias_limite", 0) == 0
    ok_estavel = (m.get("fator_lucro_1a_metade") or 0) > 1.0 and (m.get("fator_lucro_2a_metade") or 0) > 1.0
    if ok_fl and ok_estavel:
        return "PASSA no fator de lucro e é estável nos dois anos" + ("" if ok_semana else " (mas estoura o limite diário em algumas semanas)")
    if ok_fl:
        return "Fator de lucro passa no total, mas NÃO é estável entre os dois anos"
    if fl > 1.0:
        return "Ganha um pouco, mas abaixo da régua de 1,2 (não aprovaria)"
    return "PERDE dinheiro no período (não aprovaria)"


# ----------------------------------------------------------------------------
# Programa principal
# ----------------------------------------------------------------------------
def main():
    caminho_extra = sys.argv[1] if len(sys.argv) > 1 else None
    robo, caminho_bot = carregar_robo(caminho_extra)
    log(f"Robô carregado de: {caminho_bot}  ({getattr(robo, 'VERSAO_ROBO', '?')})")
    inicio_ms = _ms(INICIO) - AQUECIMENTO_DIAS * 24 * UMA_HORA
    fim_ms = _ms(FIM)
    agora_ms = int(time.time() * 1000)
    fim_ms = min(fim_ms, agora_ms - (agora_ms % UMA_HORA))

    moedas_robo = [s.split("/")[0] for s in robo.MOEDAS]
    dados, funding, fontes, faltando = {}, {}, {}, []
    log(f"\n[1/3] Baixando dados de {len(moedas_robo)} moedas (1h, {INICIO} a {FIM})...")
    for k, m in enumerate(moedas_robo, 1):
        if m in NUNCA_OPERADAS:
            faltando.append(m)
            log(f"  {k:2d}/{len(moedas_robo)} {m}: não existe em Futuros com esse nome (o robô nunca opera) — fica de fora")
            continue
        simbolo = SIMBOLO_FUTUROS.get(m, f"{m}USDT")
        velas, fonte = baixar_velas(simbolo, inicio_ms, fim_ms)
        if len(velas) < 24 * 60:
            faltando.append(m)
            log(f"  {k:2d}/{len(moedas_robo)} {m}: sem dados suficientes — fica de fora")
            continue
        fund, _ = baixar_funding(simbolo, inicio_ms, fim_ms)
        mapa = {}
        for t, taxa in fund:
            mapa[t - (t % UMA_HORA)] = taxa
        if not mapa:  # sem histórico: usa a taxa "padrão" de 0,01% a cada 8h
            for v in velas:
                if v[0] % OITO_HORAS == 0:
                    mapa[v[0]] = 0.0001
        tempos = sorted(mapa)

        def ultima_ate(t, tempos=tempos, mapa=mapa):
            j = bisect.bisect_right(tempos, t) - 1
            return mapa[tempos[j]] if j >= 0 else None

        mapa["ultima_ate"] = ultima_ate
        funding[m] = mapa
        dados[m] = {"velas": velas, "indice": {v[0]: i for i, v in enumerate(velas)}}
        fontes[m] = fonte
        log(f"  {k:2d}/{len(moedas_robo)} {m}: {len(velas)} velas ({fonte}), funding: {len(fund)} registros")

    if not dados:
        log("ERRO: nenhum dado baixado. Verifique a internet e tente de novo.")
        sys.exit(1)

    log(f"\n[2/3] Calculando indicadores (igual ao robô) para {len(dados)} moedas — pode levar alguns minutos...")
    t0 = time.time()
    tarefas = [(caminho_bot, m, dados[m]["velas"]) for m in dados]
    sinais = {}
    try:
        from multiprocessing import Pool, cpu_count
        with Pool(max(1, min(cpu_count() - 1, 8))) as pool:
            for k, (m, s) in enumerate(pool.imap_unordered(calcular_sinais_moeda, tarefas), 1):
                sinais[m] = s
                log(f"  {k:2d}/{len(tarefas)} {m} pronto ({time.time() - t0:.0f}s)")
    except Exception as e:
        log(f"  (processamento paralelo indisponível: {e}; seguindo um por vez)")
        for k, t in enumerate(tarefas, 1):
            m, s = calcular_sinais_moeda(t)
            sinais[m] = s
            log(f"  {k:2d}/{len(tarefas)} {m} pronto ({time.time() - t0:.0f}s)")

    variantes = [
        Variante("A", "v8 como está hoje (36 moedas, confluência)"),
        Variante("A1", "v8 + filtro de volume ligado (item 11)", filtro_volume=True),
        Variante("A2", "v8 + só a favor da tendência de 4h", tendencia=True),
        Variante("A3", "v8 sem a trava de ganho diário de $10", trava_ganho=False),
        Variante("B", "v8 + tendência 4h + só 10 maiores + sem trava de $10", moedas=TOP10, tendencia=True, trava_ganho=False),
        Variante("C", "Seguir tendência (Donchian+ADX do robô) nas 10 maiores", sistema="donchian", moedas=TOP10),
        Variante("C0", "Donchian+ADX só no BTC, exatamente como está no robô (12,5%, 1x)", sistema="donchian",
                 moedas=["BTC"], dc_tamanho_fixo=True, dc_alavancagem=1, max_posicoes=1),
    ]
    log(f"\n[3/3] Simulando {len(variantes)} variantes (e de novo com custos dobrados)...")
    ordem = [m for m in moedas_robo if m in dados]
    resultados = {}
    os.makedirs(PASTA_SAIDA, exist_ok=True)
    for var in variantes:
        ops, curva, dias_lim, amb = simular(var, robo, dados, sinais, funding, ordem)
        met = metricas(ops, curva, dias_lim, amb)
        var2 = Variante(var.codigo, var.nome, var.sistema, var.moedas, var.tendencia, var.filtro_volume,
                        var.trava_ganho, var.freio, 2.0, var.dc_tamanho_fixo, var.dc_alavancagem, var.max_posicoes)
        ops2, curva2, dl2, amb2 = simular(var2, robo, dados, sinais, funding, ordem)
        met2 = metricas(ops2, curva2, dl2, amb2)
        met["custos_dobrados_fator_lucro"] = met2.get("fator_lucro")
        met["custos_dobrados_retorno_pct"] = met2.get("retorno_pct")
        met["veredito"] = veredito(met)
        resultados[var.codigo] = {"nome": var.nome, **met}
        with open(os.path.join(PASTA_SAIDA, f"operacoes_{var.codigo}.csv"), "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["moeda", "lado", "abertura_utc", "fechamento_utc", "motivo", "entrada", "saida", "liquido_usdt", "r"])
            for o in ops:
                w.writerow([o["moeda"], o["lado"],
                            datetime.fromtimestamp(o["abertura"] / 1000, tz=timezone.utc).strftime("%Y-%m-%d %H:%M"),
                            datetime.fromtimestamp(o["fechamento"] / 1000, tz=timezone.utc).strftime("%Y-%m-%d %H:%M"),
                            o["motivo"], f"{o['entrada']:.6g}", f"{o['saida']:.6g}",
                            f"{o['liquido']:.2f}".replace(".", ","), f"{o['r']:.2f}".replace(".", ",")])
        log(f"  {var.codigo}: {met.get('operacoes', 0)} operações | fator de lucro {met.get('fator_lucro')} | "
            f"retorno {met.get('retorno_pct')}% | queda máx. {met.get('queda_maxima_pct')}%")

    saida = {
        "gerado_em": datetime.now().isoformat(timespec="seconds"),
        "robo": caminho_bot, "versao_robo": getattr(robo, "VERSAO_ROBO", "?"),
        "periodo": [INICIO, FIM], "banca_inicial": BANCA_INICIAL,
        "moedas_com_dados": ordem, "moedas_sem_dados": faltando, "fontes": fontes,
        "premissas": [
            "Velas de 1h; decisão a cada fechamento de vela (o robô decide a cada 2s, mas os indicadores só mudam a cada vela).",
            "Entrada no fechamento da vela + slippage; taxa 0,05% por lado; funding real de cada moeda.",
            "Se a mesma vela tocou stop e alvo, considera o STOP primeiro (conservador).",
            "Limite diário e trava de ganho zeram à meia-noite de Brasília (o robô real exige apertar L no dia seguinte).",
            "Filtro de eventos (CPI/FOMC) NÃO simulado: não há calendário histórico no robô.",
            "Moedas listadas depois do início entram quando passam a ter dados.",
        ],
        "resultados": resultados,
    }
    with open(os.path.join(PASTA_SAIDA, "resultado_backtest.json"), "w", encoding="utf-8") as f:
        json.dump(saida, f, ensure_ascii=False, indent=2, default=str)

    with open(os.path.join(PASTA_SAIDA, "resultado_backtest.txt"), "w", encoding="utf-8") as f:
        f.write(f"BACKTEST DO ROBÔ DE FUTUROS — {INICIO} a {FIM} — banca inicial ${BANCA_INICIAL:,.0f}\n")
        f.write(f"Gerado em {saida['gerado_em']} a partir de {caminho_bot}\n\n")
        f.write("Régua da Demo: fator de lucro acima de 1,2, já com taxas.\n\n")
        for cod, r in resultados.items():
            f.write(f"[{cod}] {r['nome']}\n")
            f.write(f"    Operações: {r.get('operacoes', 0)} ({r.get('operacoes_por_mes', 0)}/mês) | acerto {r.get('taxa_acerto_pct')}%\n")
            f.write(f"    Fator de lucro: {r.get('fator_lucro')} (1º ano {r.get('fator_lucro_1a_metade')}, 2º ano {r.get('fator_lucro_2a_metade')})\n")
            f.write(f"    Resultado: ${r.get('resultado_usdt')} ({r.get('retorno_pct')}%, média {r.get('retorno_medio_mes_pct')}%/mês) | queda máxima {r.get('queda_maxima_pct')}%\n")
            f.write(f"    Taxas pagas: ${r.get('taxas_pagas_usdt')} | funding: ${r.get('funding_usdt')} | meses positivos {r.get('meses_positivos')}/{r.get('meses_total')}\n")
            f.write(f"    Com custos dobrados: fator de lucro {r.get('custos_dobrados_fator_lucro')}, retorno {r.get('custos_dobrados_retorno_pct')}%\n")
            f.write(f"    VEREDITO: {r.get('veredito')}\n\n")
        if faltando:
            f.write(f"Moedas sem dados (ficaram de fora): {', '.join(faltando)}\n")
    log("\nPronto! Arquivos gerados nesta pasta:")
    log("  resultado_backtest.txt   <- resumo para ler")
    log("  resultado_backtest.json  <- envie este para o Claude Code")
    log("  operacoes_*.csv          <- todas as operações de cada variante")


if __name__ == "__main__":
    main()
