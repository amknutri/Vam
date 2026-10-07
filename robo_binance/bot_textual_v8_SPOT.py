# ============================================================
# VERSÃO 8 (v8 Spot) — refinamentos a partir da pesquisa de 07/10/2026
# Marcado no código com "[v8]". Resumo:
#  1. Filtro de custo: só entra se o alvo valer >= 3x (taxas + slippage).
#  2. Só moedas líquidas: volume 24h mínimo de 20 milhões de USDT.
#  3. Freio: 3 stops em 6h => pausa de 6h; após stop, pausa de 4h na moeda.
#  4. Proteger o lucro: ao ganhar 1x o risco, o stop vai para a entrada (+custos).
#  5. Gráfico maior: indicadores em velas de 1h (antes 15m); tendência em 4h (antes 1h).
#  6. heartbeat.json gravado a cada ciclo (prova de que o robô está vivo).
#  9. Filtro de eventos macro (eventos.json): não abre posição perto de CPI/FOMC.
# 10. [Futuros] Não abre se o funding estiver caro contra a posição.
# 11. Confirmação por volume: opção DESLIGADA (FILTRO_VOLUME_ATIVO) até o backtest provar valor.
# ============================================================
# VERSÃO v7.1: preço REAL de execução — depois de abrir/fechar, o robô consulta
# a Binance para saber o preço em que a ordem realmente executou e só então calcula
# stop e alvo (antes usava o preço da tela quando a resposta vinha sem esse dado).
# Mudanças marcadas com "[v7.1]".
# ============================================================
# VERSÃO 7 (SPOT) — adaptação para o mercado SPOT da Binance
# Motivo: a Binance não oferece Futuros/derivativos para residentes no
# Brasil (restrição da CVM). A v7 opera no Spot, que é permitido.
# Mudanças em relação à v6 (marcadas no código com "[v7 SPOT]"):
#  1. Só COMPRA (no Spot não existe venda a descoberto). Sinais de venda
#     são ignorados nas 3 estratégias e no backtest.
#  2. Sem alavancagem (1x) — não existe liquidação.
#  3. Stop registrado no Spot da Binance (STOP_LOSS, ou STOP_LOSS_LIMIT
#     quando a moeda não aceita STOP_LOSS). O alvo (take profit) é vigiado
#     pelo robô, porque no Spot o stop já "prende" as moedas.
#  4. Fechamento: cancela o stop, vende o saldo REAL da moeda; se a venda
#     falhar, recoloca o stop na hora.
#  5. Posições salvas em disco (posicoes_spot.json): no Spot não existe
#     "posição" na corretora, então o robô guarda o que comprou.
#  6. Reconciliação pelo saldo das moedas + recoloca stop que sumiu.
#  7. Taxa 0,10% (Spot) nos cálculos e no backtest.
#  8. Funding Arb, Cash-Carry, Multi-Moedas e Perfil de alavancagem
#     desligados (dependiam de Futuros).
# ============================================================
# ============================================================
# VERSÃO 6 (CORRIGIDA) — revisão de segurança e gestão de risco
# Todas as mudanças estão marcadas no código com o comentário "[v6]".
# Resumo:
#  1. Chaves da Binance e do Telegram saíram do código (variáveis de ambiente).
#  2. STOP registrado NA BINANCE a cada posição aberta (protege mesmo com o
#     PC desligado ou sem internet). Se o stop não puder ser registrado, a
#     posição é fechada na hora (EXIGIR_STOP_NA_EXCHANGE).
#  3. Fechamento só conta quando a exchange confirma — falhou, a posição
#     continua vigiada e o bot tenta de novo (antes ele "esquecia" a posição
#     e gravava um lucro/prejuízo que não aconteceu).
#  4. Tecla M bloqueada com posições abertas (antes desligava a vigilância).
#  5. Limite diário zera sozinho à meia-noite, fica salvo em disco (reiniciar
#     o bot não apaga a perda do dia) e considera as posições ainda abertas.
#  6. Setup $200 limitado a SETUP200_MAX_POSICOES_SIMULTANEAS e ao saldo livre.
#  7. Mínimo da exchange não pode mais estourar o risco planejado.
#  8. Alavancagem máxima 5x (antes até 10x); risco por trade máx. 2%.
#  9. Preço de entrada/saída = preço REAL executado pela Binance.
# 10. Indicadores calculados só com velas FECHADAS (igual ao backtest).
# 11. Backtest conservador: alvo e stop no mesmo candle conta como STOP.
# 12. Funding Arb: perna vendida travada em 1x (evita liquidação).
# ============================================================
import os
import csv
import json
import time
import asyncio
import statistics
from datetime import datetime, date

import ccxt.async_support as ccxt
from rich.table import Table as RichTable
from textual.app import App, ComposeResult
from textual.screen import Screen
from textual.containers import Horizontal, Vertical, VerticalScroll
from textual.widgets import Header, Footer, DataTable, Static, RichLog, Input, Label, Button, ProgressBar

# openpyxl é usado só pra gerar a planilha .xlsx organizada do journal — é uma
# funcionalidade "bônus" (o CSV é o que realmente importa pro funcionamento
# do bot), então importamos com cuidado: se não estiver instalado, o bot
# continua funcionando 100% normal, só sem gerar o .xlsx (avisa uma vez no log).
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import CellIsRule, DataBarRule
    from openpyxl.chart import BarChart, Reference
    OPENPYXL_DISPONIVEL = True
except ImportError:
    OPENPYXL_DISPONIVEL = False

# aiohttp é usado só pra mandar as notificações pro Telegram — mesma lógica
# de segurança do openpyxl acima: se não estiver instalado, o bot funciona
# 100% normal, só sem mandar mensagem pro grupo (avisa uma vez no log).
try:
    import aiohttp
    AIOHTTP_DISPONIVEL = True
except ImportError:
    AIOHTTP_DISPONIVEL = False

# ============================================================
# CONFIGURAÇÃO
# ============================================================
# [v6 SEGURANÇA] As chaves NUNCA ficam escritas neste arquivo. Elas são lidas
# de VARIÁVEIS DE AMBIENTE (um "cofre" do sistema operacional, fora do código).
# No Windows (PowerShell), uma única vez:
#   setx BINANCE_API_KEY "sua_chave"
#   setx BINANCE_API_SECRET "seu_segredo"
# Depois feche e abra o terminal de novo. Na Binance, crie a chave SEM
# permissão de saque e com restrição de IP.
API_KEY = os.environ.get('BINANCE_API_KEY', "")
SECRET_KEY = os.environ.get('BINANCE_API_SECRET', "")

MOEDAS = [
    'BTC/USDT', 'ETH/USDT', 'SOL/USDT', 'BNB/USDT', 'XRP/USDT',
    'ADA/USDT', 'AVAX/USDT', 'DOGE/USDT', 'DOT/USDT', 'LINK/USDT',
    'NEAR/USDT', 'SUI/USDT', 'APT/USDT', 'LTC/USDT', 'TRX/USDT',
    'POL/USDT', 'UNI/USDT', 'ATOM/USDT', 'RENDER/USDT', 'ICP/USDT',
    'ETC/USDT', 'FIL/USDT', 'ARB/USDT',
    # --- Adicionadas depois, sugeridas como moedas líquidas ainda fora da lista ---
    'TON/USDT', 'SHIB/USDT', 'INJ/USDT', 'SEI/USDT', 'TIA/USDT',
    'HBAR/USDT', 'AAVE/USDT', 'MKR/USDT', 'CRV/USDT', 'LDO/USDT',
    'BCH/USDT', 'XLM/USDT', 'EOS/USDT',
]

# --- Taxas reais da Binance USDⓈ-M Futures (nível base, VIP 0, sem BNB) ---
# https://www.binance.com/en/fee/futureFee — ordens a mercado sempre entram como TAKER.
TAKER_FEE_PCT = 0.10  # [v7 SPOT] taxa do Spot (VIP 0, sem BNB). Era 0.05 (Futuros).
MAKER_FEE_PCT = 0.02  # não usado (o bot só envia ordens a mercado), mantido de referência

# --- Taxa Taker do mercado SPOT da Binance (nível base, VIP 0, sem BNB) ---
# Diferente da taxa de Futuros acima — o Spot cobra mais. Usada na simulação
# de Funding Rate Arbitrage, que opera nas duas pontas (Spot + Futuros).
SPOT_TAKER_FEE_PCT = 0.10
# No nível base do Spot, a taxa Maker é IGUAL à Taker (0,10% os dois) — só
# quem tem BNB pra pagar taxa ou volume alto (VIP1+) consegue desconto no
# Spot. Por isso, usar ordens Maker só reduz custo de verdade no lado
# Futuros (0,05% Taker -> 0,02% Maker). Mantido separado pra ficar explícito.
SPOT_MAKER_FEE_PCT = 0.10

# --- Benchmark sem risco no Brasil, usado como referência de comparação ---
# Confira o valor atual antes de confiar demais nele — muda com decisão do
# Copom. Valor de referência usado neste projeto: ~13,9% a.a. (CDI, 09/2026).
CDI_REFERENCIA_PCT = 13.9

RSI_PERIODO = 14
RSI_TIMEFRAME = '1h'  # [v8] antes 15m: menos ruído e menos taxas
RSI_RECALC_SEGUNDOS = 60  # [v8]
RECONCILIACAO_SEGUNDOS = 30
VOLUME_MINIMO = 20_000_000  # [v8] só moedas líquidas (antes 1 milhão)
MAX_POSICOES_SIMULTANEAS = 5
CICLO_SEGUNDOS = 2.0

# --- [v8] Refinamentos (pesquisa de 07/10/2026) ---
VERSAO_ROBO = "v8 Spot"
EH_FUTUROS = False
VARIACAO_JANELA_VELAS = 1        # variação "curta" = última vela de 1h
SLIPPAGE_ESTIMADO_PCT = 0.05     # diferença média preço da tela x executado, por ordem
CUSTO_MULTIPLO_MINIMO = 3.0      # o alvo precisa valer pelo menos 3x (taxas + slippage)
FREIO_STOPS_SEGUIDOS = 3         # 3 perdas seguidas...
FREIO_JANELA_HORAS = 6           # ...dentro de 6 horas...
FREIO_PAUSA_HORAS = 6            # ...pausam novas entradas por 6 horas
COOLDOWN_MOEDA_HORAS = 4         # após perda numa moeda, não volta nela por 4 horas
BREAKEVEN_EM_R = 1.0             # ao ganhar 1x o risco, stop vai para a entrada (+custos)
CAMINHO_HEARTBEAT = "heartbeat.json"
CAMINHO_EVENTOS = "eventos.json"  # mantido pelo Claude Code (Drive) e copiado pelo Cowork
EVENTOS_JANELA_MIN_PADRAO = 60   # sem novas entradas de 60 min antes a 60 min depois
FUNDING_MAX_CONTRA_PCT = 0.03    # [Futuros] funding (por 8h) contra a posição acima disso: não entra
FILTRO_VOLUME_ATIVO = False      # [11] DESLIGADO até o backtest provar que ajuda
FILTRO_VOLUME_MULTIPLO = 1.5     # vela atual com volume >= 1,5x a média das 20 anteriores

TENDENCIA_TIMEFRAME = '4h'  # [v8] antes 1h
TENDENCIA_EMA_PERIODO = 50
TENDENCIA_RECALC_SEGUNDOS = 300  # [v8]

MAX_LONGS_SIMULTANEOS = 3
MAX_SHORTS_SIMULTANEOS = 3

# ============================================================
# MÚLTIPLAS ESTRATÉGIAS + DETECÇÃO DE REGIME DE MERCADO
# Sistema de regras (votação entre indicadores técnicos clássicos), não uma
# inteligência artificial. O bot mede se o mercado está em tendência ou
# lateralizado e escolhe quais estratégias fazem sentido, só operando quando
# várias concordam entre si (confluência).
# ============================================================
BB_PERIODO = 20
BB_DESVIOS = 2.0
MACD_RAPIDA = 12
MACD_LENTA = 26
MACD_SINAL = 9
LIMIAR_LATERAL_1H_PCT = 0.3
LIMIAR_MOMENTUM_1H_PCT = 0.5

ESTOCASTICO_PERIODO_K = 14
ESTOCASTICO_SUAVIZACAO_K = 3
ESTOCASTICO_PERIODO_D = 3

LARGURA_BB_LATERAL_PCT = 2.0
LARGURA_BB_TENDENCIA_PCT = 4.0

CONFLUENCIA_MINIMA_PADRAO = 2
CONFLUENCIA_MINIMA_BANCA_PEQUENA = 3
BANCA_PEQUENA_LIMITE_USDT = 200.0

# ============================================================
# GESTÃO DE BANCA PROFISSIONAL
# - Risco por trade como % da banca (fixed fractional position sizing)
# - Stop Loss dinâmico por ATR (Average True Range), não percentual fixo
# - Relação Risco:Retorno mínima de 1:2
# - Limite de risco total simultâneo (soma de todas as posições abertas)
# ============================================================
ATR_PERIODO = 14
SL_ATR_MULTIPLO = 1.5
RELACAO_RISCO_RETORNO = 2.0

RISCO_POR_TRADE_PCT = 1.0
LIMITE_PERDA_DIARIA_PCT = 3.0
DRAWDOWN_MAXIMO_PCT = 2.0
MAX_RISCO_TOTAL_PCT = 6.0
MAX_EXPOSICAO_POR_POSICAO_PCT = 20.0

# --- [v6] Travas de segurança novas ---
# Registra o stop (e o alvo) como ordem DENTRO da Binance a cada posição
# aberta. Se a Binance recusar o stop, a posição é fechada imediatamente
# (nunca fica posição aberta sem stop na exchange).
EXIGIR_STOP_NA_EXCHANGE = True
# Preço usado pela Binance para disparar o stop: MARK_PRICE (preço de
# referência) é mais resistente a "pavios" isolados do que o último negócio.
TIPO_PRECO_GATILHO_STOP = 'MARK_PRICE'
ALAVANCAGEM_MAXIMA_PERMITIDA = 1     # [v7 SPOT] Spot não tem alavancagem
# [v7 SPOT] posições do robô salvas em disco (Spot não tem 'posição' na corretora)
CAMINHO_ESTADO_POSICOES_SPOT = "posicoes_spot.json"
# [v7 SPOT] quando a moeda só aceita STOP_LOSS_LIMIT, a ordem de venda fica
# este % abaixo do gatilho (folga para conseguir vender numa queda rápida)
SPOT_STOP_LIMIT_FOLGA_PCT = 2.0
RISCO_POR_TRADE_MAXIMO_PCT = 2.0     # teto na tela de Configurações (antes 10%)
LIMITE_PERDA_DIARIA_MAXIMO_PCT = 5.0 # teto na tela de Configurações (antes 20%)
# Se o tamanho mínimo exigido pela Binance fizer o risco passar de 1,5x o
# planejado, a operação é PULADA (antes o bot aumentava o tamanho sozinho).
TOLERANCIA_RISCO_MINIMO_EXCHANGE = 1.5
# Estado do limite diário salvo em disco: reiniciar o bot não apaga a perda do dia.
CAMINHO_ESTADO_RISCO_DIARIO = "estado_risco_diario.json"

# --- Trava de GANHO diária, por modo (pedido explicitamente) ---
# Diferente do drawdown-do-pico acima (que só protege o que já foi ganho),
# essa trava é um TETO FIXO em dólar: assim que o lucro acumulado NAQUELE
# modo específico no dia bate esse valor, esse modo para de abrir posição
# nova até você apertar "C" (Zerar Lucro). Cada modo tem seu próprio
# contador (self.lucro_dia_por_sistema) — bater a trava de um não afeta
# os outros.
STOP_GANHO_DIARIO_CONFLUENCIA_USD = 10.0  # sistema "normal" (confluência de indicadores)

CAMINHO_JOURNAL = "trades_journal.csv"

# Planilha Excel organizada, gerada AUTOMATICAMENTE a cada trade fechado, a
# partir do mesmo trades_journal.csv -- assim o CSV continua existindo (é o
# formato mais simples e seguro de gravar, linha por linha, sem risco de
# corromper), mas quem for ABRIR pra olhar/analisar usa este .xlsx pronto,
# já formatado, com aba de resumo (win rate, PnL por moeda, por motivo etc.).
CAMINHO_JOURNAL_XLSX = "trades_journal.xlsx"

# ============================================================
# NOTIFICAÇÕES NO TELEGRAM — manda mensagem no seu grupo toda vez que uma
# posição abre ou fecha, automaticamente. Pra desligar sem apagar as
# credenciais, é só trocar TELEGRAM_NOTIFICACOES_ATIVAS pra False.
#
# ATENÇÃO DE SEGURANÇA: o TELEGRAM_BOT_TOKEN abaixo é como uma senha — quem
# tiver ele controla seu bot do Telegram (consegue mandar mensagem em seu
# nome nos chats onde o bot estiver). Se esse código for parar em algum
# lugar público (ex: GitHub), gere um token novo no @BotFather (comando
# /revoke) e cole o novo aqui.
# ============================================================
TELEGRAM_NOTIFICACOES_ATIVAS = True
# [v6 SEGURANÇA] Token e chat do Telegram também saíram do código:
#   setx TELEGRAM_BOT_TOKEN "token_novo_do_BotFather"
#   setx TELEGRAM_CHAT_ID "id_do_grupo"
TELEGRAM_BOT_TOKEN = os.environ.get('TELEGRAM_BOT_TOKEN', "")
TELEGRAM_CHAT_ID = os.environ.get('TELEGRAM_CHAT_ID', "")


def _telegram_motivo_amigavel(motivo: str) -> str:
    """Traduz o motivo técnico de fechamento (TAKE PROFIT, STOP LOSS etc.)
    pra uma frase que um iniciante entende de cara, sem precisar saber o
    que cada termo significa. Motivo não mapeado é mostrado como veio, pra
    nunca esconder informação nova que apareça no futuro."""
    m = (motivo or "").strip().upper()
    if m == "TAKE PROFIT":
        return "🎯 Bateu o alvo de lucro"
    if m == "STOP LOSS":
        return "🛑 Bateu o stop (limite de perda)"
    if m.startswith("TEMPO MAXIMO"):
        return f"⏱️ Tempo máximo atingido, fechou sozinho {motivo[motivo.find('('):]}" if "(" in motivo else "⏱️ Tempo máximo atingido, fechou sozinho"
    if "SAIU DO CANAL" in m:
        return "📉 Saiu do canal de tendência (sinal de reversão)"
    if m == "MANUAL":
        return "✋ Você fechou manualmente"
    if m in ("PÂNICO", "PANICO"):
        return "🚨 Botão de pânico — fechou tudo por segurança"
    return motivo


def _telegram_sistema_amigavel(sistema: str) -> str:
    """Nome do sistema em português claro, pra quem não decorou os nomes
    internos do bot (confluencia, setup200, donchian_adx)."""
    nomes = {
        "confluencia": "Confluência",
        "setup200": "Setup $200",
        "donchian_adx": "Donchian+ADX",
    }
    return nomes.get(sistema, sistema or "")

# --- Configuração do Funding Arbitrage AO VIVO (execução real, não backtest) ---
CAMINHO_JOURNAL_FUNDING_ARB = "funding_arb_journal.csv"
CAMINHO_ESTADO_FUNDING_ARB = "funding_arb_state.json"
ALOCACAO_ARB_PCT_BANCA = 10.0   # % da banca livre destinada a UMA posição de arb
JANELA_MM_FUNDING_ARB_PADRAO = 9  # mesmo padrão já validado no backtest (9 eventos = 3 dias)
INTERVALO_MONITOR_FUNDING_ARB_SEGUNDOS = 300  # checa a cada 5 min por novo evento de funding

# --- INTERRUPTOR DE SEGURANÇA: só mude para True quando quiser operar com
# DINHEIRO REAL. Enquanto for False, todas as ordens vão para o Demo Trading. ---
MODO_REAL = False

# ============================================================
# MODO "SETUP $200" — banca dividida em fatias fixas por moeda
# Estratégia SEPARADA do sistema de confluência acima (própria tecla,
# próprio dicionário de posições, nunca se mistura com posicoes_multiplas).
# Usa Bollinger Bands + RSI (reversão à média) — mesma lógica já testada e
# validada com dados fabricados no arquivo teste_bollinger_rsi.py.
# Diferente do resto do bot (que usa % da banca + SL dinâmico por ATR),
# aqui o tamanho da posição e o TP/SL são FIXOS EM DÓLAR, como pedido
# explicitamente.
# Reaproveita os indicadores de Bollinger/RSI que o bot JÁ calcula no
# timeframe de 15 minutos (RSI_TIMEFRAME, recalculado a cada
# RSI_RECALC_SEGUNDOS) em vez de abrir um segundo ciclo de busca de velas
# à parte — mais simples e sem carga extra na API. O ideal pedido era 5
# minutos; usar os mesmos 15 minutos que o bot já busca fica dentro do
# limite de "no máximo 15 minutos" e evita duplicar infraestrutura.
# ============================================================
SETUP200_VALOR_USD = 200.0
SETUP200_TP_USD = 3.50
SETUP200_SL_USD = 3.50              # 1:1 com o TP — mudado de 2x (que exigia 66,7% de acerto pra
                                     # empatar) porque não temos evidência de que essa entrada acerte
                                     # mais que uma moeda ao ar; 1:1 exige só 50% pra empatar.
SETUP200_RSI_SOBREVENDIDO = 30
SETUP200_RSI_SOBRECOMPRADO = 70
SETUP200_TEMPO_MAXIMO_MINUTOS = 240  # [v8] velas de 1h; antes 60  # fecha à força depois de 1h, como pedido
SETUP200_ALAVANCAGEM = 1            # travada em 1x por segurança, independente do perfil de risco geral
SETUP200_MAX_POSICOES_SIMULTANEAS = 3  # [v6] antes não tinha limite (podia abrir 36 de uma vez numa queda geral)
STOP_GANHO_DIARIO_SETUP200_USD = 3.50  # trava de ganho diária do Setup $200 (pedido explicitamente)
# Trava de PERDA diária do Setup $200 (pedido explicitamente: "perdeu uma,
# para também") — igualada ao stop-loss de 1 trade só (SETUP200_SL_USD),
# pra ficar simétrico com a trava de ganho (que é igual a 1 take-profit).
STOP_PERDA_DIARIA_SETUP200_USD = SETUP200_SL_USD

# Tabelas usadas por fechar_posicao() pra saber, genericamente, se algum
# modo bateu a própria trava de ganho OU de perda do dia — adicionar um
# modo novo aqui é só incluir mais uma linha, sem precisar mexer na lógica
# de fechar_posicao.
STOP_GANHO_DIARIO_POR_SISTEMA = {
    "confluencia": STOP_GANHO_DIARIO_CONFLUENCIA_USD,
    "setup200": STOP_GANHO_DIARIO_SETUP200_USD,
}
STOP_PERDA_DIARIA_POR_SISTEMA = {
    "setup200": STOP_PERDA_DIARIA_SETUP200_USD,
}

# ============================================================
# MODO "DONCHIAN + ADX (BTC)" — seguidor de tendência
# Estratégia validada separadamente em backtest_donchian_adx.py (e nas
# variações _bullrun.py / _lateral.py), testada em 3 regimes históricos
# de mercado (queda, lateral e bull run) SEM reajustar parâmetros entre um
# teste e outro — exatamente para evitar overfitting por regime. Ideia
# central: só entra quando o preço rompe um canal de Donchian E o ADX
# confirma que existe tendência de verdade (ADX > 25); a saída é por um
# canal de Donchian mais LARGO (deixa o lucro correr), não por um alvo
# fixo — essa foi a lição mais cara de todo o processo de validação: um
# take-profit fixo pequeno destrói o resultado numa tendência forte,
# porque corta o lucro bem antes da hora.
# Estratégia SEPARADA do sistema de confluência e do Setup $200 (própria
# tecla, próprio dicionário de posições, nunca se mistura com as outras).
# Só opera BTC/USDT (única moeda com padrão consistente nos 3 regimes
# testados — ETH/SOL/BNB pioraram ou dependeram de um único evento
# histórico não repetível, então foram deixadas de fora por enquanto).
# ============================================================
DONCHIAN_ADX_SIMBOLO = "BTC/USDT"
DONCHIAN_ADX_TIMEFRAME = "4h"      # mesmo timeframe usado na validação
DONCHIAN_ADX_ENTRADA = 20          # canal de ENTRADA (rompimento das últimas 20 velas de 4h)
DONCHIAN_ADX_SAIDA = 30            # canal de SAÍDA, mais largo (deixa o lucro correr)
DONCHIAN_ADX_PERIODO = 14          # período do ADX (padrão de Wilder)
DONCHIAN_ADX_MINIMO = 25.0         # só entra se ADX > 25 (mercado em tendência de verdade)
DONCHIAN_ADX_SL_PCT = 1.5          # stop-loss inicial fixo (%) — proteção, não é o alvo de saída
DONCHIAN_ADX_FRACAO_BANCA_PCT = 12.5  # % da banca LIVRE usada em cada posição, igual ao backtest
DONCHIAN_ADX_RECALC_SEGUNDOS = 900    # recalcula canais/ADX a cada 15min (velas de 4h mudam devagar)
DONCHIAN_ADX_ALAVANCAGEM = 1          # travada em 1x — estratégia validada SEM alavancagem


# ============================================================
# FUNÇÕES DE CÁLCULO (puras — mesma matemática já testada na versão anterior)
# ============================================================
def _ema_serie(valores, periodo):
    if len(valores) < periodo:
        return [None] * len(valores)
    k = 2 / (periodo + 1)
    serie = [None] * (periodo - 1)
    ema = sum(valores[:periodo]) / periodo
    serie.append(ema)
    for v in valores[periodo:]:
        ema = v * k + ema * (1 - k)
        serie.append(ema)
    return serie


def calcular_ema(closes, periodo):
    """Média móvel exponencial simples, usada para o filtro de tendência."""
    if len(closes) < periodo:
        return None
    k = 2 / (periodo + 1)
    ema = sum(closes[:periodo]) / periodo
    for preco in closes[periodo:]:
        ema = preco * k + ema * (1 - k)
    return ema


def calcular_rsi(closes, periodo=RSI_PERIODO):
    """RSI real (Wilder), calculado a partir de uma lista de preços de fechamento."""
    if len(closes) < periodo + 1:
        return None
    ganhos, perdas = [], []
    for i in range(1, len(closes)):
        delta = closes[i] - closes[i - 1]
        ganhos.append(max(delta, 0))
        perdas.append(max(-delta, 0))
    media_ganho = sum(ganhos[:periodo]) / periodo
    media_perda = sum(perdas[:periodo]) / periodo
    for i in range(periodo, len(ganhos)):
        media_ganho = (media_ganho * (periodo - 1) + ganhos[i]) / periodo
        media_perda = (media_perda * (periodo - 1) + perdas[i]) / periodo
    if media_perda == 0:
        return 100.0
    rs = media_ganho / media_perda
    return round(100 - (100 / (1 + rs)), 1)


def calcular_atr(ohlcv, periodo=ATR_PERIODO):
    """ATR (Average True Range, método de Wilder) — mede a volatilidade média
    do ativo em valor absoluto de preço, usado pra definir stops adaptativos."""
    if len(ohlcv) < periodo + 1:
        return None
    true_ranges = []
    for i in range(1, len(ohlcv)):
        high, low = ohlcv[i][2], ohlcv[i][3]
        close_anterior = ohlcv[i - 1][4]
        tr = max(high - low, abs(high - close_anterior), abs(low - close_anterior))
        true_ranges.append(tr)
    atr = sum(true_ranges[:periodo]) / periodo
    for tr in true_ranges[periodo:]:
        atr = (atr * (periodo - 1) + tr) / periodo
    return atr


def calcular_macd(closes, rapida=MACD_RAPIDA, lenta=MACD_LENTA, sinal=MACD_SINAL):
    """MACD clássico — detecta o CRUZAMENTO (evento), não só o estado atual."""
    if len(closes) < lenta + sinal + 1:
        return None
    ema_rapida = _ema_serie(closes, rapida)
    ema_lenta = _ema_serie(closes, lenta)
    macd_validos = [
        ema_rapida[i] - ema_lenta[i] for i in range(len(closes))
        if ema_rapida[i] is not None and ema_lenta[i] is not None
    ]
    if len(macd_validos) < sinal + 1:
        return None
    sinal_serie = _ema_serie(macd_validos, sinal)
    if len(sinal_serie) < 2 or sinal_serie[-1] is None or sinal_serie[-2] is None:
        return None
    macd_atual, macd_anterior = macd_validos[-1], macd_validos[-2]
    sinal_atual, sinal_anterior = sinal_serie[-1], sinal_serie[-2]
    hist_atual = macd_atual - sinal_atual
    hist_anterior = macd_anterior - sinal_anterior
    return {
        "macd": macd_atual, "sinal": sinal_atual, "histograma": hist_atual,
        "cruzou_alta": hist_anterior <= 0 and hist_atual > 0,
        "cruzou_baixa": hist_anterior >= 0 and hist_atual < 0,
    }


def calcular_bollinger(closes, periodo=BB_PERIODO, desvios=BB_DESVIOS):
    """Bandas de Bollinger — a largura indica se o mercado está lateral ou volátil."""
    if len(closes) < periodo:
        return None
    janela = closes[-periodo:]
    media = sum(janela) / periodo
    variancia = sum((c - media) ** 2 for c in janela) / periodo
    desvio_padrao = variancia ** 0.5
    superior = media + desvios * desvio_padrao
    inferior = media - desvios * desvio_padrao
    largura_pct = ((superior - inferior) / media * 100) if media > 0 else 0.0
    return {"superior": superior, "media": media, "inferior": inferior, "largura_pct": largura_pct}


def calcular_estocastico(ohlcv, periodo_k=ESTOCASTICO_PERIODO_K, suavizacao_k=ESTOCASTICO_SUAVIZACAO_K, periodo_d=ESTOCASTICO_PERIODO_D):
    """Oscilador Estocástico — cruzamento de %K sobre %D em zonas extremas."""
    minimo_necessario = periodo_k + suavizacao_k + periodo_d
    if len(ohlcv) < minimo_necessario:
        return None

    def k_bruto(indice):
        janela = ohlcv[indice - periodo_k + 1: indice + 1]
        maior = max(c[2] for c in janela)
        menor = min(c[3] for c in janela)
        fechamento = ohlcv[indice][4]
        if maior == menor:
            return 50.0
        return (fechamento - menor) / (maior - menor) * 100

    k_brutos = [k_bruto(i) for i in range(periodo_k - 1, len(ohlcv))]
    k_suave = [
        sum(k_brutos[i - suavizacao_k + 1:i + 1]) / suavizacao_k
        for i in range(suavizacao_k - 1, len(k_brutos))
    ]
    if len(k_suave) < periodo_d + 1:
        return None
    d_serie = [
        sum(k_suave[i - periodo_d + 1:i + 1]) / periodo_d
        for i in range(periodo_d - 1, len(k_suave))
    ]
    if len(d_serie) < 2 or len(k_suave) < 2:
        return None
    k_atual, k_anterior = k_suave[-1], k_suave[-2]
    d_atual, d_anterior = d_serie[-1], d_serie[-2]
    return {
        "k": k_atual, "d": d_atual,
        "cruzou_alta": k_anterior <= d_anterior and k_atual > d_atual and k_atual < 30,
        "cruzou_baixa": k_anterior >= d_anterior and k_atual < d_atual and k_atual > 70,
    }


def detectar_padrao_candle(ohlcv):
    """Detecta Engolfo, Martelo e Estrela Cadente na última vela."""
    if len(ohlcv) < 2:
        return None, None
    anterior, atual = ohlcv[-2], ohlcv[-1]
    o1, c1 = anterior[1], anterior[4]
    o2, h2, l2, c2 = atual[1], atual[2], atual[3], atual[4]
    corpo = abs(c2 - o2)
    sombra_inferior = min(o2, c2) - l2
    sombra_superior = h2 - max(o2, c2)

    if c1 < o1 and c2 > o2 and o2 <= c1 and c2 >= o1:
        return "COMPRA", "Engolfo de Alta"
    if c1 > o1 and c2 < o2 and o2 >= c1 and c2 <= o1:
        return "VENDA", "Engolfo de Baixa"
    if corpo > 0:
        if sombra_inferior >= corpo * 2 and sombra_superior <= corpo * 0.5:
            return "COMPRA", "Martelo"
        if sombra_superior >= corpo * 2 and sombra_inferior <= corpo * 0.5:
            return "VENDA", "Estrela Cadente"
    return None, None


def calcular_regime(bollinger):
    """Classifica o momento do mercado: TENDENCIA, LATERAL ou MISTO."""
    if bollinger is None:
        return "MISTO"
    largura = bollinger["largura_pct"]
    if largura < LARGURA_BB_LATERAL_PCT:
        return "LATERAL"
    elif largura >= LARGURA_BB_TENDENCIA_PCT:
        return "TENDENCIA"
    return "MISTO"


def calcular_donchian_canais(ohlcv, periodo_entrada=DONCHIAN_ADX_ENTRADA, periodo_saida=DONCHIAN_ADX_SAIDA):
    """Canais de Donchian: a maior máxima e a menor mínima das últimas N
    velas. Usa dois tamanhos de canal — um mais estreito pra ENTRADA
    (rompimento) e um mais largo pra SAÍDA (deixa o lucro correr até o
    preço voltar de forma mais consistente, não no primeiro solavanco).
    IMPORTANTE: a vela mais recente (ohlcv[-1], ainda em formação ou recém
    fechada) é excluída do cálculo — olhamos só para as velas ANTERIORES,
    senão estaríamos comparando o preço atual com um canal que já inclui
    o próprio preço atual (viés de olhar pro futuro / look-ahead bias)."""
    periodo_max = max(periodo_entrada, periodo_saida)
    if len(ohlcv) < periodo_max + 1:
        return None
    velas_entrada = ohlcv[-(periodo_entrada + 1):-1]
    velas_saida = ohlcv[-(periodo_saida + 1):-1]
    return {
        "maxima_entrada": max(c[2] for c in velas_entrada),
        "minima_entrada": min(c[3] for c in velas_entrada),
        "maxima_saida": max(c[2] for c in velas_saida),
        "minima_saida": min(c[3] for c in velas_saida),
    }


def calcular_adx(ohlcv, periodo=DONCHIAN_ADX_PERIODO):
    """ADX (Average Directional Index, método de Wilder) — mede a FORÇA de
    uma tendência (não a direção dela). Valores acima de ~25 costumam
    indicar tendência definida; abaixo disso, mercado sem direção clara
    (lateral). Usado aqui como um "portão": só deixa a estratégia de
    Donchian operar quando existe uma tendência de verdade pra seguir —
    mesma implementação (Wilder) validada em backtest_donchian_adx.py."""
    if len(ohlcv) < periodo * 2 + 1:
        return None
    plus_dm, minus_dm, trs = [], [], []
    for i in range(1, len(ohlcv)):
        high, low = ohlcv[i][2], ohlcv[i][3]
        high_ant, low_ant, close_ant = ohlcv[i - 1][2], ohlcv[i - 1][3], ohlcv[i - 1][4]
        up_move = high - high_ant
        down_move = low_ant - low
        plus_dm.append(up_move if (up_move > down_move and up_move > 0) else 0.0)
        minus_dm.append(down_move if (down_move > up_move and down_move > 0) else 0.0)
        trs.append(max(high - low, abs(high - close_ant), abs(low - close_ant)))

    def _suavizacao_wilder(valores, periodo):
        if len(valores) < periodo:
            return []
        suavizado = [sum(valores[:periodo])]
        for v in valores[periodo:]:
            suavizado.append(suavizado[-1] - (suavizado[-1] / periodo) + v)
        return suavizado

    tr_suave = _suavizacao_wilder(trs, periodo)
    plus_dm_suave = _suavizacao_wilder(plus_dm, periodo)
    minus_dm_suave = _suavizacao_wilder(minus_dm, periodo)
    if not tr_suave or not plus_dm_suave or not minus_dm_suave:
        return None

    dx_serie = []
    for tr_s, pdm_s, mdm_s in zip(tr_suave, plus_dm_suave, minus_dm_suave):
        if tr_s == 0:
            dx_serie.append(0.0)
            continue
        plus_di = 100 * pdm_s / tr_s
        minus_di = 100 * mdm_s / tr_s
        soma_di = plus_di + minus_di
        dx_serie.append(100 * abs(plus_di - minus_di) / soma_di if soma_di > 0 else 0.0)

    if len(dx_serie) < periodo:
        return None
    adx = sum(dx_serie[:periodo]) / periodo
    for dx in dx_serie[periodo:]:
        adx = (adx * (periodo - 1) + dx) / periodo
    return adx


# ============================================================
# JOURNAL EM PLANILHA (.xlsx) — gerado automaticamente a cada trade fechado
# ============================================================
# O trades_journal.csv já existia e continua existindo (é o formato mais
# simples e seguro pra GRAVAR, linha por linha, sem risco de corromper nada
# se o bot cair no meio). O que faltava era um jeito de LER isso de forma
# organizada — daí esta função: lê o CSV inteiro (que mistura dois formatos
# históricos diferentes, do bot evoluindo ao longo do tempo) e regenera do
# zero uma planilha .xlsx já formatada, pronta pra abrir.
_JOURNAL_COLUNAS_NOVAS = [
    'data_hora_abertura', 'data_hora_fechamento', 'duracao_min',
    'ativo', 'tipo', 'motivo_fechamento', 'resultado',
    'preco_entrada', 'preco_saida', 'quantidade', 'alavancagem', 'valor_nocional_usdt',
    'risco_planejado_usdt', 'pnl_bruto_usdt', 'taxa_abertura_usdt', 'taxa_fechamento_usdt',
    'pnl_liquido_usdt', 'r_multiplo', 'banca_no_momento_usdt', 'lucro_acumulado_dia_usdt',
]


def _ler_registros_journal(caminho_csv: str) -> list:
    """Lê o trades_journal.csv e normaliza TODAS as linhas pro mesmo formato
    em memória, não importa se a linha é do formato novo (20 colunas) ou de
    um formato antigo/menor que o bot já teve (ex.: 5 colunas, sem preço,
    taxa, R múltiplo etc.) — os campos que aquele formato não tinha ficam
    em branco, e a linha é marcada como 'Resumido' pra deixar isso visível."""
    if not os.path.exists(caminho_csv):
        return []

    def _num(v):
        try:
            return float(v) if v not in (None, "") else None
        except ValueError:
            return None

    registros = []
    with open(caminho_csv, encoding='utf-8-sig') as f:
        linhas = list(csv.reader(f))
    if not linhas:
        return []

    for linha in linhas[1:]:  # pula o cabeçalho
        if len(linha) == len(_JOURNAL_COLUNAS_NOVAS):
            (abertura, fechamento, duracao, ativo, tipo, motivo, resultado,
             preco_entrada, preco_saida, qtd, alav, valor_nocional, risco,
             pnl_bruto, taxa_ab, taxa_fe, pnl_liq, r_mult, banca, lucro_dia) = linha
            registros.append({
                "abertura": abertura or None, "fechamento": fechamento or None,
                "duracao_min": _num(duracao), "ativo": ativo, "tipo": tipo,
                "motivo": motivo, "resultado": resultado,
                "preco_entrada": _num(preco_entrada), "preco_saida": _num(preco_saida),
                "quantidade": _num(qtd), "alavancagem": _num(alav),
                "valor_nocional": _num(valor_nocional), "risco_usd": _num(risco),
                "pnl_bruto": _num(pnl_bruto), "taxa_abertura": _num(taxa_ab),
                "taxa_fechamento": _num(taxa_fe), "pnl_liquido": _num(pnl_liq),
                "r_multiplo": _num(r_mult), "banca": _num(banca),
                "lucro_dia_acum": _num(lucro_dia), "formato": "Completo",
            })
        elif len(linha) == 5:
            fechamento, ativo, tipo, pnl, motivo = linha
            pnl_val = _num(pnl) or 0.0
            registros.append({
                "abertura": None, "fechamento": fechamento or None,
                "duracao_min": None, "ativo": ativo, "tipo": tipo, "motivo": motivo,
                "resultado": "WIN" if pnl_val >= 0 else "LOSS",
                "preco_entrada": None, "preco_saida": None, "quantidade": None,
                "alavancagem": None, "valor_nocional": None, "risco_usd": None,
                "pnl_bruto": None, "taxa_abertura": None, "taxa_fechamento": None,
                "pnl_liquido": pnl_val, "r_multiplo": None, "banca": None,
                "lucro_dia_acum": None, "formato": "Resumido (versão antiga do bot)",
            })
        # linhas com outro número de colunas (corrompidas/parciais) são
        # ignoradas silenciosamente na planilha bonita — o CSV original
        # continua intacto com elas, então nada se perde de verdade.

    registros.sort(key=lambda r: r["fechamento"] or "")
    return registros


def gerar_journal_xlsx(caminho_csv: str = CAMINHO_JOURNAL, caminho_xlsx: str = CAMINHO_JOURNAL_XLSX):
    """Regenera do zero o trades_journal.xlsx a partir do trades_journal.csv
    inteiro. É chamada toda vez que uma operação fecha, então o .xlsx está
    sempre atualizado e já formatado — não precisa abrir/editar nada manual."""
    if not OPENPYXL_DISPONIVEL:
        return
    registros = _ler_registros_journal(caminho_csv)
    if not registros:
        return

    FONTE = "Arial"
    wb = openpyxl.Workbook()

    # ---------------- ABA "Operações" ----------------
    ws = wb.active
    ws.title = "Operações"
    header_font = Font(name=FONTE, bold=True, color="FFFFFF", size=10)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    normal_font = Font(name=FONTE, size=10)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    colunas = [
        ("Abertura", 18), ("Fechamento", 18), ("Duração (min)", 13),
        ("Ativo", 9), ("Tipo", 8), ("Motivo Fechamento", 20), ("Resultado", 10),
        ("Preço Entrada", 13), ("Preço Saída", 13), ("Quantidade", 13),
        ("Alavancagem", 11), ("Valor Nocional (USD)", 16), ("Risco Planejado (USD)", 16),
        ("PnL Bruto (USD)", 13), ("Taxa Abertura (USD)", 14), ("Taxa Fechamento (USD)", 15),
        ("PnL Líquido (USD)", 14), ("Múltiplo R", 10), ("Banca no Momento (USD)", 16),
        ("Lucro Acum. do Dia (USD)", 17), ("Formato do Registro", 28),
    ]
    for i, (nome, largura) in enumerate(colunas, start=1):
        c = ws.cell(row=1, column=i, value=nome)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = largura
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    def _dt(s):
        try:
            return datetime.fromisoformat(s) if s else None
        except ValueError:
            return None

    for r_idx, reg in enumerate(registros, start=2):
        valores = [
            _dt(reg["abertura"]), _dt(reg["fechamento"]), reg["duracao_min"],
            reg["ativo"], reg["tipo"], reg["motivo"], reg["resultado"],
            reg["preco_entrada"], reg["preco_saida"], reg["quantidade"],
            reg["alavancagem"], reg["valor_nocional"], reg["risco_usd"],
            reg["pnl_bruto"], reg["taxa_abertura"], reg["taxa_fechamento"],
            reg["pnl_liquido"], reg["r_multiplo"], reg["banca"],
            reg["lucro_dia_acum"], reg["formato"],
        ]
        for c_idx, v in enumerate(valores, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=v)
            cell.font = normal_font
            cell.border = border
            cell.alignment = Alignment(horizontal="center")
            if c_idx in (1, 2):
                cell.number_format = "dd/mm/yyyy hh:mm:ss"
            elif c_idx in (8, 9):
                cell.number_format = "#,##0.000000"
            elif c_idx in (12, 13, 14, 15, 16, 17, 19, 20):
                cell.number_format = "$#,##0.00;[RED]-$#,##0.00"
            elif c_idx == 18:
                cell.number_format = '0.00"R"'
            elif c_idx == 21:
                cell.alignment = Alignment(horizontal="left")

    n = len(registros)
    verde_fill, verde_font = PatternFill("solid", fgColor="C6EFCE"), Font(name=FONTE, size=10, color="006100")
    vermelho_fill, vermelho_font = PatternFill("solid", fgColor="FFC7CE"), Font(name=FONTE, size=10, color="9C0006")
    ws.conditional_formatting.add(f"G2:G{n+1}", CellIsRule(operator="equal", formula=['"WIN"'], fill=verde_fill, font=verde_font))
    ws.conditional_formatting.add(f"G2:G{n+1}", CellIsRule(operator="equal", formula=['"LOSS"'], fill=vermelho_fill, font=vermelho_font))
    ws.conditional_formatting.add(f"Q2:Q{n+1}", CellIsRule(operator="greaterThanOrEqual", formula=["0"], font=Font(name=FONTE, size=10, color="006100")))
    ws.conditional_formatting.add(f"Q2:Q{n+1}", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=FONTE, size=10, color="9C0006")))
    ws.auto_filter.ref = f"A1:U{n+1}"

    # ---------------- ABA "Resumo" ----------------
    ws_r = wb.create_sheet("Resumo", 0)
    titulo_font = Font(name=FONTE, bold=True, size=18, color="1F4E78")
    subtitulo_font = Font(name=FONTE, italic=True, size=10, color="808080")
    label_font = Font(name=FONTE, bold=True, size=11)
    valor_font = Font(name=FONTE, size=11)
    aviso_fill = PatternFill("solid", fgColor="FFF2CC")
    aviso_font = Font(name=FONTE, italic=True, size=9, color="7F6000")
    secao_font = Font(name=FONTE, bold=True, size=13, color="1F4E78")

    ws_r["B2"] = "📊 Resumo do Journal de Operações"
    ws_r["B2"].font = titulo_font
    ws_r["B3"] = "Atualizado automaticamente pelo bot a cada operação fechada — não precisa mexer em nada."
    ws_r["B3"].font = subtitulo_font
    ws_r.merge_cells("B2:M2")
    ws_r.merge_cells("B3:M3")
    ws_r.row_dimensions[2].height = 26

    rng = lambda col: f"'Operações'!{col}2:{col}{n+1}"

    # ================= CARTÕES GRANDES (visão de 3 segundos) =================
    # 4 números que realmente importam, bem grandes, pra quem abrir a planilha
    # entender o resultado geral sem precisar ler nada mais.
    total_trades = n
    total_wins = sum(1 for r in registros if r["resultado"] == "WIN")
    win_rate_val = (total_wins / total_trades) if total_trades else 0
    pnl_total_val = sum((r["pnl_liquido"] or 0.0) for r in registros)
    soma_ganhos = sum(r["pnl_liquido"] for r in registros if r["resultado"] == "WIN" and r["pnl_liquido"] is not None)
    soma_perdas = sum(r["pnl_liquido"] for r in registros if r["resultado"] == "LOSS" and r["pnl_liquido"] is not None)
    fator_lucro_val = (soma_ganhos / abs(soma_perdas)) if soma_perdas else None

    cor_pnl_card = "2E7D32" if pnl_total_val >= 0 else "C62828"  # verde se positivo, vermelho se negativo
    fill_pnl_card = "E8F5E9" if pnl_total_val >= 0 else "FDECEA"
    cor_fl_card = "2E7D32" if (fator_lucro_val or 0) >= 1 else "C62828"
    fill_fl_card = "E8F5E9" if (fator_lucro_val or 0) >= 1 else "FDECEA"

    cartoes = [
        ("TOTAL DE OPERAÇÕES", f"=COUNTA({rng('G')})", "0", "1F4E78", "E8EEF7", "B", "C"),
        ("TAXA DE ACERTO (WIN RATE)", f"=IFERROR(COUNTIF({rng('G')},\"WIN\")/COUNTA({rng('G')}),\"\")", "0.0%", "1F4E78", "E8EEF7", "D", "E"),
        ("PNL LÍQUIDO TOTAL", f"=SUM({rng('Q')})", "$#,##0.00;[RED]-$#,##0.00", cor_pnl_card, fill_pnl_card, "F", "G"),
        ("FATOR DE LUCRO", f'=IFERROR(SUMIF({rng("G")},"WIN",{rng("Q")})/ABS(SUMIF({rng("G")},"LOSS",{rng("Q")})),"")', "0.00", cor_fl_card, fill_fl_card, "H", "I"),
    ]
    linha_card_label, linha_card_valor_ini, linha_card_valor_fim = 5, 6, 8
    for titulo, formula, fmt, cor_texto, cor_fundo, col_a, col_b in cartoes:
        rotulo = f"{col_a}{linha_card_label}"
        ws_r.merge_cells(f"{col_a}{linha_card_label}:{col_b}{linha_card_label}")
        ws_r[rotulo] = titulo
        ws_r[rotulo].font = Font(name=FONTE, bold=True, size=9, color="595959")
        ws_r[rotulo].alignment = Alignment(horizontal="center")
        ws_r[rotulo].fill = PatternFill("solid", fgColor=cor_fundo)

        valor_ref = f"{col_a}{linha_card_valor_ini}"
        ws_r.merge_cells(f"{col_a}{linha_card_valor_ini}:{col_b}{linha_card_valor_fim}")
        ws_r[valor_ref] = formula
        ws_r[valor_ref].font = Font(name=FONTE, bold=True, size=22, color=cor_texto)
        ws_r[valor_ref].number_format = fmt
        ws_r[valor_ref].alignment = Alignment(horizontal="center", vertical="center")
        ws_r[valor_ref].fill = PatternFill("solid", fgColor=cor_fundo)
        # borda ao redor de todo o "cartão" (rótulo + valor)
        for row in range(linha_card_label, linha_card_valor_fim + 1):
            for col in (col_a, col_b):
                ws_r[f"{col}{row}"].border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws_r.row_dimensions[linha_card_label].height = 18
    for row in range(linha_card_valor_ini, linha_card_valor_fim + 1):
        ws_r.row_dimensions[row].height = 18

    linha_atual = linha_card_valor_fim + 2

    n_resumido = sum(1 for r in registros if r["formato"].startswith("Resumido"))
    if n_resumido:
        aviso_lin = linha_atual
        ws_r[f"B{aviso_lin}"] = (
            f"⚠ {n_resumido} dos {n} registros são do formato ANTIGO do bot (só data, moeda, tipo, "
            "resultado e motivo — sem preço, taxa, R múltiplo etc.). Os totais de PnL acima incluem "
            "esses registros; totais que dependem de campos que eles não têm usam só os "
            f"{n - n_resumido} registros completos."
        )
        ws_r[f"B{aviso_lin}"].font = aviso_font
        ws_r[f"B{aviso_lin}"].fill = aviso_fill
        ws_r[f"B{aviso_lin}"].alignment = Alignment(wrap_text=True, vertical="center")
        ws_r.merge_cells(f"B{aviso_lin}:I{aviso_lin + 1}")
        ws_r.row_dimensions[aviso_lin].height = 30
        linha_atual = aviso_lin + 3

    # ================= DETALHES (métricas completas) =================
    ws_r.cell(row=linha_atual, column=2, value="Detalhes").font = secao_font
    linha_atual += 1
    linha_ini_detalhes = linha_atual
    linhas_metricas = [
        ("Ganho médio (trade vencedor)",  f'=IFERROR(AVERAGEIF({rng("G")},"WIN",{rng("Q")}),"")',           "$#,##0.00"),
        ("Perda média (trade perdedor)",  f'=IFERROR(AVERAGEIF({rng("G")},"LOSS",{rng("Q")}),"")',          "$#,##0.00;[RED]-$#,##0.00"),
        ("Maior ganho (USD)",             f"=MAX({rng('Q')})",                                              "$#,##0.00"),
        ("Maior perda (USD)",             f"=MIN({rng('Q')})",                                              "$#,##0.00;[RED]-$#,##0.00"),
        ("Duração média em min (só formato completo)", f'=IFERROR(AVERAGE({rng("C")}),"")',                 "0.0"),
        ("Registros formato completo",    f'=COUNTIF({rng("U")},"Completo")',                               "0"),
        ("Registros formato antigo/resumido", f'=COUNTIF({rng("U")},"Resumido (versão antiga do bot)")',    "0"),
    ]
    for i, (label, formula, fmt) in enumerate(linhas_metricas):
        linha = linha_ini_detalhes + i
        ws_r.cell(row=linha, column=2, value=label).font = label_font
        ws_r.cell(row=linha, column=2).border = border
        c = ws_r.cell(row=linha, column=3, value=formula)
        c.font = valor_font
        c.number_format = fmt
        c.border = border
    linha_atual = linha_ini_detalhes + len(linhas_metricas) + 2

    ws_r.column_dimensions["A"].width = 3
    ws_r.column_dimensions["B"].width = 34
    ws_r.column_dimensions["C"].width = 16
    ws_r.column_dimensions["D"].width = 12
    ws_r.column_dimensions["E"].width = 14

    # ================= TABELAS AGRUPADAS (ordenadas do melhor pro pior) =================
    # Ordenar por PnL (do que mais rendeu pro que menos rendeu) é bem mais fácil de
    # ler de relance do que ordem alfabética -- você já vê o "vilão" no topo ou no
    # fim da lista, sem precisar procurar.
    def _pnl_grupo(chave, valor):
        return sum(r["pnl_liquido"] or 0.0 for r in registros if r[chave] == valor)

    moedas = sorted({r["ativo"] for r in registros}, key=lambda m: _pnl_grupo("ativo", m), reverse=True)
    motivos = sorted({r["motivo"] for r in registros}, key=lambda m: _pnl_grupo("motivo", m), reverse=True)

    databar_verde = DataBarRule(start_type="min", end_type="max", color="63BE7B")

    def _tabela_agrupada(titulo, col_grupo, valores, linha_atual, col_letra_ini="B"):
        cols = ["B", "C", "D", "E"] if col_letra_ini == "B" else ["G", "H", "I", "J"]
        ws_r.cell(row=linha_atual, column=2 if col_letra_ini == "B" else 7, value=titulo).font = secao_font
        linha_atual += 1
        for j, h in enumerate(["Grupo", "Trades", "Win Rate", "PnL Líquido (USD)"]):
            c = ws_r.cell(row=linha_atual, column=(2 if col_letra_ini == "B" else 7) + j, value=h)
            c.font = header_font
            c.fill = header_fill
            c.alignment = Alignment(horizontal="center")
            c.border = border
        linha_titulo_tabela = linha_atual
        linha_atual += 1
        linha_primeira = linha_atual
        for valor in valores:
            col_base = 2 if col_letra_ini == "B" else 7
            ws_r.cell(row=linha_atual, column=col_base, value=valor).font = valor_font
            ws_r.cell(row=linha_atual, column=col_base + 1, value=f'=COUNTIF({rng(col_grupo)},{cols[0]}{linha_atual})').font = valor_font
            c_wr = ws_r.cell(row=linha_atual, column=col_base + 2, value=f'=IFERROR(COUNTIFS({rng(col_grupo)},{cols[0]}{linha_atual},{rng("G")},"WIN")/{cols[1]}{linha_atual},"")')
            c_wr.font, c_wr.number_format = valor_font, "0.0%"
            c_pnl = ws_r.cell(row=linha_atual, column=col_base + 3, value=f'=SUMIF({rng(col_grupo)},{cols[0]}{linha_atual},{rng("Q")})')
            c_pnl.font, c_pnl.number_format = valor_font, "$#,##0.00;[RED]-$#,##0.00"
            for col in range(col_base, col_base + 4):
                ws_r.cell(row=linha_atual, column=col).border = border
            linha_atual += 1
        linha_ultima = linha_atual - 1
        # barrinha colorida na coluna de PnL, do menor pro maior -- dá pra ver o
        # "tamanho" do ganho/perda de cada linha só olhando, sem ler o número
        if linha_ultima >= linha_primeira:
            ws_r.conditional_formatting.add(f"{cols[3]}{linha_primeira}:{cols[3]}{linha_ultima}", databar_verde)
        # linha de TOTAL ao final da tabela
        ws_r.cell(row=linha_atual, column=col_base, value="TOTAL").font = Font(name=FONTE, bold=True, size=10)
        ws_r.cell(row=linha_atual, column=col_base + 1, value=f"=SUM({cols[1]}{linha_primeira}:{cols[1]}{linha_ultima})").font = Font(name=FONTE, bold=True, size=10)
        c_total_pnl = ws_r.cell(row=linha_atual, column=col_base + 3, value=f"=SUM({cols[3]}{linha_primeira}:{cols[3]}{linha_ultima})")
        c_total_pnl.font = Font(name=FONTE, bold=True, size=10)
        c_total_pnl.number_format = "$#,##0.00;[RED]-$#,##0.00"
        for col in range(col_base, col_base + 4):
            ws_r.cell(row=linha_atual, column=col).border = Border(top=Side(style="double"), left=thin, right=thin, bottom=thin)
        return linha_titulo_tabela, linha_primeira, linha_ultima

    linha_moeda_titulo, linha_moeda_ini, linha_moeda_fim = _tabela_agrupada("Resultado por moeda", "D", moedas, linha_atual, "B")
    _tabela_agrupada("Resultado por motivo de fechamento", "F", motivos, linha_atual, "G")

    ws_r.column_dimensions["G"].width = 16
    ws_r.column_dimensions["H"].width = 12
    ws_r.column_dimensions["I"].width = 12
    ws_r.column_dimensions["J"].width = 18

    # ================= GRÁFICO: PnL líquido por moeda =================
    if len(moedas) >= 2:
        grafico = BarChart()
        grafico.type = "col"
        grafico.title = "PnL Líquido por Moeda (USD)"
        grafico.y_axis.title = "USD"
        grafico.x_axis.title = "Moeda"
        grafico.style = 10
        grafico.height, grafico.width = 8, 24
        grafico.legend = None
        dados = Reference(ws_r, min_col=5, min_row=linha_moeda_ini, max_row=linha_moeda_fim)  # coluna E = PnL da tabela de moedas
        categorias = Reference(ws_r, min_col=2, min_row=linha_moeda_ini, max_row=linha_moeda_fim)  # coluna B = nome da moeda
        grafico.add_data(dados, titles_from_data=False)
        grafico.set_categories(categorias)
        grafico.series[0].graphicalProperties.solidFill = "4472C4"
        linha_grafico = linha_moeda_fim + 3
        ws_r.add_chart(grafico, f"B{linha_grafico}")

    ws_r.sheet_view.showGridLines = False

    # salva com um arquivo temporário + rename, pra nunca deixar o .xlsx pela
    # metade se o bot for fechado bem no meio da gravação (mesmo cuidado que
    # já existia pro funding_arb_state.json)
    caminho_tmp = caminho_xlsx + ".tmp"
    wb.save(caminho_tmp)
    os.replace(caminho_tmp, caminho_xlsx)


class TelaConfiguracoes(Screen):
    """Tela pra editar os parâmetros de risco em tempo real, sem precisar
    parar o bot e mexer no código."""

    CSS = """
    #config-container {
        padding: 2 4;
        background: #181818;
    }
    #config-titulo {
        text-style: bold;
        color: cyan;
        margin-bottom: 1;
    }
    Label {
        margin-top: 1;
    }
    Input {
        margin-bottom: 1;
    }
    #config-botoes {
        margin-top: 2;
        height: 3;
    }
    """

    BINDINGS = [("escape", "fechar", "Voltar (sem salvar)")]

    def compose(self) -> ComposeResult:
        app: "BotTraderApp" = self.app
        yield Header()
        with VerticalScroll(id="config-container"):
            yield Label("⚙️  Configurações de Risco", id="config-titulo")
            yield Label("Risco por trade (% da banca) — recomendado 0.5 a 2:")
            yield Input(value=str(app.risco_por_trade_pct), id="input-risco")
            yield Label("Multiplicador de ATR pro Stop Loss — recomendado 1 a 2:")
            yield Input(value=str(app.sl_atr_multiplo), id="input-sl-atr")
            yield Label("Relação Risco:Retorno (Take Profit = SL x isso) — recomendado 1.5 a 3:")
            yield Input(value=str(app.relacao_risco_retorno), id="input-rr")
            yield Label("Limite de perda diária (% da banca) — recomendado 2 a 3:")
            yield Input(value=str(app.limite_perda_diaria_pct), id="input-perda-diaria")
            yield Label("Drawdown máximo a partir do pico do dia (% da banca):")
            yield Input(value=str(app.drawdown_maximo_pct), id="input-drawdown")
            yield Label("Máximo de posições simultâneas:")
            yield Input(value=str(app.max_posicoes_simultaneas), id="input-max-pos")
            yield Label("Confluência mínima (quantas estratégias precisam concordar) — padrão 2:")
            yield Input(value=str(app.confluencia_minima_padrao), id="input-confluencia")
            with Horizontal(id="config-botoes"):
                yield Button("Salvar", id="btn-salvar", variant="success")
                yield Button("Cancelar", id="btn-cancelar", variant="error")
        yield Footer()

    def action_fechar(self):
        self.app.pop_screen()

    def on_button_pressed(self, event: Button.Pressed):
        if event.button.id == "btn-cancelar":
            self.app.pop_screen()
            return
        if event.button.id != "btn-salvar":
            return

        app: "BotTraderApp" = self.app
        try:
            novo_risco = float(self.query_one("#input-risco", Input).value)
            novo_sl_atr = float(self.query_one("#input-sl-atr", Input).value)
            novo_rr = float(self.query_one("#input-rr", Input).value)
            nova_perda_diaria = float(self.query_one("#input-perda-diaria", Input).value)
            novo_drawdown = float(self.query_one("#input-drawdown", Input).value)
            novo_max_pos = int(self.query_one("#input-max-pos", Input).value)
            nova_confluencia = int(self.query_one("#input-confluencia", Input).value)
        except ValueError:
            app.notify("Valores inválidos — use apenas números.", severity="error")
            return

        validacoes = [
            (0 < novo_risco <= RISCO_POR_TRADE_MAXIMO_PCT, f"Risco por trade deve ser entre 0 e {RISCO_POR_TRADE_MAXIMO_PCT}%."),  # [v6]
            (0.5 <= novo_sl_atr <= 5, "Multiplicador de ATR deve ser entre 0.5 e 5."),
            (1 <= novo_rr <= 5, "Relação R:R deve ser entre 1 e 5."),
            (0 < nova_perda_diaria <= LIMITE_PERDA_DIARIA_MAXIMO_PCT, f"Limite de perda diária deve ser entre 0 e {LIMITE_PERDA_DIARIA_MAXIMO_PCT}%."),  # [v6]
            (0 < novo_drawdown <= LIMITE_PERDA_DIARIA_MAXIMO_PCT, f"Drawdown máximo deve ser entre 0 e {LIMITE_PERDA_DIARIA_MAXIMO_PCT}%."),  # [v6]
            (1 <= novo_max_pos <= 20, "Máximo de posições deve ser entre 1 e 20."),
            (1 <= nova_confluencia <= 6, "Confluência mínima deve ser entre 1 e 6."),
        ]
        for valido, mensagem in validacoes:
            if not valido:
                app.notify(mensagem, title="Valor fora do intervalo seguro", severity="error", timeout=6)
                return

        app.risco_por_trade_pct = novo_risco
        app.sl_atr_multiplo = novo_sl_atr
        app.relacao_risco_retorno = novo_rr
        app.limite_perda_diaria_pct = nova_perda_diaria
        app.drawdown_maximo_pct = novo_drawdown
        app.max_posicoes_simultaneas = novo_max_pos
        app.confluencia_minima_padrao = nova_confluencia

        app.escrever_log(
            f"[cyan]Configurações atualizadas: risco {novo_risco}%, SL {novo_sl_atr}xATR, "
            f"R:R 1:{novo_rr}, perda diária {nova_perda_diaria}%, drawdown {novo_drawdown}%, "
            f"máx. posições {novo_max_pos}, confluência mínima {nova_confluencia}[/cyan]"
        )
        app.notify("Configurações salvas com sucesso.", severity="information")
        self.app.pop_screen()


class TelaBacktest(Screen):
    """Testa a MESMA lógica de confluência que roda ao vivo contra dados
    históricos, antes de arriscar qualquer capital. Simplificação importante:
    aproxima o filtro de tendência com uma EMA longa na própria série de 15m
    (em vez de buscar velas de 1h à parte), e não modela slippage nem funding."""

    CSS = """
    #backtest-container {
        padding: 2 4;
        background: #181818;
    }
    #backtest-titulo {
        text-style: bold;
        color: cyan;
        margin-bottom: 1;
    }
    #backtest-resultado {
        margin-top: 2;
        border: solid #444444;
        padding: 1 2;
        min-height: 12;
    }
    Label {
        margin-top: 1;
    }
    Input {
        margin-bottom: 1;
    }
    #backtest-botoes {
        margin-top: 1;
        height: 3;
    }
    """

    BINDINGS = [("escape", "fechar", "Voltar")]

    def compose(self) -> ComposeResult:
        yield Header()
        with VerticalScroll(id="backtest-container"):
            yield Label("📊 Backtest — testa a estratégia contra dados históricos", id="backtest-titulo")
            yield Static(
                "[yellow]Aproximação: usa uma EMA longa na própria vela do timeframe escolhido no lugar de "
                "buscar um timeframe maior separado, e não simula slippage nem funding rate.[/yellow]"
            )
            yield Label("Símbolo (ex: BTC/USDT):")
            yield Input(value="BTC/USDT", id="input-simbolo")
            yield Label("Timeframe (5m, 15m, 30m, 1h, 4h ou 1d):")
            yield Input(value="15m", id="input-timeframe")
            yield Label("Quantos dias de duração (recomendado 15 a 60):")
            yield Input(value="30", id="input-dias")
            yield Label("Data de início (AAAA-MM-DD, opcional — deixe em branco pra usar os dias mais recentes):")
            yield Input(value="", id="input-data-inicio")
            with Horizontal(id="backtest-botoes"):
                yield Button("Rodar Backtest", id="btn-rodar", variant="success")
                yield Button("Fechar", id="btn-fechar", variant="error")
            yield Static("", id="backtest-resultado")
        yield Footer()

    def action_fechar(self):
        self.app.pop_screen()

    TIMEFRAMES_MINUTOS = {'5m': 5, '15m': 15, '30m': 30, '1h': 60, '4h': 240, '1d': 1440}

    def on_button_pressed(self, event: Button.Pressed):
        if event.button.id == "btn-fechar":
            self.app.pop_screen()
            return
        if event.button.id != "btn-rodar":
            return

        simbolo = self.query_one("#input-simbolo", Input).value.strip().upper()
        if "/" not in simbolo:
            self.app.notify("Use o formato SIGLA/USDT, ex: BTC/USDT.", severity="error")
            return
        timeframe = self.query_one("#input-timeframe", Input).value.strip().lower()
        if timeframe not in self.TIMEFRAMES_MINUTOS:
            self.app.notify("Timeframe deve ser um de: 5m, 15m, 30m, 1h, 4h, 1d.", severity="error")
            return
        try:
            dias = int(self.query_one("#input-dias", Input).value)
        except ValueError:
            self.app.notify("Número de dias inválido.", severity="error")
            return
        if not (1 <= dias <= 730):
            self.app.notify("Use um número de dias entre 1 e 730.", severity="error")
            return
        data_inicio_str = self.query_one("#input-data-inicio", Input).value.strip()
        if data_inicio_str:
            try:
                datetime.strptime(data_inicio_str, "%Y-%m-%d")
            except ValueError:
                self.app.notify("Data de início deve ser no formato AAAA-MM-DD, ex: 2021-01-01.", severity="error")
                return

        self.query_one("#backtest-resultado", Static).update("[yellow]Buscando histórico e simulando... aguarde.[/yellow]")
        self.run_worker(self._rodar(simbolo, timeframe, dias, data_inicio_str), exclusive=True)

    async def _rodar(self, simbolo: str, timeframe: str, dias: int, data_inicio_str: str = ""):
        app: "BotTraderApp" = self.app
        resultado_widget = self.query_one("#backtest-resultado", Static)
        desde_ms, ate_ms = self._calcular_janela_temporal(dias, data_inicio_str)
        try:
            ohlcv = await self._buscar_historico(app.exchange, simbolo, timeframe, desde_ms, ate_ms)
        except Exception as e:
            resultado_widget.update(f"[red]Erro ao buscar histórico: {e}[/red]")
            return

        limite_velas = self._limite_velas_necessario()
        if len(ohlcv) < limite_velas + 50:
            resultado_widget.update(
                f"[red]Histórico insuficiente ({len(ohlcv)} velas retornadas). "
                f"Tente um período maior ou confira o símbolo.[/red]"
            )
            return

        trades, distribuicao_regimes = self._simular(ohlcv, app, timeframe)
        resultado = self._formatar_resultado(f"{simbolo} ({timeframe})", dias, len(ohlcv), trades, distribuicao_regimes)
        if data_inicio_str:
            resultado = f"[bold magenta]Período: {dias} dias a partir de {data_inicio_str}[/bold magenta]\n\n" + resultado
        resultado_widget.update(resultado)

    @staticmethod
    def _calcular_janela_temporal(dias: int, data_inicio_str: str) -> tuple:
        """Calcula (desde_ms, ate_ms). Se 'data_inicio_str' estiver vazia,
        usa os últimos 'dias' dias a partir de agora (comportamento padrão).
        Se preenchida (formato AAAA-MM-DD), usa 'dias' dias de duração a
        PARTIR dessa data específica — assim dá pra testar um período
        histórico qualquer (ex: bull run de 2021, bear market de 2022),
        não só o passado recente."""
        agora_ms = int(time.time() * 1000)
        data_inicio_str = (data_inicio_str or "").strip()
        if not data_inicio_str:
            desde_ms = agora_ms - dias * 24 * 60 * 60 * 1000
            ate_ms = agora_ms
        else:
            dt_inicio = datetime.strptime(data_inicio_str, "%Y-%m-%d")
            desde_ms = int(dt_inicio.timestamp() * 1000)
            ate_ms = min(desde_ms + dias * 24 * 60 * 60 * 1000, agora_ms)
        return desde_ms, ate_ms

    @staticmethod
    async def _buscar_historico(exchange, simbolo: str, timeframe: str, desde_ms: int, ate_ms: int) -> list:
        """Busca o histórico completo em várias chamadas paginadas (a Binance
        limita quantas velas vêm por chamada), dentro da janela [desde_ms, ate_ms]."""
        minutos_por_vela = TelaBacktest.TIMEFRAMES_MINUTOS[timeframe]
        limite_por_chamada = 1000
        tf_ms = minutos_por_vela * 60 * 1000
        candles = []
        cursor = desde_ms
        while cursor < ate_ms:
            lote = await exchange.fetch_ohlcv(simbolo, timeframe=timeframe, since=cursor, limit=limite_por_chamada)
            if not lote:
                break
            lote_filtrado = [c for c in lote if c[0] <= ate_ms]
            candles.extend(lote_filtrado)
            ultimo_ts = lote[-1][0]
            if ultimo_ts <= cursor or ultimo_ts >= ate_ms:
                break
            cursor = ultimo_ts + tf_ms
            if len(lote) < limite_por_chamada:
                break
        return candles

    @staticmethod
    def _limite_velas_necessario() -> int:
        return max(RSI_PERIODO, ATR_PERIODO, MACD_LENTA + MACD_SINAL, BB_PERIODO,
                   ESTOCASTICO_PERIODO_K + ESTOCASTICO_SUAVIZACAO_K + ESTOCASTICO_PERIODO_D) + 20

    def _simular(self, ohlcv: list, app: "BotTraderApp", timeframe: str) -> list:
        """Percorre o histórico vela a vela, recalculando os indicadores com a
        MESMA janela fixa de velas que o bot usa ao vivo (não a partir do
        início da história — isso mantém a simulação fiel ao comportamento
        real e rápida o suficiente pra rodar na hora). A janela de tendência e
        de variação curta se ajustam ao timeframe escolhido, sempre mirando
        numa referência de ~1 hora, igual ao bot ao vivo."""
        limite_velas = self._limite_velas_necessario()
        minutos_por_vela = self.TIMEFRAMES_MINUTOS[timeframe]
        velas_por_hora = max(1, round(60 / minutos_por_vela))
        periodo_tendencia = TENDENCIA_EMA_PERIODO * velas_por_hora
        janela_variacao = velas_por_hora  # ~1h de referência, igual ao bot ao vivo

        trades = []
        posicao = None
        distribuicao_regimes = {"LATERAL": 0, "TENDENCIA": 0, "MISTO": 0}

        for i in range(limite_velas + periodo_tendencia, len(ohlcv)):
            janela = ohlcv[i - limite_velas + 1: i + 1]
            closes = [c[4] for c in janela]
            preco = closes[-1]
            bb_dados = calcular_bollinger(closes)
            regime = calcular_regime(bb_dados)
            distribuicao_regimes[regime] = distribuicao_regimes.get(regime, 0) + 1

            if posicao:
                high, low = ohlcv[i][2], ohlcv[i][3]
                if posicao["tipo"] == "LONG":
                    bateu_tp, bateu_sl = high >= posicao["tp"], low <= posicao["sl"]
                else:
                    bateu_tp, bateu_sl = low <= posicao["tp"], high >= posicao["sl"]
                if bateu_tp or bateu_sl:
                    # [v6] Conservador: se alvo E stop cabem no mesmo candle, não dá pra
                    # saber qual veio primeiro — assume o PIOR caso (stop). Antes assumia
                    # o alvo, o que deixava o backtest otimista demais.
                    preco_saida = posicao["sl"] if bateu_sl else posicao["tp"]
                    if posicao["tipo"] == "LONG":
                        pnl_bruto = (preco_saida - posicao["entrada"]) * posicao["qtd"]
                    else:
                        pnl_bruto = (posicao["entrada"] - preco_saida) * posicao["qtd"]
                    taxa = (posicao["entrada"] + preco_saida) * posicao["qtd"] * TAKER_FEE_PCT / 100
                    pnl_liquido = pnl_bruto - taxa
                    r_multiplo = (pnl_liquido / posicao["risco_usdt"]) if posicao["risco_usdt"] else 0.0
                    trades.append({"pnl": pnl_liquido, "r": r_multiplo, "regime": posicao["regime"]})
                    posicao = None
                continue  # não abre outra enquanto uma está simulada como aberta

            rsi_val = calcular_rsi(closes)
            atr = calcular_atr(janela)
            macd_dados = calcular_macd(closes)
            estoc_dados = calcular_estocastico(janela)
            candle_acao, _ = detectar_padrao_candle(janela)
            variacao = (
                (closes[-1] - closes[-(janela_variacao + 1)]) / closes[-(janela_variacao + 1)] * 100
                if len(closes) > janela_variacao and closes[-(janela_variacao + 1)] else None
            )

            janela_tendencia = ohlcv[i - periodo_tendencia - 10 + 1: i + 1]
            ema_tend = calcular_ema([c[4] for c in janela_tendencia], periodo_tendencia)
            tendencia = ("ALTA" if preco >= ema_tend else "BAIXA") if ema_tend is not None else None

            acao = self._avaliar(rsi_val, macd_dados, bb_dados, estoc_dados, candle_acao, variacao, regime, preco, app.confluencia_minima_padrao)
            if acao == "COMPRA" and tendencia == "BAIXA":
                acao = "NEUTRO"
            elif acao == "VENDA" and tendencia == "ALTA":
                acao = "NEUTRO"
            if acao == "VENDA":  # [v7 SPOT] backtest só compra, igual ao robô ao vivo
                acao = "NEUTRO"

            if acao == "NEUTRO" or atr is None or atr <= 0:
                continue

            distancia_sl = atr * app.sl_atr_multiplo
            distancia_tp = distancia_sl * app.relacao_risco_retorno
            risco_usdt = 100.0 * app.risco_por_trade_pct / 100  # banca fictícia — resultado normalizado em R
            qtd = risco_usdt / distancia_sl if distancia_sl else 0
            if qtd <= 0:
                continue

            if acao == "COMPRA":
                sl, tp = preco - distancia_sl, preco + distancia_tp
            else:
                sl, tp = preco + distancia_sl, preco - distancia_tp

            posicao = {"tipo": "LONG" if acao == "COMPRA" else "SHORT", "entrada": preco, "sl": sl, "tp": tp, "qtd": qtd, "risco_usdt": risco_usdt, "regime": regime}

        return trades, distribuicao_regimes

    @staticmethod
    def _avaliar(rsi_val, macd_dados, bb_dados, estoc_dados, candle_acao, variacao, regime, preco, confluencia_minima) -> str:
        """Réplica exata das regras de avaliar_confluencia, sem depender dos
        caches ao vivo (pra poder rodar sobre uma janela histórica qualquer)."""
        def voto_rsi_reversao():
            if rsi_val is None or variacao is None or not (-LIMIAR_LATERAL_1H_PCT <= variacao <= LIMIAR_LATERAL_1H_PCT):
                return None
            if rsi_val <= 30:
                return "COMPRA"
            if rsi_val >= 70:
                return "VENDA"
            return None

        def voto_rsi_momentum():
            if rsi_val is None or variacao is None:
                return None
            if variacao > LIMIAR_MOMENTUM_1H_PCT and 40 <= rsi_val <= 55:
                return "COMPRA"
            if variacao < -LIMIAR_MOMENTUM_1H_PCT and 45 <= rsi_val <= 60:
                return "VENDA"
            return None

        def voto_macd():
            if not macd_dados:
                return None
            if macd_dados["cruzou_alta"]:
                return "COMPRA"
            if macd_dados["cruzou_baixa"]:
                return "VENDA"
            return None

        def voto_bollinger():
            if not bb_dados or rsi_val is None:
                return None
            if preco <= bb_dados["inferior"] and rsi_val <= 35:
                return "COMPRA"
            if preco >= bb_dados["superior"] and rsi_val >= 65:
                return "VENDA"
            return None

        def voto_estocastico():
            if not estoc_dados:
                return None
            if estoc_dados["cruzou_alta"]:
                return "COMPRA"
            if estoc_dados["cruzou_baixa"]:
                return "VENDA"
            return None

        if regime == "LATERAL":
            candidatos = [voto_rsi_reversao(), voto_bollinger(), voto_estocastico(), candle_acao]
        elif regime == "TENDENCIA":
            candidatos = [voto_macd(), voto_rsi_momentum(), candle_acao]
        else:
            candidatos = [voto_rsi_reversao(), voto_rsi_momentum(), voto_macd(), voto_bollinger(), voto_estocastico(), candle_acao]

        votos_compra = candidatos.count("COMPRA")
        votos_venda = candidatos.count("VENDA")

        if votos_compra >= confluencia_minima and votos_compra > votos_venda:
            return "COMPRA"
        if votos_venda >= confluencia_minima and votos_venda > votos_compra:
            return "VENDA"
        return "NEUTRO"

    @staticmethod
    def _formatar_resultado(simbolo: str, dias: int, total_velas: int, trades: list, distribuicao_regimes: dict = None) -> str:
        distribuicao_regimes = distribuicao_regimes or {}
        total_velas_regime = sum(distribuicao_regimes.values())
        linhas_distribuicao = []
        if total_velas_regime > 0:
            for nome_regime in ["LATERAL", "TENDENCIA", "MISTO"]:
                qtd = distribuicao_regimes.get(nome_regime, 0)
                pct = qtd / total_velas_regime * 100
                linhas_distribuicao.append(f"  {nome_regime:<10} {pct:5.1f}% do tempo ({qtd} velas)")
        bloco_distribuicao = (
            "\n[bold]Quanto tempo o mercado passou em cada regime nesse período:[/bold]\n" + "\n".join(linhas_distribuicao)
            if linhas_distribuicao else ""
        )

        if not trades:
            return (
                f"[yellow]Nenhum trade gerado em {dias} dias pra {simbolo} com os parâmetros atuais "
                f"({total_velas} velas analisadas).[/yellow]\n{bloco_distribuicao}"
            )

        total = len(trades)
        wins = [t for t in trades if t["pnl"] >= 0]
        losses = [t for t in trades if t["pnl"] < 0]
        taxa_acerto = len(wins) / total * 100
        r_medio = sum(t["r"] for t in trades) / total
        soma_ganhos = sum(t["pnl"] for t in wins)
        soma_perdas = abs(sum(t["pnl"] for t in losses))
        fator_lucro = (soma_ganhos / soma_perdas) if soma_perdas > 0 else float('inf')

        maior_sequencia_perdas, atual = 0, 0
        for t in trades:
            if t["pnl"] < 0:
                atual += 1
                maior_sequencia_perdas = max(maior_sequencia_perdas, atual)
            else:
                atual = 0

        cor_pf = "green" if fator_lucro >= 1.5 else ("yellow" if fator_lucro >= 1 else "red")
        fator_lucro_str = "∞" if fator_lucro == float('inf') else f"{fator_lucro:.2f}"

        # --- Quebra por regime: mostra se o resultado vem mais de um "tipo" de
        # mercado (LATERAL, TENDENCIA, MISTO) do que de outro, pra apontar
        # exatamente onde ajustar em vez de só saber que "deu ruim". ---
        linhas_regime = []
        for nome_regime in ["LATERAL", "TENDENCIA", "MISTO"]:
            trades_regime = [t for t in trades if t.get("regime") == nome_regime]
            if not trades_regime:
                continue
            n = len(trades_regime)
            wins_r = [t for t in trades_regime if t["pnl"] >= 0]
            taxa_r = len(wins_r) / n * 100
            r_medio_r = sum(t["r"] for t in trades_regime) / n
            ganhos_r = sum(t["pnl"] for t in trades_regime if t["pnl"] >= 0)
            perdas_r = abs(sum(t["pnl"] for t in trades_regime if t["pnl"] < 0))
            fl_r = (ganhos_r / perdas_r) if perdas_r > 0 else float('inf')
            fl_r_str = "∞" if fl_r == float('inf') else f"{fl_r:.2f}"
            cor_r = "green" if fl_r >= 1.5 else ("yellow" if fl_r >= 1 else "red")
            linhas_regime.append(
                f"  {nome_regime:<10} {n:>3} trades | acerto {taxa_r:5.1f}% | R médio {r_medio_r:+.2f} | "
                f"fator [{cor_r}]{fl_r_str}[/{cor_r}]"
            )
        bloco_regime = "\n[bold]Quebra por regime de mercado:[/bold]\n" + "\n".join(linhas_regime) if linhas_regime else ""

        return (
            f"[bold cyan]Resultado — {simbolo}, últimos {dias} dias ({total_velas} velas de 15m)[/bold cyan]\n\n"
            f"Total de trades simulados: [cyan]{total}[/cyan]\n"
            f"Taxa de acerto: [cyan]{taxa_acerto:.1f}%[/cyan] ({len(wins)}W / {len(losses)}L)\n"
            f"Múltiplo de R médio: [cyan]{r_medio:+.2f}R[/cyan]\n"
            f"Fator de lucro: [{cor_pf}]{fator_lucro_str}[/{cor_pf}] (soma dos ganhos ÷ soma das perdas — acima de 1.5 é considerado saudável)\n"
            f"Maior sequência de perdas seguidas: [red]{maior_sequencia_perdas}[/red]\n"
            f"{bloco_regime}\n"
            f"{bloco_distribuicao}\n\n"
            f"[grey50]Lembrete: simula 1 símbolo por vez, aproxima o filtro de tendência, e não modela "
            f"slippage nem funding rate. É uma estimativa direcional, não uma garantia.[/grey50]"
        )


class TelaFundingArb(Screen):
    """Testa FUNDING RATE ARBITRAGE (delta-neutral): compra o ativo no Spot
    e simultaneamente vende (short) o mesmo valor no Futuro Perpétuo. Como as
    duas pontas se cancelam, o resultado NÃO depende de o preço subir ou
    descer — vem só do Funding Rate que a Binance liquida a cada 8h. Quando
    o funding é positivo, quem está Short (nós) recebe; quando fica negativo
    de forma persistente, a estratégia fecha a posição pra não pagar.
    Usa o histórico REAL de funding rate da Binance, não um valor estimado."""

    CSS = """
    #funding-container {
        padding: 2 4;
        background: #181818;
    }
    #funding-titulo {
        text-style: bold;
        color: cyan;
        margin-bottom: 1;
    }
    #funding-resultado {
        margin-top: 2;
        border: solid #444444;
        padding: 1 2;
        min-height: 12;
    }
    Label {
        margin-top: 1;
    }
    Input {
        margin-bottom: 1;
    }
    #funding-botoes {
        margin-top: 1;
        height: 3;
    }
    """

    BINDINGS = [("escape", "fechar", "Voltar")]

    def compose(self) -> ComposeResult:
        yield Header()
        with VerticalScroll(id="funding-container"):
            yield Label("💰 Funding Rate Arbitrage (Delta-Neutral)", id="funding-titulo")
            yield Static(
                "[yellow]Compra no Spot + vende no Futuro Perpétuo em quantias iguais — o resultado "
                "não depende da direção do preço, só do Funding Rate pago a cada 8h. Fecha a posição "
                "quando a média móvel de 24h do funding fica negativa, reabre quando volta a ficar "
                "positiva (cada abertura/fechamento tem custo de taxa nas duas pontas).[/yellow]"
            )
            yield Label("Símbolo (ex: BTC/USDT):")
            yield Input(value="BTC/USDT", id="input-fa-simbolo")
            yield Label("Quantos dias de duração:")
            yield Input(value="180", id="input-fa-dias")
            yield Label("Data de início (AAAA-MM-DD, opcional):")
            yield Input(value="", id="input-fa-data-inicio")
            yield Label("Janela da média móvel pra decidir fechar (nº de eventos de 8h — 3 = 24h, 9 = 3 dias, 21 = 7 dias):")
            yield Input(value="9", id="input-fa-janela")
            yield Label("Usar ordens Maker (limitadas) em vez de Taker? (SIM/NAO — SIM assume execução perfeita, sem garantia real):")
            yield Input(value="NAO", id="input-fa-maker")
            with Horizontal(id="funding-botoes"):
                yield Button("Rodar Backtest", id="btn-fa-rodar", variant="success")
                yield Button("▶ Abrir Posição Real (Demo)", id="btn-fa-abrir-real", variant="warning")
                yield Button("📊 Ver Posições Abertas", id="btn-fa-ver-abertas", variant="primary")
                yield Button("Fechar", id="btn-fa-fechar", variant="error")
            yield Static(
                "[grey50]'Abrir Posição Real' envia ordens de verdade pro Demo Trading (ou dinheiro real, "
                "se MODO_REAL=True) usando o Símbolo e a Janela preenchidos acima — sem precisar rodar o "
                "backtest antes. Ela abre as duas pontas (Spot + Futuro) e, a partir daí, o próprio bot "
                "monitora o funding sozinho e decide fechar/reabrir usando a mesma regra já validada.[/grey50]"
            )
            yield Static("", id="funding-resultado")
        yield Footer()

    @staticmethod
    def _formatar_posicoes_abertas(app: "BotTraderApp") -> str:
        if not app.posicoes_arb:
            return "[yellow]Nenhuma posição de Funding Arb aberta no momento.[/yellow]"
        linhas = ["[bold cyan]Posições de Funding Arb monitoradas agora (dado em memória, atualizado a cada evento de funding):[/bold cyan]", ""]
        for ativo, pos in app.posicoes_arb.items():
            pnl_ate_agora = pos["funding_recebido_acumulado"] - pos["taxas_pagas_acumuladas"]
            cor = "green" if pnl_ate_agora >= 0 else "red"
            status_perna = "ABERTA (recebendo/pagando funding normalmente)" if pos["posicao_aberta"] else "PAUSADA (funding negativo persistente — só o Spot ficou, esperando reverter)"
            linhas.append(
                f"[bold]{ativo}[/bold] — aberta desde {pos['hora_abertura']}\n"
                f"  Status: {status_perna}\n"
                f"  Entrada: Spot @ {pos['preco_spot_entrada']:.4f} | Futuro @ {pos['preco_perp_entrada']:.4f}\n"
                f"  Notional por perna: ${pos['notional_por_perna']:.2f} | Quantidade: {pos['qtd']}\n"
                f"  Funding recebido até agora: ${pos['funding_recebido_acumulado']:+.4f}\n"
                f"  Taxas pagas até agora: ${pos['taxas_pagas_acumuladas']:.4f}\n"
                f"  [{cor}]PnL líquido até agora: ${pnl_ate_agora:+.4f}[/{cor}]\n"
            )
        linhas.append(
            "[grey50]Esse número só muda quando um evento de funding novo é processado (a cada ~8h). "
            "Feche esta tela (ESC) e abra de novo (tecla F), depois clique em 'Ver Posições Abertas' "
            "outra vez pra atualizar. Pra encerrar uma posição antes da hora, use o botão de Pânico "
            "na tela principal (ele fecha as duas pontas com segurança).[/grey50]"
        )
        return "\n".join(linhas)

    def action_fechar(self):
        self.app.pop_screen()

    def on_button_pressed(self, event: Button.Pressed):
        if event.button.id == "btn-fa-fechar":
            self.app.pop_screen()
            return
        if event.button.id == "btn-fa-abrir-real":
            simbolo = self.query_one("#input-fa-simbolo", Input).value.strip().upper()
            if "/" not in simbolo:
                self.app.notify("Use o formato SIGLA/USDT, ex: BTC/USDT.", severity="error")
                return
            ativo = simbolo.split("/")[0]
            try:
                janela_mm = int(self.query_one("#input-fa-janela", Input).value)
            except ValueError:
                self.app.notify("A janela da média móvel precisa ser um número inteiro.", severity="error")
                return
            if not (1 <= janela_mm <= 90):
                self.app.notify("Use uma janela entre 1 e 90 eventos.", severity="error")
                return
            self.query_one("#funding-resultado", Static).update(f"[yellow]Abrindo posição real em {ativo}... aguarde.[/yellow]")
            self.run_worker(self._abrir_real(ativo, janela_mm), exclusive=False)
            return
        if event.button.id == "btn-fa-ver-abertas":
            app: "BotTraderApp" = self.app
            self.query_one("#funding-resultado", Static).update(self._formatar_posicoes_abertas(app))
            return
        if event.button.id != "btn-fa-rodar":
            return

        simbolo = self.query_one("#input-fa-simbolo", Input).value.strip().upper()
        if "/" not in simbolo:
            self.app.notify("Use o formato SIGLA/USDT, ex: BTC/USDT.", severity="error")
            return
        try:
            dias = int(self.query_one("#input-fa-dias", Input).value)
        except ValueError:
            self.app.notify("Dias precisa ser um número.", severity="error")
            return
        if not (1 <= dias <= 730):
            self.app.notify("Use um número de dias entre 1 e 730.", severity="error")
            return
        data_inicio_str = self.query_one("#input-fa-data-inicio", Input).value.strip()
        if data_inicio_str:
            try:
                datetime.strptime(data_inicio_str, "%Y-%m-%d")
            except ValueError:
                self.app.notify("Data de início deve ser no formato AAAA-MM-DD, ex: 2026-01-01.", severity="error")
                return
        try:
            janela_mm = int(self.query_one("#input-fa-janela", Input).value)
        except ValueError:
            self.app.notify("A janela da média móvel precisa ser um número inteiro.", severity="error")
            return
        if not (1 <= janela_mm <= 90):
            self.app.notify("Use uma janela entre 1 e 90 eventos.", severity="error")
            return
        usar_maker_str = self.query_one("#input-fa-maker", Input).value.strip().upper()
        if usar_maker_str not in ("SIM", "NAO", "NÃO"):
            self.app.notify("O campo de ordens Maker deve ser SIM ou NAO.", severity="error")
            return
        usar_maker = usar_maker_str == "SIM"

        self.query_one("#funding-resultado", Static).update("[yellow]Buscando histórico de funding rate... aguarde.[/yellow]")
        self.run_worker(self._rodar(simbolo, dias, data_inicio_str, janela_mm, usar_maker), exclusive=True)

    async def _abrir_real(self, ativo: str, janela_mm: int):
        app: "BotTraderApp" = self.app
        resultado_widget = self.query_one("#funding-resultado", Static)
        ok = await app.abrir_posicao_funding(ativo, janela_mm)
        if ok:
            resultado_widget.update(
                f"[green]Posição real de Funding Arb aberta em {ativo}. Acompanhe pelo log principal "
                f"do bot (tela inicial, tecla ESC pra voltar) e pelo arquivo funding_arb_journal.csv "
                f"quando ela for encerrada.[/green]"
            )
        else:
            resultado_widget.update(
                f"[red]Não foi possível abrir a posição em {ativo}. Veja o log principal do bot (ESC "
                f"pra voltar) pra entender o motivo exato do erro.[/red]"
            )

    async def _rodar(self, simbolo: str, dias: int, data_inicio_str: str, janela_mm: int, usar_maker: bool):
        app: "BotTraderApp" = self.app
        resultado_widget = self.query_one("#funding-resultado", Static)
        desde_ms, ate_ms = TelaBacktest._calcular_janela_temporal(dias, data_inicio_str)
        try:
            historico_funding = await self._buscar_funding_history(app.exchange, simbolo, desde_ms, ate_ms)
        except Exception as e:
            resultado_widget.update(f"[red]Erro ao buscar histórico de funding rate: {e}[/red]")
            return

        if len(historico_funding) < janela_mm + 2:
            resultado_widget.update(
                f"[red]Histórico de funding insuficiente ({len(historico_funding)} eventos, "
                f"a cada 8h). Tente mais dias ou uma janela menor.[/red]"
            )
            return

        resultado_sim, diagnostico = self._simular(historico_funding, janela_mm, usar_maker)
        resultado_hold = self._simular_hold_continuo(historico_funding, usar_maker)
        resultado = self._formatar_resultado(
            f"{simbolo}", dias, len(historico_funding), resultado_sim, diagnostico, janela_mm, resultado_hold, usar_maker
        )
        if data_inicio_str:
            resultado = f"[bold magenta]Período: {dias} dias a partir de {data_inicio_str}[/bold magenta]\n\n" + resultado
        resultado_widget.update(resultado)

    @staticmethod
    async def _buscar_funding_history(exchange, simbolo: str, desde_ms: int, ate_ms: int) -> list:
        """Busca o histórico REAL de funding rate da Binance (paginado, igual
        ao _buscar_historico de velas). Cada evento tem 'timestamp' e
        'fundingRate' (fração, ex: 0.0001 = 0.01%). A Binance liquida a cada
        8h."""
        intervalo_funding_ms = 8 * 60 * 60 * 1000
        limite_por_chamada = 1000
        eventos = []
        cursor = desde_ms
        while cursor < ate_ms:
            lote = await exchange.fetch_funding_rate_history(simbolo, since=cursor, limit=limite_por_chamada)
            if not lote:
                break
            lote_filtrado = [e for e in lote if e["timestamp"] <= ate_ms]
            eventos.extend(lote_filtrado)
            ultimo_ts = lote[-1]["timestamp"]
            if ultimo_ts <= cursor or ultimo_ts >= ate_ms:
                break
            cursor = ultimo_ts + intervalo_funding_ms
            if len(lote) < limite_por_chamada:
                break
        return eventos

    @staticmethod
    def _simular(historico_funding: list, janela_mm: int = 3, usar_maker: bool = False) -> tuple:
        """Simula a estratégia delta-neutral vela a vela (evento a evento de
        funding). Banca fictícia de $100, dividida meio a meio: metade
        comprada no Spot, metade travada como margem no Futuro (short) —
        então o valor NOCIONAL de cada ponta é banca/2.

        Regra de saída com TOLERÂNCIA FINANCEIRA (sugestão validada por uma
        segunda IA que analisou o teste anterior): não basta a média móvel
        ficar negativa — o prejuízo PROJETADO de continuar aberto pelos
        próximos 'janela_mm' eventos, no ritmo atual, precisa ser MAIOR que
        o custo de fechar e reabrir a posição. Ruído pequeno (funding cai um
        pouco negativo mas não o suficiente pra justificar a taxa de sair)
        é absorvido em vez de disparar uma reação cara."""
        banca_ficticia = 100.0
        notional_por_ponta = banca_ficticia / 2
        taxa_futuros_pct = MAKER_FEE_PCT if usar_maker else TAKER_FEE_PCT
        taxa_spot_pct = SPOT_MAKER_FEE_PCT if usar_maker else SPOT_TAKER_FEE_PCT
        custo_abertura_fechamento = notional_por_ponta * (taxa_spot_pct + taxa_futuros_pct) / 100

        posicao_aberta = True  # começa aberta (entrada inicial)
        total_funding_recebido = 0.0
        total_taxas = custo_abertura_fechamento  # já paga a abertura inicial
        historico_janela = []
        # Limiar em FRAÇÃO de funding: a média só é considerada "ruim o
        # suficiente" se, sozinha, já supera o custo (em %) de fechar e
        # reabrir a posição. Sem essa conversão, qualquer média levemente
        # negativa "parecia" justificar o fechamento, mesmo sendo ruído.
        limiar_fracao = custo_abertura_fechamento / notional_por_ponta
        diagnostico = {
            "eventos_funding": len(historico_funding), "vezes_abriu": 1, "vezes_fechou": 0,
            "eventos_funding_negativo": 0, "eventos_funding_positivo": 0, "vezes_absorveu_ruido": 0,
        }

        for evento in historico_funding:
            taxa_funding = evento.get("fundingRate") or 0.0
            historico_janela.append(taxa_funding)
            if len(historico_janela) > janela_mm:
                historico_janela.pop(0)

            if taxa_funding >= 0:
                diagnostico["eventos_funding_positivo"] += 1
            else:
                diagnostico["eventos_funding_negativo"] += 1

            if posicao_aberta:
                total_funding_recebido += notional_por_ponta * taxa_funding

            media_movel = sum(historico_janela) / len(historico_janela)

            if (posicao_aberta and media_movel < 0 and len(historico_janela) == janela_mm):
                if media_movel < -limiar_fracao:
                    total_taxas += custo_abertura_fechamento
                    posicao_aberta = False
                    diagnostico["vezes_fechou"] += 1
                else:
                    diagnostico["vezes_absorveu_ruido"] += 1
            elif not posicao_aberta and media_movel >= 0:
                total_taxas += custo_abertura_fechamento
                posicao_aberta = True
                diagnostico["vezes_abriu"] += 1

        if posicao_aberta:
            total_taxas += custo_abertura_fechamento

        resultado_liquido = total_funding_recebido - total_taxas
        return {
            "banca_ficticia": banca_ficticia, "notional_por_ponta": notional_por_ponta,
            "total_funding_recebido": total_funding_recebido, "total_taxas": total_taxas,
            "resultado_liquido": resultado_liquido,
        }, diagnostico

    @staticmethod
    def _simular_hold_continuo(historico_funding: list, usar_maker: bool = False) -> dict:
        """Baseline de comparação: abre a posição uma única vez e NUNCA
        fecha, independente do funding ficar negativo no meio do caminho.
        Isola se o problema está na regra de entrada/saída ou no funding
        em si — sem nenhum custo de reabertura no meio."""
        banca_ficticia = 100.0
        notional_por_ponta = banca_ficticia / 2
        taxa_futuros_pct = MAKER_FEE_PCT if usar_maker else TAKER_FEE_PCT
        taxa_spot_pct = SPOT_MAKER_FEE_PCT if usar_maker else SPOT_TAKER_FEE_PCT
        custo_abertura_fechamento = notional_por_ponta * (taxa_spot_pct + taxa_futuros_pct) / 100

        total_funding_recebido = sum(
            notional_por_ponta * (evento.get("fundingRate") or 0.0) for evento in historico_funding
        )
        total_taxas = custo_abertura_fechamento * 2  # só abre uma vez e fecha uma vez, no fim
        resultado_liquido = total_funding_recebido - total_taxas
        return {
            "total_funding_recebido": total_funding_recebido, "total_taxas": total_taxas,
            "resultado_liquido": resultado_liquido,
        }

    @staticmethod
    def _formatar_resultado(simbolo, dias, total_eventos, resultado_sim, diagnostico, janela_mm,
                             resultado_hold, usar_maker) -> str:
        resultado_liquido = resultado_sim["resultado_liquido"]
        retorno_pct = resultado_liquido / resultado_sim["banca_ficticia"] * 100
        retorno_anualizado_pct = retorno_pct * (365 / dias) if dias > 0 else 0
        cor_resultado = "green" if resultado_liquido >= 0 else "red"
        pct_eventos_positivos = (
            diagnostico["eventos_funding_positivo"] / diagnostico["eventos_funding"] * 100
            if diagnostico["eventos_funding"] else 0
        )

        cor_hold = "green" if resultado_hold["resultado_liquido"] >= 0 else "red"
        horas_janela = janela_mm * 8
        linha_hold = (
            f"Comparação — se NUNCA fechasse a posição (só funding puro, sem custo de reabertura): "
            f"[{cor_hold}]{resultado_hold['resultado_liquido']:+.2f} USDT[/{cor_hold}]\n"
        )

        taxa_futuros_pct = MAKER_FEE_PCT if usar_maker else TAKER_FEE_PCT
        taxa_spot_pct = SPOT_MAKER_FEE_PCT if usar_maker else SPOT_TAKER_FEE_PCT
        nota_maker = ""
        if usar_maker:
            nota_maker = (
                "\n[yellow]Atenção: este teste assume ordens Maker (limitadas) com execução PERFEITA — "
                "na vida real, uma ordem limitada pode não ser preenchida a tempo (o preço 'foge' antes "
                "de completar), te forçando a usar Taker mesmo ou perder a janela de entrada/saída. Trate "
                "este resultado como um cenário otimista, não como o esperado.[/yellow]"
            )

        return (
            f"[bold cyan]Funding Rate Arbitrage — {simbolo}, últimos {dias} dias "
            f"({total_eventos} eventos de funding, a cada 8h, janela de saída: {janela_mm} eventos "
            f"= {horas_janela}h, ordens {'Maker' if usar_maker else 'Taker'}: "
            f"{taxa_spot_pct:.2f}% Spot + {taxa_futuros_pct:.2f}% Futuros)[/bold cyan]\n\n"
            f"Notional por ponta (Spot e Futuro): [cyan]${resultado_sim['notional_por_ponta']:.2f}[/cyan] "
            f"(banca fictícia total: $100)\n"
            f"Total recebido em funding: [green]+{resultado_sim['total_funding_recebido']:.2f} USDT[/green]\n"
            f"Total pago em taxas (abertura/fechamento): [red]-{resultado_sim['total_taxas']:.2f} USDT[/red]\n"
            f"Resultado líquido: [{cor_resultado}]{resultado_liquido:+.2f} USDT[/{cor_resultado}]\n"
            f"{linha_hold}"
            f"Retorno no período: [{cor_resultado}]{retorno_pct:+.2f}%[/{cor_resultado}] "
            f"(equivalente a [{cor_resultado}]{retorno_anualizado_pct:+.2f}% ao ano[/{cor_resultado}], anualizado)\n"
            f"Funding positivo em [cyan]{pct_eventos_positivos:.1f}%[/cyan] dos eventos "
            f"({diagnostico['eventos_funding_positivo']}/{diagnostico['eventos_funding']})\n"
            f"Vezes que abriu posição: [cyan]{diagnostico['vezes_abriu']}[/cyan] | "
            f"Vezes que fechou por funding negativo: [cyan]{diagnostico['vezes_fechou']}[/cyan] | "
            f"Vezes que absorveu ruído sem fechar: [cyan]{diagnostico['vezes_absorveu_ruido']}[/cyan]\n"
            f"{nota_maker}\n"
            f"[grey50]Simplificação importante: este teste assume que o preço Spot e o Futuro Perpétuo "
            f"se movem juntos (risco de base ≈ zero) — não modela slippage nem a possibilidade real de "
            f"a base (diferença Spot vs. Futuro) se abrir durante a vida da posição. A regra de saída só "
            f"fecha quando o prejuízo projetado supera o custo de fechar/reabrir — ruído pequeno é "
            f"absorvido.[/grey50]"
        )


class TelaCashCarry(Screen):
    """Testa CASH-AND-CARRY (Spot vs. Futuro com vencimento fixo, o
    'Quarterly'/trimestral da Binance) — uma estratégia delta-neutral
    DIFERENTE do Funding Arbitrage (tela F), mesmo os dois lucrando sem
    depender da direção do preço.

    A diferença: o Funding Arbitrage explora o Futuro PERPÉTUO (sem data de
    vencimento, funding pago a cada 8h). Esta tela explora o Futuro
    TRIMESTRAL (tem data de vencimento fixa, ex: BTCUSDT_260327). Como o
    contrato trimestral é procurado por especuladores que querem alavancar
    uma aposta de alta sem precisar segurar o ativo, o preço dele costuma
    ficar ACIMA do preço à vista (Spot) — essa diferença é chamada de
    'basis'. Comprando Spot e vendendo o Futuro Trimestral ao mesmo tempo,
    trava-se essa diferença como lucro, e ELA VIRA ZERO NO VENCIMENTO por
    definição contratual (não é uma expectativa estatística, é uma garantia
    da própria mecânica do contrato) — diferente do Funding, aqui não existe
    'risco de o funding virar negativo pra sempre'.

    LIMITAÇÃO IMPORTANTE E HONESTA: a Binance não mantém histórico de
    contratos trimestrais JÁ VENCIDOS disponível pra backtest via API do
    jeito que mantém velas de Spot/Perpétuo. Por isso, este teste só
    consegue olhar os contratos trimestrais ATUALMENTE listados (o mais
    próximo do vencimento e, se já existir, o seguinte) — não dá pra
    reconstruir vários anos de ciclos passados como fizemos com o Funding
    Arbitrage. Pra compensar isso, esta tela GRAVA um registro (CSV) toda
    vez que roda, criando um histórico real e crescente a cada trimestre que
    passar — o mesmo espírito de rigor do projeto, só que construído aos
    poucos daqui pra frente, em vez de tudo de uma vez com dado do passado.
    """

    CSS = """
    #cc-container {
        padding: 2 4;
        background: #181818;
    }
    #cc-titulo {
        text-style: bold;
        color: cyan;
        margin-bottom: 1;
    }
    #cc-resultado {
        margin-top: 2;
        border: solid #444444;
        padding: 1 2;
        min-height: 12;
    }
    Label {
        margin-top: 1;
    }
    Input {
        margin-bottom: 1;
    }
    #cc-botoes {
        margin-top: 1;
        height: 3;
    }
    """

    BINDINGS = [("escape", "fechar", "Voltar")]

    ARQUIVO_LOG = "cash_carry_log.csv"

    def compose(self) -> ComposeResult:
        yield Header()
        with VerticalScroll(id="cc-container"):
            yield Label("📅 Cash-and-Carry (Spot vs. Futuro Trimestral) [v2 — filtro de vencidos corrigido]", id="cc-titulo")
            yield Static(
                "[yellow]Compra no Spot + vende o Futuro TRIMESTRAL (com data de vencimento) em "
                "quantias iguais. A diferença de preço entre os dois ('basis') é travada como lucro "
                "e converge a zero no vencimento por definição do contrato — diferente do Funding "
                "Arbitrage (tela F), aqui não existe risco de o retorno virar negativo se a posição "
                "for mantida até o vencimento. Mas só dá pra medir os contratos ATUALMENTE listados "
                "(a Binance não guarda histórico de contratos trimestrais já vencidos), então este "
                "teste é mais um 'raio-x de agora' do que um backtest de anos.[/yellow]"
            )
            yield Label("Moeda base (ex: BTC, ETH, SOL):")
            yield Input(value="BTC", id="input-cc-base")
            yield Label("Moeda de cotação (normalmente USDT):")
            yield Input(value="USDT", id="input-cc-quote")
            yield Label("Máximo de dias de histórico a buscar por contrato (o contrato pode ter menos tempo listado):")
            yield Input(value="200", id="input-cc-dias")
            with Horizontal(id="cc-botoes"):
                yield Button("Rodar Análise", id="btn-cc-rodar", variant="success")
                yield Button("Fechar", id="btn-cc-fechar", variant="error")
            yield Static("", id="cc-resultado")
        yield Footer()

    def action_fechar(self):
        self.app.pop_screen()

    def on_button_pressed(self, event: Button.Pressed):
        if event.button.id == "btn-cc-fechar":
            self.app.pop_screen()
            return
        if event.button.id != "btn-cc-rodar":
            return

        base = self.query_one("#input-cc-base", Input).value.strip().upper()
        quote = self.query_one("#input-cc-quote", Input).value.strip().upper()
        if not base or not quote:
            self.app.notify("Preencha a moeda base e a de cotação.", severity="error")
            return
        try:
            dias_max = int(self.query_one("#input-cc-dias", Input).value)
        except ValueError:
            self.app.notify("Dias precisa ser um número.", severity="error")
            return
        if not (5 <= dias_max <= 400):
            self.app.notify("Use um número de dias entre 5 e 400.", severity="error")
            return

        self.query_one("#cc-resultado", Static).update(
            "[yellow]Buscando contratos trimestrais e dado de Spot... aguarde.[/yellow]"
        )
        self.run_worker(self._rodar(base, quote, dias_max), exclusive=True)

    async def _rodar(self, base: str, quote: str, dias_max: int):
        app: "BotTraderApp" = self.app
        resultado_widget = self.query_one("#cc-resultado", Static)

        try:
            await app.exchange.load_markets()
        except Exception as e:
            resultado_widget.update(f"[red]Erro ao carregar mercados da Binance: {e}[/red]")
            return

        contratos = self._achar_contratos_trimestrais(app.exchange, base, quote)
        if not contratos:
            resultado_widget.update(
                f"[red]Nenhum contrato trimestral encontrado para {base}/{quote}. Confira se essa "
                f"moeda tem Futuro com vencimento fixo na Binance (nem toda moeda tem — BTC e ETH "
                f"são as mais garantidas).[/red]"
            )
            return

        symbol_spot = f"{base}/{quote}"
        desde_ms = app.exchange.milliseconds() - dias_max * 24 * 60 * 60 * 1000

        try:
            velas_spot = await app.exchange.fetch_ohlcv(symbol_spot, timeframe="1d", since=desde_ms, limit=1000)
        except Exception as e:
            resultado_widget.update(f"[red]Erro ao buscar histórico do Spot ({symbol_spot}): {e}[/red]")
            return

        if not velas_spot:
            resultado_widget.update(f"[red]Sem dado de Spot pra {symbol_spot}.[/red]")
            return

        precos_spot_por_dia = {self._chave_dia(v[0]): v[4] for v in velas_spot}  # timestamp -> close

        analises = []
        for simbolo_contrato, expiry_ms in contratos:
            try:
                velas_fut = await app.exchange.fetch_ohlcv(
                    simbolo_contrato, timeframe="1d", since=desde_ms, limit=1000
                )
            except Exception as e:
                analises.append({"simbolo": simbolo_contrato, "erro": str(e)})
                continue
            if not velas_fut:
                analises.append({"simbolo": simbolo_contrato, "erro": "sem dado retornado"})
                continue
            analise = self._analisar_contrato(velas_fut, precos_spot_por_dia, expiry_ms, app.exchange)
            analise["simbolo"] = simbolo_contrato
            analises.append(analise)

        resultado = self._formatar_resultado(base, quote, analises)
        resultado_widget.update(resultado)
        self._gravar_log(base, quote, analises)

    @staticmethod
    def _chave_dia(timestamp_ms: int) -> int:
        """Arredonda um timestamp pro início do dia UTC, pra casar as velas
        diárias do Spot com as do Futuro mesmo que cheguem com pequenas
        diferenças de milissegundo."""
        return (timestamp_ms // 86_400_000) * 86_400_000

    @staticmethod
    def _achar_contratos_trimestrais(exchange, base: str, quote: str) -> list:
        """Procura, entre os mercados já carregados pelo ccxt, os contratos
        de Futuro TRIMESTRAL (com data de vencimento) pra essa moeda —
        diferentes do Futuro Perpétuo, que não tem data. No formato unificado
        do ccxt, o Perpétuo é 'BTC/USDT:USDT' (sem data no final) e o
        Trimestral é 'BTC/USDT:USDT-260327' (com a data). Retorna uma lista
        de tuplas (simbolo, timestamp_de_vencimento_em_ms), da mais próxima
        pra mais distante.

        IMPORTANTE: a Binance mantém, na lista de mercados, também contratos
        JÁ VENCIDOS (o histórico de preço deles continua consultável, mas
        ninguém mais negocia). Perto do fim da vida de um contrato morto, o
        preço fica travado e sem liquidez, o que pode gerar uma 'basis'
        completamente irreal (valores gigantescos, sem sentido econômico).
        Por isso, aqui filtramos SÓ os contratos cujo vencimento ainda está
        no FUTURO em relação a agora — os únicos que representam uma
        oportunidade de verdade."""
        agora_ms = exchange.milliseconds()
        prefixo = f"{base}/{quote}:{quote}-"
        encontrados = []
        for simbolo, mercado in exchange.markets.items():
            if not simbolo.startswith(prefixo):
                continue
            expiry = mercado.get("expiry")
            if not expiry:
                continue
            if expiry <= agora_ms:
                continue  # já venceu — descarta, dado morto/sem liquidez
            if mercado.get("active") is False:
                continue  # mercado marcado como inativo pela própria exchange
            encontrados.append((simbolo, expiry))
        encontrados.sort(key=lambda item: item[1])
        return encontrados

    @staticmethod
    def _analisar_contrato(velas_fut: list, precos_spot_por_dia: dict, expiry_ms: int, exchange) -> dict:
        """Para cada dia em que o contrato trimestral tinha preço, calcula a
        'basis' (diferença % entre o Futuro e o Spot) e ANUALIZA essa
        diferença de acordo com quantos dias faltavam pro vencimento naquele
        dia específico — porque a mesma diferença de preço vale muito mais
        anualizada se faltam 10 dias pro vencimento do que se faltam 90.

        Isso NÃO é o lucro já realizado — é 'quanto essa posição renderia ao
        ano SE fosse aberta naquele dia e mantida até o vencimento', que é
        exatamente o resultado garantido pela mecânica do contrato (a
        diferença converge a zero no vencimento, sempre)."""
        pontos_anualizados = []
        pontos_basis_pct = []
        ultimo_ponto = None

        for vela in velas_fut:
            chave = (vela[0] // 86_400_000) * 86_400_000
            preco_spot = precos_spot_por_dia.get(chave)
            if preco_spot is None or preco_spot <= 0:
                continue
            preco_fut = vela[4]
            dias_restantes = max((expiry_ms - vela[0]) / 86_400_000, 0.5)  # nunca zero, evita divisão explosiva
            basis_pct = (preco_fut - preco_spot) / preco_spot * 100
            anualizado_pct = basis_pct * (365 / dias_restantes)
            pontos_basis_pct.append(basis_pct)
            pontos_anualizados.append(anualizado_pct)
            ultimo_ponto = {
                "timestamp": vela[0], "preco_spot": preco_spot, "preco_fut": preco_fut,
                "dias_restantes": dias_restantes, "basis_pct": basis_pct, "anualizado_pct": anualizado_pct,
            }

        if not pontos_anualizados:
            return {"erro": "não foi possível casar datas do Futuro com o Spot"}

        media_anualizada = statistics.mean(pontos_anualizados)
        desvio_anualizado = statistics.pstdev(pontos_anualizados) if len(pontos_anualizados) > 1 else 0.0
        pior_dia = min(pontos_anualizados)
        pct_dias_positivos = sum(1 for p in pontos_anualizados if p > 0) / len(pontos_anualizados) * 100

        # Custo estimado de abrir + fechar as duas pontas (Spot + Futuro),
        # convertido pra taxa anualizada usando os dias restantes DO ÚLTIMO
        # ponto (a leitura mais atual de "quanto custaria entrar hoje").
        custo_pct_total = (SPOT_TAKER_FEE_PCT * 2) + (TAKER_FEE_PCT * 2)  # compra+venda Spot, abre+liquidação Futuro
        dias_restantes_atual = ultimo_ponto["dias_restantes"]
        custo_anualizado_pct = custo_pct_total * (365 / dias_restantes_atual)
        anualizado_liquido_atual = ultimo_ponto["anualizado_pct"] - custo_anualizado_pct

        return {
            "media_anualizada_pct": media_anualizada,
            "desvio_anualizado_pct": desvio_anualizado,
            "pior_dia_anualizado_pct": pior_dia,
            "pct_dias_positivos": pct_dias_positivos,
            "dias_restantes_atual": dias_restantes_atual,
            "basis_atual_pct": ultimo_ponto["basis_pct"],
            "anualizado_bruto_atual_pct": ultimo_ponto["anualizado_pct"],
            "custo_anualizado_pct": custo_anualizado_pct,
            "anualizado_liquido_atual_pct": anualizado_liquido_atual,
            "total_dias_observados": len(pontos_anualizados),
        }

    @staticmethod
    def _formatar_resultado(base: str, quote: str, analises: list) -> str:
        linhas = [
            f"[bold cyan]Cash-and-Carry — {base}/{quote}, contratos trimestrais atualmente listados[/bold cyan]",
            "",
        ]
        algum_valido = False
        for a in analises:
            if "erro" in a:
                linhas.append(f"[red]{a['simbolo']}: erro — {a['erro']}[/red]\n")
                continue
            algum_valido = True
            cor = "green" if a["anualizado_liquido_atual_pct"] > 0 else "red"
            vence_cdi = "SIM" if a["anualizado_liquido_atual_pct"] > CDI_REFERENCIA_PCT else "não"
            aviso_vencimento_proximo = ""
            if a["dias_restantes_atual"] < 14:
                aviso_vencimento_proximo = (
                    f"  [yellow]Aviso: faltam só {a['dias_restantes_atual']:.0f} dias pro vencimento — "
                    f"nesse prazo curto, o custo fixo de taxa anualizado fica artificialmente enorme e "
                    f"o número líquido parece pior do que a estratégia realmente é. Se houver um "
                    f"contrato seguinte na lista, olhe ele em vez deste.[/yellow]\n"
                )
            linhas.append(
                f"[bold]{a['simbolo']}[/bold] — {a['total_dias_observados']} dias observados, "
                f"vence em [cyan]{a['dias_restantes_atual']:.0f} dias[/cyan]\n"
                f"  Basis atual: [cyan]{a['basis_atual_pct']:+.3f}%[/cyan] "
                f"-> anualizado bruto: [{cor}]{a['anualizado_bruto_atual_pct']:+.1f}% a.a.[/{cor}]\n"
                f"  Custo estimado (abrir+fechar as 2 pontas), anualizado: "
                f"[red]-{a['custo_anualizado_pct']:.1f}% a.a.[/red]\n"
                f"  [bold]Anualizado líquido (o que sobra de verdade): "
                f"[{cor}]{a['anualizado_liquido_atual_pct']:+.1f}% a.a.[/{cor}][/bold] "
                f"— passa do CDI ({CDI_REFERENCIA_PCT}%)? [bold]{vence_cdi}[/bold]\n"
                f"  Ao longo do período observado: média [cyan]{a['media_anualizada_pct']:+.1f}% a.a.[/cyan] | "
                f"instabilidade [cyan]{a['desvio_anualizado_pct']:.1f}[/cyan] | "
                f"pior dia [cyan]{a['pior_dia_anualizado_pct']:+.1f}% a.a.[/cyan] | "
                f"dias com basis positivo: [cyan]{a['pct_dias_positivos']:.0f}%[/cyan]\n"
                f"{aviso_vencimento_proximo}"
            )

        if not algum_valido:
            linhas.append("[red]Nenhum contrato pôde ser analisado.[/red]")

        linhas.append(
            "[grey50]IMPORTANTE — leia com o mesmo ceticismo do resto do projeto: 'anualizado líquido' "
            "é o retorno que essa posição pagaria SE mantida até o vencimento, calculado com o preço "
            "de HOJE — não é garantia do que vai acontecer amanhã (a basis muda todo dia, geralmente "
            "encolhendo perto do vencimento). A vantagem estrutural real desta estratégia sobre o "
            "Funding Arbitrage é que, se mantida até o vencimento, o resultado NÃO depende de o "
            "funding continuar positivo — a convergência é garantida pela liquidação do contrato. "
            "O ponto fraco é a falta de histórico de ciclos passados: cada vez que você rodar esta "
            "tela, um novo registro é salvo em cash_carry_log.csv, construindo evidência real ao "
            "longo dos próximos trimestres. Este teste também não modela o spread (diferença entre "
            "compra e venda) do mercado trimestral, que costuma ser mais fino em liquidez que o "
            "Perpétuo — o custo real de entrar/sair pode ser maior que o estimado aqui.[/grey50]"
        )
        return "\n".join(linhas)

    def _gravar_log(self, base: str, quote: str, analises: list):
        """Acrescenta uma linha por contrato analisado no CSV de histórico,
        criando o cabeçalho se o arquivo ainda não existir. Com o tempo,
        rodar esta tela a cada poucas semanas constrói um histórico real de
        vários trimestres — o jeito possível de ter uma amostra grande aqui,
        já que a Binance não guarda contratos vencidos pra backtest."""
        existe = os.path.exists(self.ARQUIVO_LOG)
        try:
            with open(self.ARQUIVO_LOG, "a", newline="", encoding="utf-8") as f:
                escritor = csv.writer(f)
                if not existe:
                    escritor.writerow([
                        "data_hora_registro", "par", "contrato", "dias_restantes",
                        "basis_atual_pct", "anualizado_bruto_pct", "custo_anualizado_pct",
                        "anualizado_liquido_pct", "media_periodo_pct", "desvio_periodo",
                        "pior_dia_pct", "pct_dias_positivos", "dias_observados",
                    ])
                for a in analises:
                    if "erro" in a:
                        continue
                    escritor.writerow([
                        datetime.now().isoformat(timespec="seconds"), f"{base}/{quote}", a["simbolo"],
                        f"{a['dias_restantes_atual']:.1f}", f"{a['basis_atual_pct']:.4f}",
                        f"{a['anualizado_bruto_atual_pct']:.2f}", f"{a['custo_anualizado_pct']:.2f}",
                        f"{a['anualizado_liquido_atual_pct']:.2f}", f"{a['media_anualizada_pct']:.2f}",
                        f"{a['desvio_anualizado_pct']:.2f}", f"{a['pior_dia_anualizado_pct']:.2f}",
                        f"{a['pct_dias_positivos']:.1f}", a["total_dias_observados"],
                    ])
        except Exception:
            pass  # o log é um extra; nunca deve travar a tela por causa dele


class BotTraderApp(App):
    CSS = """
    Screen {
        layout: vertical;
        background: #121212;
        color: #e0e0e0;
    }
    #topo-painel {
        height: 5;
        background: #1e1e1e;
        border: solid #00ffff;
        color: #ffffff;
        padding: 0 1;
    }
    #corpo-principal {
        height: 1fr;
    }
    #tabela-container {
        width: 65%;
        border: solid #008080;
        background: #181818;
    }
    #monitor-container {
        width: 35%;
        border: solid #cca300;
        background: #181818;
        padding: 1;
    }
    #label-risco-dia {
        background: #181818;
        margin-top: 1;
    }
    #barra-risco-dia {
        background: #181818;
        margin-bottom: 1;
    }
    DataTable {
        background: #181818;
        color: #ffffff;
    }
    #painel-comandos {
        height: auto;
        min-height: 3;
        background: #1e1e1e;
        border: solid #555555;
        color: #f0d060;
        padding: 0 1;
    }
    """

    BINDINGS = [
        ("l", "toggle_robo", "Ligar/Pausar"),
        ("p", "panico", "Pânico"),
        ("z", "zerar_todos", "Zerar Todos"),
        ("c", "zerar_lucro", "Zerar Lucro"),
        ("x", "fechar_selecionado", "Fechar Selec."),
        ("b", "bloquear_selecionado", "Travar Selec."),
        ("s", "abrir_configuracoes", "Configurações"),
        ("t", "abrir_backtest", "Backtest"),
        ("h", "sincronizar_horario_manual", "Sincronizar Hora"),
        ("o", "toggle_setup200", "Setup $200 (15m)"),
        ("y", "toggle_donchian_adx", "Donchian+ADX (BTC)"),
        ("k", "toggle_confluencia", "Confluência (normal)"),
        ("q", "quit", "Sair"),
    ]

    def __init__(self):
        super().__init__()
        self.exchange = None

        self.banca_total = 0.0
        self.banca_livre = 0.0
        self.lucro_dia = 0.0
        self.pico_lucro_dia = 0.0
        self.lucro_dia_por_sistema = {}  # {"confluencia": float, "setup200": float, ...} — trava de ganho individual por modo
        self.vitorias = 0
        self.derrotas = 0
        self.ping = 0

        self.alavancagem = 1  # [v7 SPOT] sem alavancagem
        self.perfil_risco = "SPOT (1x, só compra)"
        self.multi_moedas = True
        self.robo_ligado = False
        self.robo_pausado_por_risco = False

        self.moedas_bloqueadas = []
        self.posicao_ativa = {"ativo": "NENHUMA", "tipo": "NEUTRO", "pnl": 0.0, "qtd": 0, "preco_entrada": 0.0}
        self.posicoes_multiplas = {}
        self.posicoes_arb = {}  # Funding Arb ao vivo: {ativo: {dados da posição delta-neutral}}
        self.confluencia_ativa = True  # liga/desliga o sistema normal de confluência (tecla K) — começa ligado, é o modo padrão
        self.setup200_ativo = False  # liga/desliga o modo Setup $200 (tecla O) — começa desligado por segurança
        self.posicoes_setup200 = {}  # Setup $200 (Bollinger+RSI, banca fixa): {ativo: {dados da posição}}
        self.donchian_adx_ativo = False  # liga/desliga o modo Donchian+ADX (tecla Y) — começa desligado por segurança
        self.posicoes_donchian_adx = {}  # Donchian+ADX (BTC, seguidor de tendência): {ativo: {dados da posição}}
        self.donchian_adx_cache = {}     # canais de Donchian + valor de ADX mais recentes, por símbolo
        self.ultimo_calculo_donchian_adx = 0.0
        self._arb_ja_avisado_reconciliacao = set()  # evita logar o mesmo aviso repetido a cada reconciliação
        self.trades_recentes = []

        # Caches de indicadores (mesma lógica da versão anterior)
        self.precos_anteriores = {}
        self.rsi_cache = {}
        self.atr_cache = {}
        self.macd_cache = {}
        self.bollinger_cache = {}
        self.estocastico_cache = {}
        self.candle_cache = {}
        self.variacao_curta_cache = {}
        self.regime_cache = {}
        self.tendencia_cache = {}
        self.simbolos_indisponiveis = set()

        self.ultimo_calculo_rsi = 0.0
        self.ultimo_calculo_tendencia = 0.0
        self.ultima_reconciliacao = 0.0

        self.processando = False  # evita ciclos sobrepostos
        self.linhas_criadas = set()  # quais ativos já têm linha na tabela
        self.ordem_ativos = []  # ordem das linhas, pra localizar o ativo selecionado no cursor

        # Parâmetros de risco editáveis em tempo real pela tela de Configurações
        # (tecla S). Começam com os valores padrão definidos no topo do arquivo.
        self.risco_por_trade_pct = RISCO_POR_TRADE_PCT
        self.sl_atr_multiplo = SL_ATR_MULTIPLO
        self.relacao_risco_retorno = RELACAO_RISCO_RETORNO
        self.limite_perda_diaria_pct = LIMITE_PERDA_DIARIA_PCT
        self.drawdown_maximo_pct = DRAWDOWN_MAXIMO_PCT
        self.max_posicoes_simultaneas = MAX_POSICOES_SIMULTANEAS
        self.confluencia_minima_padrao = CONFLUENCIA_MINIMA_PADRAO
        self.confluencia_minima_banca_pequena = CONFLUENCIA_MINIMA_BANCA_PEQUENA
        self.col_keys = {}
        # [v6] controle do "dia" para o limite diário zerar sozinho à meia-noite
        self.data_dia = date.today().isoformat()
        self._ultimo_aviso_falha_fechamento = {}
        self._ultimo_aviso_setup200 = 0.0
        # [v8] estado dos filtros
        self._cooldown_ate = {}
        self._stops_recentes = []
        self._freio_ate = 0.0
        self._eventos = []
        self._eventos_lidos_em = 0.0
        self._ultimo_aviso_bloqueio = {}
        self.volume_relativo_cache = {}
        self._funding_cache = {}

    def _montar_texto_comandos(self) -> RichTable:
        """Monta os atalhos de teclado como uma lista organizada em colunas
        (tecla + descrição), em vez do rodapé padrão do Textual — que corta o
        texto quando a janela não é larga o suficiente pra mostrar tudo numa
        única linha horizontal."""
        descricoes = {b[0]: b[2] for b in self.BINDINGS}
        itens = [(tecla.upper(), descricao) for tecla, descricao in descricoes.items()]

        colunas_por_linha = 3
        tabela = RichTable(show_header=False, box=None, padding=(0, 2, 0, 0), expand=True)
        for _ in range(colunas_por_linha):
            tabela.add_column(justify="left")

        for i in range(0, len(itens), colunas_por_linha):
            linha_itens = itens[i:i + colunas_por_linha]
            celulas = [f"[bold cyan]{tecla}[/bold cyan] {descricao}" for tecla, descricao in linha_itens]
            while len(celulas) < colunas_por_linha:
                celulas.append("")
            tabela.add_row(*celulas)

        return tabela

    def compose(self) -> ComposeResult:
        yield Header(show_clock=True)
        yield Static("Inicializando Painel Profissional...", id="topo-painel")
        with Horizontal(id="corpo-principal"):
            with Vertical(id="tabela-container"):
                yield DataTable(id="tabela-ativos")
            with Vertical(id="monitor-container"):
                yield Static("[bold yellow]Monitor de Ordens e PnL[/bold yellow]\n", id="monitor-texto")
                yield Label("Risco do dia utilizado:", id="label-risco-dia")
                yield ProgressBar(id="barra-risco-dia", total=100, show_eta=False, show_percentage=False)
                yield RichLog(id="log-historico", highlight=True, markup=True, wrap=True)
        yield Static(self._montar_texto_comandos(), id="painel-comandos")

    async def _sincronizar_horario(self):
        """Recalcula a diferença entre o relógio do PC e o relógio da
        Binance, nas duas conexões. Chamado no início e periodicamente,
        porque o relógio do Windows pode 'derivar' de novo com o tempo,
        mesmo depois de sincronizado uma vez."""
        await self.exchange.load_time_difference()
        await self.exchange_spot.load_time_difference()

    async def on_mount(self) -> None:
        # [v7 SPOT] uma única conexão, configurada para o mercado SPOT
        self.exchange = ccxt.binance({
            'apiKey': API_KEY,
            'secret': SECRET_KEY,
            'options': {'defaultType': 'spot', 'adjustForTimeDifference': True, 'recvWindow': 60000},
            'enableRateLimit': True,
        })
        if not MODO_REAL:
            self.exchange.enable_demo_trading(True)
        self.exchange_spot = self.exchange  # compatibilidade com o código antigo

        # CORRIGIDO: erro "Timestamp for this request was Xms ahead" —
        # força a sincronização de horário com a Binance JÁ na inicialização
        # (não espera acontecer sozinho), e aumenta a margem de tolerância
        # (recvWindow) de 5s (padrão) pra 60s (o máximo permitido pela
        # Binance) — dá bastante folga pro relógio do PC sem quebrar.
        # Isso é reforço — o ideal continua sendo manter o relógio do
        # Windows sincronizado.
        try:
            await self._sincronizar_horario()
        except Exception as e:
            self.escrever_log(f"[yellow]Aviso: não consegui sincronizar o horário com a Binance no início: {e}[/yellow]")

        # CRÍTICO: carrega qualquer posição de Funding Arb salva de uma
        # sessão anterior ANTES de reconciliar — assim a reconciliação (que
        # é do motor de indicadores clássicos, de uma perna só) já sabe
        # quais posições no Futuro pertencem ao Funding Arb e não tenta
        # "adotar" elas como se fossem um trade comum, órfão, sem hedge.
        self._carregar_estado_risco_diario()  # [v6]
        self._carregar_posicoes_spot()  # [v7 SPOT]

        table = self.query_one("#tabela-ativos", DataTable)
        table.cursor_type = "row"
        table.zebra_stripes = True
        # Largura FIXA e generosa em cada coluna — sem isso, o Textual define a
        # largura pela primeira vez que a coluna recebe conteúdo, e atualizações
        # posteriores com texto mais longo (ex: "RSI Mom.+Candle ●●○○○○ (MISTO)")
        # ficam cortadas em vez de mostrar tudo. Com a soma das larguras passando
        # do tamanho visível da tabela, a barra de rolagem horizontal aparece
        # sozinha (o DataTable já suporta isso nativamente).
        titulos_e_larguras = [
            ("ativo", "Ativo", 14),
            ("preco", "Preço ($)", 16),
            ("rsi", "RSI Real", 10),
            ("tendencia", "Tend. 1h", 12),
            ("sinal", "Sinal", 46),
            ("vol", "Vol 24h", 12),
            ("posicao", "Posição (PnL)", 22),
        ]
        self.col_keys = {}
        for nome, titulo, largura in titulos_e_larguras:
            self.col_keys[nome] = table.add_column(titulo, width=largura)

        log = self.query_one("#log-historico", RichLog)
        log.write("[cyan]Sistema iniciado. RSI real, TP/SL automático e limites de risco ativos.[/cyan]")
        log.write("[cyan]Use as teclas mostradas no rodapé — não precisa digitar número + Enter.[/cyan]")

        if not API_KEY or not SECRET_KEY:  # [v6]
            log.write("[bold red]Chaves da Binance não configuradas (BINANCE_API_KEY / BINANCE_API_SECRET). O bot só vai mostrar preços — nenhuma ordem será enviada.[/bold red]")
        log.write(f"[cyan][v6] Stop registrado na Binance: {'ATIVO' if EXIGIR_STOP_NA_EXCHANGE else 'DESLIGADO'} | Alavancagem máx.: {ALAVANCAGEM_MAXIMA_PERMITIDA}x[/cyan]")
        log.write("[cyan][v7 SPOT] Mercado Spot: só compra, sem alavancagem. Taxa considerada: 0,10%/ordem.[/cyan]")
        log.write(f"[cyan][v8] {VERSAO_ROBO}: velas de 1h, filtro de custo, freio após perdas, proteção de lucro, "
                  f"filtro de eventos ({'eventos.json encontrado' if os.path.exists(CAMINHO_EVENTOS) else 'sem eventos.json'}), "
                  f"volume mín. ${VOLUME_MINIMO/1e6:.0f}M.[/cyan]")
        log.write("[cyan]Verificando posições já abertas na exchange...[/cyan]")
        await self.reconciliar_posicoes()

        self.set_interval(CICLO_SEGUNDOS, self.atualizar_mercado)

    async def on_unmount(self) -> None:
        if self.exchange:
            await self.exchange.close()
        if getattr(self, 'exchange_spot', None) and self.exchange_spot is not self.exchange:
            await self.exchange_spot.close()

    # ------------------------------------------------------------------
    # Utilidades de log e journal
    # ------------------------------------------------------------------
    def escrever_log(self, msg: str):
        self.query_one("#log-historico", RichLog).write(msg)

    def enviar_telegram(self, mensagem: str):
        """Dispara uma mensagem pro grupo do Telegram configurado, SEM
        travar o resto do bot esperando a resposta — a mensagem é mandada
        'em paralelo' (asyncio.create_task) e qualquer falha (sem internet,
        token errado, bot removido do grupo etc.) só vira um aviso no log,
        nunca interrompe uma abertura/fechamento de posição de verdade."""
        if not TELEGRAM_NOTIFICACOES_ATIVAS:
            return
        if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
            return
        if not AIOHTTP_DISPONIVEL:
            if not getattr(self, "_avisou_aiohttp_faltando", False):
                self._avisou_aiohttp_faltando = True
                self.escrever_log(
                    "[yellow]Para mandar notificações no Telegram, instale: pip install aiohttp[/yellow]"
                )
            return
        asyncio.create_task(self._enviar_telegram_async(mensagem))

    async def _enviar_telegram_async(self, mensagem: str):
        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        payload = {"chat_id": TELEGRAM_CHAT_ID, "text": mensagem, "parse_mode": "HTML"}
        try:
            timeout = aiohttp.ClientTimeout(total=10)
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.post(url, json=payload) as resp:
                    if resp.status != 200:
                        corpo = await resp.text()
                        self.escrever_log(f"[yellow]Telegram recusou a mensagem (status {resp.status}): {corpo[:200]}[/yellow]")
        except Exception as e:
            self.escrever_log(f"[yellow]Aviso: falha ao enviar notificação pro Telegram: {e}[/yellow]")

    def registrar_trade(self, pnl_liquido: float, ativo: str):
        self.lucro_dia += pnl_liquido
        if pnl_liquido >= 0:
            self.vitorias += 1
            self.trades_recentes.insert(0, f"[green]WIN +${pnl_liquido:.2f} ({ativo})[/green]")
        else:
            self.derrotas += 1
            self.trades_recentes.insert(0, f"[red]LOSS ${pnl_liquido:.2f} ({ativo})[/red]")
        if len(self.trades_recentes) > 20:
            del self.trades_recentes[20:]

    def registrar_trade_csv(self, ativo: str, dados_p: dict, preco_saida: float, pnl_bruto: float,
                             taxa_abertura: float, taxa_fechamento: float, pnl_liquido: float, motivo: str):
        """Journal completo em CSV, no padrão de um trader profissional:
        preços, taxas separadas, duração, e o múltiplo de R."""
        try:
            agora = datetime.now()
            hora_abertura_str = dados_p.get("hora_abertura", "")
            try:
                duracao_min = round((agora - datetime.fromisoformat(hora_abertura_str)).total_seconds() / 60, 1) if hora_abertura_str else ""
            except ValueError:
                duracao_min = ""

            risco_real = dados_p.get("risco_usdt", 0.0)
            r_multiplo = round(pnl_liquido / risco_real, 2) if risco_real else ""
            resultado = "WIN" if pnl_liquido >= 0 else "LOSS"
            valor_nocional = dados_p["qtd"] * dados_p["preco_entrada"]

            arquivo_novo = not os.path.exists(CAMINHO_JOURNAL)
            with open(CAMINHO_JOURNAL, 'a', newline='', encoding='utf-8-sig') as f:
                escritor = csv.writer(f)
                if arquivo_novo:
                    escritor.writerow([
                        'data_hora_abertura', 'data_hora_fechamento', 'duracao_min',
                        'ativo', 'tipo', 'motivo_fechamento', 'resultado',
                        'preco_entrada', 'preco_saida', 'quantidade', 'alavancagem', 'valor_nocional_usdt',
                        'risco_planejado_usdt', 'pnl_bruto_usdt', 'taxa_abertura_usdt', 'taxa_fechamento_usdt',
                        'pnl_liquido_usdt', 'r_multiplo', 'banca_no_momento_usdt', 'lucro_acumulado_dia_usdt',
                    ])
                escritor.writerow([
                    hora_abertura_str, agora.isoformat(timespec='seconds'), duracao_min,
                    ativo, dados_p["tipo"], motivo, resultado,
                    f"{dados_p['preco_entrada']:.6f}", f"{preco_saida:.6f}", dados_p["qtd"],
                    dados_p.get("alavancagem", ""), f"{valor_nocional:.2f}",
                    f"{risco_real:.2f}", f"{pnl_bruto:.4f}", f"{taxa_abertura:.4f}", f"{taxa_fechamento:.4f}",
                    f"{pnl_liquido:.4f}", r_multiplo, f"{self.banca_total:.2f}", f"{self.lucro_dia:.2f}",
                ])
        except Exception as e:
            self.escrever_log(f"[red]Erro ao gravar journal: {e}[/red]")
            return

        # Regenera a planilha .xlsx já formatada (aba Resumo + Operações),
        # a partir do CSV que acabou de ser atualizado. Isso é só um "bônus"
        # de visualização — se der qualquer problema aqui (ex.: openpyxl não
        # instalado, arquivo aberto no Excel travando a escrita), o CSV já
        # foi gravado normalmente acima e o bot continua funcionando 100%.
        if OPENPYXL_DISPONIVEL:
            try:
                gerar_journal_xlsx()
            except Exception as e:
                self.escrever_log(
                    f"[yellow]Aviso: não deu pra atualizar {CAMINHO_JOURNAL_XLSX} agora "
                    f"(talvez esteja aberto no Excel?). O trades_journal.csv foi gravado "
                    f"normalmente. Detalhe: {e}[/yellow]"
                )
        elif not getattr(self, "_avisou_openpyxl_faltando", False):
            self._avisou_openpyxl_faltando = True
            self.escrever_log(
                "[yellow]Para gerar o trades_journal.xlsx (planilha organizada) automaticamente, "
                "instale a biblioteca openpyxl: pip install openpyxl[/yellow]"
            )

    # ------------------------------------------------------------------
    # Reconciliação com a exchange
    # ------------------------------------------------------------------
    def _reconstruir_posicao_orfa(self, dados_ex: dict) -> dict:
        """Monta os campos que o bot precisa (SL, TP, risco) para uma posição
        encontrada na exchange mas que o bot não tinha em memória."""
        simbolo = dados_ex["simbolo"]
        entrada = dados_ex["preco_entrada"]
        atr = self.atr_cache.get(simbolo)
        distancia_sl = (atr * self.sl_atr_multiplo) if atr else (entrada * 0.01)
        distancia_tp = distancia_sl * self.relacao_risco_retorno

        if dados_ex["tipo"] == "LONG":
            sl_preco = entrada - distancia_sl
            tp_preco = entrada + distancia_tp
        else:
            sl_preco = entrada + distancia_sl
            tp_preco = entrada - distancia_tp

        return {
            "tipo": dados_ex["tipo"], "pnl": 0.0, "qtd": dados_ex["qtd"], "preco_entrada": entrada,
            "sl_preco": sl_preco, "tp_preco": tp_preco,
            "risco_usdt": distancia_sl * dados_ex["qtd"],
            "fee_abertura": entrada * dados_ex["qtd"] * TAKER_FEE_PCT / 100,
            "alavancagem": dados_ex["alavancagem"],
            "hora_abertura": "",
        }

    async def reconciliar_posicoes(self):
        """[v7 SPOT] No Spot não existe "posição" na corretora: o robô confere
        o SALDO de cada moeda que ele comprou.
        - Saldo sumiu (stop executado pela Binance, venda manual no app):
          registra no journal e para de vigiar.
        - Saldo existe mas o stop sumiu da Binance: registra o stop de novo.
        Moedas que você já tinha na conta e que o robô não comprou são
        IGNORADAS (o robô nunca vende o que não comprou)."""
        try:
            balance = await self.exchange.fetch_balance()
        except Exception as e:
            self.escrever_log(f"[red]Erro ao reconciliar posições: {e}[/red]")
            return
        totais = balance.get('total', {}) or {}
        sistemas = (
            ("confluencia", self.posicoes_multiplas),
            ("setup200", self.posicoes_setup200),
            ("donchian_adx", self.posicoes_donchian_adx),
        )
        for sistema, dicionario in sistemas:
            for ativo, dados_p in list(dicionario.items()):
                saldo_total = float(totais.get(ativo, 0.0) or 0.0)
                if saldo_total < dados_p["qtd"] * 0.5:
                    if await self._fechamento_externo(ativo, dados_p, sistema):
                        dicionario.pop(ativo, None)
                    continue
                if not EXIGIR_STOP_NA_EXCHANGE or self._idade_segundos(dados_p) < 60:
                    continue
                simbolo_par = f"{ativo}/USDT"
                try:
                    ordens_abertas = await self.exchange.fetch_open_orders(simbolo_par)
                except Exception:
                    continue
                if not ordens_abertas:
                    self.escrever_log(f"[yellow]{ativo}: stop não encontrado na Binance — registrando de novo.[/yellow]")
                    await self._colocar_protecoes(simbolo_par, "LONG", dados_p["qtd"], dados_p["sl_preco"], None)
        self._salvar_posicoes_spot()
        self.ultima_reconciliacao = time.time()

    # ------------------------------------------------------------------
    # Indicadores (recalculados periodicamente, não em todo ciclo)
    # ------------------------------------------------------------------
    async def _buscar_e_processar_candles_rsi(self, simbolo: str, semaforo: asyncio.Semaphore, limite_velas: int):
        async with semaforo:
            try:
                ohlcv = await self.exchange.fetch_ohlcv(simbolo, timeframe=RSI_TIMEFRAME, limit=limite_velas + 1)
                # [v6] descarta a vela ainda em formação: os sinais param de "piscar"
                # e o robô ao vivo passa a decidir igual ao backtest (velas fechadas).
                ohlcv = ohlcv[:-1]
                closes = [c[4] for c in ohlcv]
                self.rsi_cache[simbolo] = calcular_rsi(closes)
                self.atr_cache[simbolo] = calcular_atr(ohlcv)
                self.macd_cache[simbolo] = calcular_macd(closes)
                self.bollinger_cache[simbolo] = calcular_bollinger(closes)
                self.estocastico_cache[simbolo] = calcular_estocastico(ohlcv)
                self.candle_cache[simbolo] = detectar_padrao_candle(ohlcv)
                self.regime_cache[simbolo] = calcular_regime(self.bollinger_cache[simbolo])
                # [v8] volume relativo (para o filtro opcional de volume)
                vols = [c[5] for c in ohlcv if len(c) > 5 and c[5] is not None]
                if len(vols) >= 21 and sum(vols[-21:-1]) > 0:
                    self.volume_relativo_cache[simbolo] = vols[-1] / (sum(vols[-21:-1]) / 20)
                n_var = VARIACAO_JANELA_VELAS  # [v8] janela da variação curta em velas
                if len(closes) > n_var and closes[-1 - n_var] > 0:
                    self.variacao_curta_cache[simbolo] = (closes[-1] - closes[-1 - n_var]) / closes[-1 - n_var] * 100
                else:
                    self.variacao_curta_cache[simbolo] = None
            except Exception as e:
                msg = str(e)
                if "does not have market symbol" in msg or "BadSymbol" in type(e).__name__:
                    self.simbolos_indisponiveis.add(simbolo)
                    self.escrever_log(f"[yellow]{simbolo} indisponível nesse ambiente — removido da lista.[/yellow]")
                # mantém os valores anteriores, se houver

    async def atualizar_rsi_real(self):
        """Recalcula RSI/ATR/MACD/Bollinger/Estocástico/Candle pra todas as
        moedas EM PARALELO (limitado por semáforo, pra não estourar o rate
        limit da Binance) — bem mais rápido que buscar uma de cada vez."""
        agora = time.time()
        if agora - self.ultimo_calculo_rsi < RSI_RECALC_SEGUNDOS and self.rsi_cache:
            return
        self.ultimo_calculo_rsi = agora
        limite_velas = max(RSI_PERIODO, ATR_PERIODO, MACD_LENTA + MACD_SINAL, BB_PERIODO,
                            ESTOCASTICO_PERIODO_K + ESTOCASTICO_SUAVIZACAO_K + ESTOCASTICO_PERIODO_D) + 20
        semaforo = asyncio.Semaphore(8)
        tarefas = [
            self._buscar_e_processar_candles_rsi(simbolo, semaforo, limite_velas)
            for simbolo in MOEDAS if simbolo not in self.simbolos_indisponiveis
        ]
        if tarefas:
            await asyncio.gather(*tarefas)

    async def _buscar_e_processar_tendencia(self, simbolo: str, semaforo: asyncio.Semaphore):
        async with semaforo:
            try:
                ohlcv = await self.exchange.fetch_ohlcv(simbolo, timeframe=TENDENCIA_TIMEFRAME, limit=TENDENCIA_EMA_PERIODO + 11)
                ohlcv = ohlcv[:-1]  # [v6] só velas fechadas
                closes = [c[4] for c in ohlcv]
                ema = calcular_ema(closes, TENDENCIA_EMA_PERIODO)
                if ema is None or not closes:
                    return
                self.tendencia_cache[simbolo] = "ALTA" if closes[-1] >= ema else "BAIXA"
            except Exception:
                pass

    async def atualizar_tendencia(self):
        agora = time.time()
        if agora - self.ultimo_calculo_tendencia < TENDENCIA_RECALC_SEGUNDOS and self.tendencia_cache:
            return
        self.ultimo_calculo_tendencia = agora
        semaforo = asyncio.Semaphore(8)
        tarefas = [
            self._buscar_e_processar_tendencia(simbolo, semaforo)
            for simbolo in MOEDAS if simbolo not in self.simbolos_indisponiveis
        ]
        if tarefas:
            await asyncio.gather(*tarefas)

    async def atualizar_donchian_adx(self):
        """Busca velas de 4h do BTC e recalcula o canal de Donchian (entrada
        20 / saída 30) e o ADX(14) — mesmos parâmetros validados nos
        backtests (backtest_donchian_adx.py e as variações _bullrun/_lateral).
        Velas de 4h mudam devagar, então recalcular a cada 15 minutos já é
        mais que suficiente (não precisa do ritmo do RSI de 15m)."""
        agora = time.time()
        if agora - self.ultimo_calculo_donchian_adx < DONCHIAN_ADX_RECALC_SEGUNDOS and self.donchian_adx_cache:
            return
        self.ultimo_calculo_donchian_adx = agora
        try:
            limite_velas = max(DONCHIAN_ADX_ENTRADA, DONCHIAN_ADX_SAIDA, DONCHIAN_ADX_PERIODO * 2) + 20
            ohlcv = await self.exchange.fetch_ohlcv(
                DONCHIAN_ADX_SIMBOLO, timeframe=DONCHIAN_ADX_TIMEFRAME, limit=limite_velas
            )
            canais = calcular_donchian_canais(ohlcv, DONCHIAN_ADX_ENTRADA, DONCHIAN_ADX_SAIDA)
            adx_val = calcular_adx(ohlcv, DONCHIAN_ADX_PERIODO)
            if canais is not None:
                self.donchian_adx_cache[DONCHIAN_ADX_SIMBOLO] = {**canais, "adx": adx_val}
        except Exception as e:
            self.escrever_log(f"[yellow]Donchian+ADX: falha ao atualizar indicadores de {DONCHIAN_ADX_SIMBOLO}: {e}[/yellow]")

    def avaliar_confluencia(self, simbolo: str, preco: float, volume: float):
        """Vota entre as estratégias aplicáveis ao regime atual e só libera um
        sinal se um número mínimo delas concordar (mais exigente p/ banca pequena)."""
        if volume < VOLUME_MINIMO:
            return "NEUTRO", "AGUARDANDO (volume baixo)", "—"

        rsi_val = self.rsi_cache.get(simbolo)
        macd_dados = self.macd_cache.get(simbolo)
        bb_dados = self.bollinger_cache.get(simbolo)
        estoc_dados = self.estocastico_cache.get(simbolo)
        candle_acao, _candle_nome = self.candle_cache.get(simbolo, (None, None))
        variacao = self.variacao_curta_cache.get(simbolo)
        regime = self.regime_cache.get(simbolo, "MISTO")

        def voto_rsi_reversao():
            if rsi_val is None or variacao is None or not (-LIMIAR_LATERAL_1H_PCT <= variacao <= LIMIAR_LATERAL_1H_PCT):
                return None
            if rsi_val <= 30:
                return "COMPRA"
            if rsi_val >= 70:
                return "VENDA"
            return None

        def voto_rsi_momentum():
            if rsi_val is None or variacao is None:
                return None
            if variacao > LIMIAR_MOMENTUM_1H_PCT and 40 <= rsi_val <= 55:
                return "COMPRA"
            if variacao < -LIMIAR_MOMENTUM_1H_PCT and 45 <= rsi_val <= 60:
                return "VENDA"
            return None

        def voto_macd():
            if not macd_dados:
                return None
            if macd_dados["cruzou_alta"]:
                return "COMPRA"
            if macd_dados["cruzou_baixa"]:
                return "VENDA"
            return None

        def voto_bollinger():
            if not bb_dados or rsi_val is None:
                return None
            if preco <= bb_dados["inferior"] and rsi_val <= 35:
                return "COMPRA"
            if preco >= bb_dados["superior"] and rsi_val >= 65:
                return "VENDA"
            return None

        def voto_estocastico():
            if not estoc_dados:
                return None
            if estoc_dados["cruzou_alta"]:
                return "COMPRA"
            if estoc_dados["cruzou_baixa"]:
                return "VENDA"
            return None

        def voto_candle():
            return candle_acao

        if regime == "LATERAL":
            candidatos = [
                ("RSI Rev.", voto_rsi_reversao()), ("Bollinger", voto_bollinger()),
                ("Estocástico", voto_estocastico()), ("Candle", voto_candle()),
            ]
        elif regime == "TENDENCIA":
            candidatos = [("MACD", voto_macd()), ("RSI Mom.", voto_rsi_momentum()), ("Candle", voto_candle())]
        else:
            candidatos = [
                ("RSI Rev.", voto_rsi_reversao()), ("RSI Mom.", voto_rsi_momentum()),
                ("MACD", voto_macd()), ("Bollinger", voto_bollinger()),
                ("Estocástico", voto_estocastico()), ("Candle", voto_candle()),
            ]

        votos_compra = [nome for nome, v in candidatos if v == "COMPRA"]
        votos_venda = [nome for nome, v in candidatos if v == "VENDA"]

        confluencia_necessaria = (
            self.confluencia_minima_banca_pequena if self.banca_total < BANCA_PEQUENA_LIMITE_USDT
            else self.confluencia_minima_padrao
        )

        def _bolinhas(qtd_votos: int, total: int) -> str:
            return "●" * qtd_votos + "○" * (total - qtd_votos)

        total_candidatos = len(candidatos)
        if len(votos_compra) >= confluencia_necessaria and len(votos_compra) > len(votos_venda):
            bolinhas = _bolinhas(len(votos_compra), total_candidatos)
            return "COMPRA", f"{'+'.join(votos_compra)} {bolinhas} ({regime})", regime
        if len(votos_venda) >= confluencia_necessaria and len(votos_venda) > len(votos_compra):
            bolinhas = _bolinhas(len(votos_venda), total_candidatos)
            return "VENDA", f"{'+'.join(votos_venda)} {bolinhas} ({regime})", regime
        melhor_contagem = max(len(votos_compra), len(votos_venda))
        bolinhas = _bolinhas(melhor_contagem, total_candidatos)
        return "NEUTRO", f"AGUARDANDO {bolinhas} ({regime})", regime

    # ------------------------------------------------------------------
    # [v6] Utilidades de segurança: estado diário, proteções na exchange
    # ------------------------------------------------------------------
    def _salvar_estado_risco_diario(self):
        """Grava o resultado do dia em disco — reiniciar o bot não apaga mais
        a perda do dia (antes, fechar e abrir o bot "zerava" o limite)."""
        try:
            with open(CAMINHO_ESTADO_RISCO_DIARIO, "w", encoding="utf-8") as f:
                json.dump({
                    "data": self.data_dia,
                    "lucro_dia": self.lucro_dia,
                    "pico_lucro_dia": self.pico_lucro_dia,
                    "lucro_dia_por_sistema": self.lucro_dia_por_sistema,
                    "robo_pausado_por_risco": self.robo_pausado_por_risco,
                }, f)
        except Exception as e:
            self.escrever_log(f"[yellow]Aviso: falha ao salvar o estado do limite diário: {e}[/yellow]")

    def _carregar_estado_risco_diario(self):
        if not os.path.exists(CAMINHO_ESTADO_RISCO_DIARIO):
            return
        try:
            with open(CAMINHO_ESTADO_RISCO_DIARIO, "r", encoding="utf-8") as f:
                estado = json.load(f)
        except Exception as e:
            self.escrever_log(f"[yellow]Aviso: falha ao ler o estado do limite diário: {e}[/yellow]")
            return
        if estado.get("data") != date.today().isoformat():
            return  # é de outro dia: começa o dia zerado
        self.lucro_dia = float(estado.get("lucro_dia", 0.0))
        self.pico_lucro_dia = float(estado.get("pico_lucro_dia", 0.0))
        self.lucro_dia_por_sistema = dict(estado.get("lucro_dia_por_sistema", {}))
        self.robo_pausado_por_risco = bool(estado.get("robo_pausado_por_risco", False))
        self.escrever_log(
            f"[cyan]Resultado de hoje recuperado do disco: ${self.lucro_dia:+.2f} "
            f"(o limite diário continua valendo mesmo após reiniciar o bot).[/cyan]"
        )

    def _verificar_virada_do_dia(self):
        """Zera o resultado do dia automaticamente à meia-noite (antes só
        zerava apertando C — rodando vários dias, o 'diário' virava 'total')."""
        hoje = date.today().isoformat()
        if hoje == self.data_dia:
            return
        self.data_dia = hoje
        self.lucro_dia = 0.0
        self.pico_lucro_dia = 0.0
        self.lucro_dia_por_sistema = {}
        if self.robo_pausado_por_risco:
            self.robo_pausado_por_risco = False
            self.escrever_log("[cyan]Novo dia: pausa por risco liberada. O robô continua DESLIGADO — aperte L para religar.[/cyan]")
        self.escrever_log("[cyan]Novo dia: resultado diário e travas de ganho/perda zerados automaticamente.[/cyan]")
        self._salvar_estado_risco_diario()

    def _pnl_aberto_total(self) -> float:
        """Soma o lucro/prejuízo ainda NÃO realizado de todas as posições."""
        total = sum(d.get("pnl", 0.0) for d in self.posicoes_multiplas.values())
        if self.posicao_ativa.get("ativo") != "NENHUMA":
            total += self.posicao_ativa.get("pnl", 0.0)
        total += sum(d.get("pnl", 0.0) for d in self.posicoes_setup200.values())
        total += sum(d.get("pnl", 0.0) for d in self.posicoes_donchian_adx.values())
        return total

    def _ajustar_qtd(self, simbolo: str, qtd: float) -> float:
        """Arredonda a quantidade para o passo aceito pela Binance (em vez de
        round(..., 4), que podia gerar quantidades recusadas ou diferentes da real)."""
        try:
            return float(self.exchange.amount_to_precision(simbolo, qtd))
        except Exception:
            return 0.0

    @staticmethod
    def _idade_segundos(dados_p: dict) -> float:
        try:
            return (datetime.now() - datetime.fromisoformat(dados_p.get("hora_abertura", ""))).total_seconds()
        except (ValueError, TypeError):
            return float("inf")  # posição adotada de sessão anterior: idade desconhecida

    def _remover_posicao(self, sistema: str, ativo: str):
        """Tira a posição do controle local do sistema certo."""
        if sistema == "setup200":
            self.posicoes_setup200.pop(ativo, None)
        elif sistema == "donchian_adx":
            self.posicoes_donchian_adx.pop(ativo, None)
        else:
            self.posicoes_multiplas.pop(ativo, None)
            if self.posicao_ativa.get("ativo") == ativo:
                self.posicao_ativa = {"ativo": "NENHUMA", "tipo": "NEUTRO", "pnl": 0.0, "qtd": 0, "preco_entrada": 0.0}
        self._salvar_posicoes_spot()  # [v7 SPOT]

    async def _cancelar_protecoes(self, simbolo_par: str):
        """[v7 SPOT] Cancela as ordens abertas desse par (o stop do robô).
        Atenção: cancela também ordens manuais que você tenha nesse mesmo par."""
        try:
            await self.exchange.cancel_all_orders(simbolo_par)
        except Exception:
            pass  # sem ordens pra cancelar também cai aqui — não é erro grave

    async def _colocar_protecoes(self, simbolo_par: str, tipo_pos: str, qtd: float, sl_preco: float, tp_preco) -> bool:
        """[v7 SPOT] Registra o STOP de venda no Spot da Binance. Assim, se a
        internet cair ou o PC desligar, a própria Binance vende no stop.
        Usa STOP_LOSS (venda a mercado ao tocar o gatilho); se a moeda não
        aceitar, usa STOP_LOSS_LIMIT com folga de SPOT_STOP_LIMIT_FOLGA_PCT.
        O ALVO não é registrado: no Spot o stop já reserva as moedas, então
        o robô vigia o alvo e, ao atingir, cancela o stop e vende.
        Retorna True se o stop ficou registrado."""
        if not EXIGIR_STOP_NA_EXCHANGE:
            return True
        if tipo_pos != "LONG":
            return False  # no Spot só existe posição comprada
        await self._cancelar_protecoes(simbolo_par)
        try:
            qtd_ok = self._ajustar_qtd(simbolo_par, qtd)
            sl = float(self.exchange.price_to_precision(simbolo_par, sl_preco))
            tipos_aceitos = ((self.exchange.market(simbolo_par) or {}).get('info') or {}).get('orderTypes') or []
            if 'STOP_LOSS' in tipos_aceitos:
                await self.exchange.create_order(simbolo_par, 'market', 'sell', qtd_ok, None,
                                                 params={'stopLossPrice': sl})
            else:
                limite = float(self.exchange.price_to_precision(simbolo_par, sl_preco * (1 - SPOT_STOP_LIMIT_FOLGA_PCT / 100)))
                await self.exchange.create_order(simbolo_par, 'limit', 'sell', qtd_ok, limite,
                                                 params={'stopLossPrice': sl, 'timeInForce': 'GTC'})
        except Exception as e:
            self.escrever_log(f"[bold red]{simbolo_par}: a Binance recusou o STOP de proteção: {e}[/bold red]")
            return False
        return True

    async def _garantir_protecao_ou_fechar(self, ativo: str, simbolo_par: str, dados_p: dict, sistema: str) -> bool:
        """Logo após comprar: registra o stop na Binance. Se não conseguir,
        VENDE na hora — nunca fica moeda comprada sem stop na exchange.
        Retorna True se a posição continua aberta (e deve seguir vigiada)."""
        if await self._colocar_protecoes(simbolo_par, dados_p["tipo"], dados_p["qtd"], dados_p["sl_preco"], None):
            return True
        self.escrever_log(f"[bold red]{ativo}: sem stop na Binance — vendendo por segurança.[/bold red]")
        fechou = await self.fechar_posicao(ativo, dados_p, "SEM STOP NA EXCHANGE", sistema=sistema)
        return not fechou

    async def _saldo_moeda(self, ativo: str):
        """(livre, total) da moeda na carteira Spot, ou None se não deu pra consultar."""
        try:
            balance = await self.exchange.fetch_balance()
        except Exception:
            return None
        livre = float((balance.get('free', {}) or {}).get(ativo, 0.0) or 0.0)
        total = float((balance.get('total', {}) or {}).get(ativo, 0.0) or 0.0)
        return livre, total

    async def _posicao_existe_na_exchange(self, ativo: str, dados_p: dict = None):
        """[v7 SPOT] True = ainda tem a moeda, False = não tem mais, None = não deu pra saber."""
        saldo = await self._saldo_moeda(ativo)
        if saldo is None:
            return None
        qtd_esperada = (dados_p or {}).get("qtd", 0.0)
        if qtd_esperada <= 0:
            return saldo[1] > 0
        return saldo[1] >= qtd_esperada * 0.5

    def _qtd_liquida_compra(self, simbolo: str, ordem: dict, qtd_pedida: float) -> float:
        """[v7 SPOT] Quantidade que REALMENTE ficou na carteira após a compra.
        No Spot, sem BNB, a Binance cobra a taxa na própria moeda comprada
        (compra 1 BTC, recebe 0,999). Vender a quantidade "cheia" daria erro."""
        ordem = ordem or {}
        executada = float(ordem.get('filled') or qtd_pedida)
        base = simbolo.split('/')[0]
        taxas = ordem.get('fees') or ([ordem['fee']] if ordem.get('fee') else [])
        if taxas:
            taxa_na_moeda = sum(float(t.get('cost') or 0.0) for t in taxas if t and t.get('currency') == base)
        else:
            taxa_na_moeda = executada * TAKER_FEE_PCT / 100  # sem informação: assume o pior caso
        return self._ajustar_qtd(simbolo, executada - taxa_na_moeda)

    def _valor_posicoes_abertas(self) -> float:
        """Valor atual (em USDT) das moedas compradas pelo robô."""
        total = 0.0
        for dicionario in (self.posicoes_multiplas, self.posicoes_setup200, self.posicoes_donchian_adx):
            for ativo, d in dicionario.items():
                preco = self.precos_anteriores.get(f"{ativo}/USDT") or d.get("preco_entrada", 0.0)
                total += d.get("qtd", 0.0) * preco
        return total

    def _salvar_posicoes_spot(self):
        """[v7 SPOT] Grava as posições do robô em disco."""
        def limpar(d):
            return {a: {k: v for k, v in p.items() if not k.startswith('_')} for a, p in d.items()}
        try:
            with open(CAMINHO_ESTADO_POSICOES_SPOT, "w", encoding="utf-8") as f:
                json.dump({
                    "confluencia": limpar(self.posicoes_multiplas),
                    "setup200": limpar(self.posicoes_setup200),
                    "donchian_adx": limpar(self.posicoes_donchian_adx),
                }, f)
        except Exception as e:
            self.escrever_log(f"[yellow]Aviso: falha ao salvar posições em disco: {e}[/yellow]")

    def _carregar_posicoes_spot(self):
        """[v7 SPOT] Recupera as posições do robô ao reiniciar."""
        if not os.path.exists(CAMINHO_ESTADO_POSICOES_SPOT):
            return
        try:
            with open(CAMINHO_ESTADO_POSICOES_SPOT, "r", encoding="utf-8") as f:
                estado = json.load(f)
        except Exception as e:
            self.escrever_log(f"[yellow]Aviso: falha ao ler posições salvas: {e}[/yellow]")
            return
        self.posicoes_multiplas = dict(estado.get("confluencia", {}))
        self.posicoes_setup200 = dict(estado.get("setup200", {}))
        self.posicoes_donchian_adx = dict(estado.get("donchian_adx", {}))
        qtd = len(self.posicoes_multiplas) + len(self.posicoes_setup200) + len(self.posicoes_donchian_adx)
        if qtd:
            self.escrever_log(f"[cyan]{qtd} posição(ões) do robô recuperada(s) do disco — conferindo com a Binance...[/cyan]")

    async def _fechamento_externo(self, ativo: str, dados_p: dict, sistema: str) -> bool:
        """A posição sumiu da Binance sem o bot fechar (stop/alvo executados
        pela própria Binance, liquidação, ou fechamento manual no app).
        Registra no journal com preço ESTIMADO. Ignora posições com menos de
        60s (a Binance pode demorar um pouco pra listar uma posição nova)."""
        if self._idade_segundos(dados_p) < 60:
            return False
        simbolo_par = f"{ativo}/USDT"
        await self._cancelar_protecoes(simbolo_par)
        preco = self.precos_anteriores.get(simbolo_par) or dados_p["preco_entrada"]
        self._registrar_fechamento(ativo, dados_p, preco, "FECHADA NA EXCHANGE (stop/alvo/manual — preço estimado)", sistema)
        return True

    # ------------------------------------------------------------------
    # Execução de ordens
    # ------------------------------------------------------------------
    async def _preco_execucao_real(self, simbolo: str, ordem: dict, preco_ref):
        """[v7.1] Preço médio REAL em que a ordem executou. Em Futuros a Binance
        responde à ordem a mercado ANTES de informar o preço executado; então o
        robô consulta a ordem (e, em Futuros, a posição) até obter o preço real.
        Na Demo, o preço real chegou a ficar 0,18% longe do preço da tela."""
        ordem = ordem or {}
        if ordem.get('average'):
            return float(ordem['average'])
        oid = ordem.get('id')
        if oid:
            for _ in range(3):
                try:
                    consulta = await self.exchange.fetch_order(oid, simbolo)
                    if consulta and consulta.get('average'):
                        return float(consulta['average'])
                except Exception:
                    pass
                await asyncio.sleep(0.5)
        try:
            for pos in await self.exchange.fetch_positions([simbolo]):
                if float(pos.get('contracts') or 0) and pos.get('entryPrice'):
                    return float(pos['entryPrice'])
        except Exception:
            pass  # Spot não tem posições — normal cair aqui
        if preco_ref is not None:
            self.escrever_log(f"[yellow]{simbolo}: preço real de execução não confirmado; usando preço da tela.[/yellow]")
        return preco_ref

    def _log_slippage(self, ativo: str, preco_tela: float, preco_real: float):
        """[v7.1] Registra a diferença entre o preço da tela e o executado."""
        if preco_tela and preco_real:
            dif = (preco_real - preco_tela) / preco_tela * 100
            cor = "yellow" if abs(dif) >= 0.1 else "grey50"
            self.escrever_log(f"[{cor}]{ativo}: preço da tela {preco_tela:.6g} → executado {preco_real:.6g} ({dif:+.3f}%).[/{cor}]")

    async def fechar_posicao(self, ativo: str, dados_p: dict, motivo: str = "manual", preco_saida: float = None, sistema: str = None) -> bool:
        """[v7 SPOT] Para vender: (1) cancela o stop, que está "prendendo" as
        moedas; (2) vende o saldo REAL disponível; (3) se a venda falhar,
        recoloca o stop na hora. Só considera fechada quando a Binance
        confirma — senão a posição continua vigiada e o bot tenta de novo.
        Retorna True se fechou; quem chamou só remove a posição nesse caso."""
        if dados_p.get("_fechando"):
            return False  # já tem um fechamento em andamento pra essa posição
        dados_p["_fechando"] = True
        simbolo_par = f"{ativo}/USDT"
        preco_executado = None
        try:
            await self._cancelar_protecoes(simbolo_par)
            try:
                saldo = await self._saldo_moeda(ativo)
                qtd_venda = dados_p["qtd"] if saldo is None else min(dados_p["qtd"], saldo[0])
                qtd_venda = self._ajustar_qtd(simbolo_par, qtd_venda)
                if qtd_venda <= 0:
                    raise Exception("sem saldo livre da moeda para vender")
                ordem = await self.exchange.create_market_order(simbolo_par, 'sell', qtd_venda)
                preco_executado = await self._preco_execucao_real(simbolo_par, ordem, None)  # [v7.1]
            except Exception as e:
                existe = await self._posicao_existe_na_exchange(ativo, dados_p)
                if existe is False:
                    # A Binance já tinha vendido (o stop registrado lá executou).
                    motivo = f"{motivo} (já executado pela Binance)"
                else:
                    # Venda falhou e a moeda continua na carteira: recoloca o stop JÁ.
                    await self._colocar_protecoes(simbolo_par, "LONG", dados_p["qtd"], dados_p["sl_preco"], None)
                    agora = time.time()
                    if agora - self._ultimo_aviso_falha_fechamento.get(ativo, 0.0) > 30:
                        self._ultimo_aviso_falha_fechamento[ativo] = agora
                        self.escrever_log(
                            f"[bold red]FALHA ao vender {ativo} ({motivo}): {e}. A posição CONTINUA ABERTA — "
                            f"o stop foi recolocado na Binance e o bot tenta de novo a cada ciclo.[/bold red]"
                        )
                        self.notify(f"Falha ao vender {ativo} — tentando de novo.", title="⚠️ Fechamento", severity="error", timeout=8)
                        self.enviar_telegram(f"⚠️ <b>{ativo}</b> — falha ao vender ({motivo}). Robô tentando de novo.")
                    return False
        finally:
            dados_p["_fechando"] = False

        self._ultimo_aviso_falha_fechamento.pop(ativo, None)
        preco_saida = float(preco_executado or preco_saida or self.precos_anteriores.get(simbolo_par) or dados_p["preco_entrada"])
        self._registrar_fechamento(ativo, dados_p, preco_saida, motivo, sistema)
        return True

    # ------------------------------------------------------------------
    # [v8] Filtros de entrada e batimento cardíaco
    # ------------------------------------------------------------------
    def _carregar_eventos(self):
        """Lê eventos.json (no máximo a cada 5 min). Formato:
        {"eventos": [{"nome": "CPI EUA", "data_hora": "2026-10-15T09:30", "janela_min": 60}]}
        Horários no fuso do PC (Brasília)."""
        agora = time.time()
        if agora - self._eventos_lidos_em < 300:
            return
        self._eventos_lidos_em = agora
        if not os.path.exists(CAMINHO_EVENTOS):
            self._eventos = []
            return
        try:
            with open(CAMINHO_EVENTOS, "r", encoding="utf-8") as f:
                dados = json.load(f)
            lista = dados.get("eventos", []) if isinstance(dados, dict) else dados
            eventos = []
            for e in lista:
                eventos.append((datetime.fromisoformat(e["data_hora"]),
                                int(e.get("janela_min", EVENTOS_JANELA_MIN_PADRAO)),
                                e.get("nome", "evento")))
            self._eventos = eventos
        except Exception as ex:
            self.escrever_log(f"[yellow]Aviso: não consegui ler {CAMINHO_EVENTOS}: {ex}[/yellow]")

    def _evento_ativo(self):
        self._carregar_eventos()
        agora = datetime.now()
        for quando, janela, nome in self._eventos:
            if abs((agora - quando).total_seconds()) <= janela * 60:
                return f"{nome} ({quando:%d/%m %H:%M})"
        return None

    def _motivo_bloqueio_entrada(self, ativo: str):
        agora = time.time()
        if agora < self._freio_ate:
            return f"freio após {FREIO_STOPS_SEGUIDOS} perdas seguidas (até {datetime.fromtimestamp(self._freio_ate):%H:%M})"
        ate = self._cooldown_ate.get(ativo, 0.0)
        if agora < ate:
            return f"pausa da moeda após perda (até {datetime.fromtimestamp(ate):%H:%M})"
        evento = self._evento_ativo()
        if evento:
            return f"evento macro: {evento}"
        return None

    def _bloqueado(self, ativo: str) -> bool:
        motivo = self._motivo_bloqueio_entrada(ativo)
        if not motivo:
            return False
        chave = f"{ativo}|{motivo}"
        agora = time.time()
        if agora - self._ultimo_aviso_bloqueio.get(chave, 0.0) > 900:
            self._ultimo_aviso_bloqueio[chave] = agora
            self.escrever_log(f"[yellow]{ativo}: entrada bloqueada — {motivo}.[/yellow]")
        return True

    def _registrar_resultado_para_freio(self, ativo: str, pnl_liquido: float):
        agora = time.time()
        if pnl_liquido >= 0:
            self._stops_recentes = []  # um ganho quebra a sequência de perdas
            return
        self._cooldown_ate[ativo] = agora + COOLDOWN_MOEDA_HORAS * 3600
        self._stops_recentes = [t for t in self._stops_recentes if agora - t < FREIO_JANELA_HORAS * 3600] + [agora]
        if len(self._stops_recentes) >= FREIO_STOPS_SEGUIDOS:
            self._freio_ate = agora + FREIO_PAUSA_HORAS * 3600
            self._stops_recentes = []
            self.escrever_log(
                f"[bold red]FREIO: {FREIO_STOPS_SEGUIDOS} perdas em {FREIO_JANELA_HORAS}h — sem novas entradas por "
                f"{FREIO_PAUSA_HORAS}h (até {datetime.fromtimestamp(self._freio_ate):%H:%M}). Posições abertas seguem protegidas.[/bold red]"
            )
            self.enviar_telegram(f"⏸️ Freio acionado: {FREIO_STOPS_SEGUIDOS} perdas seguidas. Sem novas entradas por {FREIO_PAUSA_HORAS}h.")

    def _custo_compensa(self, ativo: str, preco: float, distancia_tp: float) -> bool:
        if preco <= 0:
            return False
        custo_pct = 2 * (TAKER_FEE_PCT + SLIPPAGE_ESTIMADO_PCT)
        alvo_pct = distancia_tp / preco * 100
        if alvo_pct >= CUSTO_MULTIPLO_MINIMO * custo_pct:
            return True
        chave = f"{ativo}|custo"
        agora = time.time()
        if agora - self._ultimo_aviso_bloqueio.get(chave, 0.0) > 900:
            self._ultimo_aviso_bloqueio[chave] = agora
            self.escrever_log(
                f"[grey50]{ativo}: alvo de {alvo_pct:.2f}% não cobre {CUSTO_MULTIPLO_MINIMO:.0f}x o custo "
                f"({custo_pct:.2f}%). Pulando.[/grey50]"
            )
        return False

    def _volume_confirma(self, simbolo: str) -> bool:
        if not FILTRO_VOLUME_ATIVO:
            return True
        relativo = self.volume_relativo_cache.get(simbolo)
        return relativo is not None and relativo >= FILTRO_VOLUME_MULTIPLO

    async def _funding_ok(self, ativo: str, simbolo_par: str, tipo_pos: str) -> bool:
        """[Futuros] Funding é uma tarifa cobrada a cada 8h. Taxa positiva: LONG paga.
        Não abre se a taxa contra a posição passar de FUNDING_MAX_CONTRA_PCT."""
        if not EH_FUTUROS:
            return True
        agora = time.time()
        cache = self._funding_cache.get(simbolo_par)
        if not cache or agora - cache[1] > 300:
            try:
                info = await self.exchange.fetch_funding_rate(simbolo_par)
                taxa = float(info.get('fundingRate') or 0.0) * 100
            except Exception:
                return True  # sem o dado, não bloqueia
            self._funding_cache[simbolo_par] = (taxa, agora)
        taxa = self._funding_cache[simbolo_par][0]
        contra = taxa if tipo_pos == "LONG" else -taxa
        if contra <= FUNDING_MAX_CONTRA_PCT:
            return True
        chave = f"{ativo}|funding"
        if agora - self._ultimo_aviso_bloqueio.get(chave, 0.0) > 900:
            self._ultimo_aviso_bloqueio[chave] = agora
            self.escrever_log(f"[grey50]{ativo}: funding de {taxa:+.4f}% contra {tipo_pos}. Pulando.[/grey50]")
        return False

    def _gravar_heartbeat(self):
        """Arquivo regravado a cada ciclo: se ficar mais de 2 min sem mudar, o robô parou."""
        try:
            posicoes = len(self.posicoes_multiplas) + len(self.posicoes_setup200) + len(self.posicoes_donchian_adx)
            with open(CAMINHO_HEARTBEAT, "w", encoding="utf-8") as f:
                json.dump({
                    "versao": VERSAO_ROBO,
                    "hora": datetime.now().isoformat(timespec="seconds"),
                    "ligado": self.robo_ligado,
                    "pausado_por_risco": self.robo_pausado_por_risco,
                    "freio_ate": datetime.fromtimestamp(self._freio_ate).isoformat(timespec="seconds") if self._freio_ate > time.time() else None,
                    "posicoes_abertas": posicoes,
                    "banca": round(self.banca_total, 2),
                    "lucro_dia": round(self.lucro_dia, 2),
                }, f)
        except Exception:
            pass

    def _registrar_fechamento(self, ativo: str, dados_p: dict, preco_saida: float, motivo: str, sistema: str = None):
        """Contabiliza um fechamento CONFIRMADO: journal, log, Telegram e travas diárias."""
        entrada = dados_p["preco_entrada"]
        if dados_p["tipo"] == "LONG":
            pnl_bruto = (preco_saida - entrada) * dados_p["qtd"]
        else:
            pnl_bruto = (entrada - preco_saida) * dados_p["qtd"]
        dados_p["pnl"] = pnl_bruto
        taxa_abertura = dados_p.get("fee_abertura", entrada * dados_p["qtd"] * TAKER_FEE_PCT / 100)
        taxa_fechamento = preco_saida * dados_p["qtd"] * TAKER_FEE_PCT / 100
        pnl_liquido = pnl_bruto - taxa_abertura - taxa_fechamento

        self.registrar_trade(pnl_liquido, ativo)
        self._registrar_resultado_para_freio(ativo, pnl_liquido)  # [v8]
        self.registrar_trade_csv(ativo, dados_p, preco_saida, pnl_bruto, taxa_abertura, taxa_fechamento, pnl_liquido, motivo)
        cor = "green" if pnl_liquido >= 0 else "red"
        self.escrever_log(f"[{cor}]Fechado {ativo} ({motivo}) PnL líquido: ${pnl_liquido:+.2f} (taxas: ${taxa_abertura + taxa_fechamento:.3f})[/{cor}]")

        severidade = "information" if pnl_liquido >= 0 else "warning"
        resultado_txt = "GANHOU" if pnl_liquido >= 0 else "PERDEU"
        self.notify(
            f"{ativo} — {resultado_txt} ${abs(pnl_liquido):.2f} ({motivo})",
            title="Posição fechada", severity=severidade, timeout=6,
        )
        emoji_resultado = "🟢" if pnl_liquido >= 0 else "🔴"
        sistema_txt = f" ({_telegram_sistema_amigavel(sistema)})" if sistema else ""
        resultado_amigavel = "✅ Lucro" if pnl_liquido >= 0 else "❌ Prejuízo"
        self.enviar_telegram(
            f"{emoji_resultado} <b>{ativo}</b> — Operação encerrada{sistema_txt}\n"
            f"{resultado_amigavel}: ${abs(pnl_liquido):.2f}\n"
            f"Motivo: {_telegram_motivo_amigavel(motivo)}"
        )

        # Trava de ganho/perda diária POR MODO (mesma regra da v5).
        if sistema:
            antes = self.lucro_dia_por_sistema.get(sistema, 0.0)
            depois = antes + pnl_liquido
            self.lucro_dia_por_sistema[sistema] = depois
            limite = STOP_GANHO_DIARIO_POR_SISTEMA.get(sistema)
            if limite is not None and antes < limite <= depois:
                self.escrever_log(
                    f"[bold cyan]🎯 Trava de ganho diária do modo '{sistema}' atingida: ${depois:.2f} "
                    f"(limite ${limite:.2f}). Esse modo não abre mais posição nova hoje — libera sozinho "
                    f"à meia-noite (ou tecla 'C').[/bold cyan]"
                )
                self.notify(
                    f"Modo '{sistema}': meta diária de ${limite:.2f} batida (${depois:.2f}). Sem novas entradas hoje.",
                    title="🎯 Trava de ganho batida", severity="information", timeout=10,
                )
            limite_perda = STOP_PERDA_DIARIA_POR_SISTEMA.get(sistema)
            if limite_perda is not None:
                limite_perda_neg = -abs(limite_perda)
                if antes > limite_perda_neg >= depois:
                    self.escrever_log(
                        f"[bold red]🛑 Trava de perda diária do modo '{sistema}' atingida: ${depois:.2f} "
                        f"(limite -${abs(limite_perda):.2f}). Esse modo não abre mais posição nova hoje — "
                        f"libera sozinho à meia-noite.[/bold red]"
                    )
                    self.notify(
                        f"Modo '{sistema}': limite de perda diária de -${abs(limite_perda):.2f} atingido (${depois:.2f}). Sem novas entradas hoje.",
                        title="🛑 Trava de perda batida", severity="error", timeout=10,
                    )
        self._salvar_estado_risco_diario()

    async def abrir_posicao(self, ativo: str, tipo_sinal: str, preco_atual: float, simbolo: str):
        if preco_atual <= 0:
            return
        atr = self.atr_cache.get(simbolo)
        if atr is None or atr <= 0:
            self.escrever_log(f"[yellow]{ativo}: ATR ainda não calculado, aguardando próximo ciclo.[/yellow]")
            return

        simbolo_par = f"{ativo}/USDT"
        alavancagem = min(self.alavancagem, ALAVANCAGEM_MAXIMA_PERMITIDA)  # [v6] teto absoluto
        if self._bloqueado(ativo) or not self._volume_confirma(simbolo):  # [v8]
            return
        try:
            mercado = self.exchange.market(simbolo_par)
            min_notional = mercado.get('limits', {}).get('cost', {}).get('min') or 5.0

            distancia_sl = atr * self.sl_atr_multiplo
            distancia_tp = distancia_sl * self.relacao_risco_retorno
            risco_usdt = self.banca_total * self.risco_por_trade_pct / 100
            risco_alvo_usdt = risco_usdt
            if not self._custo_compensa(ativo, preco_atual, distancia_tp):  # [v8]
                return
            if not await self._funding_ok(ativo, simbolo_par, "LONG" if tipo_sinal == "COMPRA" else "SHORT"):  # [v8]
                return

            if self.multi_moedas:
                risco_ja_comprometido = sum(p.get("risco_usdt", 0.0) for p in self.posicoes_multiplas.values())
            else:
                risco_ja_comprometido = self.posicao_ativa.get("risco_usdt", 0.0) if self.posicao_ativa["ativo"] != "NENHUMA" else 0.0
            limite_risco_total = self.banca_total * MAX_RISCO_TOTAL_PCT / 100
            if risco_ja_comprometido + risco_usdt > limite_risco_total:
                self.escrever_log(f"[yellow]{ativo}: risco total simultâneo passaria de {MAX_RISCO_TOTAL_PCT}% da banca. Pulando.[/yellow]")
                return

            qtd = self._ajustar_qtd(simbolo_par, risco_usdt / distancia_sl)
            valor_nocional = qtd * preco_atual

            max_valor_nocional = self.banca_total * alavancagem * MAX_EXPOSICAO_POR_POSICAO_PCT / 100
            if valor_nocional > max_valor_nocional:
                qtd = self._ajustar_qtd(simbolo_par, max_valor_nocional / preco_atual)
                valor_nocional = qtd * preco_atual
                risco_real = distancia_sl * qtd
                self.escrever_log(
                    f"[yellow]{ativo}: ATR muito baixo, posição limitada a ${valor_nocional:.2f} "
                    f"({MAX_EXPOSICAO_POR_POSICAO_PCT:.0f}% do poder de compra). Risco real: ${risco_real:.2f} "
                    f"(menor que o alvo de ${risco_usdt:.2f}).[/yellow]"
                )
                risco_usdt = risco_real

            if valor_nocional < min_notional:
                # [v6] Antes: aumentava a posição até o mínimo, mesmo que o risco
                # passasse muito de 1%. Agora: só aceita se o risco ficar até
                # 1,5x o planejado; senão, pula a operação.
                qtd_minima = self._ajustar_qtd(simbolo_par, min_notional * 1.02 / preco_atual)
                risco_com_minimo = distancia_sl * qtd_minima
                if qtd_minima <= 0 or risco_com_minimo > risco_alvo_usdt * TOLERANCIA_RISCO_MINIMO_EXCHANGE:
                    self.escrever_log(
                        f"[yellow]{ativo}: pulando — o tamanho mínimo da Binance levaria o risco a "
                        f"${risco_com_minimo:.2f} (planejado: ${risco_alvo_usdt:.2f}).[/yellow]"
                    )
                    return
                qtd = qtd_minima
                valor_nocional = qtd * preco_atual
                risco_usdt = risco_com_minimo
                self.escrever_log(f"[yellow]{ativo}: ajustado ao mínimo da exchange (${valor_nocional:.2f}), risco real: ${risco_usdt:.2f}.[/yellow]")

            if qtd <= 0:
                return

            margem_necessaria = valor_nocional / alavancagem
            margem_livre_com_seguranca = self.banca_livre * 0.95
            if margem_necessaria > margem_livre_com_seguranca:
                self.escrever_log(f"[red]{ativo}: margem necessária (${margem_necessaria:.2f}) maior que o saldo livre (${self.banca_livre:.2f}). Pulando.[/red]")
                return

            # [v7 SPOT] sem set_leverage: o Spot não tem alavancagem
            lado = 'buy' if tipo_sinal == "COMPRA" else 'sell'
            tipo_pos = "LONG" if tipo_sinal == "COMPRA" else "SHORT"
            ordem = await self.exchange.create_market_order(simbolo_par, lado, qtd)
            # [v6] usa o preço REAL executado (antes: preço da tela, sem slippage)
            preco_tela = preco_atual
            preco_atual = float(await self._preco_execucao_real(simbolo_par, ordem, preco_atual))  # [v7.1]
            self._log_slippage(ativo, preco_tela, preco_atual)
            qtd = self._qtd_liquida_compra(simbolo_par, ordem, qtd)  # [v7 SPOT] desconta a taxa cobrada na moeda

            if tipo_pos == "LONG":
                sl_preco = preco_atual - distancia_sl
                tp_preco = preco_atual + distancia_tp
            else:
                sl_preco = preco_atual + distancia_sl
                tp_preco = preco_atual - distancia_tp

            fee_abertura = qtd * preco_atual * TAKER_FEE_PCT / 100
            nova_posicao = {
                "tipo": tipo_pos, "pnl": 0.0, "qtd": qtd, "preco_entrada": preco_atual,
                "sl_preco": sl_preco, "tp_preco": tp_preco,
                "risco_usdt": risco_usdt, "fee_abertura": fee_abertura,
                "alavancagem": alavancagem,
                "hora_abertura": datetime.now().isoformat(timespec='seconds'),
            }
            if self.multi_moedas:
                self.posicoes_multiplas[ativo] = nova_posicao
            else:
                self.posicao_ativa = {"ativo": ativo, **nova_posicao}
            dados_registrados = self.posicoes_multiplas[ativo] if self.multi_moedas else self.posicao_ativa

            # [v6] stop registrado na Binance — sem stop, fecha na hora
            if not await self._garantir_protecao_ou_fechar(ativo, simbolo_par, dados_registrados, "confluencia"):
                self._remover_posicao("confluencia", ativo)
                return
            self._salvar_posicoes_spot()  # [v7 SPOT] salva logo após a compra

            cor = "green" if tipo_pos == "LONG" else "red"
            self.escrever_log(f"[{cor}]Aberto {tipo_pos} em {ativo} | Risco: ${risco_usdt:.2f} | SL: {sl_preco:,.4f} | TP: {tp_preco:,.4f}[/{cor}]")
            self.notify(f"{tipo_pos} em {ativo} — risco ${risco_usdt:.2f}", title="Posição aberta", severity="information", timeout=5)
            emoji_lado = "🟢" if tipo_pos == "LONG" else "🔴"
            direcao_txt = "Comprado (apostando que vai subir)" if tipo_pos == "LONG" else "Vendido (apostando que vai cair)"
            self.enviar_telegram(
                f"{emoji_lado} <b>{ativo}</b> — {direcao_txt}\n"
                f"Sistema: Confluência\n"
                f"💰 Entrada: {preco_atual:,.4f}\n"
                f"🛑 Stop (se bater aqui, fecha a perda): {sl_preco:,.4f}\n"
                f"🎯 Alvo (se bater aqui, fecha o lucro): {tp_preco:,.4f}\n"
                f"⚠️ Perda máxima nessa operação: ${risco_usdt:.2f}"
            )
        except Exception as e:
            self.escrever_log(f"[red]Erro ao abrir posição em {ativo}: {e}[/red]")

    # ------------------------------------------------------------------
    # Modo Setup $200 (Bollinger + RSI, banca dividida em fatias fixas)
    # Sistema INDEPENDENTE do resto do bot: própria tecla (O), próprio
    # dicionário de posições (posicoes_setup200), nunca conta pro limite
    # MAX_POSICOES_SIMULTANEAS nem pro risco em % da banca. O "limite de
    # operações simultâneas" pedido aqui é natural: no máximo 1 posição
    # por moeda (usa a fatia de $200 daquela moeda), e até N moedas ao
    # mesmo tempo, uma por ativo monitorado.
    # ------------------------------------------------------------------
    async def avaliar_e_abrir_setup200(self, ativo_nome: str, simbolo: str, preco: float):
        """Usa os MESMOS indicadores de Bollinger/RSI que o bot já calcula
        no timeframe de 15m (self.bollinger_cache / self.rsi_cache) e
        aplica a regra de reversão à média: só entra quando o preço está
        fora da banda E o RSI confirma o extremo (mesma lógica já testada
        em teste_bollinger_rsi.py)."""
        bb = self.bollinger_cache.get(simbolo)
        rsi_val = self.rsi_cache.get(simbolo)
        if not bb or rsi_val is None:
            return

        sinal = None
        if preco < bb["inferior"] and rsi_val < SETUP200_RSI_SOBREVENDIDO:
            sinal = "COMPRA"
        elif preco > bb["superior"] and rsi_val > SETUP200_RSI_SOBRECOMPRADO:
            sinal = "VENDA"

        if sinal == "VENDA":
            return  # [v7 SPOT] no Spot não existe venda a descoberto — só compra
        if sinal:
            # [v6] limite de posições simultâneas e de saldo (antes: sem limite)
            if len(self.posicoes_setup200) >= SETUP200_MAX_POSICOES_SIMULTANEAS:
                return
            if self.banca_livre < SETUP200_VALOR_USD * 1.05:
                if time.time() - self._ultimo_aviso_setup200 > 300:
                    self._ultimo_aviso_setup200 = time.time()
                    self.escrever_log(f"[yellow]Setup $200: saldo livre (${self.banca_livre:.2f}) insuficiente para nova posição de ${SETUP200_VALOR_USD:.0f}.[/yellow]")
                return
            await self.abrir_posicao_setup200(ativo_nome, sinal, preco, simbolo)

    async def abrir_posicao_setup200(self, ativo: str, tipo_sinal: str, preco_atual: float, simbolo: str):
        """Abre posição de tamanho FIXO ($200) com TP fixo ($3,50) e SL fixo
        ($3,50 = 1:1 com o TP), independente da gestão de risco em % da banca
        usada pelo resto do bot — conforme pedido explicitamente."""
        if preco_atual <= 0:
            return
        if self._bloqueado(ativo) or not await self._funding_ok(ativo, simbolo, "LONG" if tipo_sinal == "COMPRA" else "SHORT"):  # [v8]
            return
        try:
            mercado = self.exchange.market(simbolo)
            min_notional = mercado.get('limits', {}).get('cost', {}).get('min') or 5.0
            if SETUP200_VALOR_USD < min_notional:
                self.escrever_log(
                    f"[red]Setup $200 {ativo}: valor da posição (${SETUP200_VALOR_USD}) é menor que "
                    f"o mínimo exigido pela exchange (${min_notional}). Pulando essa moeda.[/red]"
                )
                return

            qtd = float(self.exchange.amount_to_precision(simbolo, SETUP200_VALOR_USD / preco_atual))
            if qtd <= 0:
                return

            lado = 'buy' if tipo_sinal == "COMPRA" else 'sell'
            tipo_pos = "LONG" if tipo_sinal == "COMPRA" else "SHORT"

            # [v7 SPOT] sem set_leverage: o Spot não tem alavancagem
            ordem = await self.exchange.create_market_order(simbolo, lado, qtd)
            preco_tela = preco_atual
            preco_atual = float(await self._preco_execucao_real(simbolo, ordem, preco_atual))  # [v7.1]
            self._log_slippage(ativo, preco_tela, preco_atual)
            qtd = self._qtd_liquida_compra(simbolo, ordem, qtd)  # [v7 SPOT] desconta a taxa cobrada na moeda

            variacao_tp_pct = SETUP200_TP_USD / SETUP200_VALOR_USD
            variacao_sl_pct = SETUP200_SL_USD / SETUP200_VALOR_USD
            if tipo_pos == "LONG":
                sl_preco = preco_atual * (1 - variacao_sl_pct)
                tp_preco = preco_atual * (1 + variacao_tp_pct)
            else:
                sl_preco = preco_atual * (1 + variacao_sl_pct)
                tp_preco = preco_atual * (1 - variacao_tp_pct)

            fee_abertura = SETUP200_VALOR_USD * TAKER_FEE_PCT / 100
            self.posicoes_setup200[ativo] = {
                "tipo": tipo_pos, "pnl": 0.0, "qtd": qtd, "preco_entrada": preco_atual,
                "sl_preco": sl_preco, "tp_preco": tp_preco,
                "risco_usdt": SETUP200_SL_USD, "fee_abertura": fee_abertura,
                "alavancagem": SETUP200_ALAVANCAGEM,
                "hora_abertura": datetime.now().isoformat(timespec='seconds'),
            }
            # [v6] stop registrado na Binance — sem stop, fecha na hora
            if not await self._garantir_protecao_ou_fechar(ativo, simbolo, self.posicoes_setup200[ativo], "setup200"):
                self._remover_posicao("setup200", ativo)
                return
            self._salvar_posicoes_spot()  # [v7 SPOT] salva logo após a compra
            cor = "green" if tipo_pos == "LONG" else "red"
            self.escrever_log(
                f"[{cor}]Setup $200: aberto {tipo_pos} em {ativo} @ {preco_atual:,.4f} | "
                f"TP {tp_preco:,.4f} | SL {sl_preco:,.4f} | fecha em até {SETUP200_TEMPO_MAXIMO_MINUTOS}min[/{cor}]"
            )
            self.notify(f"Setup $200: {tipo_pos} em {ativo}", title="Posição aberta", severity="information", timeout=5)
            emoji_lado = "🟢" if tipo_pos == "LONG" else "🔴"
            direcao_txt = "Comprado (apostando que vai subir)" if tipo_pos == "LONG" else "Vendido (apostando que vai cair)"
            self.enviar_telegram(
                f"{emoji_lado} <b>{ativo}</b> — {direcao_txt}\n"
                f"Sistema: Setup $200\n"
                f"💰 Entrada: {preco_atual:,.4f}\n"
                f"🛑 Stop (perda máxima, ~${SETUP200_SL_USD:.2f}): {sl_preco:,.4f}\n"
                f"🎯 Alvo (lucro esperado, ~${SETUP200_TP_USD:.2f}): {tp_preco:,.4f}\n"
                f"⏱️ Fecha sozinho em até {SETUP200_TEMPO_MAXIMO_MINUTOS} min se não bater nenhum dos dois"
            )
        except ccxt.InsufficientFunds as e:
            self.escrever_log(f"[red]Setup $200 {ativo}: saldo insuficiente: {e}[/red]")
        except ccxt.InvalidOrder as e:
            self.escrever_log(f"[red]Setup $200 {ativo}: ordem rejeitada (confira LOT_SIZE/MIN_NOTIONAL/PRICE_FILTER): {e}[/red]")
        except Exception as e:
            self.escrever_log(f"[red]Setup $200 {ativo}: erro ao abrir posição: {e}[/red]")

    async def checar_setup200(self, ativo: str, dados_p: dict, preco_atual: float):
        """Confere TP, SL, e o limite de 1 hora (critério de saída específico
        desse modo) — reaproveita a função genérica fechar_posicao(), a
        mesma usada pelo resto do bot, então o journal (trades_journal.csv)
        e o log ficam no mesmo padrão de sempre."""
        entrada = dados_p.get("preco_entrada", 0.0)
        if entrada <= 0 or preco_atual <= 0:
            return

        if dados_p["tipo"] == "LONG":
            dados_p["pnl"] = (preco_atual - entrada) * dados_p["qtd"]
            bateu_tp = preco_atual >= dados_p["tp_preco"]
            bateu_sl = preco_atual <= dados_p["sl_preco"]
        else:
            dados_p["pnl"] = (entrada - preco_atual) * dados_p["qtd"]
            bateu_tp = preco_atual <= dados_p["tp_preco"]
            bateu_sl = preco_atual >= dados_p["sl_preco"]

        tempo_esgotado = False
        hora_abertura_str = dados_p.get("hora_abertura", "")
        if hora_abertura_str:
            try:
                minutos_aberto = (datetime.now() - datetime.fromisoformat(hora_abertura_str)).total_seconds() / 60
                tempo_esgotado = minutos_aberto >= SETUP200_TEMPO_MAXIMO_MINUTOS
            except ValueError:
                pass

        # [v6] só remove do controle se a Binance confirmar o fechamento
        motivo = None
        if bateu_tp:
            motivo = "TAKE PROFIT"
        elif bateu_sl:
            motivo = "STOP LOSS"
        elif tempo_esgotado:
            motivo = f"TEMPO MAXIMO ({SETUP200_TEMPO_MAXIMO_MINUTOS}min)"
        if motivo and await self.fechar_posicao(ativo, dados_p, motivo, preco_saida=preco_atual, sistema="setup200"):
            self.posicoes_setup200.pop(ativo, None)

    async def avaliar_e_abrir_donchian_adx(self, ativo_nome: str, simbolo: str, preco: float):
        """Regra de entrada: só considera operar se o ADX confirmar que
        existe uma tendência forte o suficiente (> 25); dentro disso, entra
        COMPRADO se o preço rompeu para cima o canal de Donchian de entrada,
        ou VENDIDO se rompeu pra baixo (a estratégia foi validada nos dois
        sentidos, já que no mercado de Futuros dá pra vender antes de
        comprar). Mesma lógica de backtest_donchian_adx.py."""
        dados = self.donchian_adx_cache.get(simbolo)
        if not dados:
            return
        adx_val = dados.get("adx")
        if adx_val is None or adx_val < DONCHIAN_ADX_MINIMO:
            return  # sem tendência forte o suficiente -- não opera esse regime

        sinal = None
        if preco > dados["maxima_entrada"]:
            sinal = "COMPRA"
        elif preco < dados["minima_entrada"]:
            sinal = "VENDA"

        if sinal == "VENDA":
            return  # [v7 SPOT] só compra
        if sinal:
            await self.abrir_posicao_donchian_adx(ativo_nome, sinal, preco, simbolo)

    async def abrir_posicao_donchian_adx(self, ativo: str, tipo_sinal: str, preco_atual: float, simbolo: str):
        """Abre a posição com tamanho = % da banca LIVRE (12,5%, igual ao
        backtest) e alavancagem travada em 1x. Tem um stop-loss inicial
        fixo (1,5%) só como PROTEÇÃO contra um movimento violento contra a
        posição — a saída "normal" (deixar o lucro correr) é feita depois,
        em checar_donchian_adx, pelo canal de Donchian mais largo. SEM
        take-profit fixo, de propósito: foi essa mudança, validada no
        backtest, que fez a estratégia parar de cortar o lucro cedo demais
        durante tendências fortes."""
        if preco_atual <= 0:
            return
        if self._bloqueado(ativo) or not await self._funding_ok(ativo, simbolo, "LONG" if tipo_sinal == "COMPRA" else "SHORT"):  # [v8]
            return
        try:
            mercado = self.exchange.market(simbolo)
            min_notional = mercado.get('limits', {}).get('cost', {}).get('min') or 5.0
            valor_posicao = self.banca_livre * DONCHIAN_ADX_FRACAO_BANCA_PCT / 100
            if valor_posicao < min_notional:
                self.escrever_log(
                    f"[red]Donchian+ADX {ativo}: valor da posição (${valor_posicao:.2f}, "
                    f"{DONCHIAN_ADX_FRACAO_BANCA_PCT}% da banca livre) é menor que o mínimo "
                    f"exigido pela exchange (${min_notional}). Pulando.[/red]"
                )
                return

            qtd = float(self.exchange.amount_to_precision(simbolo, valor_posicao / preco_atual))
            if qtd <= 0:
                return

            lado = 'buy' if tipo_sinal == "COMPRA" else 'sell'
            tipo_pos = "LONG" if tipo_sinal == "COMPRA" else "SHORT"

            # [v7 SPOT] sem set_leverage: o Spot não tem alavancagem
            ordem = await self.exchange.create_market_order(simbolo, lado, qtd)
            preco_tela = preco_atual
            preco_atual = float(await self._preco_execucao_real(simbolo, ordem, preco_atual))  # [v7.1]
            self._log_slippage(ativo, preco_tela, preco_atual)
            qtd = self._qtd_liquida_compra(simbolo, ordem, qtd)  # [v7 SPOT] desconta a taxa cobrada na moeda

            variacao_sl_pct = DONCHIAN_ADX_SL_PCT / 100
            if tipo_pos == "LONG":
                sl_preco = preco_atual * (1 - variacao_sl_pct)
            else:
                sl_preco = preco_atual * (1 + variacao_sl_pct)

            fee_abertura = qtd * preco_atual * TAKER_FEE_PCT / 100
            self.posicoes_donchian_adx[ativo] = {
                "tipo": tipo_pos, "pnl": 0.0, "qtd": qtd, "preco_entrada": preco_atual,
                "sl_preco": sl_preco, "fee_abertura": fee_abertura,
                "alavancagem": DONCHIAN_ADX_ALAVANCAGEM,
                "hora_abertura": datetime.now().isoformat(timespec='seconds'),
            }
            # [v6] stop registrado na Binance (sem alvo fixo, por desenho da estratégia)
            if not await self._garantir_protecao_ou_fechar(ativo, simbolo, self.posicoes_donchian_adx[ativo], "donchian_adx"):
                self._remover_posicao("donchian_adx", ativo)
                return
            self._salvar_posicoes_spot()  # [v7 SPOT] salva logo após a compra
            cor = "green" if tipo_pos == "LONG" else "red"
            self.escrever_log(
                f"[{cor}]Donchian+ADX: aberto {tipo_pos} em {ativo} @ {preco_atual:,.2f} "
                f"(ADX={self.donchian_adx_cache.get(simbolo, {}).get('adx', 0):.1f}) | "
                f"SL inicial {sl_preco:,.2f} | saída pelo canal de Donchian (sem alvo fixo)[/{cor}]"
            )
            self.notify(f"Donchian+ADX: {tipo_pos} em {ativo}", title="Posição aberta", severity="information", timeout=5)
            emoji_lado = "🟢" if tipo_pos == "LONG" else "🔴"
            direcao_txt = "Comprado (apostando que vai subir)" if tipo_pos == "LONG" else "Vendido (apostando que vai cair)"
            self.enviar_telegram(
                f"{emoji_lado} <b>{ativo}</b> — {direcao_txt}\n"
                f"Sistema: Donchian+ADX\n"
                f"💰 Entrada: {preco_atual:,.2f}\n"
                f"🛑 Stop inicial (perda máxima por enquanto): {sl_preco:,.2f}\n"
                f"🎯 Sem alvo fixo — deixa o lucro correr enquanto a tendência durar"
            )
        except ccxt.InsufficientFunds as e:
            self.escrever_log(f"[red]Donchian+ADX {ativo}: saldo insuficiente: {e}[/red]")
        except ccxt.InvalidOrder as e:
            self.escrever_log(f"[red]Donchian+ADX {ativo}: ordem rejeitada (confira LOT_SIZE/MIN_NOTIONAL/PRICE_FILTER): {e}[/red]")
        except Exception as e:
            self.escrever_log(f"[red]Donchian+ADX {ativo}: erro ao abrir posição: {e}[/red]")

    async def checar_donchian_adx(self, ativo: str, dados_p: dict, preco_atual: float, simbolo: str):
        """Duas formas de sair: (1) stop-loss inicial de proteção (1,5%),
        pra limitar o estrago se o preço virar contra a posição logo depois
        de entrar; (2) saída "normal", quando o preço fecha do lado errado
        do canal de Donchian mais LARGO (saída) — é essa segunda que deixa
        o lucro correr durante uma tendência, em vez de sair no primeiro
        solavanco. Reaproveita fechar_posicao(), a mesma função de sempre,
        então o journal (trades_journal.csv) fica no mesmo padrão."""
        entrada = dados_p.get("preco_entrada", 0.0)
        if entrada <= 0 or preco_atual <= 0:
            return

        canais = self.donchian_adx_cache.get(simbolo, {})
        minima_saida = canais.get("minima_saida")
        maxima_saida = canais.get("maxima_saida")

        if dados_p["tipo"] == "LONG":
            dados_p["pnl"] = (preco_atual - entrada) * dados_p["qtd"]
            bateu_sl = preco_atual <= dados_p["sl_preco"]
            saiu_do_canal = minima_saida is not None and preco_atual < minima_saida
        else:
            dados_p["pnl"] = (entrada - preco_atual) * dados_p["qtd"]
            bateu_sl = preco_atual >= dados_p["sl_preco"]
            saiu_do_canal = maxima_saida is not None and preco_atual > maxima_saida

        motivo = "STOP LOSS" if bateu_sl else ("SAIU DO CANAL (Donchian)" if saiu_do_canal else None)
        if motivo and await self.fechar_posicao(ativo, dados_p, motivo, preco_saida=preco_atual, sistema="donchian_adx"):  # [v6]
            self.posicoes_donchian_adx.pop(ativo, None)

    async def checar_tp_sl(self, ativo: str, dados_p: dict, preco_atual: float):
        entrada = dados_p.get("preco_entrada", 0.0)
        sl_preco = dados_p.get("sl_preco")
        tp_preco = dados_p.get("tp_preco")
        if entrada <= 0 or preco_atual <= 0:
            return

        if dados_p["tipo"] == "LONG":
            dados_p["pnl"] = (preco_atual - entrada) * dados_p["qtd"]
            bateu_tp = tp_preco is not None and preco_atual >= tp_preco
            bateu_sl = sl_preco is not None and preco_atual <= sl_preco
        else:
            dados_p["pnl"] = (entrada - preco_atual) * dados_p["qtd"]
            bateu_tp = tp_preco is not None and preco_atual <= tp_preco
            bateu_sl = sl_preco is not None and preco_atual >= sl_preco

        # [v8] Proteger o lucro: ao ganhar BREAKEVEN_EM_R x o risco, leva o stop para a entrada (+custos).
        risco = dados_p.get("risco_usdt") or 0.0
        if (not dados_p.get("breakeven") and risco > 0 and not (bateu_tp or bateu_sl)
                and dados_p["pnl"] >= BREAKEVEN_EM_R * risco):
            folga = entrada * (2 * TAKER_FEE_PCT + SLIPPAGE_ESTIMADO_PCT) / 100
            novo_sl = entrada + folga if dados_p["tipo"] == "LONG" else entrada - folga
            if await self._colocar_protecoes(f"{ativo}/USDT", dados_p["tipo"], dados_p["qtd"], novo_sl, tp_preco):
                dados_p["sl_preco"] = novo_sl
                dados_p["breakeven"] = True
                self.escrever_log(f"[green]{ativo}: lucro protegido — stop movido para {novo_sl:,.6g} (entrada + custos).[/green]")
            else:
                # não conseguiu o stop novo: recoloca o original na hora
                await self._colocar_protecoes(f"{ativo}/USDT", dados_p["tipo"], dados_p["qtd"], sl_preco, tp_preco)

        motivo = "TAKE PROFIT" if bateu_tp else ("STOP LOSS" if bateu_sl else None)
        if motivo and await self.fechar_posicao(ativo, dados_p, motivo, preco_saida=preco_atual, sistema="confluencia"):  # [v6]
            self._remover_posicao("confluencia", ativo)

    async def _fechar_e_remover(self, ativo: str, dados_p: dict, motivo: str, sistema: str):
        """[v6] Fecha e só tira do controle se a Binance confirmar."""
        if await self.fechar_posicao(ativo, dados_p, motivo, sistema=sistema):
            self._remover_posicao(sistema, ativo)

    async def fechar_tudo(self, motivo: str):
        """[v6] Fecha TODAS as posições de todos os sistemas, independente do
        modo Multi-Moedas, e só remove do controle as que a Binance confirmar.
        As que falharem continuam vigiadas e o bot avisa."""
        for ativo, dados_p in list(self.posicoes_multiplas.items()):
            await self._fechar_e_remover(ativo, dados_p, motivo, "confluencia")
        if self.posicao_ativa["ativo"] != "NENHUMA":
            await self._fechar_e_remover(self.posicao_ativa["ativo"], self.posicao_ativa, motivo, "confluencia")
        for ativo in list(self.posicoes_arb.keys()):
            await self.fechar_posicao_funding(ativo, motivo)
        for ativo, dados_p in list(self.posicoes_setup200.items()):
            await self._fechar_e_remover(ativo, dados_p, motivo, "setup200")
        for ativo, dados_p in list(self.posicoes_donchian_adx.items()):
            await self._fechar_e_remover(ativo, dados_p, motivo, "donchian_adx")

        restantes = (list(self.posicoes_multiplas) + list(self.posicoes_setup200) + list(self.posicoes_donchian_adx)
                     + ([self.posicao_ativa["ativo"]] if self.posicao_ativa["ativo"] != "NENHUMA" else [])
                     + list(self.posicoes_arb))
        if restantes:
            self.escrever_log(
                f"[bold red]ATENÇÃO: {len(restantes)} posição(ões) NÃO fecharam: {', '.join(restantes)}. "
                f"Continuam vigiadas — confira na Binance e tente de novo.[/bold red]"
            )
            self.enviar_telegram(f"⚠️ Fechar tudo ({motivo}): não fecharam {', '.join(restantes)}. Conferir na Binance.")

    def _salvar_estado_funding_arb(self):
        """Grava self.posicoes_arb em disco. Chamado toda vez que uma
        posição abre, fecha, ou muda de estado (fecha/reabre a perna) — pra
        sobreviver a um reinício do bot sem se perder."""
        try:
            with open(CAMINHO_ESTADO_FUNDING_ARB, "w", encoding="utf-8") as f:
                json.dump(self.posicoes_arb, f)
        except Exception as e:
            self.escrever_log(f"[yellow]Aviso: falha ao salvar o estado do Funding Arb em disco: {e}[/yellow]")

    def _carregar_estado_funding_arb(self):
        """Lê o estado salvo (se existir) de volta pra self.posicoes_arb,
        logo na inicialização — antes de qualquer reconciliação rodar."""
        if not os.path.exists(CAMINHO_ESTADO_FUNDING_ARB):
            return
        try:
            with open(CAMINHO_ESTADO_FUNDING_ARB, "r", encoding="utf-8") as f:
                estado = json.load(f)
            if estado:
                self.posicoes_arb = estado
                self.escrever_log(
                    f"[cyan]Retomando {len(estado)} posição(ões) de Funding Arb de uma sessão anterior: "
                    f"{', '.join(estado.keys())}.[/cyan]"
                )
        except Exception as e:
            self.escrever_log(f"[yellow]Aviso: falha ao carregar o estado salvo do Funding Arb: {e}[/yellow]")

    async def _executar_pernas_arb(self, ativo: str, qtd: float, abrir: bool) -> tuple:
        """Executa as DUAS pontas do Funding Arb juntas (Spot + Futuro
        Perpétuo), na direção pedida:
          abrir=True  -> compra Spot + vende (short) o Futuro (ABRE a posição)
          abrir=False -> vende Spot + recompra o Futuro (FECHA a posição)

        Isso replica exatamente a suposição do backtest já validado
        (TelaFundingArb._simular): abrir/fechar sempre mexe nas DUAS pontas
        ao mesmo tempo, nunca só uma.

        SEGURANÇA: se a segunda perna falhar depois que a primeira já foi
        executada, tenta desfazer (rollback) a primeira imediatamente, pra
        nunca deixar a banca com uma exposição direcional destravada (sem
        hedge). Se até o rollback falhar, avisa em CAIXA ALTA pra você
        verificar manualmente na exchange — isso nunca deve ficar em
        silêncio."""
        simbolo_spot = f"{ativo}/USDT"
        simbolo_perp = f"{ativo}/USDT:USDT"
        lado_spot = 'buy' if abrir else 'sell'
        lado_perp = 'sell' if abrir else 'buy'
        params_perp = {} if abrir else {'reduceOnly': True}

        if lado_spot == 'sell':
            # CORRIGIDO: a quantidade guardada (pos["qtd"]) pode ser levemente
            # maior que o saldo REAL disponível no Spot — sobra de
            # arredondamento de quando a ordem de compra original foi
            # executada (a Binance pode preencher um pouquinho menos do que
            # o pedido). Antes de vender, confere o saldo de verdade e nunca
            # pede mais do que existe — evita o erro "insufficient balance"
            # travando o fechamento por causa de uma diferença de poeira
            # (frações de centavo).
            try:
                saldo = await self.exchange_spot.fetch_balance()
                disponivel = float(saldo.get('free', {}).get(ativo, 0.0) or 0.0)
                if disponivel < qtd:
                    qtd_ajustada = float(self.exchange_spot.amount_to_precision(simbolo_spot, disponivel))
                    if qtd_ajustada <= 0:
                        self.escrever_log(
                            f"[bold red]Funding Arb {ativo}: saldo real no Spot ({disponivel}) é praticamente "
                            f"zero — nada pra vender. Verifique manualmente se a posição já não foi fechada "
                            f"fora do bot.[/bold red]"
                        )
                        return False, {}
                    self.escrever_log(
                        f"[yellow]Funding Arb {ativo}: ajustando quantidade de venda de {qtd} para {qtd_ajustada} "
                        f"(saldo real disponível no Spot é um pouco menor, por arredondamento).[/yellow]"
                    )
                    qtd = qtd_ajustada
            except Exception as e:
                self.escrever_log(f"[yellow]Funding Arb {ativo}: não consegui conferir o saldo real antes de vender ({e}). Tentando com a quantidade original.[/yellow]")

        try:
            ordem_spot = await self.exchange_spot.create_market_order(simbolo_spot, lado_spot, qtd)
        except Exception as e:
            self.escrever_log(f"[red]Funding Arb {ativo}: a perna Spot ({lado_spot}) falhou: {e}. Nenhuma ordem foi enviada.[/red]")
            return False, {}

        try:
            ordem_perp = await self.exchange.create_market_order(simbolo_perp, lado_perp, qtd, params=params_perp)
        except Exception as e:
            self.escrever_log(
                f"[red]Funding Arb {ativo}: a perna Spot executou, mas o Futuro FALHOU ({e}). "
                f"Tentando desfazer a perna Spot para não sobrar exposição sem hedge...[/red]"
            )
            lado_rollback = 'sell' if lado_spot == 'buy' else 'buy'
            try:
                await self.exchange_spot.create_market_order(simbolo_spot, lado_rollback, qtd)
                self.escrever_log(f"[yellow]Rollback da perna Spot em {ativo} feito com sucesso — sem exposição residual.[/yellow]")
            except Exception as e2:
                self.escrever_log(
                    f"[bold red]CRÍTICO: o rollback da perna Spot em {ativo} TAMBÉM falhou ({e2}). "
                    f"Pode ter sobrado uma posição Spot solta, sem hedge. VERIFIQUE MANUALMENTE na Binance agora.[/bold red]"
                )
                self.notify(f"CRÍTICO: posição Spot solta em {ativo}, sem hedge! Verifique a exchange.",
                            title="Funding Arb", severity="error", timeout=None)
            return False, {}

        preco_spot = ordem_spot.get('average') or ordem_spot.get('price') or 0.0
        preco_perp = ordem_perp.get('average') or ordem_perp.get('price') or 0.0
        if not preco_spot:
            try:
                preco_spot = float((await self.exchange_spot.fetch_ticker(simbolo_spot)).get('last') or 0.0)
            except Exception:
                pass
        if not preco_perp:
            try:
                preco_perp = float((await self.exchange.fetch_ticker(simbolo_perp)).get('last') or 0.0)
            except Exception:
                pass
        return True, {"preco_spot": float(preco_spot or 0.0), "preco_perp": float(preco_perp or 0.0)}

    async def abrir_posicao_funding(self, ativo: str, janela_mm: int = JANELA_MM_FUNDING_ARB_PADRAO) -> bool:
        """Abre uma posição REAL (na sua conta Demo, ou em dinheiro real se
        MODO_REAL=True) de Funding Arbitrage: compra Spot + vende Futuro
        Perpétuo, valor igual nas duas pontas."""
        if ativo in self.posicoes_arb:
            self.escrever_log(f"[yellow]{ativo}: já existe uma posição de Funding Arb monitorada. Feche-a antes de abrir outra.[/yellow]")
            return False

        simbolo_spot = f"{ativo}/USDT"
        simbolo_perp = f"{ativo}/USDT:USDT"
        try:
            await self.exchange_spot.load_markets()
            mercado_spot = self.exchange_spot.market(simbolo_spot)
            mercado_perp = self.exchange.market(simbolo_perp)
        except Exception as e:
            self.escrever_log(f"[red]{ativo}: mercado Spot ou Futuro Perpétuo não encontrado ({e}). Confira se essa moeda tem os dois na Binance.[/red]")
            return False

        try:
            ticker_spot = await self.exchange_spot.fetch_ticker(simbolo_spot)
        except Exception as e:
            self.escrever_log(f"[red]{ativo}: erro ao buscar preço do Spot: {e}[/red]")
            return False
        preco_spot = float(ticker_spot.get('last') or 0.0)
        if preco_spot <= 0:
            self.escrever_log(f"[red]{ativo}: preço do Spot veio inválido, abortando.[/red]")
            return False

        min_notional_spot = mercado_spot.get('limits', {}).get('cost', {}).get('min') or 5.0
        min_notional_perp = mercado_perp.get('limits', {}).get('cost', {}).get('min') or 5.0
        notional_alvo = max(self.banca_livre * ALOCACAO_ARB_PCT_BANCA / 100, min_notional_spot, min_notional_perp)

        if notional_alvo > self.banca_livre * 0.95:
            self.escrever_log(
                f"[red]{ativo}: banca livre (${self.banca_livre:.2f}) não é suficiente pro tamanho mínimo "
                f"de ${notional_alvo:.2f} por perna nessa moeda.[/red]"
            )
            return False

        qtd_bruta = notional_alvo / preco_spot
        try:
            qtd = min(
                float(self.exchange_spot.amount_to_precision(simbolo_spot, qtd_bruta)),
                float(self.exchange.amount_to_precision(simbolo_perp, qtd_bruta)),
            )
        except Exception:
            qtd = round(qtd_bruta, 4)
        if qtd <= 0:
            self.escrever_log(f"[red]{ativo}: a quantidade calculada ficou zero. Aumente a alocação ou escolha outra moeda.[/red]")
            return False

        self.escrever_log(
            f"[cyan]Abrindo Funding Arb em {ativo}: comprando Spot + vendendo Futuro Perpétuo, "
            f"~{qtd} {ativo} por perna (~${qtd * preco_spot:.2f}).[/cyan]"
        )
        # [v6] trava a perna vendida (Futuro) em 1x: com alavancagem herdada de
        # outras operações, uma alta forte da moeda podia LIQUIDAR essa perna
        # mesmo com o Spot "protegendo" o conjunto.
        try:
            await self.exchange.set_leverage(1, simbolo_perp)
        except Exception as e:
            self.escrever_log(f"[red]{ativo}: não consegui travar o Futuro em 1x ({e}). Abortando por segurança.[/red]")
            return False
        ok, precos = await self._executar_pernas_arb(ativo, qtd, abrir=True)
        if not ok:
            return False

        notional_por_perna = qtd * precos["preco_spot"]
        custo_abertura_fechamento = notional_por_perna * (SPOT_TAKER_FEE_PCT + TAKER_FEE_PCT) / 100
        limiar_fracao = custo_abertura_fechamento / notional_por_perna if notional_por_perna > 0 else 0.0

        self.posicoes_arb[ativo] = {
            "qtd": qtd,
            "preco_spot_entrada": precos["preco_spot"],
            "preco_perp_entrada": precos["preco_perp"],
            "notional_por_perna": notional_por_perna,
            "custo_abertura_fechamento": custo_abertura_fechamento,
            "limiar_fracao": limiar_fracao,
            "janela_mm": janela_mm,
            "historico_janela": [],
            "posicao_aberta": True,
            "funding_recebido_acumulado": 0.0,
            "taxas_pagas_acumuladas": custo_abertura_fechamento,
            "ultimo_funding_ts_processado": None,
            "hora_abertura": datetime.now().isoformat(timespec='seconds'),
        }
        self._salvar_estado_funding_arb()
        self.escrever_log(
            f"[green]Funding Arb em {ativo} aberto — Spot @ {precos['preco_spot']:.4f}, "
            f"Futuro @ {precos['preco_perp']:.4f}. Monitorando a cada "
            f"{INTERVALO_MONITOR_FUNDING_ARB_SEGUNDOS // 60} min.[/green]"
        )
        self.notify(f"Funding Arb aberto em {ativo} (~${notional_por_perna:.2f} por perna)",
                    title="Funding Arb", severity="information", timeout=6)
        return True

    async def monitorar_posicoes_funding_arb(self):
        """Roda a cada poucos minutos (ver INTERVALO_MONITOR_FUNDING_ARB_SEGUNDOS).
        Para cada posição de Funding Arb aberta, busca o evento de funding
        REAL mais recente da Binance e aplica A MESMA regra de decisão já
        validada no backtest (TelaFundingArb._simular): fecha as duas pontas
        se a média móvel do funding ficar negativa além da tolerância
        financeira (custo de sair), reabre quando volta a ficar positiva."""
        for ativo, pos in list(self.posicoes_arb.items()):
            simbolo_perp = f"{ativo}/USDT:USDT"
            try:
                historico = await self.exchange.fetch_funding_rate_history(simbolo_perp, limit=1)
            except Exception as e:
                self.escrever_log(f"[yellow]Funding Arb {ativo}: erro ao checar funding ({e}). Tento de novo no próximo ciclo.[/yellow]")
                continue
            if not historico:
                continue

            evento = historico[-1]
            ts = evento.get('timestamp')
            if ts is None or ts == pos["ultimo_funding_ts_processado"]:
                continue  # nenhum evento novo desde a última checagem
            pos["ultimo_funding_ts_processado"] = ts

            taxa_funding = evento.get('fundingRate') or 0.0
            pos["historico_janela"].append(taxa_funding)
            if len(pos["historico_janela"]) > pos["janela_mm"]:
                pos["historico_janela"].pop(0)

            if pos["posicao_aberta"]:
                pos["funding_recebido_acumulado"] += pos["notional_por_perna"] * taxa_funding

            media_movel = sum(pos["historico_janela"]) / len(pos["historico_janela"])
            janela_completa = len(pos["historico_janela"]) == pos["janela_mm"]

            # A condição abaixo é IDÊNTICA à de TelaFundingArb._simular —
            # de propósito, pra garantir que o que roda ao vivo é o que foi
            # testado, não uma reinterpretação minha da mesma ideia.
            if pos["posicao_aberta"] and media_movel < 0 and janela_completa:
                if media_movel < -pos["limiar_fracao"]:
                    self.escrever_log(f"[yellow]Funding Arb {ativo}: funding negativo persistente (média {media_movel*100:.4f}%). Fechando as duas pontas...[/yellow]")
                    ok, _ = await self._executar_pernas_arb(ativo, pos["qtd"], abrir=False)
                    if ok:
                        pos["taxas_pagas_acumuladas"] += pos["custo_abertura_fechamento"]
                        pos["posicao_aberta"] = False
                        self.escrever_log(f"[yellow]Funding Arb {ativo}: fechado. Aguardando o funding voltar a ficar positivo pra reabrir sozinho.[/yellow]")
                # senão: ruído pequeno, absorve sem reagir — igual ao backtest
            elif not pos["posicao_aberta"] and media_movel >= 0:
                self.escrever_log(f"[cyan]Funding Arb {ativo}: funding voltou a ficar positivo. Reabrindo as duas pontas...[/cyan]")
                ok, precos = await self._executar_pernas_arb(ativo, pos["qtd"], abrir=True)
                if ok:
                    pos["taxas_pagas_acumuladas"] += pos["custo_abertura_fechamento"]
                    pos["posicao_aberta"] = True
                    pos["preco_spot_entrada"] = precos["preco_spot"]
                    pos["preco_perp_entrada"] = precos["preco_perp"]
                    self.escrever_log(f"[green]Funding Arb {ativo}: reaberto.[/green]")

            self._salvar_estado_funding_arb()

    async def fechar_posicao_funding(self, ativo: str, motivo: str = "manual"):
        """Encerra de vez uma posição de Funding Arb (usado pelo Pânico,
        pelo 'Zerar Todos', ou por você mesmo manualmente)."""
        pos = self.posicoes_arb.get(ativo)
        if not pos:
            return
        if pos["posicao_aberta"]:
            ok, _ = await self._executar_pernas_arb(ativo, pos["qtd"], abrir=False)
            if not ok:
                self.escrever_log(f"[bold red]Funding Arb {ativo}: falha ao encerrar ({motivo}). A posição continua sendo monitorada — verifique manualmente.[/bold red]")
                return
            pos["taxas_pagas_acumuladas"] += pos["custo_abertura_fechamento"]

        pnl_liquido = pos["funding_recebido_acumulado"] - pos["taxas_pagas_acumuladas"]
        self.registrar_trade_funding_arb_csv(ativo, pos, pnl_liquido, motivo)
        cor = "green" if pnl_liquido >= 0 else "red"
        self.escrever_log(f"[{cor}]Funding Arb {ativo} encerrado ({motivo}). PnL líquido acumulado: ${pnl_liquido:+.2f}[/{cor}]")
        self.notify(f"Funding Arb {ativo} encerrado — PnL ${pnl_liquido:+.2f}", title="Funding Arb", severity="information", timeout=6)
        emoji_resultado = "🟢" if pnl_liquido >= 0 else "🔴"
        resultado_amigavel = "✅ Lucro" if pnl_liquido >= 0 else "❌ Prejuízo"
        self.enviar_telegram(
            f"{emoji_resultado} <b>{ativo}</b> — Funding Arb encerrado\n"
            f"Motivo: {motivo}\n"
            f"{resultado_amigavel} acumulado: ${abs(pnl_liquido):.2f}"
        )
        del self.posicoes_arb[ativo]
        self._salvar_estado_funding_arb()

    def registrar_trade_funding_arb_csv(self, ativo: str, pos: dict, pnl_liquido: float, motivo: str):
        existe = os.path.exists(CAMINHO_JOURNAL_FUNDING_ARB)
        try:
            with open(CAMINHO_JOURNAL_FUNDING_ARB, "a", newline="", encoding="utf-8") as f:
                escritor = csv.writer(f)
                if not existe:
                    escritor.writerow([
                        "data_hora_fechamento", "ativo", "motivo", "qtd", "preco_spot_entrada",
                        "preco_perp_entrada", "notional_por_perna", "funding_recebido_acumulado",
                        "taxas_pagas_acumuladas", "pnl_liquido", "hora_abertura",
                    ])
                escritor.writerow([
                    datetime.now().isoformat(timespec="seconds"), ativo, motivo, pos["qtd"],
                    f"{pos['preco_spot_entrada']:.6f}", f"{pos['preco_perp_entrada']:.6f}",
                    f"{pos['notional_por_perna']:.2f}", f"{pos['funding_recebido_acumulado']:.4f}",
                    f"{pos['taxas_pagas_acumuladas']:.4f}", f"{pnl_liquido:.4f}", pos["hora_abertura"],
                ])
        except Exception as e:
            self.escrever_log(f"[yellow]Aviso: falha ao gravar journal de Funding Arb: {e}[/yellow]")

    async def _buscar_tickers_paralelo(self, simbolos: list) -> dict:
        """Busca o ticker de todas as moedas ao mesmo tempo (limitado por
        semáforo), em vez de uma de cada vez — reduz bastante a duração do ciclo."""
        semaforo = asyncio.Semaphore(8)
        resultado = {}

        async def _buscar_um(simbolo):
            async with semaforo:
                try:
                    resultado[simbolo] = await self.exchange.fetch_ticker(simbolo)
                except Exception as e:
                    resultado[simbolo] = None
                    msg = str(e)
                    if "does not have market symbol" in msg or "BadSymbol" in type(e).__name__:
                        self.simbolos_indisponiveis.add(simbolo)

        if simbolos:
            await asyncio.gather(*[_buscar_um(s) for s in simbolos])
        return resultado

    # ------------------------------------------------------------------
    # Ciclo principal
    # ------------------------------------------------------------------
    async def atualizar_mercado(self):
        if self.processando:
            return
        self.processando = True
        try:
            await self._atualizar_mercado_impl()
        finally:
            self.processando = False

    async def _atualizar_mercado_impl(self):
        table = self.query_one("#tabela-ativos", DataTable)

        inicio_ping = time.time()
        try:
            await self.exchange.fetch_time()
            self.ping = int((time.time() - inicio_ping) * 1000)
        except Exception as e:
            self.ping = 999
            self.escrever_log(f"[red]Erro de ping: {e}[/red]")

        try:
            balance = await self.exchange.fetch_balance()
            # [v7 SPOT] banca = USDT + valor atual das moedas compradas pelo robô
            self.banca_livre = float(balance.get('free', {}).get('USDT', 0.0) or 0.0)
            self.banca_total = float(balance.get('total', {}).get('USDT', 0.0) or 0.0) + self._valor_posicoes_abertas()
        except Exception as e:
            self.escrever_log(f"[red]Erro ao buscar saldo: {e}[/red]")

        if time.time() - self.ultima_reconciliacao >= RECONCILIACAO_SEGUNDOS:
            try:
                await self._sincronizar_horario()
            except Exception as e:
                # ANTES isso ficava em silêncio total — o que significa que,
                # se a resincronização automática estivesse falhando toda
                # hora, você nunca ia saber, e o offset ficaria desatualizado
                # pra sempre. Agora aparece no log, mesmo que só 1 vez por
                # ciclo, pra ajudar a diagnosticar a causa real do -1021.
                self.escrever_log(f"[yellow]Resincronização automática de horário falhou: {e}[/yellow]")
            await self.reconciliar_posicoes()

        await self.atualizar_rsi_real()
        await self.atualizar_tendencia()
        await self.atualizar_donchian_adx()
        self._verificar_virada_do_dia()  # [v6] zera o "dia" sozinho à meia-noite

        # Controle de risco: perda diária (% da banca)
        # [v6] considera também o prejuízo das posições AINDA ABERTAS
        limite_perda_diaria_usdt = -abs(self.banca_total * self.limite_perda_diaria_pct / 100)
        resultado_dia_com_abertas = self.lucro_dia + self._pnl_aberto_total()
        if self.robo_ligado and resultado_dia_com_abertas <= limite_perda_diaria_usdt:
            self.robo_ligado = False
            self.robo_pausado_por_risco = True
            self._salvar_estado_risco_diario()
            self.escrever_log(
                f"[red]LIMITE DE PERDA DIÁRIA ATINGIDO (realizado ${self.lucro_dia:.2f} + aberto "
                f"${self._pnl_aberto_total():.2f}, limite {self.limite_perda_diaria_pct}% da banca). Robô pausado.[/red]"
            )
            self.enviar_telegram("🛑 Limite de perda diária atingido — robô pausado (sem novas entradas hoje).")
            self.notify("Limite de perda diária atingido — robô pausado automaticamente.", title="⚠️ Proteção de risco", severity="error", timeout=10)

        # Controle de risco: drawdown a partir do pico do dia
        if self.lucro_dia > self.pico_lucro_dia:
            self.pico_lucro_dia = self.lucro_dia
        queda_do_pico = self.pico_lucro_dia - self.lucro_dia
        drawdown_maximo_usdt = self.banca_total * self.drawdown_maximo_pct / 100
        if self.robo_ligado and queda_do_pico >= drawdown_maximo_usdt:
            self.robo_ligado = False
            self.robo_pausado_por_risco = True
            self._salvar_estado_risco_diario()  # [v6]
            self.escrever_log(f"[red]DRAWDOWN DE ${queda_do_pico:.2f} A PARTIR DO PICO (${self.pico_lucro_dia:.2f}). Robô pausado.[/red]")
            self.notify("Drawdown a partir do pico do dia atingido — robô pausado automaticamente.", title="⚠️ Proteção de lucro", severity="error", timeout=10)

        total_posicoes_abertas = (
            len(self.posicoes_multiplas) if self.multi_moedas
            else (1 if self.posicao_ativa["ativo"] != "NENHUMA" else 0)
        )
        if self.multi_moedas:
            longs_abertos = sum(1 for p in self.posicoes_multiplas.values() if p["tipo"] == "LONG")
            shorts_abertos = sum(1 for p in self.posicoes_multiplas.values() if p["tipo"] == "SHORT")
        else:
            pos_unica = self.posicao_ativa
            longs_abertos = 1 if pos_unica["ativo"] != "NENHUMA" and pos_unica["tipo"] == "LONG" else 0
            shorts_abertos = 1 if pos_unica["ativo"] != "NENHUMA" and pos_unica["tipo"] == "SHORT" else 0

        simbolos_validos = [s for s in MOEDAS if s not in self.simbolos_indisponiveis]
        tickers_por_simbolo = await self._buscar_tickers_paralelo(simbolos_validos)

        ordem_atual = []
        for simbolo in simbolos_validos:
            ativo_nome = simbolo.split('/')[0]
            ordem_atual.append(ativo_nome)
            ticker = tickers_por_simbolo.get(simbolo)

            try:
                if ticker is None:
                    raise ValueError("ticker não retornado")
                preco = float(ticker.get('last', 0.0) or 0.0)
                volume = float(ticker.get('quoteVolume', 0.0) or 0.0)

                rsi_val = self.rsi_cache.get(simbolo)
                rsi_str = f"{rsi_val:.1f}" if rsi_val is not None else "..."

                acao, estrategia, _regime = self.avaliar_confluencia(simbolo, preco, volume)

                tendencia = self.tendencia_cache.get(simbolo)
                if acao == "COMPRA" and tendencia == "BAIXA":
                    estrategia, acao = estrategia + " [bloq. tendência]", "NEUTRO"
                elif acao == "VENDA" and tendencia == "ALTA":
                    estrategia, acao = estrategia + " [bloq. tendência]", "NEUTRO"
                if acao == "VENDA":  # [v7 SPOT] só compra
                    estrategia, acao = estrategia + " [Spot: só compra]", "NEUTRO"

                if acao == "COMPRA":
                    sinal_fmt = f"[green]{estrategia}[/green]"
                elif acao == "VENDA":
                    sinal_fmt = f"[red]{estrategia}[/red]"
                else:
                    sinal_fmt = f"[yellow]{estrategia}[/yellow]"

                vol_str = f"${volume/1_000_000:.1f}M" if volume > 1_000_000 else f"${volume/1_000:.1f}K"

                preco_ant = self.precos_anteriores.get(simbolo, preco)
                if preco > preco_ant:
                    preco_display = f"[green]{preco:,.4f} ▲[/green]"
                elif preco < preco_ant:
                    preco_display = f"[red]{preco:,.4f} ▼[/red]"
                else:
                    preco_display = f"{preco:,.4f}"
                if preco > 0:
                    self.precos_anteriores[simbolo] = preco

                bloqueado = ativo_nome in self.moedas_bloqueadas
                ativo_display = f"{ativo_nome} 🔒" if bloqueado else ativo_nome

                if tendencia == "ALTA":
                    tendencia_fmt = "[green]ALTA ▲[/green]"
                elif tendencia == "BAIXA":
                    tendencia_fmt = "[red]BAIXA ▼[/red]"
                else:
                    tendencia_fmt = "[grey50]...[/grey50]"

                posicao_str = "[grey50]Neutro[/grey50]"
                tipo_posicao_linha = None
                if self.multi_moedas:
                    if ativo_nome in self.posicoes_multiplas:
                        dados_p = self.posicoes_multiplas[ativo_nome]
                        await self.checar_tp_sl(ativo_nome, dados_p, preco)
                        if ativo_nome in self.posicoes_multiplas:
                            pnl_m = dados_p["pnl"]
                            cor_p = "green" if pnl_m >= 0 else "red"
                            posicao_str = f"[{cor_p}]{dados_p['tipo']} ${pnl_m:+.2f}[/{cor_p}]"
                            tipo_posicao_linha = dados_p["tipo"]
                else:
                    if self.posicao_ativa["ativo"] == ativo_nome:
                        dados_p = self.posicao_ativa
                        await self.checar_tp_sl(ativo_nome, dados_p, preco)
                        if self.posicao_ativa["ativo"] == ativo_nome:
                            pnl_s = dados_p["pnl"]
                            cor_p = "green" if pnl_s >= 0 else "red"
                            posicao_str = f"[{cor_p}]{dados_p['tipo']} ${pnl_s:+.2f}[/{cor_p}]"
                            tipo_posicao_linha = dados_p["tipo"]
                if bloqueado:
                    posicao_str = "[grey50]BLOQUEADO[/grey50]"

                # --- Modo Setup $200 (Bollinger+RSI, banca fixa) — totalmente
                # independente do sistema de confluência acima. Roda sempre,
                # mesmo em multi_moedas=False, porque usa seu próprio
                # dicionário de posições. ---
                if ativo_nome in self.posicoes_setup200:
                    await self.checar_setup200(ativo_nome, self.posicoes_setup200[ativo_nome], preco)
                    if ativo_nome in self.posicoes_setup200:
                        dados_s200 = self.posicoes_setup200[ativo_nome]
                        pnl_s200 = dados_s200["pnl"]
                        cor_s200 = "green" if pnl_s200 >= 0 else "red"
                        # Substitui o "Neutro" inteiro (mais curto, cabe na coluna sem
                        # truncar) — como os dois sistemas nunca operam a mesma moeda
                        # ao mesmo tempo, não tem risco de sobrescrever nada relevante.
                        posicao_str = f"[{cor_s200}]S200 {dados_s200['tipo']} ${pnl_s200:+.2f}[/{cor_s200}]"
                elif (self.setup200_ativo and self.robo_ligado and not bloqueado
                      and ativo_nome not in self.posicoes_multiplas
                      and ativo_nome not in self.posicoes_donchian_adx
                      and ativo_nome not in self.posicoes_arb
                      and self.lucro_dia_por_sistema.get("setup200", 0.0) < STOP_GANHO_DIARIO_SETUP200_USD
                      and self.lucro_dia_por_sistema.get("setup200", 0.0) > -STOP_PERDA_DIARIA_SETUP200_USD
                      and self.posicao_ativa["ativo"] != ativo_nome):
                    # Não abre se o sistema de confluência já estiver com uma
                    # posição aberta nessa MESMA moeda — os dois nunca podem
                    # operar o mesmo ativo ao mesmo tempo (evitaria duas
                    # ordens reais se misturando numa única posição na
                    # exchange, cada sistema achando que controla uma parte).
                    await self.avaliar_e_abrir_setup200(ativo_nome, simbolo, preco)

                # --- Modo Donchian+ADX (BTC, seguidor de tendência) ---
                # totalmente independente dos outros dois sistemas — só
                # opera BTC/USDT, própria tecla (Y), próprio dicionário de
                # posições. Nunca abre na mesma moeda que outro sistema já
                # esteja operando.
                if simbolo == DONCHIAN_ADX_SIMBOLO and ativo_nome in self.posicoes_donchian_adx:
                    await self.checar_donchian_adx(ativo_nome, self.posicoes_donchian_adx[ativo_nome], preco, simbolo)
                    if ativo_nome in self.posicoes_donchian_adx:
                        dados_dc = self.posicoes_donchian_adx[ativo_nome]
                        pnl_dc = dados_dc["pnl"]
                        cor_dc = "green" if pnl_dc >= 0 else "red"
                        posicao_str = f"[{cor_dc}]DC {dados_dc['tipo']} ${pnl_dc:+.2f}[/{cor_dc}]"
                elif (simbolo == DONCHIAN_ADX_SIMBOLO and self.donchian_adx_ativo and self.robo_ligado
                      and not bloqueado and ativo_nome not in self.posicoes_multiplas
                      and ativo_nome not in self.posicoes_setup200
                      and ativo_nome not in self.posicoes_arb
                      and self.posicao_ativa["ativo"] != ativo_nome):
                    await self.avaliar_e_abrir_donchian_adx(ativo_nome, simbolo, preco)

                valores = {
                    "ativo": ativo_display, "preco": preco_display, "rsi": rsi_str,
                    "tendencia": tendencia_fmt, "sinal": sinal_fmt, "vol": vol_str, "posicao": posicao_str,
                }
                self._atualizar_linha_tabela(table, ativo_nome, valores, tipo_posicao_linha)

                pode_abrir = (
                    self.robo_ligado and not bloqueado and acao != "NEUTRO"
                    and total_posicoes_abertas < self.max_posicoes_simultaneas
                )
                if acao == "COMPRA":
                    pode_abrir = pode_abrir and longs_abertos < MAX_LONGS_SIMULTANEOS
                elif acao == "VENDA":
                    pode_abrir = pode_abrir and shorts_abertos < MAX_SHORTS_SIMULTANEOS
                if self.multi_moedas:
                    pode_abrir = pode_abrir and ativo_nome not in self.posicoes_multiplas
                else:
                    pode_abrir = pode_abrir and self.posicao_ativa["ativo"] == "NENHUMA"
                # Mesma proteção do outro lado: não abre pela confluência se
                # o Setup $200, o Donchian+ADX ou o Funding Arb já estiverem
                # operando essa moeda (o Funding Arb usa essa MESMA posição de
                # Futuros como perna do hedge — abrir outra coisa em cima
                # bagunçaria o delta-neutro dele).
                pode_abrir = pode_abrir and ativo_nome not in self.posicoes_setup200
                pode_abrir = pode_abrir and ativo_nome not in self.posicoes_donchian_adx
                pode_abrir = pode_abrir and ativo_nome not in self.posicoes_arb
                # Trava de ganho diária do sistema de confluência (pedido
                # explicitamente): já bateu a meta do dia? Não abre mais.
                pode_abrir = pode_abrir and self.lucro_dia_por_sistema.get("confluencia", 0.0) < STOP_GANHO_DIARIO_CONFLUENCIA_USD
                # "Só 1 modo por vez" (pedido explicitamente, tecla K liga/desliga):
                # se a confluência estiver desligada, ela nunca abre posição nova
                # (mas continua GERENCIANDO as que já estavam abertas antes de desligar).
                pode_abrir = pode_abrir and self.confluencia_ativa

                if pode_abrir:
                    await self.abrir_posicao(ativo_nome, acao, preco, simbolo)
                    total_posicoes_abertas += 1
                    if acao == "COMPRA":
                        longs_abertos += 1
                    else:
                        shorts_abertos += 1

            except Exception as e:
                valores = {
                    "ativo": ativo_nome, "preco": "—", "rsi": "...", "tendencia": "[grey50]...[/grey50]",
                    "sinal": "[grey50]erro[/grey50]", "vol": "—", "posicao": "[grey50]Neutro[/grey50]",
                }
                self._atualizar_linha_tabela(table, ativo_nome, valores)
                self.escrever_log(f"[red]Erro ao processar {ativo_nome}: {e}[/red]")

        # Remove linhas de moedas que ficaram indisponíveis desde a última rodada
        for ativo_orfa in list(self.linhas_criadas - set(ordem_atual)):
            try:
                table.remove_row(ativo_orfa)
            except Exception:
                pass
            self.linhas_criadas.discard(ativo_orfa)

        self._atualizar_topo()
        self._atualizar_monitor()
        self._gravar_heartbeat()  # [v8]
        self._salvar_posicoes_spot()  # [v7 SPOT] guarda o estado a cada ciclo

    def _atualizar_linha_tabela(self, table: DataTable, ativo_nome: str, valores: dict, tipo_posicao: str = None):
        """Atualiza uma linha existente, ou cria se ainda não existir. Nunca usa
        table.clear() — isso preservaria zero rolagem/seleção do usuário.
        Quando há posição aberta nessa moeda, destaca a coluna do ativo com um
        fundo sólido (verde pra LONG, vermelho pra SHORT) — mais fácil de achar
        de relance as moedas com dinheiro em jogo numa tabela grande. (Colorir
        a LINHA inteira não dá pra fazer de forma sólida, porque a largura das
        colunas varia com o conteúdo — só o texto ficaria colorido, não a
        célula inteira, dando um efeito remendado.)"""
        if tipo_posicao == "LONG":
            fundo = "#1f5c1f"
        elif tipo_posicao == "SHORT":
            fundo = "#5c1f1f"
        else:
            fundo = None

        valores_finais = dict(valores)
        if fundo:
            # Preenche com espaços até uma largura generosa, pra cobrir a
            # célula inteira mesmo com nomes curtos tipo "BTC" ou "ETH".
            ativo_com_espaco = valores["ativo"].ljust(12)
            valores_finais["ativo"] = f"[on {fundo}]{ativo_com_espaco}[/on {fundo}]"

        if ativo_nome not in self.linhas_criadas:
            table.add_row(
                valores_finais["ativo"], valores_finais["preco"], valores_finais["rsi"], valores_finais["tendencia"],
                valores_finais["sinal"], valores_finais["vol"], valores_finais["posicao"], key=ativo_nome,
            )
            self.linhas_criadas.add(ativo_nome)
            self.ordem_ativos.append(ativo_nome)
        else:
            for nome, chave in self.col_keys.items():
                table.update_cell(ativo_nome, chave, valores_finais[nome], update_width=False)

    def _atualizar_topo(self):
        status_cor = "red" if self.robo_pausado_por_risco else ("green" if self.robo_ligado else "yellow")
        status_txt = "PAUSADO (risco)" if self.robo_pausado_por_risco else ("LIGADO" if self.robo_ligado else "PAUSADO")
        cor_lucro_dia = "green" if self.lucro_dia >= 0 else "red"
        multi_status = "[green]ON[/green]" if self.multi_moedas else "[red]OFF[/red]"

        total_posicoes = (
            len(self.posicoes_multiplas) if self.multi_moedas
            else (1 if self.posicao_ativa["ativo"] != "NENHUMA" else 0)
        )
        if self.multi_moedas:
            n_longs = sum(1 for p in self.posicoes_multiplas.values() if p["tipo"] == "LONG")
            n_shorts = sum(1 for p in self.posicoes_multiplas.values() if p["tipo"] == "SHORT")
        else:
            pos_unica = self.posicao_ativa
            n_longs = 1 if pos_unica["ativo"] != "NENHUMA" and pos_unica["tipo"] == "LONG" else 0
            n_shorts = 1 if pos_unica["ativo"] != "NENHUMA" and pos_unica["tipo"] == "SHORT" else 0

        modo_txt = "[bold red]DINHEIRO REAL[/bold red]" if MODO_REAL else "[bold cyan]Demo Trading[/bold cyan]"
        risco_calc_usdt = self.banca_total * self.risco_por_trade_pct / 100
        conf_min = self.confluencia_minima_banca_pequena if self.banca_total < BANCA_PEQUENA_LIMITE_USDT else self.confluencia_minima_padrao
        conf_label = "conta pequena" if self.banca_total < BANCA_PEQUENA_LIMITE_USDT else "padrão"

        texto = (
            f"Modo: {modo_txt} | Status: [bold {status_cor}]{status_txt}[/bold {status_cor}] | Ping: [yellow]{self.ping}ms[/yellow] | "
            f"Perfil: [cyan]{self.perfil_risco}[/cyan] | Multi-Moedas: {multi_status}\n"
            f"Posições: [cyan]{total_posicoes}/{self.max_posicoes_simultaneas}[/cyan] "
            f"(L:{n_longs}/{MAX_LONGS_SIMULTANEOS} S:{n_shorts}/{MAX_SHORTS_SIMULTANEOS}) | "
            f"Risco/trade: [cyan]{self.risco_por_trade_pct}% (~${risco_calc_usdt:.2f})[/cyan] | "
            f"R:R [green]1:{self.relacao_risco_retorno:.0f}[/green] | Taxa: [yellow]{TAKER_FEE_PCT}%/perna[/yellow]\n"
            f"💰 Banca: [cyan]${self.banca_total:,.2f}[/cyan]  |  🎯 Lucro do Dia: [{cor_lucro_dia}]${self.lucro_dia:.2f}[/{cor_lucro_dia}] "
            f"(pico: ${self.pico_lucro_dia:.2f})  |  🔒 Bloqueados: [red]{len(self.moedas_bloqueadas)}[/red]  |  "
            f"Confluência mín.: [cyan]{conf_min} ({conf_label})[/cyan]"
        )
        self.query_one("#topo-painel", Static).update(texto)
        self._atualizar_barra_risco()

    def _atualizar_barra_risco(self):
        """Mostra visualmente o quanto já foi 'gasto' do orçamento de risco do
        dia — o pior entre a perda direta e o drawdown a partir do pico —
        pra dar uma noção rápida de proximidade da pausa automática."""
        limite_perda_usdt = self.banca_total * self.limite_perda_diaria_pct / 100
        perda_atual_usdt = max(0.0, -self.lucro_dia)
        pct_perda = (perda_atual_usdt / limite_perda_usdt * 100) if limite_perda_usdt > 0 else 0.0

        drawdown_maximo_usdt = self.banca_total * self.drawdown_maximo_pct / 100
        queda_do_pico = max(0.0, self.pico_lucro_dia - self.lucro_dia)
        pct_drawdown = (queda_do_pico / drawdown_maximo_usdt * 100) if drawdown_maximo_usdt > 0 else 0.0

        progresso = min(100.0, max(pct_perda, pct_drawdown))
        try:
            barra = self.query_one("#barra-risco-dia", ProgressBar)
            barra.update(progress=progresso)
        except Exception:
            pass

        label = self.query_one("#label-risco-dia", Label)
        cor = "red" if progresso >= 80 else ("yellow" if progresso >= 50 else "green")
        label.update(f"[{cor}]Risco do dia: {progresso:.0f}%[/{cor}] [grey50](perda + drawdown)[/grey50]")

    def _atualizar_monitor(self):
        if self.multi_moedas:
            total_pnl_multi = sum(d["pnl"] for d in self.posicoes_multiplas.values())
            cor_pnl_m = "green" if total_pnl_multi >= 0 else "red"
            ativos_abertos = ", ".join(self.posicoes_multiplas.keys()) or "Nenhum"
            resumo = (
                f"Modo: [bold green]MULTI-MOEDAS[/bold green]\n"
                f"Abertos: [cyan]{ativos_abertos}[/cyan]\n"
                f"PnL Total Aberto: [{cor_pnl_m}]${total_pnl_multi:.2f}[/{cor_pnl_m}]\n"
            )
        else:
            pos = self.posicao_ativa
            cor_pnl = "green" if pos["pnl"] >= 0 else "red"
            resumo = f"Ativo Atual: [bold cyan]{pos['ativo']}[/bold cyan] ({pos['tipo']}) | PnL: [{cor_pnl}]${pos['pnl']:.2f}[/{cor_pnl}]\n"

        historico_trades = "\n".join(self.trades_recentes[:6]) if self.trades_recentes else "[grey50]Nenhum trade fechado ainda[/grey50]"
        texto = (
            "[bold yellow]Monitor de Ordens e PnL[/bold yellow]\n\n"
            + resumo
            + "\n[bold yellow]--- Histórico de Trades ---[/bold yellow]\n" + historico_trades
            + f"\n\n[grey50]Journal completo em: {CAMINHO_JOURNAL}[/grey50]"
            + "\n[grey50]Selecione uma linha na tabela: [X] fecha, [B] trava/destrava[/grey50]"
        )
        self.query_one("#monitor-texto", Static).update(texto)

    # ------------------------------------------------------------------
    # Atalhos de teclado
    # ------------------------------------------------------------------
    def action_toggle_robo(self):
        self.robo_ligado = not self.robo_ligado
        self.robo_pausado_por_risco = False
        estado = "LIGADO" if self.robo_ligado else "PAUSADO"
        cor = "green" if self.robo_ligado else "yellow"
        self.escrever_log(f"[{cor}]Robô {estado}[/{cor}]")

    def action_panico(self):
        self.robo_ligado = False
        self.escrever_log("[red]⚠️ PÂNICO: fechando tudo e pausando o robô.[/red]")
        self.notify("Fechando todas as posições e pausando o robô.", title="⚠️ PÂNICO", severity="error", timeout=8)
        self.run_worker(self.fechar_tudo("PÂNICO"), exclusive=False)

    def action_zerar_todos(self):
        self.escrever_log("[yellow]Fechando todas as posições (robô continua ligado)...[/yellow]")
        self.run_worker(self.fechar_tudo("manual"), exclusive=False)

    def action_zerar_lucro(self):
        # [v6] A tecla C serve para liberar travas de GANHO. Ela não pode mais
        # apagar uma PERDA do dia (isso burlava o limite de perda diária).
        if self.lucro_dia < 0:
            self.escrever_log(
                "[yellow]Não é possível zerar uma PERDA do dia — ela zera sozinha à meia-noite. "
                "(A tecla C serve para liberar as travas de GANHO.)[/yellow]"
            )
            return
        self.lucro_dia = 0.0
        self.pico_lucro_dia = 0.0
        # mantém as travas de PERDA por modo; libera só as de ganho
        self.lucro_dia_por_sistema = {k: v for k, v in self.lucro_dia_por_sistema.items() if v < 0}
        self._salvar_estado_risco_diario()
        self.vitorias = 0
        self.derrotas = 0
        self.trades_recentes.clear()
        self.escrever_log("[cyan]Lucro do dia zerado (inclui as travas de ganho por modo).[/cyan]")

    def action_trocar_risco(self):
        # [v6] 10x removido. Ciclo: 2x -> 3x -> 5x (teto ALAVANCAGEM_MAXIMA_PERMITIDA).
        if self.alavancagem == 2:
            self.perfil_risco, self.alavancagem = "MODERADO (3x)", 3
        elif self.alavancagem == 3:
            self.perfil_risco, self.alavancagem = "AGRESSIVO (5x)", 5
        else:
            self.perfil_risco, self.alavancagem = "CONSERVADOR (2x)", 2
        self.alavancagem = min(self.alavancagem, ALAVANCAGEM_MAXIMA_PERMITIDA)
        self.escrever_log(f"[cyan]Perfil alterado para: {self.perfil_risco}[/cyan]")

    def action_toggle_multi(self):
        # [v6] Trocar de modo com posições abertas fazia o bot PARAR de vigiar
        # o stop delas. Agora a troca só é permitida sem posições da Confluência.
        if self.posicoes_multiplas or self.posicao_ativa["ativo"] != "NENHUMA":
            self.escrever_log("[yellow]Feche as posições da Confluência antes de trocar o modo Multi-Moedas.[/yellow]")
            self.notify("Feche as posições da Confluência antes de trocar o modo Multi-Moedas.",
                        title="Multi-Moedas", severity="warning", timeout=6)
            return
        self.multi_moedas = not self.multi_moedas
        self.escrever_log(f"[cyan]Multi-Moedas: {'HABILITADO' if self.multi_moedas else 'DESABILITADO'}[/cyan]")

    def _ativo_selecionado(self):
        table = self.query_one("#tabela-ativos", DataTable)
        if table.cursor_row is None or table.cursor_row >= len(self.ordem_ativos):
            return None
        try:
            row_key, _ = table.coordinate_to_cell_key(table.cursor_coordinate)
            return row_key.value
        except Exception:
            return None

    def action_fechar_selecionado(self):
        """Fecha manualmente a posição da linha selecionada, seja ela de
        qual sistema for (confluência, Setup $200, Donchian+ADX ou Funding
        Arb) — CORRIGIDO: antes só olhava pra posicoes_multiplas/posicao_ativa
        (confluência), então apertar X numa moeda do Setup $200 ou do
        Donchian+ADX não fazia nada."""
        ativo = self._ativo_selecionado()
        if not ativo:
            return
        # [v6] não tira a posição do controle ANTES de a Binance confirmar
        if self.multi_moedas and ativo in self.posicoes_multiplas:
            self.escrever_log(f"[cyan]Fechando {ativo} manualmente (confluência)...[/cyan]")
            self.run_worker(self._fechar_e_remover(ativo, self.posicoes_multiplas[ativo], "manual", "confluencia"), exclusive=False)
        elif not self.multi_moedas and self.posicao_ativa["ativo"] == ativo:
            self.escrever_log(f"[cyan]Fechando {ativo} manualmente (confluência)...[/cyan]")
            self.run_worker(self._fechar_e_remover(ativo, self.posicao_ativa, "manual", "confluencia"), exclusive=False)
        elif ativo in self.posicoes_setup200:
            self.escrever_log(f"[cyan]Fechando {ativo} manualmente (Setup $200)...[/cyan]")
            self.run_worker(self._fechar_e_remover(ativo, self.posicoes_setup200[ativo], "manual", "setup200"), exclusive=False)
        elif ativo in self.posicoes_donchian_adx:
            self.escrever_log(f"[cyan]Fechando {ativo} manualmente (Donchian+ADX)...[/cyan]")
            self.run_worker(self._fechar_e_remover(ativo, self.posicoes_donchian_adx[ativo], "manual", "donchian_adx"), exclusive=False)
        elif ativo in self.posicoes_arb:
            self.escrever_log(f"[cyan]Fechando {ativo} manualmente (Funding Arb)...[/cyan]")
            self.run_worker(self.fechar_posicao_funding(ativo, "manual"), exclusive=False)
        else:
            self.escrever_log(f"[yellow]{ativo} não tem posição aberta.[/yellow]")

    def action_bloquear_selecionado(self):
        ativo = self._ativo_selecionado()
        if not ativo:
            return
        if ativo in self.moedas_bloqueadas:
            self.moedas_bloqueadas.remove(ativo)
            self.escrever_log(f"[cyan]{ativo} DESBLOQUEADO ✅[/cyan]")
        else:
            self.moedas_bloqueadas.append(ativo)
            self.escrever_log(f"[cyan]{ativo} BLOQUEADO 🔒[/cyan]")

    def action_abrir_configuracoes(self):
        self.push_screen(TelaConfiguracoes())

    def action_abrir_backtest(self):
        self.push_screen(TelaBacktest())

    def action_abrir_funding_arb(self):
        self.push_screen(TelaFundingArb())

    def action_abrir_cash_carry(self):
        self.push_screen(TelaCashCarry())

    def _desligar_outros_modos(self, exceto: str):
        """"Só 1 modo por vez" (pedido explicitamente): ao ligar um modo,
        desliga os outros dois automaticamente — evita operar sem querer
        mais de um sistema ao mesmo tempo. Isso só bloqueia ABERTURA de
        posição nova em cada modo desligado; uma posição que já estava
        aberta continua sendo gerenciada (TP/SL/canal) normalmente até
        fechar sozinha."""
        desligados = []
        if exceto != "confluencia" and self.confluencia_ativa:
            self.confluencia_ativa = False
            desligados.append("Confluência")
        if exceto != "setup200" and self.setup200_ativo:
            self.setup200_ativo = False
            desligados.append("Setup $200")
        if exceto != "donchian_adx" and self.donchian_adx_ativo:
            self.donchian_adx_ativo = False
            desligados.append("Donchian+ADX")
        if desligados:
            self.escrever_log(
                f"[yellow]Modo exclusivo: desligando automaticamente {', '.join(desligados)} "
                f"pra deixar só um modo ativo por vez.[/yellow]"
            )

    def action_toggle_confluencia(self):
        self.confluencia_ativa = not self.confluencia_ativa
        if self.confluencia_ativa:
            self._desligar_outros_modos(exceto="confluencia")
        estado = "LIGADA" if self.confluencia_ativa else "DESLIGADA"
        cor = "green" if self.confluencia_ativa else "yellow"
        self.escrever_log(f"[{cor}]Confluência (sistema normal de indicadores): {estado}.[/{cor}]")
        self.notify(f"Confluência: {estado}", title="Modo Confluência", severity="information", timeout=5)

    def action_toggle_setup200(self):
        self.setup200_ativo = not self.setup200_ativo
        if self.setup200_ativo:
            self._desligar_outros_modos(exceto="setup200")
        estado = "LIGADO" if self.setup200_ativo else "DESLIGADO"
        cor = "green" if self.setup200_ativo else "yellow"
        self.escrever_log(
            f"[{cor}]Modo Setup $200 (Bollinger+RSI, banca fixa por moeda): {estado}.[/{cor}] "
            f"{'Lembre-se: robô precisa estar Ligado (tecla L) para esse modo operar de verdade.' if self.setup200_ativo else ''}"
        )
        self.notify(f"Setup $200: {estado}", title="Modo Setup $200", severity="information", timeout=5)

    def action_toggle_donchian_adx(self):
        self.donchian_adx_ativo = not self.donchian_adx_ativo
        if self.donchian_adx_ativo:
            self._desligar_outros_modos(exceto="donchian_adx")
        estado = "LIGADO" if self.donchian_adx_ativo else "DESLIGADO"
        cor = "green" if self.donchian_adx_ativo else "yellow"
        self.escrever_log(
            f"[{cor}]Modo Donchian+ADX (BTC, seguidor de tendência, validado em 3 regimes históricos): "
            f"{estado}.[/{cor}] "
            f"{'Lembre-se: robô precisa estar Ligado (tecla L) para esse modo operar de verdade. Confirme que está em Demo Trading antes de deixar rodando.' if self.donchian_adx_ativo else ''}"
        )
        self.notify(f"Donchian+ADX (BTC): {estado}", title="Modo Donchian+ADX", severity="information", timeout=5)

    async def action_sincronizar_horario_manual(self):
        self.escrever_log("[cyan]Sincronizando horário com a Binance...[/cyan]")
        try:
            await self._sincronizar_horario()
            self.escrever_log("[green]Horário sincronizado com sucesso nas duas conexões (Futuro e Spot).[/green]")
            self.notify("Horário sincronizado com a Binance.", title="Sincronização", severity="information", timeout=4)
        except Exception as e:
            self.escrever_log(f"[red]Falha ao sincronizar horário: {e}[/red]")
            self.notify(f"Falha ao sincronizar horário: {e}", title="Sincronização", severity="error", timeout=6)


if __name__ == "__main__":
    # [v6] sem chave configurada, não inicia (evita rodar "no escuro")
    if not API_KEY or not SECRET_KEY:
        print("=" * 60)
        print("Chaves da Binance NÃO configuradas.")
        print("No PowerShell do Windows, rode (uma única vez):")
        print('  setx BINANCE_API_KEY "sua_chave"')
        print('  setx BINANCE_API_SECRET "seu_segredo"')
        print("Depois FECHE e abra o terminal de novo e rode o bot outra vez.")
        print("=" * 60)
        raise SystemExit(1)
    if MODO_REAL:
        print("=" * 60)
        print("ATENÇÃO: MODO_REAL = True")
        print("Este bot vai operar com DINHEIRO REAL na sua conta da Binance,")
        print("não mais no ambiente de testes (Demo Trading).")
        print("=" * 60)
        resposta = input("Digite CONFIRMO em maiúsculas para continuar, ou qualquer outra coisa para cancelar: ").strip()
        if resposta != "CONFIRMO":
            print("Operação cancelada. Nenhuma ordem será executada.")
            raise SystemExit(0)

    app = BotTraderApp()
    app.run()






