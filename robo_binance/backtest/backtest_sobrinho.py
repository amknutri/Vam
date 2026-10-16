"""
Backtest da estratégia do robô MORS-VIDA-SEM-FOME.py ("Cibele") — 2 anos, dados públicos.

Reproduz a lógica do robô do sobrinho do Vamberto, lida do código-fonte:
  - Velas de 5 minutos, 11 moedas de Futuros.
  - Sinal "contra o RSI" na vela fechada:
      VENDA  se RSI(14) >= 70, %B(20) >= 0,98, fechamento > anterior e volume >= 75% da média de 20;
      COMPRA se RSI(14) <= 30, %B(20) <= 0,02, fechamento < anterior e volume >= 75% da média de 20.
  - Até 2 entradas por moeda no mesmo sentido (preço médio); stop e alvo refeitos a partir da última entrada.
  - Stop manual de 8% e alvo manual de 0,33% (padrão do código e configuração da tela),
    ou stop automático (topo/fundo de 20 velas +- 0,5 ATR) com alvo de 1,5x o risco.
  - "Proteção de superávit": quando a soma dos GANHOS (perdas não contam) chega a R$ 333,33,
    fecha tudo e pausa 4 horas; depois da pausa a soma volta a zero.
  - Caixa fixo de R$ 10.000 (o robô não soma o resultado ao caixa); margem cruzada.

NÃO envia ordem nenhuma e NÃO usa chave: só baixa dados públicos.

Uso (PowerShell), na mesma pasta do backtest anterior:
    cd C:\\RoboBacktest
    python backtest_sobrinho.py
"""
import csv
import json
import os
import sys
import time
from datetime import datetime, timedelta, timezone

import backtest_estrategias as base  # mesma pasta: reaproveita download de funding e utilidades

INICIO = os.environ.get("BT_INICIO", "2024-10-01")
FIM = os.environ.get("BT_FIM", "2026-10-01")
AQUECIMENTO_DIAS = 3
MOEDAS = os.environ.get("BT_MOEDAS", "BTC,ETH,SOL,BNB,XRP,ADA,AVAX,DOT,LINK,LTC,TRX").split(",")
PASTA_DADOS = "dados"
CINCO_MIN = 5 * 60 * 1000
UMA_HORA = 3600 * 1000

CAIXA_BRL = 10000.0
DOLAR = 5.20                 # cotação padrão do código (o robô atualiza; aqui fica fixa)
GATILHO_SUPERAVIT_BRL = 333.33
PAUSA_MS = 4 * UMA_HORA
MAX_ENTRADAS = 2


def log(m):
    print(m, flush=True)


def baixar_velas_5m(simbolo, inicio_ms, fim_ms):
    os.makedirs(PASTA_DADOS, exist_ok=True)
    arq = os.path.join(PASTA_DADOS, f"{simbolo}_5m.csv")
    if os.path.exists(arq):
        with open(arq, encoding="utf-8") as f:
            v = [[int(r[0])] + [float(x) for x in r[1:6]] for r in csv.reader(f)]
        if v and v[0][0] <= inicio_ms + CINCO_MIN and v[-1][0] >= fim_ms - 2 * CINCO_MIN:
            return v, "cache"
    velas, t = [], inicio_ms
    try:
        while t < fim_ms:
            url = (f"https://fapi.binance.com/fapi/v1/klines?symbol={simbolo}&interval=5m"
                   f"&startTime={t}&endTime={fim_ms - 1}&limit=1500")
            lote = base._get_json(url)
            if not lote:
                break
            velas += [[int(k[0]), float(k[1]), float(k[2]), float(k[3]), float(k[4]), float(k[5])] for k in lote]
            t = int(lote[-1][0]) + CINCO_MIN
            time.sleep(0.12)
        fonte = "fapi.binance.com"
    except Exception as e:
        log(f"   {simbolo}: API indisponível ({type(e).__name__}); tentando arquivos públicos...")
        velas, fonte = [], None
        for mes in base._meses(inicio_ms, fim_ms):
            url = f"https://data.binance.vision/data/futures/um/monthly/klines/{simbolo}/5m/{simbolo}-5m-{mes}.zip"
            try:
                for k in base._get_zip_csv(url):
                    velas.append([int(k[0]), float(k[1]), float(k[2]), float(k[3]), float(k[4]), float(k[5])])
            except Exception:
                continue
        fonte = "data.binance.vision" if velas else None
    vistos, limpas = set(), []
    for v in sorted(velas):
        if v[0] not in vistos and inicio_ms <= v[0] < fim_ms:
            vistos.add(v[0])
            limpas.append(v)
    if limpas:
        with open(arq, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerows(limpas)
    return limpas, fonte


def sinais_moeda(velas):
    """RSI de Wilder (14), %B de Bollinger (20, desvio populacional) e volume, como no robô.
    O robô semeia o RSI no início de uma janela de 200 velas; depois de 185 passos o efeito
    da semente é menor que 0,0001%, então a série contínua dá o mesmo resultado."""
    n = len(velas)
    closes = [v[4] for v in velas]
    vols = [v[5] for v in velas]
    out = [0] * n           # +1 compra, -1 venda, 0 nada
    ag = al = None
    soma_c = soma_c2 = soma_v = 0.0
    for i in range(n):
        c = closes[i]
        soma_c += c; soma_c2 += c * c; soma_v += vols[i]
        if i >= 20:
            co = closes[i - 20]; soma_c -= co; soma_c2 -= co * co; soma_v -= vols[i - 20]
        if i >= 1:
            d = c - closes[i - 1]
            g, p = (d if d > 0 else 0.0), (-d if d < 0 else 0.0)
            if i <= 14:
                ag = (ag or 0.0) + g / 14.0
                al = (al or 0.0) + p / 14.0
            else:
                ag = (ag * 13 + g) / 14.0
                al = (al * 13 + p) / 14.0
        if i < 30:
            continue
        rsi = 100.0 if al == 0 else 100.0 - 100.0 / (1.0 + ag / al)
        media = soma_c / 20.0
        var = max(0.0, soma_c2 / 20.0 - media * media)
        dp = var ** 0.5
        larg = 4.0 * dp
        pct_b = (c - (media - 2 * dp)) / larg if larg > 0 else 0.5
        vol_ok = vols[i] >= (soma_v / 20.0) * 0.75
        if rsi >= 70 and pct_b >= 0.98 and c > closes[i - 1] and vol_ok:
            out[i] = -1
        elif rsi <= 30 and pct_b <= 0.02 and c < closes[i - 1] and vol_ok:
            out[i] = 1
    return out


def stop_auto(velas, i, lado, preco):
    """Stop automático do robô: fundo/topo das últimas 20 velas -+ 0,5 ATR(14)."""
    janela = velas[max(0, i - 19): i + 1]
    trs = []
    for k in range(max(1, i - 13), i + 1):
        h, l, pc = velas[k][2], velas[k][3], velas[k - 1][4]
        trs.append(max(h - l, abs(h - pc), abs(l - pc)))
    atr = sum(trs) / len(trs) if trs else 0.0
    if lado > 0:
        s = (min(v[3] for v in janela) - 0.5 * atr) * 0.999
        return s if s < preco else preco * 0.995
    s = (max(v[2] for v in janela) + 0.5 * atr) * 1.001
    return s if s > preco else preco * 1.005


def simular(cfg, dados, sinais, funding):
    taxa, slip = cfg["taxa"], cfg["slip"]
    alav, margem_usdt = cfg["alav"], cfg["margem_usdt"]
    ini, fim = base._ms(INICIO), base._ms(FIM)
    linha = sorted({v[0] for m in dados for v in dados[m]["velas"] if ini <= v[0] < fim})
    pos = {}
    ops, curva = [], []
    saldo_ganhos, pausado_ate = 0.0, 0
    realizado = 0.0
    pausas = 0
    ambig = 0
    max_comprometido = 0.0

    def fechar(m, preco_ref, motivo, ts):
        nonlocal realizado, saldo_ganhos
        p = pos.pop(m)
        lado = p["lado"]
        preco = preco_ref * (1 - slip * lado)
        bruto = lado * (preco - p["entrada"]) * p["qtd"]
        taxas = p["taxa_ent"] + preco * p["qtd"] * taxa
        liq_usdt = bruto - taxas + p["funding"]
        liq_brl = liq_usdt * DOLAR
        realizado += liq_brl
        if liq_brl > 0:
            saldo_ganhos += liq_brl
        ops.append({"moeda": m, "lado": "LONG" if lado > 0 else "SHORT", "abertura": p["ts"], "fechamento": ts,
                    "motivo": motivo, "entradas": p["n"], "liq_brl": liq_brl, "taxas_brl": taxas * DOLAR,
                    "funding_brl": p["funding"] * DOLAR, "nocional_usdt": p["qtd"] * p["entrada"]})

    for ts in linha:
        tfecha = ts + CINCO_MIN
        agora = {}
        for m in dados:
            i = dados[m]["idx"].get(ts)
            if i is not None:
                agora[m] = i
        # funding (00/08/16 UTC)
        for m, p in pos.items():
            if m in agora:
                tf = funding[m].get(ts)
                if tf is not None:
                    p["funding"] -= p["lado"] * p["qtd"] * dados[m]["velas"][agora[m]][1] * tf
        # saídas dentro da vela (stop primeiro se a vela tocou os dois)
        fechou_algo = False
        for m in list(pos):
            if m not in agora:
                continue
            p = pos[m]
            _, o, h, l, c = dados[m]["velas"][agora[m]][:5]
            lado = p["lado"]
            if lado > 0:
                if o <= p["sl"]:
                    fechar(m, o, "STOP", tfecha)
                elif l <= p["sl"]:
                    ambig += h >= p["tp"]
                    fechar(m, p["sl"], "STOP", tfecha)
                elif o >= p["tp"]:
                    fechar(m, o, "ALVO", tfecha)
                elif h >= p["tp"]:
                    fechar(m, p["tp"], "ALVO", tfecha)
                else:
                    continue
            else:
                if o >= p["sl"]:
                    fechar(m, o, "STOP", tfecha)
                elif h >= p["sl"]:
                    ambig += l <= p["tp"]
                    fechar(m, p["sl"], "STOP", tfecha)
                elif o <= p["tp"]:
                    fechar(m, o, "ALVO", tfecha)
                elif l <= p["tp"]:
                    fechar(m, p["tp"], "ALVO", tfecha)
                else:
                    continue
            fechou_algo = True
        # proteção de superávit: fecha tudo e pausa 4h
        if fechou_algo and saldo_ganhos >= GATILHO_SUPERAVIT_BRL and tfecha >= pausado_ate:
            pausas += 1
            pausado_ate = tfecha + PAUSA_MS
            for m in list(pos):
                ref = dados[m]["velas"][agora[m]][4] if m in agora else pos[m]["entrada"]
                fechar(m, ref, "PROTEÇÃO (fecha tudo e pausa 4h)", tfecha)
        if pausado_ate and tfecha >= pausado_ate:
            pausado_ate, saldo_ganhos = 0, 0.0
        # marca a mercado (a cada hora, para a curva)
        if tfecha % UMA_HORA == 0:
            aberto = sum(p["lado"] * (dados[m]["velas"][agora[m]][4] - p["entrada"]) * p["qtd"]
                         for m, p in pos.items() if m in agora) * DOLAR
            curva.append((tfecha, CAIXA_BRL + realizado + aberto))
        if pausado_ate:
            continue
        # entradas na vela fechada
        for m in MOEDAS:
            if m not in agora:
                continue
            i = agora[m]
            s = sinais[m][i]
            if not s:
                continue
            p = pos.get(m)
            if p and (p["lado"] != s or p["n"] >= MAX_ENTRADAS):
                continue
            comprometido = sum(x["margem_brl"] for x in pos.values())
            margem_brl = margem_usdt * DOLAR
            if margem_brl > CAIXA_BRL - comprometido:
                continue
            c = dados[m]["velas"][i][4]
            preco = c * (1 + 0.0002 * s) if cfg["slip_entrada_robo"] else c * (1 + slip * s)
            qtd = margem_usdt * alav / c
            if cfg["stop_manual"]:
                sl = c * (1 - cfg["sl_pct"] / 100 * s)
                tp = c * (1 + cfg["tp_pct"] / 100 * s)
            else:
                sl = stop_auto(dados[m]["velas"], i, s, c)
                tp = c + (c - sl) * 1.5
            if p:
                nq = p["qtd"] + qtd
                p.update(entrada=(p["entrada"] * p["qtd"] + preco * qtd) / nq, qtd=nq, sl=sl, tp=tp, n=p["n"] + 1,
                         taxa_ent=p["taxa_ent"] + preco * qtd * taxa, margem_brl=p["margem_brl"] + margem_brl)
            else:
                pos[m] = {"lado": s, "entrada": preco, "qtd": qtd, "sl": sl, "tp": tp, "n": 1, "ts": tfecha,
                          "taxa_ent": preco * qtd * taxa, "funding": 0.0, "margem_brl": margem_brl}
            max_comprometido = max(max_comprometido, sum(x["margem_brl"] for x in pos.values()))
    for m in list(pos):
        ult = dados[m]["velas"][-1][4]
        fechar(m, ult, "FIM DO TESTE", linha[-1] + CINCO_MIN)
    return ops, curva, pausas, ambig, max_comprometido


def metricas(ops, curva, pausas, ambig, max_comp):
    r = {"operacoes": len(ops), "pausas_protecao": pausas, "velas_ambiguas": ambig,
         "margem_maxima_comprometida_brl": round(max_comp, 2)}
    if not ops:
        return r
    g = [o["liq_brl"] for o in ops if o["liq_brl"] > 0]
    p = [o["liq_brl"] for o in ops if o["liq_brl"] <= 0]
    sg, sp = sum(g), -sum(p)
    r["fator_lucro"] = round(sg / sp, 3) if sp else None
    r["taxa_acerto_pct"] = round(100 * len(g) / len(ops), 2)
    r["resultado_brl"] = round(sum(o["liq_brl"] for o in ops), 2)
    r["retorno_sobre_caixa_pct"] = round(100 * r["resultado_brl"] / CAIXA_BRL, 2)
    r["ganho_medio_brl"] = round(sg / len(g), 2) if g else 0
    r["perda_media_brl"] = round(-sp / len(p), 2) if p else 0
    r["pior_operacao_brl"] = round(min(o["liq_brl"] for o in ops), 2)
    r["ganhos_para_cobrir_1_perda_media"] = round((sp / len(p)) / (sg / len(g)), 1) if g and p else None
    r["taxas_pagas_brl"] = round(sum(o["taxas_brl"] for o in ops), 2)
    r["funding_brl"] = round(sum(o["funding_brl"] for o in ops), 2)
    mot = {}
    for o in ops:
        mot[o["motivo"]] = mot.get(o["motivo"], 0) + 1
    r["motivos_saida"] = mot
    pico, dd = -1e18, 0.0
    for _, v in curva:
        pico = max(pico, v)
        dd = max(dd, pico - v)
    r["queda_maxima_brl"] = round(dd, 2)
    zerou = next((t for t, v in curva if v <= 0), None)
    r["caixa_zerado_em"] = (datetime.fromtimestamp(zerou / 1000, tz=timezone.utc).strftime("%Y-%m-%d")
                            if zerou else "nunca zerou")
    r["queda_maxima_pct_do_caixa"] = round(100 * dd / CAIXA_BRL, 2)
    meses = {}
    for o in ops:
        k = datetime.fromtimestamp(o["fechamento"] / 1000, tz=timezone.utc).strftime("%Y-%m")
        meses[k] = meses.get(k, 0.0) + o["liq_brl"]
    r["resultado_por_mes_brl"] = {k: round(v, 2) for k, v in sorted(meses.items())}
    r["meses_positivos"] = sum(1 for v in meses.values() if v > 0)
    r["meses_total"] = len(meses)
    pm = {}
    for o in ops:
        pm[o["moeda"]] = pm.get(o["moeda"], 0.0) + o["liq_brl"]
    r["resultado_por_moeda_brl"] = {k: round(v, 2) for k, v in sorted(pm.items(), key=lambda x: x[1])}
    for lado in ("LONG", "SHORT"):
        sub = [o["liq_brl"] for o in ops if o["lado"] == lado]
        gg = sum(x for x in sub if x > 0); pp = -sum(x for x in sub if x <= 0)
        r[f"{lado.lower()}_operacoes"] = len(sub)
        r[f"{lado.lower()}_fator_lucro"] = round(gg / pp, 3) if pp else None
    meio = curva[0][0] + (curva[-1][0] - curva[0][0]) / 2 if curva else 0
    for nome, f in (("1o_ano", lambda o: o["fechamento"] < meio), ("2o_ano", lambda o: o["fechamento"] >= meio)):
        sub = [o["liq_brl"] for o in ops if f(o)]
        gg = sum(x for x in sub if x > 0); pp = -sum(x for x in sub if x <= 0)
        r[f"fator_lucro_{nome}"] = round(gg / pp, 3) if pp else None
        r[f"resultado_{nome}_brl"] = round(gg - pp, 2)
    return r


def main():
    ini = base._ms(INICIO) - AQUECIMENTO_DIAS * 24 * UMA_HORA
    fim = base._ms(FIM)
    agora_ms = int(time.time() * 1000)
    fim = min(fim, agora_ms - agora_ms % CINCO_MIN)
    dados, funding, sinais = {}, {}, {}
    log(f"[1/3] Baixando velas de 5 minutos de {len(MOEDAS)} moedas ({INICIO} a {FIM}). Pode levar de 5 a 15 minutos...")
    for k, m in enumerate(MOEDAS, 1):
        sim = f"{m}USDT"
        v, fonte = baixar_velas_5m(sim, ini, fim)
        if len(v) < 2000:
            log(f"  {k:2d}/{len(MOEDAS)} {m}: sem dados suficientes — fica de fora")
            continue
        fund, _ = base.baixar_funding(sim, ini, fim)
        mapa = {t - t % CINCO_MIN: x for t, x in fund}
        if not mapa:
            mapa = {vv[0]: 0.0001 for vv in v if vv[0] % (8 * UMA_HORA) == 0}
        funding[m] = mapa
        dados[m] = {"velas": v, "idx": {vv[0]: i for i, vv in enumerate(v)}}
        log(f"  {k:2d}/{len(MOEDAS)} {m}: {len(v)} velas ({fonte}), funding: {len(fund)}")
    if not dados:
        log("ERRO: nenhum dado baixado.")
        sys.exit(1)
    log("\n[2/3] Calculando RSI, Bollinger e volume (igual ao robô)...")
    for m in dados:
        sinais[m] = sinais_moeda(dados[m]["velas"])
    variantes = [
        ("S1", "Configuração da tela: 9x, stop 8%, alvo 0,33%, margem US$ 235 por entrada (≈ R$ 1.220)",
         dict(alav=9, margem_usdt=235, stop_manual=True, sl_pct=8.0, tp_pct=0.33, taxa=0.0005, slip=0.0002, slip_entrada_robo=True)),
        ("S2", "Padrão do código: 3x, stop 8%, alvo 0,33%, mesma margem",
         dict(alav=3, margem_usdt=235, stop_manual=True, sl_pct=8.0, tp_pct=0.33, taxa=0.0005, slip=0.0002, slip_entrada_robo=True)),
        ("S3", "Stop automático (topo/fundo + ATR) e alvo 1,5x o risco, 9x",
         dict(alav=9, margem_usdt=235, stop_manual=False, sl_pct=0, tp_pct=0, taxa=0.0005, slip=0.0002, slip_entrada_robo=True)),
        ("S1-taxa-robo", "Igual à S1, com a taxa que o robô supõe (0,04% por lado)",
         dict(alav=9, margem_usdt=235, stop_manual=True, sl_pct=8.0, tp_pct=0.33, taxa=0.0004, slip=0.0002, slip_entrada_robo=True)),
        ("S1-conservador", "Igual à S1, com taxa 0,05% e slippage 0,05% por lado",
         dict(alav=9, margem_usdt=235, stop_manual=True, sl_pct=8.0, tp_pct=0.33, taxa=0.0005, slip=0.0005, slip_entrada_robo=False)),
    ]
    log("\n[3/3] Simulando variantes...")
    res = {}
    for cod, nome, cfg in variantes:
        ops, curva, pausas, amb, mc = simular(cfg, dados, sinais, funding)
        mt = metricas(ops, curva, pausas, amb, mc)
        res[cod] = {"nome": nome, "config": cfg, **mt}
        with open(f"operacoes_sobrinho_{cod}.csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["moeda", "lado", "abertura_utc", "fechamento_utc", "motivo", "entradas", "liquido_brl"])
            for o in ops:
                w.writerow([o["moeda"], o["lado"],
                            datetime.fromtimestamp(o["abertura"] / 1000, tz=timezone.utc).strftime("%Y-%m-%d %H:%M"),
                            datetime.fromtimestamp(o["fechamento"] / 1000, tz=timezone.utc).strftime("%Y-%m-%d %H:%M"),
                            o["motivo"], o["entradas"], f"{o['liq_brl']:.2f}".replace(".", ",")])
        log(f"  {cod}: {mt.get('operacoes', 0)} operações | acerto {mt.get('taxa_acerto_pct')}% | fator de lucro {mt.get('fator_lucro')} "
            f"| resultado R$ {mt.get('resultado_brl')} | queda máx. R$ {mt.get('queda_maxima_brl')}")
    saida = {"gerado_em": datetime.now().isoformat(timespec="seconds"), "periodo": [INICIO, FIM],
             "moedas": list(dados), "caixa_brl": CAIXA_BRL, "dolar_fixo": DOLAR, "resultados": res}
    with open("resultado_sobrinho.json", "w", encoding="utf-8") as f:
        json.dump(saida, f, ensure_ascii=False, indent=2)
    with open("resultado_sobrinho.txt", "w", encoding="utf-8") as f:
        f.write(f"BACKTEST DO ROBÔ DO SOBRINHO (MORS-VIDA-SEM-FOME) — {INICIO} a {FIM} — caixa R$ {CAIXA_BRL:,.0f}\n\n")
        for cod, r in res.items():
            f.write(f"[{cod}] {r['nome']}\n")
            f.write(f"    Operações: {r.get('operacoes')} | acerto {r.get('taxa_acerto_pct')}% | fator de lucro {r.get('fator_lucro')} "
                    f"(1º ano {r.get('fator_lucro_1o_ano')}, 2º ano {r.get('fator_lucro_2o_ano')})\n")
            f.write(f"    Resultado: R$ {r.get('resultado_brl')} ({r.get('retorno_sobre_caixa_pct')}% do caixa) | queda máxima R$ {r.get('queda_maxima_brl')}\n")
            f.write(f"    Ganho médio R$ {r.get('ganho_medio_brl')} | perda média R$ {r.get('perda_media_brl')} | pior operação R$ {r.get('pior_operacao_brl')}\n")
            f.write(f"    Meses positivos {r.get('meses_positivos')}/{r.get('meses_total')} | pausas de proteção {r.get('pausas_protecao')} | caixa de R$ 10 mil zerado em: {r.get('caixa_zerado_em')}\n\n")
    log("\nPronto! Arquivos gerados: resultado_sobrinho.txt, resultado_sobrinho.json e operacoes_sobrinho_*.csv")


if __name__ == "__main__":
    main()
