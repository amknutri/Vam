# Robô de Trade Binance — contexto do projeto

Leia este arquivo inteiro antes de qualquer ação neste projeto.

## Quem é o usuário

- **Vamberto Barbosa**, fundador da Mais Talentto. Empreendedor **não técnico**.
- Responder sempre em **português do Brasil**, linguagem direta, chamá-lo de **Vamberto**.
- Explicar cada conceito antes de executar. **Um passo por vez**, pedindo confirmação antes de criar, alterar ou apagar arquivos.
- Documentos para ele: **DOCX**. Planilhas: Google Sheets no Drive (conta amknutri@gmail.com).
- Honestidade total sobre riscos: nenhum robô garante lucro. Nunca prometer resultado.
- **Nada roda com dinheiro real sem aprovação explícita dele**, e só depois de a Demo cumprir os critérios.
- Ele **nunca** cola chaves de API ou senhas na conversa. Se aparecer chave em arquivo ou print, alertar e orientar a revogar.
- O código original é desenvolvido pelo **Lucas** (filho dele). Um **sobrinho** prepara outro código, que será comparado com este.

## Status atual (atualizar a cada sessão)

- **07/10/2026, 14:21:** dois robôs rodando na **Demo** da Binance, no PC do Vamberto:
  - **v7.1 SPOT** em `C:\RoboBinance` (Confluência, 1x, só compra, banca $5.000). Sem posições ainda (mercado em queda).
  - **v6.1 FUTUROS** em `C:\RoboBinanceFuturos` (Confluência, 3x, compra e venda, banca ~$5.000). A carteira de Futuros da Demo foi ativada abrindo demo.binance.com com o site em **English**.
- **Stop na exchange validado na prática:** a internet caiu com SHORT em TRX aberto; a Binance executou o stop sozinha (LOSS -$7,75). Ordens Stop Market/Take Profit Market aparecem em Open Orders > Conditional.
- **Bug encontrado na Demo e corrigido (v6.1/v7.1):** em Futuros a ordem a mercado volta sem preço médio; o robô usava o preço da tela (0,18% de diferença na TRX). Agora consulta o preço real.
- **Meta do Vamberto:** conta real em **05/11/2026**. Plano aprovado: v8 até 12/10, estratégia congelada depois; real só com critérios cumpridos, começando com 10–20% do valor.
- **07/10/2026: v8 pronta e no GitHub** (`bot_textual_v8_FUTUROS.py` e `bot_textual_v8_SPOT.py`), aprovada pelo Vamberto (itens 1–5, heartbeat, 8 Telegram, 9 eventos, 10 funding, 11 volume desligado). **Instalada no PC:** Futuros v8 às 17:19 e Spot v8 às 19:36 (v6.1/v7.1 continuam nas pastas como reserva). Futuros ficou PAUSADO (risco) no dia 07/10: 4 stops em vendas (TRX, LDO, RENDER, LDO) somaram -$111,16, acima do drawdown de 2% — comportamento correto. No dia seguinte o Vamberto precisa apertar L (não religa sozinho).
  - Cada robô v8 grava `heartbeat.json` na própria pasta a cada ciclo. Para o Cowork: robô vivo = campo `hora` com menos de 2 minutos.
  - `eventos.json` (na pasta de cada robô) lista CPI/FOMC; perto deles o robô não abre posição. Falta criar o arquivo com datas oficiais.

## Backtest de 07/10/2026 (resultado central do projeto)

- `backtest/backtest_estrategias.py` roda no PC (a nuvem não acessa a Binance), usa as funções do próprio robô, 2 anos (out/2024–set/2026), taxas, slippage, funding real. Resultados em `backtest/resultado_backtest_2026-10-07.*`.
- **Nenhuma variante ganhou.** v8 confluência: fator de lucro 0,78, -96,7% em 2 anos, 352 dias batendo limite diário. Filtros (volume, tendência 4h, 10 maiores, sem trava de $10) ficaram entre 0,79 e 0,84. Donchian 10 moedas: 0,83 (LONG 1,24, mas só por um trade de XRP de +80R; 2º ano 0,50). Donchian BTC do robô: 0,92.
- Velas ambíguas (stop e alvo na mesma hora): 0 a 3 por variante → a premissa conservadora não distorce o resultado.
- Antes dos custos a confluência fica perto de zero; os custos (~0,2% por operação) a tornam fortemente negativa.
- **Recomendação dada ao Vamberto:** não ir para dinheiro real com nenhuma dessas estratégias; aguardando decisão dele sobre o caminho (pesquisa sem prazo, testar código do sobrinho, ou encerrar trading ativo).

## Decisões tomadas (não rediscutir sem motivo novo)

1. **v5 → v6:** 13 correções de segurança (chaves fora do código, stop registrado na exchange, fechamento só com confirmação, limite diário que zera à meia-noite e persiste em disco etc.).
2. **Futuros → Spot (v7):** a Binance **não oferece Futuros a residentes no Brasil** (restrição da CVM, vigente em 2026). A Demo de Futuros redireciona para o Spot e o menu da conta real não tem "Derivativos".
3. **VPN descartada:** a conta é verificada como brasileira, há risco de bloqueio com saldo dentro, não há recurso legal e a VPN impede a restrição de IP da chave. Não orientar uso de VPN.
4. **Futuros (decisão do Vamberto, 07/10):** o robô deve operar Futuros na Binance e na BitGet. Começa pela Demo (Binance já rodando com a v6.1; BitGet = v8 a criar). Dinheiro real em Futuros só com aprovação explícita dele, ciente dos riscos: Binance via site em outro idioma é zona cinzenta perante a CVM; BitGet não tem autorização da CVM. VPN continua descartada.
5. **Critérios de aprovação da Demo** (definidos antes de começar, não mudar a régua depois):
   - **Zero** compras sem stop na Binance;
   - pelo menos **30 operações** fechadas;
   - **fator de lucro acima de 1,2**, já com taxas;
   - no máximo **2 dias por semana** batendo o limite de perda;
   - mínimo de **4 semanas** (início 07/10/2026, fim mínimo 04/11/2026).
6. **Régua de resultado (banca de $5.000, 30 dias):**
   - Ruim: abaixo do CDI (~1,1%/mês ≈ $55) ou queda máxima acima de 6%.
   - Boa: 1,5% a 3% ao mês.
   - Muito boa: 3% a 5% ao mês.
   - Acima de 5%: desconfiar (provável sorte).

## Arquivos (pasta `robo_binance/` deste repositório)

| Arquivo | O que é |
|---|---|
| `bot_textual_v8_FUTUROS.py` / `bot_textual_v8_SPOT.py` | **Próxima versão** (marcas `[v8]`). Velas 1h, filtros de custo/funding/eventos, freio, breakeven, heartbeat. Testes 17/17 |
| `bot_textual_v7_1_SPOT.py` | **Em uso (Spot).** v7 + preço real de execução |
| `bot_textual_v6_1_FUTUROS.py` | **Em uso (Futuros).** v6 + preço real de execução |
| `bot_textual_v7_SPOT.py` | Versão anterior do Spot. Spot, só compra, 1x, stop STOP_LOSS / STOP_LOSS_LIMIT na Binance, posições salvas em `posicoes_spot.json` |
| `bot_textual_v6_CORRIGIDA.py` | Versão Futuros com as correções de segurança (não utilizável no Brasil) |
| `Relatorio_Robo_v7_SPOT.docx` | Relatório da v7 para o Lucas |
| `Relatorio_Melhorias_Robo_v6.docx` | Relatório da v6 para o Lucas |

- O arquivo original **v5 (bot_textual_v5_ANTIGA.py) NUNCA deve ser commitado**: ele contém chaves e token do Telegram.
- Todas as mudanças no código estão marcadas com os comentários `[v6]` e `[v7 SPOT]`.
- Os testes foram feitos com simuladores de exchange (v6: 19/19; v7: 18/18). **Nunca** foram validados contra a Binance real: a Demo é o teste.

## Ambiente no PC do Vamberto (Windows)

- Pasta do robô: `C:\RoboBinance`.
- Chaves da Demo nas variáveis de ambiente `BINANCE_API_KEY` e `BINANCE_API_SECRET` (configuradas com `setx`; 64 caracteres cada). Criadas em demo.binance.com > Gerenciamento de API. Na Demo as permissões são fixas.
- Python **3.13.16**; bibliotecas ccxt **4.5.85**, textual, aiohttp, openpyxl.
- O PC não suspende (suspensão = Nunca; vídeo desliga em 15 min).
- Para rodar: PowerShell → `cd C:\RoboBinance` → `python bot_textual_v7_SPOT.py` → tecla **L** liga.
- Para sair: tecla **Q**. Nunca reiniciar o PC sem antes apertar Q.

## Acompanhamento (Google Drive, conta amknutri@gmail.com)

- Pasta: **Robô Binance** (id `1rnDwn5vtP0ol7uO5h_AacsIL7V5C3I1z`).
- Planilha: **Robô Binance — Acompanhamento Demo** (id `1sEg3BMg9kpptqt3HB9tZy89jmUtiG4XbK8HpqjC5kWQ`).
- **Mural — Comunicação entre Projetos (Vamberto)** (id `1XjZz0vTW4TtT_ewgHN82XYC3mHviulEORNlbmAORG2Y`, agora na raiz do Meu Drive): desde 07/10/2026 é o canal COMUM entre todos os projetos e agentes. Colunas: Data/hora, Projeto, De, Para, Assunto, Mensagem, Status, Resposta. Abas: Mensagens, Protocolo, Projetos.
- **Rotinas do robô DESATIVADAS em 07/10/2026** (projeto pausado): `trig_01W5dPYg2w2xgBU26Pb3xB15` (09h15) e `trig_01A9cMjrGErzRaUy4N5bjXgH` (18h22) — desligadas, não apagadas. Vamberto pediu ao Cowork para cancelar as dele (08h49 e 17h59).
- **Painel (artefato do Cowork):** subpasta Drive "Robô Binance > Painel" (id `11L_RGjOv8_FQ62n-qOW6C_ErpmvBCK7e`); instruções `INSTRUCOES_CLAUDE_CODE_PAINEL.md` (id `1kpsdSRm1bEIYB6Q6RdtxQ0OTjX_hix5Z`), código atual `robo-binance-painel.html` (id `1mpwWD_Hs6sGMHcFjs2bu8N2VDUuHjhxW`). Entregar como `robo-binance-painel-v3.html` (não sobrescrever) e avisar no Mural. Prioridade depois da v8.
- **Janelas dos robôs:** devem ficar VISÍVEIS (não minimizadas) até a v8 ter heartbeat.json — o Cowork não consegue restaurar janela minimizada.
- Abas: Painel (semáforo + veredito), Diário, Operações, Decisões, Banco de Ideias, Instruções Cowork.
- O **Cowork** (no PC do Vamberto) lê `C:\RoboBinance\trades_journal.csv` e atualiza a planilha. O Claude Code roda na nuvem e **não acessa o PC**: lê a planilha pelo conector do Drive.
- O Vamberto planeja um **mural** no Drive compartilhado entre projetos (cada projeto sem interferir no outro).

## Pendências

- [ ] Vamberto apagar o txt das chaves e o `bot_textual_v5_ANTIGA.py`, e esvaziar a Lixeira.
- [ ] Encaminhar a v7 e o relatório ao Lucas.
- [ ] Configurar a tarefa do Cowork para atualizar a planilha.
- [x] Stop na exchange validado (TRX, Futuros, durante queda de internet).
- [x] Vamberto aprovou a v8; código pronto e testado.
- [x] v8 instalada nos dois robôs (07/10). Mural avisado (heartbeat dos dois, limite de 2 min).
- [ ] Balanço da v8 na sexta 10/10 (separar azar de defeito).
- [ ] Criar `eventos.json` com CPI/FOMC de out–nov/2026 (datas oficiais) e avisar o Cowork no Mural.
- [ ] Telegram: novo token no BotFather + `setx TELEGRAM_BOT_TOKEN` / `TELEGRAM_CHAT_ID`.
- [x] Backtest do filtro de volume (item 11): não salva a estratégia (FL 0,82).
- [x] 07/10 (noite): Vamberto escolheu o caminho C (pausar trading ativo; foco na Mais Talentto). Capital que aceitaria arriscar: no máximo 5 mil. Meta dele de R$ 3.000/mês com robô é inviável com esse capital (exigiria ~60%/mês) — explicado. Robô fica como projeto de aprendizado: Demo pode seguir, backtest pronto para testar o código do sobrinho (caminho B).
- [ ] Resolver a verificação pendente da conta REAL da Binance (comprovante de residência; trade e depósito restritos).
- [ ] Criar conta Demo na BitGet para a v8 BitGet.
- [ ] Comparar com o código do sobrinho quando chegar.

## Regras de trabalho neste projeto

- Mudanças de estratégia ou escopo (moeda nova, corretora nova, alavancagem): **sempre perguntar**, nunca decidir sozinho.
- Mudança de código: arquivo **novo** com versão nova (v8, …). Não sobrescrever a versão em uso. Testar com simulador antes de entregar.
- Commits no branch indicado pela sessão. Nunca commitar arquivo com chave, token ou senha.
