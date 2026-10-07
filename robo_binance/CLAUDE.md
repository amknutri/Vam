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

- **07/10/2026, 01:09:** robô **v7 SPOT** ligado na **Demo Spot** da Binance, estratégia **Confluência**, banca fictícia de **$5.000 USDT**.
- Mercado em queda (tendência BAIXA): nenhuma compra até o momento. O filtro de tendência bloqueia compras contra a tendência (comportamento esperado).
- Próximo marco: **1ª compra**. Conferir com o Vamberto se a ordem de **Stop** apareceu em demo.binance.com > Ordens > Ordem Spot.

## Decisões tomadas (não rediscutir sem motivo novo)

1. **v5 → v6:** 13 correções de segurança (chaves fora do código, stop registrado na exchange, fechamento só com confirmação, limite diário que zera à meia-noite e persiste em disco etc.).
2. **Futuros → Spot (v7):** a Binance **não oferece Futuros a residentes no Brasil** (restrição da CVM, vigente em 2026). A Demo de Futuros redireciona para o Spot e o menu da conta real não tem "Derivativos".
3. **VPN descartada:** a conta é verificada como brasileira, há risco de bloqueio com saldo dentro, não há recurso legal e a VPN impede a restrição de IP da chave. Não orientar uso de VPN.
4. **BitGet Futuros:** fica no Banco de Ideias. Só será avaliada se a v7 Spot aprovar na Demo. A BitGet não tem autorização da CVM.
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
| `bot_textual_v7_SPOT.py` | **Versão em uso.** Spot, só compra, 1x, stop STOP_LOSS / STOP_LOSS_LIMIT na Binance, posições salvas em `posicoes_spot.json` |
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
- Abas: Painel (semáforo + veredito), Diário, Operações, Decisões, Banco de Ideias, Instruções Cowork.
- O **Cowork** (no PC do Vamberto) lê `C:\RoboBinance\trades_journal.csv` e atualiza a planilha. O Claude Code roda na nuvem e **não acessa o PC**: lê a planilha pelo conector do Drive.
- O Vamberto planeja um **mural** no Drive compartilhado entre projetos (cada projeto sem interferir no outro).

## Pendências

- [ ] Vamberto apagar o txt das chaves e o `bot_textual_v5_ANTIGA.py`, e esvaziar a Lixeira.
- [ ] Encaminhar a v7 e o relatório ao Lucas.
- [ ] Configurar a tarefa do Cowork para atualizar a planilha.
- [ ] Na 1ª compra: validar o stop na Binance Demo.
- [ ] Comparar com o código do sobrinho quando chegar.

## Regras de trabalho neste projeto

- Mudanças de estratégia ou escopo (moeda nova, corretora nova, alavancagem): **sempre perguntar**, nunca decidir sozinho.
- Mudança de código: arquivo **novo** com versão nova (v8, …). Não sobrescrever a versão em uso. Testar com simulador antes de entregar.
- Commits no branch indicado pela sessão. Nunca commitar arquivo com chave, token ou senha.
