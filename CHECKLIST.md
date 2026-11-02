# CHECKLIST — sessão Claude Code (nuvem · Nanda)

Frentes desta sessão: agente Nanda, Núcleo Padrão e manutenção do Mural.
Repositório público: só resumo e onde procurar. Detalhes no Mural e no Drive (ver `CLAUDE.md`).

## Onde paramos

Atualizado em 02/11/2026.

### Estado atual

1. **Fluxo da Nanda (n8n)**
   - V29 entregue pelo Claude Code em 26/10 e instalada pelo Projeto Nanda em 27/10, com 2 ajustes deles (fluxo importado desligado e sem o id da V28; chave da Evolution mantida nos nós HTTP até a troca da chave).
   - Roteiro de testes T1 a T13: todos aprovados (Mural Nº 89).
   - Documento de controle e estudos: Drive, pasta "AGENTE NANDA" > "V29 — Revisão Claude Code (26-10-2026)".
2. **Prompt da Nanda**
   - V29.1 (Pix na mesma resposta da confirmação): revisada e aprovada (Nº 84).
   - V29.2 (regra geral contra copiar frase-modelo + horário aberto/fechado): instalada pelo Projeto Nanda em 27/10. **Aguardando o arquivo, que vem pelo Vamberto, para revisar** (Nº 89).
3. **Troca do modelo de IA:** o gpt-4.1-mini sai da API da OpenAI em 14/01/2027 (confirmado pelo Projeto Nanda). Pedido de teste de substitutos (Nº 85), ainda sem resposta. Vale também para Bia e Cris.
4. **Memória (fim da conversa):** estudo de prós e contras enviado (Nº 82). Aguarda a opinião do Projeto Nanda; a decisão é do Vamberto.
5. **Mural:** manutenção feita em 28/10 (coluna "Nº", aba "Arquivo", Protocolo atualizado). Assinaturas das sessões do Code registradas no Protocolo.
6. **Inventário das sessões do Code:** enviado ao Chat Diretor (Nº 87), respondido. O Piloto IPTV da Cris é parte da Mais Talentto.
7. **Repositórios públicos (02/11):** este repositório é público. Achados e uma proposta de convivência (páginas públicas + material interno em lugar privado) foram enviados ao responsável pela segurança, "Claude Code (nuvem · Sobra Real)", na Nº 106 e por mensagem direta. **A decisão final sobre segurança é dele.** Nada de outra sessão foi apagado ou movido.

### Decisões do Vamberto já tomadas

- Escalonamento para atendente já no 1º pedido do cliente.
- Memória de 20 trocas.
- "#pausar" visível ao cliente: mantido por enquanto; voltar depois com algo mais simples para o lojista.
- Modelo de IA: melhor desempenho sem quebrar o orçamento e preservando a margem.
- Opus 5.5 com raciocínio "alto" é suficiente; avisar antes de tarefa que justifique subir.

### Próximos passos, em ordem

1. Acompanhar a resposta do responsável pela segurança (Nº 106) e executar o que ele pedir. Se ele mover o material interno para um repositório privado, este CHECKLIST.md e o CLAUDE.md vão junto.
2. Receber o PROMPT_NANDA_PRIME_V29.2 e revisar; responder no Mural (Nº 89).
3. Acompanhar os testes novos (T14, T15, T16) e os refeitos (T6, T7, T13) com a V29.2.
4. Teste dos modelos substitutos pelo Projeto Nanda até 15/12/2026; comparar e levar a recomendação ao Vamberto; trocar Nanda, Bia e Cris antes de 14/01/2027.
5. Próxima fase do fluxo: mensagens repetidas e em rajada (desenho do Claude Code; mudança no Supabase só com o ok do Vamberto).
6. Fechar a decisão da memória (Nº 82) com a opinião do Projeto Nanda.
7. Próximo JSON entregue: sem id, desligado, funcionando ao importar; corrigir o "55" na legenda do comprovante.
8. Núcleo Padrão: só depois da Nanda estável e da correção de segurança do banco aprovada (trava do Chat Diretor).
9. Manutenção do Mural quando ele voltar a crescer, sempre com aviso prévio.

### Lembretes

- Segurança da Estrutura é da sessão "Claude Code (nuvem · Sobra Real)". Não registrar falhas de segurança neste repositório.
- Lições para o Núcleo Padrão (documento de controle V29): um caminho só até o cliente; dado importante vem do WhatsApp e do n8n, não do texto da IA; prompt e fluxo dizendo a mesma coisa; uma tag por pedido; data e hora no prompt; histórico de versões fora do prompt.
