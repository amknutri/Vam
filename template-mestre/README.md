# Template Mestre de Alta Conversão · Mais Talentto

Landing page reutilizável para clientes da Mais Talentto. Um único arquivo (`index.html`), sem servidor, sem instalação. Você troca os textos marcados, publica e pronto.

---

## Configuração em 5 passos (cerca de 10 minutos)

1. **Copie** a pasta `template-mestre/` e renomeie com o nome do cliente (ex.: `lp-joao-silva/`).
2. **Abra** o `index.html` num editor de texto (VS Code, Sublime ou até o Bloco de Notas).
3. **Substitua** cada marcação da tabela abaixo usando *Localizar e Substituir* (`Ctrl + H` no Windows, `Cmd + Option + F` no Mac) e clique em **Substituir tudo**. Algumas marcações aparecem em vários lugares; o "Substituir tudo" troca todas de uma vez.
4. **Abra o arquivo no navegador** (dois cliques nele). Se aparecer uma **faixa vermelha no topo**, ela lista as marcações que ainda faltam preencher.
5. **Publique** (veja a seção "Publicação" abaixo).

> Exemplo pronto: veja `exemplos/vam-barbosa/index.html`, o template preenchido com os dados do Vam Barbosa (MRP Mobi).

> Dica: guarde as respostas do cliente num bloco de notas, no mesmo formato da tabela. Assim o próximo cliente fica ainda mais rápido.

---

## Lista de marcações

### Dados gerais e SEO

| Marcação | O que colocar | Exemplo |
|---|---|---|
| `{TITULO_PAGINA}` | Título da aba do navegador e do link compartilhado | `Vam Barbosa · MRP Mobi` |
| `{DESCRICAO_PAGINA}` | Frase curta que aparece no Google e no WhatsApp (até 155 caracteres) | `Cadastre-se grátis na MRP Mobi e ganhe cashback que vira renda.` |
| `{IMAGEM_COMPARTILHAMENTO}` | Link de uma imagem (1200×630 px) para a prévia do link | `https://tvpronta.com/capa-mrp.jpg` |
| `{NOME_CLIENTE}` | Nome de quem atende (aparece nos botões de WhatsApp) | `Vam Barbosa` |
| `{NOME_MARCA}` | Nome da empresa/produto | `MRP Mobi` |
| `{LOGO_URL}` | Link da imagem do logo | `https://tvpronta.com/logo-mrp.png` |
| `{ANO}` | Ano do copyright | `2026` |

### Links de conversão

| Marcação | O que colocar | Exemplo |
|---|---|---|
| `{WHATSAPP_LINK}` | Link completo do WhatsApp, **com 55 + DDD + número, só dígitos** | `https://wa.me/5522999921997?text=Olá! Vi sua página e tenho uma dúvida.` |
| `{CTA_PRINCIPAL_LINK}` | Link da ação principal (cadastro, compra, agendamento) | `https://convite.mrpmobi.com.br/register/747` |
| `{CTA_PRINCIPAL_TEXTO}` | Texto do botão principal (verbo + benefício) | `Quero me cadastrar grátis` |

> O formulário e o botão flutuante usam o mesmo `{WHATSAPP_LINK}`. Preencheu uma vez, funciona em todo lugar.

### Seção 2 · Hero (primeira dobra)

| Marcação | O que colocar | Exemplo |
|---|---|---|
| `{SELO_HERO}` | Selinho acima do título (gatilho rápido) | `Cadastro 100% gratuito` |
| `{OFERTA_PRINCIPAL}` | Título principal: a promessa da oferta | `Carteira MRP Mobi: cashback que vira renda` |
| `{SUBTITULO_HERO}` | Uma ou duas frases explicando a promessa | `Cadastre-se gratuitamente e comece a ganhar hoje.` |
| `{MICROCOPY_HERO}` | Frase pequena abaixo dos botões que tira o medo | `Sem mensalidade. Leva menos de 2 minutos.` |

### Seção 3 · Benefícios (3 cards)

| Marcação | Exemplo |
|---|---|
| `{BENEFICIOS_TITULO}` | `O que é a MRP Mobi?` |
| `{BENEFICIOS_SUBTITULO}` | `Mobilidade, carteira digital e renda em um só app.` |
| `{BENEFICIO_1_ICONE}` / `_2_` / `_3_` | Um emoji: `📱` `💰` `🌍` |
| `{BENEFICIO_1_TITULO}` / `_2_` / `_3_` | `Aplicativo de Mobilidade` |
| `{BENEFICIO_1_TEXTO}` / `_2_` / `_3_` | `Motoristas ficam com 79% da corrida.` |

### Seção 4 · Vídeos (opcional)

| Marcação | O que colocar |
|---|---|
| `{VIDEOS_TITULO}` | `Entenda na prática` |
| `{VIDEO_1_URL}` / `_2_` / `_3_` | Link do arquivo `.mp4` |
| `{VIDEO_1_CAPA}` / `_2_` / `_3_` | Link de uma imagem de capa (`.jpg`). **Sem capa?** Apague o trecho `poster="..."` daquele vídeo |
| `{VIDEO_1_LEGENDA}` / `_2_` / `_3_` | `Como funciona o cashback` |

Os vídeos só começam a baixar quando o visitante clica em play. Isso deixa a página bem mais rápida no 4G.

### Seção 5 · Como funciona (3 passos)

| Marcação | Exemplo |
|---|---|
| `{PASSOS_TITULO}` | `Como funciona o cashback?` |
| `{PASSO_1_TITULO}` / `_2_` / `_3_` | `Empresário oferece desconto` |
| `{PASSO_1_TEXTO}` / `_2_` / `_3_` | `O parceiro define o percentual de desconto sobre a compra.` |

### Seção 6 · Prova social

| Marcação | Exemplo |
|---|---|
| `{PROVA_TITULO}` | `Quem já está dentro` |
| `{PROVA_SUBTITULO}` | `Resultados reais de quem começou antes de você.` |
| `{NUMERO_1_VALOR}` / `_2_` / `_3_` | `+30 mil` |
| `{NUMERO_1_LEGENDA}` / `_2_` / `_3_` | `cadastros realizados` |
| `{DEPOIMENTO_1_TEXTO}` / `_2_` / `_3_` | `O escritório físico já está em funcionamento. Isso é real.` (sem aspas: o template já coloca) |
| `{DEPOIMENTO_1_AUTOR}` / `_2_` / `_3_` | `Ed Carlos` |
| `{DEPOIMENTO_1_CARGO}` / `_2_` / `_3_` | `Fundador` |

### Seção 7 · FAQ (5 perguntas)

| Marcação | Exemplo |
|---|---|
| `{FAQ_1_PERGUNTA}` … `{FAQ_5_PERGUNTA}` | `Preciso pagar alguma coisa?` |
| `{FAQ_1_RESPOSTA}` … `{FAQ_5_RESPOSTA}` | `Não. O cadastro é 100% gratuito.` |

Sugestão de perguntas que mais destravam a venda: preço/custo, prazo para ver resultado, segurança/garantia, "é para mim?", e como funciona o suporte.

### Seção 8 · CTA final e formulário

| Marcação | Exemplo |
|---|---|
| `{CTA_FINAL_TITULO}` | `Pronto para começar?` |
| `{CTA_FINAL_TEXTO}` | `Cadastro 100% gratuito. Você não paga nada para participar.` |

O formulário **não precisa de servidor**: quando o visitante preenche o nome e clica em enviar, o WhatsApp do cliente abre já com a mensagem pronta (*"Olá! Meu nome é Maria. Vim pela página da MRP Mobi..."*).

---

## Estrutura da página

| # | Seção | Obrigatória? | Função na conversão |
|---|---|---|---|
| 1 | Topo (logo) | Sim | Identidade e confiança |
| 2 | Hero | Sim | Promessa + ação imediata (cadastro ou WhatsApp) |
| 3 | Benefícios | Sim | Explica o "o que eu ganho" |
| 4 | Vídeos | Opcional | Demonstração e confiança |
| 5 | Como funciona | Sim | Tira a sensação de complicado |
| 6 | Prova social | Sim | Números + depoimentos |
| 7 | FAQ | Sim | Quebra objeções |
| 8 | CTA final + formulário | Sim | Última chance de conversão |
| 9 | Assinatura Mais Talentto | Fixo | Gera novos clientes para a agência |
| 10 | Rodapé | Fixo | Copyright |
| — | Botão flutuante WhatsApp | Fixo | Contato a qualquer momento |

**Para remover uma seção opcional:** no `index.html`, cada seção começa com um comentário assim:
`<!-- ===================== 4. VÍDEOS (OPCIONAL ...) ===================== -->`
Apague desse comentário até o `</section>` correspondente.

**Para ter 2 ou 4 cards/depoimentos/perguntas:** apague ou copie um bloco `<article class="card">…</article>` (ou `<figure>`, ou `<details>`). A grade se ajusta sozinha.

---

## Trocar as cores (identidade do cliente)

No topo do `index.html` existe o **PAINEL DE CORES** (dentro de `:root`). Troque só os códigos:

```css
--cor-principal: #FF6600;        /* cor da marca do cliente */
--cor-principal-escura: #E65C00; /* um tom mais escuro da mesma cor */
```

Todo o resto da página acompanha automaticamente. Para descobrir o código da cor do logo do cliente, use um site como o *imagecolorpicker.com*.

---

## O que já está otimizado

- **Celular primeiro:** textos e títulos se ajustam ao tamanho da tela, botões ocupam a largura inteira no celular e têm altura mínima de toque (52px), campos do formulário não dão zoom no iPhone.
- **Carregamento rápido:** um único arquivo, sem bibliotecas externas, fontes com pesos reduzidos e conexão antecipada, vídeos carregam só no play, imagens com tamanho declarado (a página não "pula").
- **FAQ sem JavaScript:** abre e fecha mesmo se o script falhar.
- **Compartilhamento:** título, descrição e imagem aparecem bonitos quando o link é enviado no WhatsApp.
- **Acessibilidade:** respeita quem desativa animações, foco visível para teclado, contraste alto.
- **Trava de segurança:** faixa vermelha se alguma marcação ficou sem preencher.

---

## Publicação

O template funciona em qualquer hospedagem de site estático. Opções gratuitas:

- **GitHub Pages:** suba a pasta num repositório e ative em *Settings → Pages*.
- **Netlify Drop:** arraste a pasta em `app.netlify.com/drop` e receba o link na hora.
- **Hospedagem própria (ex.: tvpronta.com):** envie o `index.html` por FTP para a pasta do cliente.

## Checklist antes de entregar ao cliente

- [ ] Nenhuma faixa vermelha no topo da página
- [ ] Botão principal abre o link certo de cadastro/compra
- [ ] Botões de WhatsApp (hero, CTA final, flutuante e formulário) abrem o número certo
- [ ] Testado no celular (abrir o link no próprio WhatsApp)
- [ ] Prévia do link aparece com imagem e título ao enviar no WhatsApp
- [ ] Pixel/Analytics colado no espaço indicado no `<head>` (se o cliente usar)
