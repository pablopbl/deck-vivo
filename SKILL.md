---
name: deck-vivo
description: Cria apresentações PowerPoint (.pptx) na identidade visual da Vivo (operadora Telefônica Brasil), com roxo #660099, logo vivo + vivinho, cards arredondados e gráficos nativos. Use quando o usuário pedir um deck, apresentação, relatório em slides ou material para a Vivo, ou "no padrão Vivo", ou quiser refazer um .pptx existente na identidade Vivo.
---

# Deck Vivo

Gera decks `.pptx` na identidade visual da Vivo com a biblioteca `scripts/vivo.js` (pptxgenjs). O exemplo completo, com todos os padrões, está em `scripts/exemplo.js`, e o resultado dele é `assets/template.pptx`.

## Identidade

| Papel | Nome | Hex | Uso |
|---|---|---|---|
| Dominante | Roxo Vivo | `660099` | Capa, conclusão, títulos, card de destaque, série principal |
| Apoio | Roxo médio | `8533AD` | Rótulo de seção, círculo de ícone sobre roxo |
| Apoio | Lilás | `C9A7E0` | Série de comparação (o "antes", o grupo de referência), texto secundário sobre roxo |
| Fundo | Lilás claro | `F3ECF8` | Cards em slides claros, zebra de tabela |
| Acento | Magenta | `CD1E5A` | Um ponto de atenção por slide: selo, alerta |
| Positivo | Verde | `32A04B` | Variação favorável em tabela |
| Texto | Preto / Cinza | `1A1A1A` / `666666` | Corpo / legendas |

- **Proporção:** o roxo domina (60–70%), o lilás apoia e o magenta aparece no máximo uma vez por slide.
- **Fonte:** Arial. A fonte da marca, Vivo Type, é proprietária e não vem instalada, e a Arial abre igual em qualquer máquina.
- **Estrutura "sanduíche":** capa, conclusão e encerramento em fundo roxo com o vivinho como marca d'água; o conteúdo vai em fundo branco.
- **Motivo visual:** cards com cantos arredondados (como na interface do vivo.com.br) e ícones brancos em círculo roxo.
- **Código de cor nas comparações:** lilás = grupo de referência ou antigo; roxo = grupo em destaque ou novo. Mantenha o mesmo código no deck inteiro e declare-o num slide no início.
- **Não usar:** linha sob o título, faixas ou barras decorativas nas bordas, fundo bege, texto centralizado em parágrafos, gradiente.

Assets em `assets/`: `logo-vivo-purple.png`, `logo-vivo-white.png` (proporção 105:36), `vivinho-white.png`, `vivinho-purple.png` (25:34) e `logo-parceiro.png`, que é o logo de quem apresenta e aparece no rodapé. Para outro apresentador, passe `partnerLogo: "/caminho/logo.png"` ao `createDeck`, ou `partnerLogo: null` para tirar o logo.

## Preparação

```bash
npm install pptxgenjs react react-dom react-icons sharp
```

Rode o script com `NODE_PATH` apontando para esse `node_modules` se ele não estiver ao lado do script:

```bash
NODE_PATH=/caminho/node_modules node meu_deck.js
```

## Como montar um deck

```js
const vivo = require("<diretório desta skill>/scripts/vivo.js");
(async () => {
  const d = vivo.createDeck({ title: "Nome do arquivo", author: "Autor" });

  d.cover({ kicker: "Vivo · Projeto", title: "Título curto", subtitle: "Subtítulo", meta: "Período · Autor", date: "Outubro / 2026", notes: "..." });

  const s = d.content({ section: "Síntese", kicker: "01 · Síntese", title: "O título é a mensagem do slide", notes: "..." });
  await d.statCards(s, [
    { value: "64%", label: "de conversão", detail: "contra 25% no grupo antigo", icon: "FaTrophy" },
    { value: "−40%", label: "no custo médio", detail: "R$ 2.400 contra R$ 4.000", icon: "FaArrowTrendDown" },
  ]);

  d.statement({ section: "Conclusão", kicker: "Conclusão", runs: [{ text: "Frase com " }, { text: "o número que importa", highlight: true }], body: "Apoio." });
  d.closing({ title: "Obrigado.", subtitle: "Empresa" });
  await d.save("saida.pptx"); // grava e aplica as cores do tema
})();
```

`section` cria as seções do PowerPoint; slides seguidos com a mesma seção ficam juntos. As notas do apresentador vão em `notes`, nunca numa caixa de texto.

## Slides

| Função | Layout | Uso |
|---|---|---|
| `cover({kicker, title, subtitle, meta, date})` | VIVO_DESTAQUE (roxo) | Capa. Título com até ~28 caracteres numa linha |
| `content({section, kicker, title})` | VIVO_CONTEUDO (branco) | Todo slide de conteúdo. Devolve o slide para compor |
| `statement({kicker, runs, body})` | VIVO_DESTAQUE | Conclusão em frase grande, até ~180 caracteres. `highlight: true` pinta o trecho de lilás |
| `closing({title, subtitle})` | VIVO_DESTAQUE | Encerramento |

### Grade dos slides de conteúdo (polegadas)

- **Canvas:** 13,333 × 7,5 (16:9 wide).
- **Margem lateral:** `d.M` = 0,6. A largura útil é `d.CW` = 12,133.
- **Topo:** rótulo de seção em y 0,42, título em y 0,78, com até 2 linhas de 28 pt.
- **Conteúdo:** de `d.TOP` = 2,2 até `d.BOTTOM` = 6,6. Use espaçamento `d.GAP` = 0,3 entre blocos.
- **Legenda de gráfico:** em y 2,0, acima do gráfico, que começa em y 2,4. Nota de rodapé em y 6,25.
- **Rodapé (do layout, não mexer):** logo Vivo à esquerda, número da página e logo do parceiro à direita, em y 6,8 a 7,25.

## Composições

Todas recebem o slide devolvido por `content()`. As funções com ícone são `async`.

| Função | O que desenha |
|---|---|
| `await statCards(slide, items, {heroIndex, y, h})` | 2–4 cards de número grande (`value`, `label`, `detail`, `icon`). O card `heroIndex` (padrão 0, `-1` para nenhum) sai em roxo sólido |
| `await chartCards(slide, items, {heroIndex, y, h})` | Variação gráfica dos `statCards` (2–3 cards). `type: "columns"`: duas colunas (`categories`, `values`, `format`: `"pct"`, `"brl"` ou `"int"`), com `badge` magenta sobre a coluna mais baixa. `type: "donut"`: anel (`value` de 0 a 1, `display`, `centerLabel`) e barrinhas de evolução em `compare: [{label, value, display}]`. A primeira categoria sai em lilás (referência) e a segunda em roxo (destaque). Prefira este formato quando o número é uma comparação. Veja o slide 3 de `exemplo.js` |
| `compareBars(slide, rows, {badge, footnote, y, h})` | Barras horizontais com número grande ao lado; `badge: {value: "2,5×", label: "maior"}` gera o selo magenta |
| `await iconCards(slide, items, {cols, y, h, heroIndex})` | Grade de cards com ícone, título e `text` ou `bullets`. Com altura menor que 2,2" o card fica compacto (texto ao lado do ícone) |
| `await heroPanel(slide, {x, y, w, h, value, label, detail, icon, statementText})` | Painel roxo com número grande, para pôr ao lado de tabela ou gráfico |
| `table(slide, header, rows, {x, y, w, colW, rowH, emphasisCol, positiveCol})` | Tabela com cabeçalho roxo e zebra lilás; `emphasisCol` em roxo negrito, `positiveCol` em verde |
| `barChart(slide, series, opts)` | Gráfico nativo de colunas. `percent: true` formata em %; `barGrouping: "stacked"` empilha; `barDir: "bar"` + `catAxisOrientation: "maxMin"` gera barras horizontais em ordem. Qualquer opção do pptxgenjs sobrepõe o padrão |
| `legend(slide, x, y, [{label, color, width}])` | Legenda própria com quadradinhos; use as cores de `d.col` |
| `await callout(slide, text, {y, h, tone, icon, label})` | Faixa de destaque: `tone: "alerta"` (magenta) ou `"info"` (lilás) |
| `note(slide, text)` | Nota de rodapé em itálico cinza em y 6,25 |
| `card`, `pill`, `chip`, `text`, `await iconDot` | Primitivas para composições livres |

- **Ícones:** use nomes do `react-icons/fa6`, como `FaTrophy`, `FaChartPie`, `FaArrowTrendDown`, `FaHandshake`, `FaScaleBalanced`, `FaCalendarCheck`, `FaMagnifyingGlass`, `FaUserShield`, `FaLayerGroup`, `FaTriangleExclamation`, `FaCircleInfo`, `FaWifi`, `FaSatelliteDish`, `FaFilter` e `FaClockRotateLeft`.
- **Cores:** em formas e textos use `d.col.roxo`, `d.col.lilas`, `d.col.fundo` e as demais, que são cores do tema e acompanham uma troca de tema. Em opções que só aceitam hex (`chartColors`, `valGridLine`, `shadow`), use `d.HEX.roxo` etc.

## Regras de conteúdo

- **Uma mensagem por slide.** O título é a mensagem numa frase completa, sem ponto final. O rótulo de seção (`kicker`) é numerado: "01 · Síntese".
- **Abertura em pirâmide:** capa, síntese com 3 números, desenvolvimento por seção, conclusão em frase e encerramento.
- **Nada de slide só de texto.** Todo slide tem número grande, gráfico, tabela ou cards com ícone. Varie as composições.
- **Números de destaque:** de 44 a 64 pt, em português (`64,5%`, `R$ 2.400`, `−40%` com o sinal de menos tipográfico).
- **Tamanho mínimo de texto:** corpo 14 pt, legendas 11–12 pt. Se não couber, corte palavras ou divida o slide; não reduza a fonte.
- **Gráficos sempre nativos** (`barChart` ou `pres.addChart`), nunca imagem. Em empilhado, não use `dataLabelPosition: "outEnd"`, porque corrompe o arquivo.
- **Rótulos de porcentagem:** `dataLabelFormatCode: "0.0%"` (o PowerPoint em pt-BR mostra a vírgula). Não use `"0,0%"`.
- **Ao refazer um deck existente:** preserve todos os números, textos e notas. Corrija só erros evidentes de digitação e avise o usuário.

## Revisão visual (obrigatória)

Converta para PDF e olhe cada slide. Procure texto estourando caixas, sobreposição, título em alturas diferentes, rótulo de gráfico encavalado e baixo contraste.

- **Com LibreOffice:** `soffice --headless --convert-to pdf saida.pptx`, depois `pdftoppm -jpeg -r 80 saida.pdf slide`.
- **No Mac com PowerPoint:** o app só grava dentro do próprio container. Copie o deck para `~/Library/Containers/com.microsoft.Powerpoint/Data/` e rode:

```bash
osascript -e 'tell application "Microsoft PowerPoint"
  open POSIX file "/Users/<user>/Library/Containers/com.microsoft.Powerpoint/Data/deck.pptx"
  delay 3
  save active presentation in POSIX file "/Users/<user>/Library/Containers/com.microsoft.Powerpoint/Data/deck.pdf" as save as PDF
  close active presentation saving no
end tell'
```

Se o PowerPoint abrir o arquivo sem pedir reparo, a estrutura está íntegra.

## Uso com python-pptx (alternativa)

`assets/template.pptx` traz os layouts prontos para quem preferir python-pptx:

| Índice | Layout | Placeholders |
|---|---|---|
| 1 | VIVO_CONTEUDO | idx 102 (BODY): rótulo de seção em x 0,6, y 0,42, w 10, h 0,32; idx 103 (TITLE): título em x 0,6, y 0,78, w 12,13, h 1,05; número da página |
| 2 | VIVO_DESTAQUE | sem placeholders; fundo roxo, logo branco, vivinho e logo do parceiro já no layout |

Apague os slides de exemplo antes de adicionar os seus:

```python
from pptx import Presentation
prs = Presentation("assets/template.pptx")
while len(prs.slides):
    rId = prs.slides._sldIdLst[0].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[0]
s = prs.slides.add_slide(prs.slide_layouts[1])
for ph in s.placeholders:
    if ph.placeholder_format.idx == 102: ph.text = "01 · SÍNTESE"
    if ph.placeholder_format.idx == 103: ph.text = "Título do slide"
```

O conteúdo vai na área de y 2,2 a 6,6, com x de 0,6 a 12,733. O caminho principal continua sendo o `vivo.js`, que já traz as composições prontas.
