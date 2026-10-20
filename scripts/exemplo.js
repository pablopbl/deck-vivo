// Gera assets/template.pptx: um deck de exemplo com todos os padrões da skill.
// Os números são fictícios. Uso: node scripts/exemplo.js [saida.pptx]
const path = require("path");
const vivo = require("./vivo");

(async () => {
  const out = process.argv[2] || path.join(__dirname, "..", "assets", "template.pptx");
  const d = vivo.createDeck({ title: "Modelo de apresentação Vivo" });

  d.cover({
    kicker: "Vivo · Nome do projeto",
    title: "Título da apresentação",
    subtitle: "Subtítulo com a promessa do deck",
    meta: "Período · Nome do autor ou da empresa",
    date: "Mês / Ano",
    notes: "Capa: fundo roxo, logo Vivo branco no topo, vivinho como marca d'água.",
  });

  let s = d.content({ section: "Síntese", kicker: "01 · Síntese", title: "O título do slide é a mensagem, numa frase completa" });
  await d.statCards(s, [
    { value: "64%", label: "indicador principal", detail: "contexto do número: base de comparação e período", icon: "FaTrophy" },
    { value: "−38%", label: "segundo indicador", detail: "contexto do número", icon: "FaArrowTrendDown" },
    { value: "2,4 mil", label: "terceiro indicador", detail: "contexto do número", icon: "FaChartPie" },
  ]);

  s = d.content({ section: "Diagnóstico", kicker: "02 · Comparação", title: "Comparação entre dois grupos com selo de destaque" });
  d.compareBars(s, [
    { label: "Grupo A", value: 25, display: "25%", sub: "100 de 400" },
    { label: "Grupo B", value: 62, display: "62%", sub: "124 de 200" },
  ], { badge: { value: "2,5×", label: "maior" }, footnote: "Definição do indicador em uma linha" });
  await d.iconCards(s, [
    { title: "Critério alternativo", text: "Quando a métrica muda de definição, o resultado se mantém.", icon: "FaScaleBalanced" },
    { title: "Em todos os recortes", text: "A vantagem aparece por período, tipo e fase.", icon: "FaLayerGroup" },
  ], { y: 5.05, h: 1.55 });

  s = d.content({ section: "Diagnóstico", kicker: "02 · Evolução mensal", title: "Gráfico de colunas empilhadas com legenda própria" });
  d.legend(s, d.M, 2.0, [{ label: "Base antiga", color: d.col.lilas }, { label: "Base nova", color: d.col.roxo }]);
  const meses = ["jan", "fev", "mar", "abr", "mai", "jun"];
  d.barChart(s, [
    { name: "Base antiga", labels: meses, values: [900, 820, 700, 610, 520, 430] },
    { name: "Base nova", labels: meses, values: [100, 210, 330, 420, 540, 650] },
  ], { barGrouping: "stacked", barGapWidthPct: 35 });

  s = d.content({ section: "Diagnóstico", kicker: "02 · Por período", title: "Colunas agrupadas com rótulos em porcentagem" });
  d.legend(s, d.M, 2.0, [{ label: "Grupo A", color: d.col.lilas }, { label: "Grupo B", color: d.col.roxo }]);
  const tri = ["1º tri", "2º tri", "3º tri", "Acumulado"];
  d.barChart(s, [
    { name: "Grupo A", labels: tri, values: [0.22, 0.30, 0.26, 0.27] },
    { name: "Grupo B", labels: tri, values: [0.58, 0.72, 0.65, 0.66] },
  ], { percent: true, valAxisMaxVal: 0.8, valAxisMajorUnit: 0.2, h: 3.75 });
  d.note(s, "Nota de rodapé do gráfico: fonte, recorte ou ressalva");

  s = d.content({ section: "Resultados", kicker: "03 · Tabela", title: "Painel de destaque ao lado de uma tabela" });
  await d.heroPanel(s, { w: 3.5, h: 3.72, value: "−60%", label: "indicador que a tabela explica", detail: "R$ 1.200 contra R$ 3.000 no grupo de comparação" });
  d.table(s, ["Indicador", "Grupo A", "Grupo B", "Diferença"], [
    ["Primeiro indicador", "R$ 4.000", "R$ 2.400", "−40%"],
    ["Segundo indicador", "70%", "30%", "−40 p.p."],
    ["Terceiro indicador", "R$ 3.000", "R$ 1.200", "−60%"],
  ], { x: d.M + 3.8, w: d.CW - 3.8, colW: [3.83, 1.4, 1.4, 1.7], rowH: 0.62, emphasisCol: 2, positiveCol: 3 });
  await d.callout(s, "frase de contexto que complementa a tabela, em uma linha.", { y: 6.05, h: 0.55, tone: "info", icon: "FaCircleInfo", label: "Destaque:" });

  s = d.content({ section: "Resultados", kicker: "03 · Alerta", title: "Gráfico de barras horizontais com painel e alerta" });
  await d.heroPanel(s, { h: 3.2, value: "38", label: "ocorrências no período", icon: "FaTriangleExclamation" });
  d.card(s, d.M + 3.8, d.TOP, d.CW - 3.8, 3.2);
  d.barChart(s, [{ name: "Ocorrências", labels: ["Categoria um", "Categoria dois", "Categoria três", "Categoria quatro"], values: [15, 13, 7, 3] }], {
    x: d.M + 4.0, y: d.TOP + 0.15, w: d.CW - 4.2, h: 2.9, barDir: "bar", catAxisOrientation: "maxMin", valAxisHidden: true, valGridLine: { style: "none" }, barGapWidthPct: 45, catAxisLabelFontSize: 13,
  });
  await d.callout(s, "Frase de alerta: o que esse número significa para a decisão.", { y: 5.75 });

  s = d.content({ section: "Atuação", kicker: "04 · Frentes", title: "Grade de cards com ícone" });
  await d.iconCards(s, [
    { title: "Primeira frente", text: "Descrição curta do que é feito e por quem.", icon: "FaScaleBalanced" },
    { title: "Segunda frente", text: "Descrição curta do que é feito e por quem.", icon: "FaHandshake" },
    { title: "Terceira frente", text: "Descrição curta do que é feito e por quem.", icon: "FaCalendarCheck" },
    { title: "Quarta frente", text: "Descrição curta do que é feito e por quem.", icon: "FaSatelliteDish" },
  ]);

  s = d.content({ section: "Atuação", kicker: "04 · Duas colunas", title: "Duas colunas com listas, uma em destaque" });
  await d.iconCards(s, [
    { title: "Coluna informativa", bullets: ["Primeiro ponto da lista.", "Segundo ponto da lista.", "Terceiro ponto da lista."], icon: "FaMagnifyingGlass" },
    { title: "Coluna em destaque", bullets: ["Primeiro ponto da lista.", "Segundo ponto da lista."], icon: "FaUserShield" },
  ], { heroIndex: 1 });

  d.statement({
    section: "Conclusão", kicker: "05 · Conclusão",
    runs: [{ text: "A frase de conclusão cabe em até quatro linhas e destaca " }, { text: "o número que importa", highlight: true }, { text: "." }],
    body: "Parágrafo de apoio com a consequência ou o próximo passo.",
  });
  d.closing({ section: "Conclusão", title: "Obrigado.", subtitle: "Nome da empresa" });

  await d.save(out);
  console.log("ok", out);
})().catch((e) => { console.error(e); process.exit(1); });
