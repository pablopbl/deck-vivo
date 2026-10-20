# deck-vivo

Skill para Claude que gera apresentações PowerPoint (`.pptx`) na identidade visual da Vivo: roxo `#660099`, logo vivo + vivinho, cards arredondados, ícones e gráficos nativos editáveis.

![Prévia dos slides](docs/preview.jpg)

## Instalação

```bash
npx skills add pablopbl/deck-vivo -g -y
```

Ou copie esta pasta para `~/.claude/skills/deck-vivo`.

Dependências para gerar os decks:

```bash
npm install pptxgenjs react react-dom react-icons sharp
```

## Uso

Peça ao Claude algo como *"monta um deck no padrão Vivo com estes números"* ou *"refaz este .pptx na identidade da Vivo"*. As instruções completas estão em [SKILL.md](SKILL.md).

Para gerar o deck de exemplo à mão:

```bash
node scripts/exemplo.js exemplo.pptx
```

## Conteúdo

| Caminho | O que é |
|---|---|
| `SKILL.md` | Instruções da skill: identidade, grade, composições e revisão |
| `scripts/vivo.js` | Biblioteca (pptxgenjs) com layouts, slides e composições |
| `scripts/exemplo.js` | Deck de exemplo com todos os padrões |
| `assets/template.pptx` | Resultado do exemplo; também serve de template para python-pptx |
| `assets/*.png` | Logos usados nos layouts |

## Marcas

O código é MIT. As marcas e logos em `assets/` (Vivo e o logo do parceiro) pertencem aos seus titulares e estão aqui só para compor os layouts; a licença MIT não se aplica a eles. A identidade foi levantada do site público vivo.com.br; a fonte proprietária da marca (Vivo Type) foi substituída por Arial.
