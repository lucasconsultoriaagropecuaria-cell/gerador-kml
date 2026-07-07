# Melhorias de maior alavancagem no setup Claude Code — Diagnóstico

**Data:** 2026-07-07 · **Status:** apenas diagnóstico, nada foi executado.

## ⚠️ Limitação importante encontrada antes de tudo

**Não existem transcrições de sessões anteriores acessíveis neste ambiente.** Esta sessão
roda num container remoto (claude.ai/code) que é criado do zero a cada sessão. Varri o
container inteiro: só existe a transcrição da sessão atual.

Onde as suas sessões realmente ficam:

| Origem da sessão | Onde fica a transcrição | Acessível daqui? |
|---|---|---|
| Celular / web (claude.ai) | Container efêmero na nuvem, descartado após a sessão | ❌ |
| Claude Code CLI em cada computador | `~/.claude/projects/` **daquele computador** (não sincroniza) | ❌ |
| Esta sessão | Container atual | ✅ (só ela) |

Também não há vestígios no repositório: 1 commit único ("Add files via upload"), zero PRs,
zero issues, zero branches de sessões anteriores.

**Consequência:** o diagnóstico abaixo foi extraído dos artefatos disponíveis — o código do
app, o histórico git e a forma como o repositório foi montado — que já contam bastante
sobre o fluxo de trabalho. O item 1 descreve como fazer a mineração de transcrições de
verdade.

---

## Candidatos, ordenados por impacto

### 1. Mineração de transcrições rodando localmente — **skill nova**
**Sinal:** o pedido desta sessão é irrealizável da nuvem; as transcrições ricas estão nos
seus computadores, em `~/.claude/projects/*.jsonl`.
**Proposta:** uma skill/script para rodar **em cada computador** que: lê os `.jsonl` locais,
extrai sinais brutos (comandos repetidos, erros recorrentes, permissões negadas, pedidos
que você reformula várias vezes), e grava um resumo anônimo num arquivo que você commita
neste repositório. Com os resumos de todas as máquinas agregados, refaço esta análise com
dados reais.
**Impacto:** alto — destrava o objetivo original e vira hábito recorrente.

### 2. Publicar o app no GitHub Pages — **automação**
**Sinal:** o app é um HTML estático único, e o commit "Add files via upload" indica que o
arquivo foi gerado fora do git e subido manualmente pela interface web. Você usa celular e
vários computadores — hoje precisa levar o arquivo para cada um.
**Proposta:** ativar GitHub Pages no repositório. O gerador passa a ter uma URL fixa que
funciona em qualquer aparelho, sempre na versão mais recente. Custo: minutos, zero
manutenção.
**Impacto:** alto — é a melhoria com melhor relação esforço/benefício para o uso diário.

### 3. Funcionar offline no campo — **correção**
**Sinal:** o app promete "tudo roda no seu navegador", mas carrega `pdf.js` e `jszip` de um
CDN externo. Sem internet (situação comum em área rural), o app abre mas **não processa
nenhum PDF**, falhando de forma silenciosa/confusa.
**Proposta:** embutir as duas bibliotecas no próprio HTML (ou vendorizar no repositório).
O arquivo cresce ~1 MB e passa a funcionar 100% offline.
**Impacto:** alto para uso em campo; médio se você só usa no escritório.

### 4. Proteção contra mudança de layout do PDF do BB Pl@ntar — **correção**
**Sinal:** a extração depende de regexes acopladas ao layout atual do PDF
(`Capturada (ha)`, `Gleba ... Área Capturada`, formato das coordenadas). Se o banco mudar
o relatório, a extração quebra ou — pior — extrai área/coordenadas erradas sem avisar.
**Proposta:** (a) pasta `exemplos/` com 2–3 PDFs reais anonimizados servindo de casos de
teste; (b) validações de sanidade no app: coordenadas dentro do Brasil, anel com área
compatível com os hectares declarados, alerta visível quando algum campo não for
encontrado.
**Impacto:** médio-alto — é risco de erro em documento de crédito rural, não só bug.

### 5. Trazer o desenvolvimento para dentro do Claude Code + CLAUDE.md — **correção**
**Sinal:** o código foi construído fora do git (provavelmente em conversas no claude.ai) e
subido pronto. Sem histórico de versões, cada sessão nova começa sem contexto do projeto.
**Proposta:** (a) editar o app via sessões Claude Code que commitam e fazem push — ganha
histórico, diff e rollback; (b) criar um `CLAUDE.md` curto com as regras do domínio:
nomenclatura "NOME - área ha", estilo amarelo (linha 5 / 20% opacidade), KML consolidado,
polígono sempre completo. Toda sessão futura (celular ou computador) já parte disso.
**Impacto:** médio — reduz retrabalho de explicar o contexto a cada sessão.

### 6. Skill de domínio "glebas/KML" — **aguardando dados (por ora, nada)**
**Sinal:** se as transcrições locais mostrarem tarefas recorrentes (ajustar estilo,
converter formatos, validar no Google Earth), vale codificar isso numa skill. Sem as
transcrições, seria chute.
**Decisão:** reavaliar depois do item 1.

### 7. Arquitetura do app — **nada**
Single-file HTML sem framework é a escolha certa para este escopo. Não refatorar.

---

## Próximo passo sugerido
Me diga se quer que eu prepare o script do item 1 (você rodaria em cada computador) e/ou
execute os itens 2, 3 e 5 — todos de baixo risco. Nada será executado sem sua confirmação.
