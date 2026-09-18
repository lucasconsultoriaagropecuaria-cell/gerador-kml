#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Minerador de sessões do Claude Code — roda LOCALMENTE em cada computador.

Lê as transcrições em ~/.claude/projects/*/*.jsonl e extrai sinais brutos:
  - sessões por projeto e período coberto
  - comandos Bash mais usados
  - erros de ferramenta recorrentes
  - permissões negadas / interrupções do usuário
  - amostra dos pedidos (primeira frase de cada prompt)

Gera: diagnostico-local/resumo-<hostname>.md (para commitar neste repositório).
Nada é enviado para a internet; revise o arquivo antes de commitar.

Como usar em cada computador:
    python3 diagnostico-local/minerar-sessoes.py
ou simplesmente peça ao Claude Code local:
    "rode diagnostico-local/minerar-sessoes.py e commite o resumo gerado"
"""

import json
import os
import re
import socket
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

CLAUDE_DIR = Path(os.environ.get("CLAUDE_CONFIG_DIR", Path.home() / ".claude"))
PROJECTS = CLAUDE_DIR / "projects"
OUT_DIR = Path(__file__).resolve().parent
MAX_PROMPT_CHARS = 120

INTERRUPT_MARKERS = ("[Request interrupted by user", "[Solicitação interrompida")
DENIAL_MARKERS = ("doesn't want to proceed", "user declined", "permission denied by user",
                  "The user denied", "rejected the tool")
RETRY_WORDS = re.compile(r"\b(de novo|novamente|não funcionou|nao funcionou|errado|erro de novo|"
                         r"tente outra vez|continua|ainda não|ainda nao|mesma coisa)\b", re.I)


def iter_transcripts():
    if not PROJECTS.is_dir():
        return
    for proj in sorted(PROJECTS.iterdir()):
        if proj.is_dir():
            for f in sorted(proj.glob("*.jsonl")):
                yield proj.name, f


def text_of(content):
    """Extrai texto de message.content (string ou lista de blocos)."""
    if isinstance(content, str):
        return content
    out = []
    if isinstance(content, list):
        for block in content:
            if isinstance(block, dict):
                if block.get("type") == "text":
                    out.append(block.get("text", ""))
                elif block.get("type") == "tool_result":
                    c = block.get("content")
                    if isinstance(c, str):
                        out.append(c)
                    elif isinstance(c, list):
                        for cc in c:
                            if isinstance(cc, dict) and cc.get("type") == "text":
                                out.append(cc.get("text", ""))
    return "\n".join(out)


def first_line(s, limit=MAX_PROMPT_CHARS):
    s = " ".join(s.strip().split())
    return (s[:limit] + "…") if len(s) > limit else s


def main():
    stats = defaultdict(lambda: {
        "sessions": set(), "user_prompts": [], "first_ts": None, "last_ts": None,
        "bash_cmds": Counter(), "tools": Counter(), "errors": Counter(),
        "denials": 0, "interrupts": 0, "retry_signals": 0,
    })
    n_files = 0

    for proj_name, path in iter_transcripts():
        n_files += 1
        st = stats[proj_name]
        try:
            lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
        except OSError as e:
            print(f"  aviso: não consegui ler {path}: {e}", file=sys.stderr)
            continue
        for line in lines:
            line = line.strip()
            if not line:
                continue
            try:
                rec = json.loads(line)
            except json.JSONDecodeError:
                continue
            sid = rec.get("sessionId")
            if sid:
                st["sessions"].add(sid)
            ts = rec.get("timestamp")
            if ts:
                st["first_ts"] = min(st["first_ts"] or ts, ts)
                st["last_ts"] = max(st["last_ts"] or ts, ts)

            msg = rec.get("message") or {}
            rtype = rec.get("type")

            if rtype == "user":
                txt = text_of(msg.get("content"))
                if txt:
                    if any(m in txt for m in INTERRUPT_MARKERS):
                        st["interrupts"] += 1
                    elif not rec.get("isMeta") and "tool_use_id" not in txt[:200] \
                            and not txt.startswith("<") and not rec.get("toolUseResult"):
                        st["user_prompts"].append(first_line(txt))
                        if RETRY_WORDS.search(txt):
                            st["retry_signals"] += 1
                # resultados de ferramenta com erro
                if isinstance(msg.get("content"), list):
                    for block in msg["content"]:
                        if isinstance(block, dict) and block.get("type") == "tool_result" \
                                and block.get("is_error"):
                            err = first_line(text_of(block.get("content") or ""), 100)
                            if err:
                                if any(d in err for d in DENIAL_MARKERS):
                                    st["denials"] += 1
                                else:
                                    st["errors"][err] += 1

            elif rtype == "assistant" and isinstance(msg.get("content"), list):
                for block in msg["content"]:
                    if isinstance(block, dict) and block.get("type") == "tool_use":
                        name = block.get("name", "?")
                        st["tools"][name] += 1
                        if name == "Bash":
                            cmd = (block.get("input") or {}).get("command", "")
                            head = cmd.strip().split()
                            if head:
                                st["bash_cmds"][head[0]] += 1

    host = socket.gethostname()
    out_path = OUT_DIR / f"resumo-{host}.md"
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    lines = [f"# Resumo de sessões Claude Code — máquina `{host}`",
             f"\nGerado em {now} · {n_files} arquivo(s) de transcrição em `{PROJECTS}`\n"]

    if not stats:
        lines.append("**Nenhuma transcrição encontrada nesta máquina.**\n")
    for proj, st in sorted(stats.items()):
        prompts = st["user_prompts"]
        lines.append(f"\n## Projeto `{proj}`")
        lines.append(f"- Sessões: {len(st['sessions']) or 1} · Prompts do usuário: {len(prompts)}"
                     f" · Período: {st['first_ts'] or '?'} → {st['last_ts'] or '?'}")
        lines.append(f"- Interrupções: {st['interrupts']} · Permissões negadas: {st['denials']}"
                     f" · Sinais de retrabalho (\"de novo\", \"errado\"...): {st['retry_signals']}")
        if st["tools"]:
            top = ", ".join(f"{t} ({n})" for t, n in st["tools"].most_common(8))
            lines.append(f"- Ferramentas mais usadas: {top}")
        if st["bash_cmds"]:
            top = ", ".join(f"`{c}` ({n})" for c, n in st["bash_cmds"].most_common(10))
            lines.append(f"- Comandos Bash mais usados: {top}")
        if st["errors"]:
            lines.append("- Erros de ferramenta recorrentes:")
            for err, n in st["errors"].most_common(8):
                lines.append(f"  - ({n}×) {err}")
        if prompts:
            lines.append(f"- Amostra de pedidos (até 20 de {len(prompts)}):")
            for p in prompts[:20]:
                lines.append(f"  - {p}")

    out_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"✔ Resumo gravado em: {out_path}")
    print("Revise o conteúdo (pode conter trechos dos seus pedidos) e depois commite:")
    print(f"  git add '{out_path.name}' && git commit -m 'Resumo de sessões: {host}' && git push")


if __name__ == "__main__":
    main()
