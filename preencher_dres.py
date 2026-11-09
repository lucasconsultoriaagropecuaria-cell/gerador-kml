# -*- coding: utf-8 -*-
"""
PREENCHER PLANILHA DRES A PARTIR DOS PRINTS DO "RELATÓRIO DE COLETAS"
=====================================================================

Lê os prints (fotos de tela) dos cartões do Relatório de coletas do portal
— aqueles com "Coleta 1002-B01", talhão, data e hora, coordenadas etc. —
e preenche a planilha de campo DRES (aba "LEITURAS DDMM").

O que é preenchido em cada ponto:
    A  PONTO        número do ponto (só se estiver vazio)
    B  Código       ex.: 1002-B01
    C  OBS          "T 39." (talhão da coleta) na frente da observação
    D  hora         hora da coleta
    E  altitude     opcional (--altitude), pelo modelo de relevo SRTM 30 m
    F  latitude     ex.: 18 49 52,58 S
    G  longitude    ex.: 48 34 27,87 O
    H  E            espessura de cada camada (cm), uma camada por linha  } avaliação DRES
    I  Q            nota de qualidade estrutural Qe (1 a 6)               } feita pelas
    J  IQEA         soma(E x Q) / soma(E)                                 } fotos
    K  Leitura      MUITO BOM / BOM / REGULAR / RUIM / MUITO RUIM

Avaliação DRES (método Embrapa): o Claude olha as fotos do cartão (trincheira,
bloco, bandeja), separa as camadas, dá a espessura e a nota Qe de cada uma e
escreve uma observação curta (ex.: "Raízes superficiais. Crescimento lateral.").
Se E e Q já foram anotados no campo, eles são mantidos e o IQEA/Leitura são
calculados a partir deles. A nota pela foto é uma ESTIMATIVA: confira antes de
fechar o relatório. Fotos originais em boa resolução avaliam melhor que o print.
Use --sem-dres para só transcrever os dados.

Regras:
  * A aba é escolhida pela data do print (04/11/2026 -> "LEITURAS 0411").
    Se a aba não existir, é criada copiando o modelo da primeira aba LEITURAS.
  * Se o código já estiver na aba, completa a mesma linha; senão usa o
    próximo ponto vazio.
  * Por padrão só escreve em células vazias (não apaga o que foi digitado
    no campo). Use --sobrescrever para substituir.
  * A planilha original não é alterada: o resultado vai para um arquivo novo.
  * A leitura de cada print fica guardada em <saida>.leituras.json; rodar de
    novo não lê (nem cobra) o mesmo print outra vez. Dá para corrigir algum
    valor nesse arquivo e rodar de novo.

Requisitos (uma vez):
    pip install anthropic openpyxl pydantic

Chave da API da Anthropic (uma vez por janela do Prompt de Comando):
    set ANTHROPIC_API_KEY=sk-ant-...

Uso:
    python preencher_dres.py --planilha DRES_Planilha_de_Campo_AROEIRA.xlsx --fotos C:\\prints
    python preencher_dres.py --planilha DRES.xlsx --fotos print1.png print2.png --altitude
"""

import argparse
import base64
import copy
import datetime as dt
import json
import mimetypes
import re
import sys
import unicodedata
import urllib.request
from pathlib import Path
from typing import Optional

import anthropic
import openpyxl
from pydantic import BaseModel, Field

MODELO_PADRAO = "claude-opus-5-5"
EXTENSOES = {".png", ".jpg", ".jpeg", ".webp", ".gif"}
PONTOS_POR_PAGINA = 7
LINHAS_POR_PONTO = 4

# Colunas da planilha
COL_PONTO, COL_COD, COL_OBS, COL_HORA, COL_ALT, COL_LAT, COL_LON = 1, 2, 3, 4, 5, 6, 7
COL_E, COL_Q, COL_IQEA, COL_LEITURA = 8, 9, 10, 11

CLASSES_IQEA = [(5.0, "MUITO BOM"), (4.0, "BOM"), (3.0, "REGULAR"), (2.0, "RUIM"), (0.0, "MUITO RUIM")]


# --------------------------------------------------------------------------
# 1) Leitura do print com o Claude
# --------------------------------------------------------------------------
class Camada(BaseModel):
    espessura_cm: float = Field(description="Espessura da camada em cm (a soma das camadas é a profundidade do bloco, normalmente 25 cm)")
    nota_q: int = Field(description="Nota de qualidade estrutural Qe da camada, inteiro de 1 a 6, pelo método DRES")
    descricao: str = Field(description="Por que essa nota: estrutura, agregados, porosidade, raízes, sinais de compactação")


class AvaliacaoDRES(BaseModel):
    avaliavel: bool = Field(description="false se as fotos não permitem avaliar o perfil (bloco não visível, foto borrada etc.)")
    camadas: list[Camada] = Field(description="Camadas de cima para baixo (1 a 4)")
    observacao_campo: str = Field(description='Observação curta no estilo do relatório de campo, ex.: "Raízes superficiais. Crescimento lateral."')
    confianca: str = Field(description='"alta", "media" ou "baixa"')


class Coleta(BaseModel):
    codigo: Optional[str] = Field(None, description='Código da coleta, ex.: "1002-B01" (do título "Coleta 1002-B01")')
    fazenda: Optional[str] = Field(None, description="Nome da fazenda/estado da linha azul do topo, ex.: MATO GROSSO")
    talhao: Optional[str] = Field(None, description='Valor de "TALHÃO DA COLETA"')
    area_bloco_ha: Optional[float] = Field(None, description='Valor de "ÁREA DO BLOCO" em hectares, como número')
    talhoes_bloco: Optional[str] = Field(None, description='Valor de "TALHÕES DO BLOCO"')
    data: Optional[str] = Field(None, description='Data de "DATA E HORA" no formato DD/MM/AAAA')
    hora: Optional[str] = Field(None, description='Hora de "DATA E HORA" no formato HH:MM (24 h)')
    latitude: Optional[float] = Field(None, description='Primeiro número de "COORDENADAS", em graus decimais, com sinal')
    longitude: Optional[float] = Field(None, description='Segundo número de "COORDENADAS", em graus decimais, com sinal')
    origem_coordenada: Optional[str] = Field(None, description='Texto entre parênteses após as coordenadas, ex.: "ponto planejado"')
    observacao: Optional[str] = Field(None, description='Texto de "OBSERVAÇÃO"')
    coletado_por: Optional[str] = Field(None, description='Texto de "COLETADO POR"')
    dres: Optional[AvaliacaoDRES] = Field(None, description="Avaliação DRES feita a partir das fotos do cartão (nula se não foi pedida)")


METODO_DRES = """Você é um agrônomo experiente no método DRES - Diagnóstico Rápido da Estrutura do Solo (Embrapa).

Como o DRES funciona:
- Abre-se uma pequena trincheira e retira-se com a pá reta um bloco de solo de cerca de 25 cm de profundidade,
  20 cm de largura e 10 cm de espessura, que é colocado na bandeja.
- O avaliador separa o bloco em camadas (geralmente 1 a 3) de aspecto estrutural diferente, mede a espessura (E, em cm)
  de cada uma e dá a cada camada uma nota de qualidade estrutural Qe de 1 a 6:
    6 - Estrutura grumosa/granular, agregados pequenos e arredondados que se desfazem facilmente; muitos poros e
        canais biológicos; raízes abundantes, finas, distribuídas por toda a camada; sem sinais de compactação.
    5 - Predomínio de grumos e pequenos blocos subangulares friáveis; boa porosidade; raízes abundantes e bem distribuídas.
    4 - Blocos subangulares médios, que se rompem com pouca força em agregados menores; porosidade e raízes ainda
        presentes por toda a camada, com alguma redução.
    3 - Blocos maiores e mais firmes, angulares/subangulares; porosidade visível reduzida; raízes concentradas entre os
        agregados e nas fissuras; início de compactação.
    2 - Camada compactada: blocos grandes, angulosos, de faces lisas, que exigem força para romper; poucos poros;
        raízes raras, tortas, achatadas ou com crescimento lateral (desviando da camada); fissuras horizontais.
    1 - Estrutura maciça ou laminar, muito compactada, sem agregação nem porosidade visível; raízes ausentes ou
        totalmente desviadas.
- Índice do ponto: IQEA = soma(E x Qe) / soma(E). Ex.: camadas 5 cm Qe 3 e 20 cm Qe 2 -> (15+40)/25 = 2,2.
- Classes do IQEA: 5,0-6,0 Muito bom | 4,0-4,9 Bom | 3,0-3,9 Regular | 2,0-2,9 Ruim | 1,0-1,9 Muito ruim.

Sinais típicos nas fotos: camada superficial solta/grumosa de poucos cm sobre camada adensada em blocos; faces lisas e
fissuras horizontais indicam compactação (Qe 1-2); raízes que correm na horizontal ("crescimento lateral",
"raízes superficiais") indicam impedimento físico; torrões que se desfazem em grumos com raízes finas indicam Qe 5-6.
Seja realista e conservador: na dúvida entre duas notas, escolha a que melhor explica o que se vê no bloco e informe
a confiança. Não invente camadas que não aparecem."""

INSTRUCOES = (
    "Esta imagem é um print de um cartão do 'Relatório de coletas' de amostragem de solo. "
    "Transcreva exatamente os campos de texto do cartão (lado esquerdo). Ignore o texto pequeno sobreposto às fotos. "
    "Copie os números exatamente como aparecem, convertendo a vírgula decimal da área em ponto. "
    "Se um campo não aparecer ou estiver ilegível, deixe nulo."
)
INSTRUCOES_DRES = (
    " Depois, use as fotos do cartão (trincheira, bloco e bandeja) para fazer a avaliação DRES do ponto: "
    "camadas de cima para baixo com espessura e nota Qe, e uma observação curta de campo."
)


def ler_print(client: anthropic.Anthropic, caminho: Path, modelo: str, avaliar: bool) -> dict:
    tipo = mimetypes.guess_type(caminho.name)[0] or "image/png"
    if tipo == "image/jpg":
        tipo = "image/jpeg"
    dados = base64.standard_b64encode(caminho.read_bytes()).decode("utf-8")
    resp = client.beta.messages.parse(
        model=modelo,
        max_tokens=16000,
        output_config={"effort": "high" if avaliar else "low"},
        system=METODO_DRES if avaliar else "Você transcreve dados de prints de relatórios agrícolas.",
        betas=["server-side-fallback-2026-07-01"],
        fallbacks="default",
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": tipo, "data": dados}},
                {"type": "text", "text": INSTRUCOES + (INSTRUCOES_DRES if avaliar else "")},
            ],
        }],
        output_format=Coleta,
    )
    if resp.stop_reason == "refusal" or resp.parsed_output is None:
        raise RuntimeError(f"não foi possível ler (stop_reason={resp.stop_reason})")
    dados = resp.parsed_output.model_dump()
    if not avaliar:
        dados["dres"] = None
    return dados


# --------------------------------------------------------------------------
# 2) Conversões
# --------------------------------------------------------------------------
def decimal_para_gms(valor: float, eixo: str) -> str:
    """-18.831272 (lat) -> '18 49 52,58 S'  |  -48.574409 (lon) -> '48 34 27,87 O'"""
    if eixo == "lat":
        hemi = "S" if valor < 0 else "N"
    else:
        hemi = "O" if valor < 0 else "L"
    v = abs(valor)
    graus = int(v)
    minutos_f = (v - graus) * 60
    minutos = int(minutos_f)
    segundos = round((minutos_f - minutos) * 60, 2)
    if segundos >= 60:
        segundos -= 60
        minutos += 1
    if minutos >= 60:
        minutos -= 60
        graus += 1
    return f"{graus} {minutos:02d} {segundos:05.2f} {hemi}".replace(".", ",")


def texto_para_hora(txt: Optional[str]) -> Optional[dt.time]:
    m = re.search(r"(\d{1,2})[:h](\d{2})", txt or "")
    return dt.time(int(m.group(1)), int(m.group(2))) if m else None


def normalizar_codigo(txt) -> str:
    return re.sub(r"\s+", "", str(txt or "")).upper()


def sem_acento(txt: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", txt) if unicodedata.category(c) != "Mn")


def buscar_altitude(lat: float, lon: float) -> Optional[int]:
    url = f"https://api.opentopodata.org/v1/srtm30m?locations={lat},{lon}"
    try:
        with urllib.request.urlopen(url, timeout=20) as r:
            elev = json.load(r)["results"][0]["elevation"]
        return round(elev) if elev is not None else None
    except Exception as e:  # sem internet, limite do serviço etc.
        print(f"    (altitude não obtida: {e})")
        return None


def classe_iqea(iqea: float) -> str:
    return next(nome for limite, nome in CLASSES_IQEA if round(iqea, 1) >= limite)


def numero(v) -> Optional[float]:
    try:
        return float(str(v).replace(",", "."))
    except (TypeError, ValueError):
        return None


# --------------------------------------------------------------------------
# 3) Planilha
# --------------------------------------------------------------------------
def eh_cabecalho(ws, linha: int) -> bool:
    v = ws.cell(linha, 1).value
    return isinstance(v, str) and v.strip().upper() == "PONTO"


def linhas_dos_pontos(ws) -> list:
    """Primeira linha de cada ponto: abaixo de cada cabeçalho 'PONTO', de 4 em 4 linhas."""
    cabecalhos = [r for r in range(1, ws.max_row + 1) if eh_cabecalho(ws, r)]
    linhas = []
    for i, h in enumerate(cabecalhos):
        prox = cabecalhos[i + 1] if i + 1 < len(cabecalhos) else ws.max_row + 2
        for k in range(PONTOS_POR_PAGINA):
            r = h + 1 + k * LINHAS_POR_PONTO
            if r + LINHAS_POR_PONTO - 1 >= prox:
                break
            # para em rodapés como "*plantada", "Anexo I ..."
            if any(isinstance(ws.cell(x, 1).value, str) for x in range(r, r + LINHAS_POR_PONTO)):
                break
            linhas.append(r)
    return linhas


def nome_aba(data_txt: str) -> Optional[str]:
    m = re.search(r"(\d{1,2})/(\d{1,2})", data_txt or "")
    return f"LEITURAS {int(m.group(1)):02d}{int(m.group(2)):02d}" if m else None


def obter_aba(wb, nome: str, data_txt: str):
    for ws in wb.worksheets:
        if ws.title.strip().upper() == nome:
            return ws
    modelo = next((ws for ws in wb.worksheets if sem_acento(ws.title).upper().startswith("LEITURAS")), None)
    if modelo is None:
        raise SystemExit("A planilha não tem nenhuma aba 'LEITURAS ...' para servir de modelo.")
    nova = wb.copy_worksheet(modelo)
    nova.title = nome
    for r in linhas_dos_pontos(nova):
        for x in range(r, r + LINHAS_POR_PONTO):
            for c in range(COL_COD, 12):  # B..K
                if not isinstance(nova.cell(x, c), openpyxl.cell.cell.MergedCell):
                    nova.cell(x, c).value = None
    titulo = nova.cell(1, 1).value
    if isinstance(titulo, str):
        nova.cell(1, 1).value = re.sub(r"Data:\s*\S+", f"Data: {data_txt}", titulo)
    print(f"  * aba '{nome}' criada a partir de '{modelo.title}'")
    return nova


def escrever(ws, linha, col, valor, sobrescrever, log, rotulo):
    if valor in (None, ""):
        return
    cel = ws.cell(linha, col)
    if cel.value not in (None, "") and not sobrescrever:
        if str(cel.value).strip() != str(valor).strip():
            fmt = lambda v: v.strftime("%H:%M") if isinstance(v, dt.time) else str(v)
            log.append(f"{rotulo} mantido ({fmt(cel.value)}; print: {fmt(valor)})")
        return
    cel.value = valor
    if isinstance(valor, dt.time):
        cel.number_format = "h:mm"


def montar_obs(atual, c: dict, incluir_obs_app: bool) -> Optional[str]:
    partes = []
    if c.get("talhao"):
        partes.append(f"T {c['talhao']}.")
    if incluir_obs_app and c.get("observacao"):
        partes.append(c["observacao"].strip())
    dres = c.get("dres") or {}
    if atual in (None, "") and dres.get("avaliavel") and dres.get("observacao_campo"):
        partes.append(dres["observacao_campo"].strip())
    prefixo = " ".join(partes)
    if not prefixo:
        return None
    if atual in (None, ""):
        return prefixo
    if c.get("talhao") and re.match(rf"\s*T\s*{re.escape(str(c['talhao']))}\b", str(atual)):
        return None  # já tem o talhão
    return f"{prefixo} {atual}"


def lancar_dres(ws, linha: int, dres: Optional[dict], sobrescrever: bool) -> list:
    """Escreve E e Q das camadas (uma por linha do ponto) e calcula IQEA e Leitura."""
    log = []
    linhas = range(linha, linha + LINHAS_POR_PONTO)
    ja_tem = any(ws.cell(r, COL_E).value not in (None, "") or ws.cell(r, COL_Q).value not in (None, "")
                 for r in linhas)
    camadas = (dres or {}).get("camadas") or []
    if dres and dres.get("avaliavel") and camadas and (sobrescrever or not ja_tem):
        for r in linhas:
            ws.cell(r, COL_E).value = None
            ws.cell(r, COL_Q).value = None
        for r, cam in zip(linhas, camadas):
            e = round(float(cam["espessura_cm"]))
            q = min(6, max(1, int(cam["nota_q"])))
            ws.cell(r, COL_E).value = e
            ws.cell(r, COL_Q).value = q
        if len(camadas) > LINHAS_POR_PONTO:
            log.append(f"só {LINHAS_POR_PONTO} de {len(camadas)} camadas cabem no ponto")
        log.append(f"DRES pela foto (confiança {dres.get('confianca')})")
    elif ja_tem and dres and dres.get("avaliavel"):
        log.append("E/Q de campo mantidos")
    elif dres and not dres.get("avaliavel"):
        log.append("fotos não permitem avaliar o DRES")

    # IQEA e Leitura a partir do que está na planilha (de campo ou da foto)
    pares = [(numero(ws.cell(r, COL_E).value), numero(ws.cell(r, COL_Q).value)) for r in linhas]
    pares = [(e, q) for e, q in pares if e and q]
    if pares:
        iqea = round(sum(e * q for e, q in pares) / sum(e for e, _ in pares), 1)
        if sobrescrever or ws.cell(linha, COL_IQEA).value in (None, ""):
            ws.cell(linha, COL_IQEA).value = iqea
            ws.cell(linha, COL_IQEA).number_format = "0.0"
        if sobrescrever or ws.cell(linha, COL_LEITURA).value in (None, ""):
            ws.cell(linha, COL_LEITURA).value = classe_iqea(iqea)
        log.append(f"IQEA {str(iqea).replace('.', ',')} {classe_iqea(iqea)}")
    return log


def lancar(wb, c: dict, args) -> str:
    if not c.get("codigo"):
        return "sem código no print — ignorado"
    aba = nome_aba(c.get("data"))
    if not aba:
        return "sem data no print — ignorado"
    ws = obter_aba(wb, aba, c["data"])
    pontos = linhas_dos_pontos(ws)
    cod = normalizar_codigo(c["codigo"])

    linha = next((r for r in pontos if normalizar_codigo(ws.cell(r, COL_COD).value) == cod), None)
    novo = linha is None
    if novo:
        linha = next((r for r in pontos
                      if ws.cell(r, COL_COD).value in (None, "") and ws.cell(r, COL_HORA).value in (None, "")), None)
        if linha is None:
            return f"aba {aba} sem ponto vazio — ignorado"

    log = []
    sob = args.sobrescrever
    if ws.cell(linha, COL_PONTO).value in (None, ""):
        ws.cell(linha, COL_PONTO).value = pontos.index(linha) + 1
    escrever(ws, linha, COL_COD, c["codigo"].strip(), sob and not novo, log, "código")
    obs = montar_obs(ws.cell(linha, COL_OBS).value, c, args.obs_completa)
    if obs:
        ws.cell(linha, COL_OBS).value = obs
    escrever(ws, linha, COL_HORA, texto_para_hora(c.get("hora")), sob, log, "hora")
    if c.get("latitude") is not None and c.get("longitude") is not None:
        escrever(ws, linha, COL_LAT, decimal_para_gms(c["latitude"], "lat"), sob, log, "latitude")
        escrever(ws, linha, COL_LON, decimal_para_gms(c["longitude"], "lon"), sob, log, "longitude")
        if args.altitude and (sob or ws.cell(linha, COL_ALT).value in (None, "")):
            escrever(ws, linha, COL_ALT, buscar_altitude(c["latitude"], c["longitude"]), True, log, "altitude")

    log += lancar_dres(ws, linha, c.get("dres"), sob)

    acao = "novo ponto" if novo else "completado"
    msg = f"{aba} linha {linha} ({acao})"
    if log:
        msg += " | " + "; ".join(log)
    return msg


# --------------------------------------------------------------------------
# 4) Programa
# --------------------------------------------------------------------------
def listar_fotos(entradas) -> list:
    fotos = []
    for e in entradas:
        p = Path(e)
        if p.is_dir():
            fotos += sorted(f for f in p.iterdir() if f.suffix.lower() in EXTENSOES)
        elif p.is_file():
            fotos.append(p)
        else:
            print(f"! não encontrado: {e}")
    return fotos


def main():
    ap = argparse.ArgumentParser(description="Preenche a planilha DRES a partir dos prints do Relatório de coletas.")
    ap.add_argument("--planilha", required=True, help="planilha DRES (.xlsx)")
    ap.add_argument("--fotos", nargs="+", default=[], help="prints e/ou pastas com prints")
    ap.add_argument("--saida", help="arquivo de saída (padrão: <planilha>_preenchida.xlsx)")
    ap.add_argument("--altitude", action="store_true", help="preenche a altitude pelo relevo SRTM (internet)")
    ap.add_argument("--sobrescrever", action="store_true", help="substitui valores já existentes")
    ap.add_argument("--obs-completa", action="store_true", help="copia também a observação do app para a coluna OBS")
    ap.add_argument("--sem-dres", action="store_true", help="não avalia o DRES pelas fotos (só transcreve os dados)")
    ap.add_argument("--so-calcular", action="store_true",
                    help="não lê prints: só calcula IQEA e Leitura dos pontos que já têm E e Q na planilha")
    ap.add_argument("--modelo", default=MODELO_PADRAO, help=f"modelo Claude (padrão {MODELO_PADRAO})")
    args = ap.parse_args()

    entrada = Path(args.planilha)
    saida = Path(args.saida) if args.saida else entrada.with_name(entrada.stem + "_preenchida.xlsx")
    cache_arq = saida.with_suffix(".leituras.json")
    cache = json.loads(cache_arq.read_text(encoding="utf-8")) if cache_arq.exists() else {}

    if args.so_calcular:
        wb = openpyxl.load_workbook(entrada)
        for ws in wb.worksheets:
            if sem_acento(ws.title).upper().startswith("LEITURAS"):
                for r in linhas_dos_pontos(ws):
                    log = lancar_dres(ws, r, None, args.sobrescrever)
                    if log:
                        print(f"  {ws.title} linha {r} {ws.cell(r, COL_COD).value or ''}: {'; '.join(log)}")
        wb.save(saida)
        print(f"\nPronto: {saida}")
        return

    fotos = listar_fotos(args.fotos)
    if not fotos:
        sys.exit("Nenhum print encontrado (use --fotos).")

    client = None
    coletas = []
    print(f"Lendo {len(fotos)} print(s)...")
    for f in fotos:
        chave = f.name
        if chave in cache and not args.sem_dres and cache[chave].get("dres") is None:
            del cache[chave]  # lido antes sem DRES: lê de novo com a avaliação
        if chave not in cache:
            client = client or anthropic.Anthropic()
            try:
                cache[chave] = ler_print(client, f, args.modelo, not args.sem_dres)
            except anthropic.AuthenticationError:
                sys.exit("Chave da API inválida ou ausente. Rode: set ANTHROPIC_API_KEY=sk-ant-...")
            except (anthropic.APIError, RuntimeError) as e:
                print(f"  ! {f.name}: erro na leitura ({e})")
                continue
            cache_arq.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")
        c = cache[chave]
        print(f"  {f.name}: {c.get('codigo')}  {c.get('data')} {c.get('hora')}  "
              f"talhão {c.get('talhao')}  ({c.get('latitude')}, {c.get('longitude')})")
        d = c.get("dres") or {}
        if d.get("avaliavel"):
            cams = "  ".join(f"{cam['espessura_cm']:g} cm Q{cam['nota_q']}" for cam in d.get("camadas", []))
            print(f"      DRES: {cams}  | {d.get('observacao_campo')} (confiança {d.get('confianca')})")
        coletas.append((f.name, c))

    # ordena por data e hora para numerar os pontos na ordem do dia
    def ordem(item):
        c = item[1]
        try:
            d = dt.datetime.strptime(f"{c.get('data')} {c.get('hora')}", "%d/%m/%Y %H:%M")
        except (TypeError, ValueError):
            d = dt.datetime.max
        return d
    coletas.sort(key=ordem)

    wb = openpyxl.load_workbook(entrada)
    print("\nLançando na planilha...")
    for nome, c in coletas:
        print(f"  {nome} -> {lancar(wb, c, args)}")

    wb.save(saida)
    print(f"\nPronto: {saida}")
    print(f"Leituras guardadas em: {cache_arq}")


if __name__ == "__main__":
    main()
