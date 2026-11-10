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
    próximo ponto vazio. Se os pontos da aba acabarem, acrescenta uma página
    (cópia da última) e empurra para baixo o anexo, as fotos e as assinaturas.
  * Por padrão só escreve em células vazias (não apaga o que foi digitado
    no campo). Use --sobrescrever para substituir.
  * A planilha original não é alterada: o resultado vai para um arquivo novo.
    O arquivo é editado direto no XML, como o Excel faz, então fotos do anexo,
    logo do cabeçalho de impressão e configurações de página ficam intactos.
  * Também calcula IQEA e Leitura dos pontos que já têm E e Q anotados
    (use --so-coletas para não mexer nos outros pontos).
  * É a mesma lógica da Planilha DRES do Portal da equipe.
  * A leitura de cada print fica guardada em <saida>.leituras.json; rodar de
    novo não lê (nem cobra) o mesmo print outra vez. Dá para corrigir algum
    valor nesse arquivo e rodar de novo.

Requisitos (uma vez):
    pip install anthropic pydantic

Chave da API da Anthropic (uma vez por janela do Prompt de Comando):
    set ANTHROPIC_API_KEY=sk-ant-...

Uso:
    python preencher_dres.py --planilha DRES_Planilha_de_Campo_AROEIRA.xlsx --fotos C:\\prints
    python preencher_dres.py --planilha DRES.xlsx --fotos print1.png print2.png --altitude
"""

import argparse
import base64
import datetime as dt
import json
import mimetypes
import re
import sys
import unicodedata
import urllib.request
import zipfile
from pathlib import Path
from xml.dom import minidom
from typing import Optional

import anthropic
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
#
# A planilha é editada direto no XML do .xlsx (como o Excel faz), para não perder
# nada do arquivo original: fotos do anexo, logo do cabeçalho de impressão,
# configurações de página. É a mesma lógica da Planilha DRES do Portal da equipe.
# --------------------------------------------------------------------------
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
DECLARACAO = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


def _filhos(el, nome: str) -> list:
    return [n for n in el.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName == nome]


def _texto(el) -> str:
    partes = []
    for n in el.childNodes:
        if n.nodeType in (n.TEXT_NODE, n.CDATA_SECTION_NODE):
            partes.append(n.data)
        elif n.nodeType == n.ELEMENT_NODE:
            partes.append(_texto(n))
    return "".join(partes)


def _definir_texto(el, txt: str):
    while el.firstChild:
        el.removeChild(el.firstChild)
    el.appendChild(el.ownerDocument.createTextNode(txt))


def _remover_attr(el, nome: str):
    if el.hasAttribute(nome):
        el.removeAttribute(nome)


def _serializar(doc) -> bytes:
    return (DECLARACAO + doc.documentElement.toxml()).encode("utf-8")


def letra_coluna(c: int) -> str:
    s = ""
    while c > 0:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def numero_coluna(ref: str) -> int:
    n = 0
    for ch in re.match(r"[A-Z]+", ref).group(0):
        n = n * 26 + ord(ch) - 64
    return n


def deslocar_refs(txt: str, a_partir: int, n: int) -> str:
    """Empurra para baixo as referências de linha >= a_partir (ao inserir linhas na aba)."""
    return re.sub(r"(\$?[A-Z]{1,3}\$?)(\d+)",
                  lambda m: m.group(1) + str(int(m.group(2)) + n if int(m.group(2)) >= a_partir else int(m.group(2))),
                  txt)


def deslocar_formula(f: str, a_partir: int, n: int) -> str:
    """Igual, dentro de fórmula, sem tocar em nomes de função (LOG10) nem em outras abas (Aba!A1)."""
    return re.sub(r"(?<![A-Za-z0-9_.!$'])(\$?[A-Z]{1,3}\$?)(\d+)(?![0-9A-Za-z_(!])",
                  lambda m: m.group(1) + str(int(m.group(2)) + n if int(m.group(2)) >= a_partir else int(m.group(2))),
                  f)


def nome_aba(data_txt: str) -> Optional[str]:
    m = re.search(r"(\d{1,2})/(\d{1,2})", data_txt or "")
    return f"LEITURAS {int(m.group(1)):02d}{int(m.group(2)):02d}" if m else None


def hora_para_fracao(txt: Optional[str]) -> Optional[float]:
    t = texto_para_hora(txt)
    return (t.hour * 60 + t.minute) / 1440 if t else None


def fracao_para_hora(v: float) -> str:
    minutos = round((v % 1) * 1440)
    return f"{minutos // 60:02d}:{minutos % 60:02d}"


def vazio(v) -> bool:
    return v is None or v == ""


class Aba:
    def __init__(self, nome: str, caminho: str, doc, compartilhadas: list):
        self.nome, self.caminho, self.doc, self.compartilhadas = nome, caminho, doc, compartilhadas
        self.sd = doc.getElementsByTagName("sheetData")[0]

    def _linhas(self) -> list:
        return _filhos(self.sd, "row")

    def _linha(self, r: int, criar: bool):
        depois = None
        for l in self._linhas():
            n = int(l.getAttribute("r"))
            if n == r:
                return l
            if n > r:
                depois = l
                break
        if not criar:
            return None
        nova = self.doc.createElementNS(NS, "row")
        nova.setAttribute("r", str(r))
        self.sd.insertBefore(nova, depois)
        return nova

    def _celula(self, r: int, c: int, criar: bool):
        linha = self._linha(r, criar)
        if linha is None:
            return None
        depois = None
        for cel in _filhos(linha, "c"):
            cc = numero_coluna(cel.getAttribute("r"))
            if cc == c:
                return cel
            if cc > c:
                depois = cel
                break
        if not criar:
            return None
        _remover_attr(linha, "spans")  # dica de colunas usadas; o Excel recalcula
        nova = self.doc.createElementNS(NS, "c")
        nova.setAttribute("r", f"{letra_coluna(c)}{r}")
        linha.insertBefore(nova, depois)
        return nova

    def ler(self, r: int, c: int):
        el = self._celula(r, c, False)
        if el is None:
            return None
        t = el.getAttribute("t")
        vs = _filhos(el, "v")
        v = _texto(vs[0]) if vs else None
        if t == "s":
            return None if v is None else self.compartilhadas[int(v)]
        if t == "inlineStr":
            return _texto(el)
        if t in ("str", "e"):
            return v
        if v in (None, ""):
            return None
        n = float(v)
        return int(n) if n.is_integer() else n

    def escrever(self, r: int, c: int, valor):
        el = self._celula(r, c, not vazio(valor))
        if el is None:
            return
        while el.firstChild:
            el.removeChild(el.firstChild)
        _remover_attr(el, "t")
        if vazio(valor):
            return
        if isinstance(valor, (int, float)):
            v = self.doc.createElementNS(NS, "v")
            v.appendChild(self.doc.createTextNode(str(int(valor)) if float(valor).is_integer() else repr(float(valor))))
            el.appendChild(v)
        else:
            el.setAttribute("t", "inlineStr")
            is_ = self.doc.createElementNS(NS, "is")
            t = self.doc.createElementNS(NS, "t")
            t.setAttribute("xml:space", "preserve")
            t.appendChild(self.doc.createTextNode(str(valor)))
            is_.appendChild(t)
            el.appendChild(is_)

    def ultima_linha(self) -> int:
        linhas = self._linhas()
        return int(linhas[-1].getAttribute("r")) if linhas else 0

    def paginas(self) -> list:
        """Cada cabeçalho 'PONTO' com a primeira linha de cada ponto abaixo dele (de 4 em 4)."""
        fim = self.ultima_linha()
        cab = [r for r in range(1, fim + 1)
               if isinstance(self.ler(r, 1), str) and self.ler(r, 1).strip().upper() == "PONTO"]
        res = []
        for i, h in enumerate(cab):
            prox = cab[i + 1] if i + 1 < len(cab) else fim + 2
            pontos = []
            for k in range(PONTOS_POR_PAGINA):
                r = h + 1 + k * LINHAS_POR_PONTO
                if r + LINHAS_POR_PONTO - 1 >= prox:
                    break
                if any(isinstance(self.ler(x, 1), str) for x in range(r, r + LINHAS_POR_PONTO)):
                    break  # "*plantada", "Anexo I ..."
                pontos.append(r)
            res.append((h, pontos))
        return res

    def pontos(self) -> list:
        return [r for _, ps in self.paginas() for r in ps]

    @staticmethod
    def _mover_linha(row, nova: int):
        row.setAttribute("r", str(nova))
        for c in _filhos(row, "c"):
            c.setAttribute("r", re.sub(r"\d+$", str(nova), c.getAttribute("r")))

    def inserir_pagina(self):
        """Copia a última página (do fim da penúltima até o último ponto da última) logo abaixo
        dela e empurra o que vem depois. Devolve (linha onde inseriu, quantas linhas) ou None."""
        pags = [(h, ps) for h, ps in self.paginas() if ps]
        if not pags:
            return None
        cab_ult, ult = pags[-1]
        pen = pags[-2][1] if len(pags) > 1 else None
        fim = ult[-1] + LINHAS_POR_PONTO - 1
        ini = pen[-1] + LINHAS_POR_PONTO if pen else cab_ult
        n = fim - ini + 1
        ins = fim + 1

        linhas = self._linhas()
        modelo = [l for l in linhas if ini <= int(l.getAttribute("r")) <= fim]
        depois = next((l for l in linhas if int(l.getAttribute("r")) >= ins), None)

        for l in linhas:  # empurra as linhas de baixo
            r = int(l.getAttribute("r"))
            if r >= ins:
                self._mover_linha(l, r + n)
        for f in self.doc.getElementsByTagName("f"):  # corrige as fórmulas da aba
            if _texto(f):
                _definir_texto(f, deslocar_formula(_texto(f), ins, n))
            if f.hasAttribute("ref"):
                f.setAttribute("ref", deslocar_refs(f.getAttribute("ref"), ins, n))

        for l in modelo:  # cópia da página, sem fórmulas
            copia = l.cloneNode(True)
            self._mover_linha(copia, int(l.getAttribute("r")) - ini + ins)
            for c in _filhos(copia, "c"):
                if _filhos(c, "f"):
                    while c.firstChild:
                        c.removeChild(c.firstChild)
                    _remover_attr(c, "t")
            self.sd.insertBefore(copia, depois)

        mcs = self.doc.getElementsByTagName("mergeCells")
        if mcs:  # células mescladas: empurra e duplica as da página copiada
            mc = mcs[0]
            for m in _filhos(mc, "mergeCell"):
                ref = m.getAttribute("ref")
                a, _, b = ref.partition(":")
                ra, rb = int(re.search(r"\d+", a).group()), int(re.search(r"\d+", b or a).group())
                m.setAttribute("ref", deslocar_refs(ref, ins, n))
                if ra >= ini and rb <= fim:
                    novo = self.doc.createElementNS(NS, "mergeCell")
                    novo.setAttribute("ref", re.sub(r"\d+", lambda x: str(int(x.group()) - ini + ins), ref))
                    mc.appendChild(novo)
            mc.setAttribute("count", str(len(_filhos(mc, "mergeCell"))))

        for tag, attr in (("conditionalFormatting", "sqref"), ("dataValidation", "sqref"), ("hyperlink", "ref"),
                          ("ignoredError", "sqref"), ("protectedRange", "sqref"), ("dimension", "ref")):
            for el in self.doc.getElementsByTagName(tag):
                if el.hasAttribute(attr):
                    el.setAttribute(attr, deslocar_refs(el.getAttribute(attr), ins, n))

        rbs = self.doc.getElementsByTagName("rowBreaks")
        if rbs:  # quebras de página manuais
            rbk = rbs[0]
            for b in _filhos(rbk, "brk"):
                i = int(b.getAttribute("id"))
                if i >= ins:
                    b.setAttribute("id", str(i + n))
                elif ini <= i <= fim:
                    novo = b.cloneNode(True)
                    novo.setAttribute("id", str(i - ini + ins))
                    rbk.appendChild(novo)
            brks = sorted(_filhos(rbk, "brk"), key=lambda x: int(x.getAttribute("id")))
            for b in brks:
                rbk.appendChild(b)
            rbk.setAttribute("count", str(len(brks)))
            rbk.setAttribute("manualBreakCount", str(len(brks)))
        return ins, n


class Pasta:
    def __init__(self, caminho: Path):
        with zipfile.ZipFile(caminho) as zf:
            self.ordem = zf.namelist()
            self.arquivos = {n: zf.read(n) for n in self.ordem}
        self.wb = self.xml("xl/workbook.xml")
        self.rels = self.xml("xl/_rels/workbook.xml.rels")
        self.desenhos = {}

        compartilhadas = []
        ss = self.xml("xl/sharedStrings.xml")
        if ss is not None:
            compartilhadas = [_texto(si) for si in ss.getElementsByTagName("si")]

        alvos = {r.getAttribute("Id"): r.getAttribute("Target") for r in self.rels.getElementsByTagName("Relationship")}
        self.abas = []
        for sh in self.wb.getElementsByTagName("sheet"):
            rid = sh.getAttributeNS(NS_REL, "id") or sh.getAttribute("r:id")
            caminho_aba = "xl/" + re.sub(r"^/?(xl/)?", "", alvos[rid])
            self.abas.append(Aba(sh.getAttribute("name"), caminho_aba, self.xml(caminho_aba), compartilhadas))

    def xml(self, caminho: str):
        dados = self.arquivos.get(caminho)
        return minidom.parseString(dados) if dados is not None else None

    def aba(self, nome: str) -> Optional[Aba]:
        return next((a for a in self.abas if a.nome.strip().upper() == nome.upper()), None)

    def abas_leituras(self) -> list:
        return [a for a in self.abas if sem_acento(a.nome).upper().startswith("LEITURAS")]

    def _editar_tipos(self, funcao):
        ct = self.xml("[Content_Types].xml")
        funcao(ct)
        self.arquivos["[Content_Types].xml"] = _serializar(ct)

    def criar_aba(self, nome: str, data_txt: str) -> Aba:
        """Cria a aba do dia copiando a primeira aba LEITURAS, com os pontos limpos.
        Fotos e imagem do cabeçalho não são copiadas."""
        modelo = (self.abas_leituras() or [None])[0]
        if modelo is None:
            raise SystemExit("A planilha não tem nenhuma aba 'LEITURAS ...' para servir de modelo.")
        k = 1
        while f"xl/worksheets/sheet{k}.xml" in self.arquivos:
            k += 1
        caminho = f"xl/worksheets/sheet{k}.xml"

        doc = minidom.parseString(_serializar(modelo.doc))
        raiz = doc.documentElement
        _remover_attr(raiz, "xr:uid")
        for tag in ("drawing", "legacyDrawing", "legacyDrawingHF", "picture"):
            for el in list(raiz.getElementsByTagName(tag)):
                el.parentNode.removeChild(el)
        for ps in raiz.getElementsByTagName("pageSetup"):
            _remover_attr(ps, "r:id")
        for sv in raiz.getElementsByTagName("sheetView"):
            _remover_attr(sv, "tabSelected")

        aba = Aba(nome, caminho, doc, modelo.compartilhadas)
        for r in aba.pontos():
            for x in range(r, r + LINHAS_POR_PONTO):
                for c in range(COL_COD, COL_LEITURA + 1):
                    aba.escrever(x, c, None)
        titulo = aba.ler(1, 1)
        if isinstance(titulo, str):
            aba.escrever(1, 1, re.sub(r"Data:\s*\S+", f"Data: {data_txt}", titulo))

        def registrar(ct):
            ov = ct.createElementNS(ct.documentElement.namespaceURI, "Override")
            ov.setAttribute("PartName", "/" + caminho)
            ov.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
            ct.documentElement.appendChild(ov)
        self._editar_tipos(registrar)
        self.ordem.append(caminho)

        ids = {r.getAttribute("Id") for r in self.rels.getElementsByTagName("Relationship")}
        j = 1
        while f"rId{j}" in ids:
            j += 1
        rel = self.rels.createElementNS(NS_PKG_REL, "Relationship")
        rel.setAttribute("Id", f"rId{j}")
        rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
        rel.setAttribute("Target", f"worksheets/sheet{k}.xml")
        self.rels.documentElement.appendChild(rel)

        sheets = self.wb.getElementsByTagName("sheet")
        max_id = max(int(s.getAttribute("sheetId")) for s in sheets)
        ultima = None
        for s in sheets:
            if sem_acento(s.getAttribute("name")).upper().startswith("LEITURAS"):
                ultima = s
        sh = self.wb.createElementNS(NS, "sheet")
        sh.setAttribute("name", nome)
        sh.setAttribute("sheetId", str(max_id + 1))
        sh.setAttributeNS(NS_REL, "r:id", f"rId{j}")
        sheets[0].parentNode.insertBefore(sh, ultima.nextSibling if ultima else None)
        # nomes com "localSheetId" (áreas de impressão) apontam por posição
        pos = list(self.wb.getElementsByTagName("sheet")).index(sh)
        for dn in self.wb.getElementsByTagName("definedName"):
            if dn.hasAttribute("localSheetId") and int(dn.getAttribute("localSheetId")) >= pos:
                dn.setAttribute("localSheetId", str(int(dn.getAttribute("localSheetId")) + 1))

        self.abas.append(aba)
        print(f"  * aba '{nome}' criada a partir de '{modelo.nome}'")
        return aba

    def acrescentar_pagina(self, aba: Aba) -> bool:
        """Acrescenta uma página de pontos no fim da aba (cópia da última), limpa e numerada."""
        antes = set(aba.pontos())
        res = aba.inserir_pagina()
        if res is None:
            return False
        inicio, linhas = res

        # fotos e figuras ancoradas abaixo da página descem junto (linhas do desenho começam em 0)
        pasta, arquivo = aba.caminho.rsplit("/", 1)
        rels = self.xml(f"{pasta}/_rels/{arquivo}.rels")
        if rels is not None:
            for rel in rels.getElementsByTagName("Relationship"):
                if not rel.getAttribute("Type").endswith("/drawing"):
                    continue
                partes = []
                for p in (pasta + "/" + rel.getAttribute("Target")).split("/"):
                    if p == "..":
                        partes.pop()
                    elif p and p != ".":
                        partes.append(p)
                alvo = "/".join(partes)
                doc = self.desenhos.get(alvo) or self.xml(alvo)
                if doc is None:
                    continue
                for tag in ("from", "to"):
                    for el in doc.getElementsByTagNameNS(NS_XDR, tag):
                        rows = el.getElementsByTagNameNS(NS_XDR, "row")
                        if rows and int(_texto(rows[0])) + 1 >= inicio:
                            _definir_texto(rows[0], str(int(_texto(rows[0])) + linhas))
                self.desenhos[alvo] = doc

        # nomes definidos que apontam para a aba (ex.: área de impressão)
        prefixo = re.compile(rf"^'?{re.escape(aba.nome)}'?!")
        for dn in self.wb.getElementsByTagName("definedName"):
            partes = [re.sub(r"!(.*)$", lambda m: "!" + deslocar_refs(m.group(1), inicio, linhas), p)
                      if prefixo.match(p) else p for p in _texto(dn).split(",")]
            _definir_texto(dn, ",".join(partes))

        # a cadeia de cálculo guarda endereços de fórmulas que mudaram de linha: o Excel refaz ao abrir
        if "xl/calcChain.xml" in self.arquivos:
            del self.arquivos["xl/calcChain.xml"]
            for rel in list(self.rels.getElementsByTagName("Relationship")):
                if rel.getAttribute("Type").endswith("/calcChain"):
                    rel.parentNode.removeChild(rel)

            def tirar(ct):
                for ov in list(ct.getElementsByTagName("Override")):
                    if ov.getAttribute("PartName") == "/xl/calcChain.xml":
                        ov.parentNode.removeChild(ov)
            self._editar_tipos(tirar)

        # pontos da página nova: sem os dados copiados e com a numeração seguindo
        pontos = aba.pontos()
        for r in [p for p in pontos if p not in antes]:
            for x in range(r, r + LINHAS_POR_PONTO):
                for c in range(COL_COD, COL_LEITURA + 1):
                    aba.escrever(x, c, None)
            for x in range(r + 1, r + LINHAS_POR_PONTO):
                aba.escrever(x, COL_PONTO, None)
            aba.escrever(r, COL_PONTO, pontos.index(r) + 1)
        return True

    def salvar(self, saida: Path):
        for a in self.abas:
            self.arquivos[a.caminho] = _serializar(a.doc)
        for caminho, doc in self.desenhos.items():
            self.arquivos[caminho] = _serializar(doc)
        for calc in self.wb.getElementsByTagName("calcPr"):
            calc.setAttribute("fullCalcOnLoad", "1")  # o Excel recalcula as fórmulas ao abrir
        self.arquivos["xl/workbook.xml"] = _serializar(self.wb)
        self.arquivos["xl/_rels/workbook.xml.rels"] = _serializar(self.rels)
        with zipfile.ZipFile(saida, "w", zipfile.ZIP_DEFLATED) as zf:
            for nome in self.ordem:
                if nome in self.arquivos:
                    zf.writestr(nome, self.arquivos[nome])


def escrever_se_vazio(aba: Aba, r: int, c: int, valor, sobrescrever: bool, log: list, rotulo: str, fmt=str):
    if vazio(valor):
        return
    atual = aba.ler(r, c)
    if not vazio(atual) and not sobrescrever:
        if fmt(atual).strip() != fmt(valor).strip():
            log.append(f"{rotulo} mantido ({fmt(atual)}; print: {fmt(valor)})")
        return
    aba.escrever(r, c, valor)


def montar_obs(atual, c: dict, incluir_obs_app: bool) -> Optional[str]:
    partes = []
    if c.get("talhao"):
        partes.append(f"T {c['talhao']}.")
    if incluir_obs_app and c.get("observacao"):
        partes.append(c["observacao"].strip())
    dres = c.get("dres") or {}
    if vazio(atual) and dres.get("avaliavel") and dres.get("observacao_campo"):
        partes.append(dres["observacao_campo"].strip())
    prefixo = " ".join(partes)
    if not prefixo:
        return None
    if vazio(atual):
        return prefixo
    if c.get("talhao") and re.match(rf"\s*T\s*{re.escape(str(c['talhao']))}\b", str(atual)):
        return None  # já tem o talhão
    return f"{prefixo} {atual}"


def lancar_dres(aba: Aba, linha: int, dres: Optional[dict], sobrescrever: bool) -> list:
    """Escreve E e Q das camadas (uma por linha do ponto) e calcula IQEA e Leitura."""
    log = []
    linhas = range(linha, linha + LINHAS_POR_PONTO)
    ja_tem = any(not vazio(aba.ler(r, COL_E)) or not vazio(aba.ler(r, COL_Q)) for r in linhas)
    camadas = (dres or {}).get("camadas") or []
    if dres and dres.get("avaliavel") and camadas and (sobrescrever or not ja_tem):
        for r in linhas:
            aba.escrever(r, COL_E, None)
            aba.escrever(r, COL_Q, None)
        for r, cam in zip(linhas, camadas):
            aba.escrever(r, COL_E, round(float(cam["espessura_cm"])))
            aba.escrever(r, COL_Q, min(6, max(1, int(cam["nota_q"]))))
        if len(camadas) > LINHAS_POR_PONTO:
            log.append(f"só {LINHAS_POR_PONTO} de {len(camadas)} camadas cabem no ponto")
        log.append(f"DRES pela foto (confiança {dres.get('confianca')})")
    elif ja_tem and dres and dres.get("avaliavel"):
        log.append("E/Q de campo mantidos")
    elif dres and not dres.get("avaliavel"):
        log.append("fotos não permitem avaliar o DRES")

    # IQEA e Leitura a partir do que está na planilha (de campo ou da foto)
    pares = [(numero(aba.ler(r, COL_E)), numero(aba.ler(r, COL_Q))) for r in linhas]
    pares = [(e, q) for e, q in pares if e and q]
    if pares:
        iqea = round(sum(e * q for e, q in pares) / sum(e for e, _ in pares) + 1e-9, 1)
        if sobrescrever or vazio(aba.ler(linha, COL_IQEA)):
            aba.escrever(linha, COL_IQEA, iqea)
        if sobrescrever or vazio(aba.ler(linha, COL_LEITURA)):
            aba.escrever(linha, COL_LEITURA, classe_iqea(iqea))
        log.append(f"IQEA {f'{iqea:g}'.replace('.', ',')} {classe_iqea(iqea)}")
    return log


def lancar(pasta: Pasta, c: dict, args) -> str:
    if not c.get("codigo"):
        return "sem código no print — ignorado"
    nome = nome_aba(c.get("data"))
    if not nome:
        return "sem data no print — ignorado"
    aba = pasta.aba(nome)
    extra = ""
    if aba is None:
        aba = pasta.criar_aba(nome, c["data"])
        extra = f" (aba {nome} criada)"
    pontos = aba.pontos()
    cod = normalizar_codigo(c["codigo"])

    linha = next((r for r in pontos if normalizar_codigo(aba.ler(r, COL_COD)) == cod), None)
    novo = linha is None
    if novo:
        def livre():
            return next((r for r in pontos if vazio(aba.ler(r, COL_COD)) and vazio(aba.ler(r, COL_HORA))), None)
        linha = livre()
        if linha is None:
            if not pasta.acrescentar_pagina(aba):
                return f"aba {nome} sem ponto vazio — ignorado"
            pontos = aba.pontos()
            linha = livre()
            if linha is None:
                return f"aba {nome} sem ponto vazio — ignorado"
            extra += " (página nova acrescentada)"

    log = []
    sob = args.sobrescrever
    if vazio(aba.ler(linha, COL_PONTO)):
        aba.escrever(linha, COL_PONTO, pontos.index(linha) + 1)
    if novo:
        aba.escrever(linha, COL_COD, c["codigo"].strip())
    obs = montar_obs(aba.ler(linha, COL_OBS), c, args.obs_completa)
    if obs:
        aba.escrever(linha, COL_OBS, obs)
    fmt_hora = lambda v: fracao_para_hora(v) if isinstance(v, (int, float)) else str(v)
    escrever_se_vazio(aba, linha, COL_HORA, hora_para_fracao(c.get("hora")), sob, log, "hora", fmt_hora)
    if c.get("latitude") is not None and c.get("longitude") is not None:
        escrever_se_vazio(aba, linha, COL_LAT, decimal_para_gms(c["latitude"], "lat"), sob, log, "latitude")
        escrever_se_vazio(aba, linha, COL_LON, decimal_para_gms(c["longitude"], "lon"), sob, log, "longitude")
        if args.altitude and (sob or vazio(aba.ler(linha, COL_ALT))):
            escrever_se_vazio(aba, linha, COL_ALT, buscar_altitude(c["latitude"], c["longitude"]), True, log, "altitude")

    log += lancar_dres(aba, linha, c.get("dres"), sob)
    msg = f"{nome} linha {linha} ({'novo ponto' if novo else 'completado'}){extra}"
    if log:
        msg += " — " + "; ".join(log)
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


def calcular_todos(pasta: Pasta, tocados: set, sobrescrever: bool):
    """Calcula IQEA e Leitura de todos os pontos que já têm E e Q (exceto os já lançados agora)."""
    for aba in pasta.abas_leituras():
        for r in aba.pontos():
            if (aba.nome, r) in tocados:
                continue
            log = lancar_dres(aba, r, None, sobrescrever)
            if log:
                print(f"  {aba.nome} linha {r} {aba.ler(r, COL_COD) or ''}: {'; '.join(log)}")


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
    ap.add_argument("--so-coletas", action="store_true",
                    help="não calcula IQEA e Leitura dos outros pontos que já têm E e Q na planilha")
    ap.add_argument("--modelo", default=MODELO_PADRAO, help=f"modelo Claude (padrão {MODELO_PADRAO})")
    args = ap.parse_args()

    entrada = Path(args.planilha)
    saida = Path(args.saida) if args.saida else entrada.with_name(entrada.stem + "_preenchida.xlsx")
    cache_arq = saida.with_suffix(".leituras.json")
    cache = json.loads(cache_arq.read_text(encoding="utf-8")) if cache_arq.exists() else {}

    if args.so_calcular:
        pasta = Pasta(entrada)
        calcular_todos(pasta, set(), args.sobrescrever)
        pasta.salvar(saida)
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

    pasta = Pasta(entrada)
    print("\nLançando na planilha...")
    tocados = set()
    for nome, c in coletas:
        msg = lancar(pasta, c, args)
        print(f"  {nome} -> {msg}")
        m = re.match(r"(LEITURAS \d{4}) linha (\d+)", msg)
        if m:
            tocados.add((m.group(1), int(m.group(2))))
    if not args.so_coletas:
        print("\nIQEA e Leitura dos pontos já anotados:")
        calcular_todos(pasta, tocados, args.sobrescrever)

    pasta.salvar(saida)
    print(f"\nPronto: {saida}")
    print(f"Leituras guardadas em: {cache_arq}")


if __name__ == "__main__":
    main()
