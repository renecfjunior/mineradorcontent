# -*- coding: utf-8 -*-
"""
ELEMENTAR — Minerador de Pautas
Bot que coleta notícias do G1, Nexo e Agência IBGE, pontua cada uma pela
Matriz de Validação do dossiê do canal e sugere um ângulo de vídeo.

Rode com:  python app.py
Depois abra:  http://127.0.0.1:5000
"""

import re
import io
import html
from datetime import datetime

import requests
import feedparser
from flask import Flask, render_template, jsonify, request, send_file

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__, template_folder=".")

# ---------------------------------------------------------------------------
# 1. FONTES (cada fonte tem candidatos de feed; usa o primeiro que responder)
# ---------------------------------------------------------------------------
FONTES = {
    "G1": [
        "https://g1.globo.com/rss/g1/economia/",
        "https://g1.globo.com/rss/g1/politica/",
        "https://g1.globo.com/rss/g1/brasil/",
        "http://g1.globo.com/dynamo/economia/rss2.xml",
    ],
    "Nexo": [
        "https://www.nexojornal.com.br/feed/",
        "https://www.nexojornal.com.br/rss",
        "https://www.nexojornal.com.br/feed.rss",
    ],
    "IBGE": [
        "https://agenciadenoticias.ibge.gov.br/agencia-noticias/2012-agencia-de-noticias/noticias?format=feed&type=rss",
        "https://agenciadenoticias.ibge.gov.br/component/obrss/rss-agencia",
    ],
}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/124.0 Safari/537.36"
}

# ---------------------------------------------------------------------------
# 2. DICIONÁRIOS DE PONTUAÇÃO (extraídos do dossiê do canal)
# ---------------------------------------------------------------------------

# Categorias temáticas + palavras-chave (dossiê, seção B)
CATEGORIAS = {
    "Monopólios e Poder Econômico": [
        "cade", "antitruste", "monopólio", "monopolio", "cartel", "fusão", "fusao",
        "aquisição", "aquisicao", "fraude fiscal", "abuso de poder", "concentração",
        "concentracao", "domina", "domínio", "dominio", "oligopólio", "oligopolio",
        "combinação de preço", "conluio", "bilionária", "bilionaria",
    ],
    "Infraestrutura e Engenharia": [
        "tcu", "auditoria", "obra", "obras", "atraso", "pac", "abnt", "incc",
        "infraestrutura", "logística", "logistica", "licitação", "licitacao",
        "concessão", "concessao", "metro quadrado", "custo da obra", "ferrovia",
        "rodovia", "saneamento", "porto", "construção", "construcao",
    ],
    "Indústria do Consumo e Margens": [
        "margem de lucro", "custo de produção", "custo de producao", "falsificação",
        "falsificacao", "pirataria", "obsolescência", "obsolescencia", "apreensão",
        "apreensao", "contrabando", "markup", "marca", "preço final", "preco final",
        "lucro abusivo", "cosméticos", "cosmeticos", "eletrônicos", "eletronicos",
    ],
    "Declínio Urbano e Imobiliário": [
        "imóveis vagos", "imoveis vagos", "censo", "ibge", "fechamento de comércio",
        "fechamento de comercio", "crise habitacional", "êxodo", "exodo",
        "gentrificação", "gentrificacao", "enchente", "seca", "vacância", "vacancia",
        "aluguel", "habitação", "habitacao", "esvaziamento", "centro", "imobiliário",
        "imobiliario", "cracolândia", "cracolandia",
    ],
}

# Critério 1 — Poder de Gancho (palavras que geram indignação/curiosidade)
GANCHO = [
    "domina", "domínio", "dominio", "monopólio", "monopolio", "escândalo", "escandalo",
    "esquema", "secreto", "oculto", "bilionária", "bilionaria", "ninguém", "ninguem",
    "por que", "porque", "recorde", "maior", "verdade", "bastidores", "cartel",
    "fraude", "golpe", "colapso", "explode", "dispara", "histórico", "historico",
    "sem precedente", "inédito", "inedito", "polêmica", "polemica",
    "revela", "revelado", "expõe", "expoe", "exposto", "atraso", "salto",
    "disparada", "abusiva", "abusivo", "esconde", "por dentro",
]

# Critério 3 — Rastreabilidade da Prova (documentos/autoridade)
PROVA = [
    "operação", "operacao", "cade", "tcu", "anatel", "anvisa", "processo",
    "multa", "auditoria", "relatório", "relatorio", "investigação", "investigacao",
    "ministério público", "ministerio publico", "mpf", "justiça", "justica",
    "stf", "decisão", "decisao", "sentença", "sentenca", "denúncia", "denuncia",
    "polícia federal", "policia federal", "pf", "receita federal", "inquérito",
    "inquerito", "censo", "pesquisa", "levantamento", "dados oficiais",
]

# Critério 4 — Apelo de Massa (afeta o bolso/vida do cidadão)
MASSA = [
    "preço", "preco", "conta", "bolso", "consumidor", "aluguel", "imóvel", "imovel",
    "salário", "salario", "emprego", "desemprego", "cidade", "comida", "alimento",
    "energia", "combustível", "combustivel", "gasolina", "luz", "água", "agua",
    "inflação", "inflacao", "mercado", "tarifa", "imposto", "renda", "família",
    "familia", "trabalhador", "população", "populacao",
]

# Regex p/ Critério 2 — Densidade Estatística (números, %, valores, bi/mi)
RE_PERCENT = re.compile(r"\d+[\.,]?\d*\s?%")
RE_DINHEIRO = re.compile(r"r\$\s?\d", re.IGNORECASE)
RE_BIMI = re.compile(r"\d+[\.,]?\d*\s?(bilh|milh|trilh|mil\b)", re.IGNORECASE)
RE_NUM = re.compile(r"\d{2,}")


# ---------------------------------------------------------------------------
# 3. MOTOR DE PONTUAÇÃO
# ---------------------------------------------------------------------------
def _conta_termos(texto, termos):
    """Conta quantos termos distintos da lista aparecem no texto."""
    t = texto.lower()
    return sum(1 for termo in termos if termo in t)


def _nota_0a3(qtd):
    """Converte uma contagem em nota de 0 a 3."""
    if qtd <= 0:
        return 0
    if qtd == 1:
        return 1
    if qtd == 2:
        return 2
    return 3


def detectar_categoria(texto):
    """Retorna a categoria temática dominante (e a contagem)."""
    t = texto.lower()
    melhor, melhor_qtd = "Geral", 0
    for cat, termos in CATEGORIAS.items():
        q = sum(1 for termo in termos if termo in t)
        if q > melhor_qtd:
            melhor, melhor_qtd = cat, q
    return melhor, melhor_qtd


def pontuar(titulo, resumo):
    """Aplica a Matriz de Validação. Retorna dict com notas e total."""
    titulo = titulo or ""
    resumo = resumo or ""
    full = f"{titulo} {resumo}"

    # 1. Gancho — pesa mais o que está no TÍTULO
    g = _conta_termos(titulo, GANCHO) * 2 + _conta_termos(resumo, GANCHO)
    nota_gancho = _nota_0a3(g)

    # 2. Densidade estatística
    dens = 0
    dens += len(RE_PERCENT.findall(full))
    dens += len(RE_DINHEIRO.findall(full))
    dens += len(RE_BIMI.findall(full))
    if RE_NUM.search(full):
        dens += 1
    nota_dens = _nota_0a3(dens)

    # 3. Rastreabilidade da prova
    nota_prova = _nota_0a3(_conta_termos(full, PROVA))

    # 4. Apelo de massa
    nota_massa = _nota_0a3(_conta_termos(full, MASSA))

    total = nota_gancho + nota_dens + nota_prova + nota_massa
    categoria, _ = detectar_categoria(full)

    return {
        "gancho": nota_gancho,
        "densidade": nota_dens,
        "prova": nota_prova,
        "massa": nota_massa,
        "total": total,
        "aprovada": total >= 9,       # regra do dossiê: soma >= 9
        "categoria": categoria,
    }


# ---------------------------------------------------------------------------
# 4. GERADOR DE ÂNGULO DE VÍDEO (estilo Elementar)
# ---------------------------------------------------------------------------
ANGULOS = {
    "Monopólios e Poder Econômico":
        "Gancho de contraste: mostre o mercado quando havia concorrência vs. hoje "
        "dominado por poucos. Tese: revele a fusão/manobra que criou o domínio e "
        "como ela trava preços e barra novos entrantes.",
    "Infraestrutura e Engenharia":
        "Gancho de contraste: compare o tempo/custo da obra no Brasil vs. no "
        "exterior. Tese: exponha o gargalo burocrático ou de licitação por trás do "
        "atraso, usando o relatório/auditoria como prova.",
    "Indústria do Consumo e Margens":
        "Gancho de contraste: o custo real de produção vs. o preço de prateleira. "
        "Tese: desmonte a margem inflada e o mecanismo (marca, escassez fabricada, "
        "obsolescência) que justifica o sobrepreço.",
    "Declínio Urbano e Imobiliário":
        "Gancho de contraste: a região cheia no passado vs. vazia hoje (dados do "
        "censo). Tese: explique a dinâmica econômica (êxodo, vacância, crise) e o "
        "efeito cascata no comércio e no valor dos imóveis.",
    "Geral":
        "Gancho de contraste: parta de um número chocante da matéria. Tese: rastreie "
        "o sistema oculto (lei, decisão, esquema) que produz o fenômeno e feche com "
        "uma reflexão sobre o futuro daquele mercado.",
}


def gerar_angulo(score, titulo):
    base = ANGULOS.get(score["categoria"], ANGULOS["Geral"])
    reforco = []
    if score["prova"] >= 2:
        reforco.append("Há documento/operação citável — ancore a autoridade do vídeo nele.")
    if score["densidade"] >= 2:
        reforco.append("Os números já dão o gancho de abertura; comece por eles na tela.")
    if score["massa"] >= 2:
        reforco.append("Tema mexe no bolso do público — explore o impacto pessoal no fechamento.")
    if score["gancho"] == 0:
        reforco.append("Título é morno: procure o contraste/indignação antes de pautar.")
    extra = (" " + " ".join(reforco)) if reforco else ""
    return base + extra


# ---------------------------------------------------------------------------
# 5. COLETA
# ---------------------------------------------------------------------------
def limpar(texto):
    if not texto:
        return ""
    texto = re.sub(r"<[^>]+>", " ", texto)      # tira HTML
    texto = html.unescape(texto)
    return re.sub(r"\s+", " ", texto).strip()


def buscar_fonte(nome, urls):
    """Tenta cada candidato de URL até um responder com itens."""
    for url in urls:
        try:
            r = requests.get(url, headers=HEADERS, timeout=20)
            if r.status_code != 200:
                continue
            d = feedparser.parse(r.content)
            if d.entries:
                return d.entries, url, None
        except Exception as e:
            ultimo_erro = str(e)
            continue
    return [], None, "nenhum feed respondeu"


def coletar(fontes_escolhidas, limite_por_fonte=25):
    itens = []
    status = {}
    for nome in fontes_escolhidas:
        urls = FONTES.get(nome, [])
        entries, url_ok, erro = buscar_fonte(nome, urls)
        status[nome] = {
            "ok": bool(entries),
            "url": url_ok,
            "qtd": len(entries),
            "erro": erro,
        }
        for e in entries[:limite_por_fonte]:
            titulo = limpar(getattr(e, "title", ""))
            resumo = limpar(getattr(e, "summary", "") or getattr(e, "description", ""))
            link = getattr(e, "link", "")
            data = getattr(e, "published", "") or getattr(e, "updated", "")
            if not titulo:
                continue
            score = pontuar(titulo, resumo)
            itens.append({
                "fonte": nome,
                "titulo": titulo,
                "resumo": resumo[:400],
                "link": link,
                "data": data,
                "score": score,
                "angulo": gerar_angulo(score, titulo),
            })
    # ordena pela maior nota total
    itens.sort(key=lambda x: x["score"]["total"], reverse=True)
    return itens, status


# ---------------------------------------------------------------------------
# 6. ROTAS
# ---------------------------------------------------------------------------
@app.route("/")
def index():
    return render_template("index.html", fontes=list(FONTES.keys()))


@app.route("/api/buscar", methods=["POST"])
def api_buscar():
    dados = request.get_json(force=True)
    fontes = dados.get("fontes") or list(FONTES.keys())
    nota_min = int(dados.get("nota_min", 0))
    so_aprovadas = bool(dados.get("so_aprovadas", False))

    itens, status = coletar(fontes)

    filtrados = []
    for it in itens:
        if so_aprovadas and not it["score"]["aprovada"]:
            continue
        if it["score"]["total"] < nota_min:
            continue
        filtrados.append(it)

    return jsonify({"itens": filtrados, "status": status, "total": len(filtrados)})


@app.route("/api/exportar", methods=["POST"])
def api_exportar():
    dados = request.get_json(force=True)
    itens = dados.get("itens", [])

    doc = Document()

    # Título
    t = doc.add_heading("ELEMENTAR — Pautas Mineradas", level=0)
    sub = doc.add_paragraph()
    run = sub.add_run("Relatório gerado em " +
                      datetime.now().strftime("%d/%m/%Y às %H:%M"))
    run.italic = True
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
    doc.add_paragraph("Pautas ordenadas pela Matriz de Validação do dossiê "
                      "(gancho · densidade · prova · massa). Soma ≥ 9 = recomendada.")

    for i, it in enumerate(itens, 1):
        s = it["score"]
        doc.add_heading(f"{i}. {it['titulo']}", level=1)

        meta = doc.add_paragraph()
        mrun = meta.add_run(f"Fonte: {it['fonte']}   |   Categoria: {s['categoria']}"
                            f"   |   {it.get('data','')}")
        mrun.font.size = Pt(9)
        mrun.font.color.rgb = RGBColor(0x70, 0x70, 0x70)

        # Notas
        nota = doc.add_paragraph()
        nrun = nota.add_run(
            f"NOTA TOTAL: {s['total']}/12  "
            f"({'RECOMENDADA' if s['aprovada'] else 'abaixo do corte'})"
        )
        nrun.bold = True
        nrun.font.color.rgb = (RGBColor(0x1B, 0x7A, 0x3C) if s['aprovada']
                               else RGBColor(0xA0, 0x40, 0x00))
        doc.add_paragraph(
            f"Gancho {s['gancho']}/3   ·   Densidade estatística {s['densidade']}/3"
            f"   ·   Rastreabilidade {s['prova']}/3   ·   Apelo de massa {s['massa']}/3"
        )

        if it.get("resumo"):
            doc.add_paragraph(it["resumo"])

        ang = doc.add_paragraph()
        ang.add_run("Ângulo de vídeo sugerido: ").bold = True
        ang.add_run(it.get("angulo", ""))

        if it.get("link"):
            lk = doc.add_paragraph()
            lrun = lk.add_run(it["link"])
            lrun.font.size = Pt(9)
            lrun.font.color.rgb = RGBColor(0x1A, 0x5C, 0xA8)

        doc.add_paragraph("—" * 30)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    nome = "pautas_elementar_" + datetime.now().strftime("%Y%m%d_%H%M") + ".docx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=nome,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    print("\n  ELEMENTAR rodando em  http://127.0.0.1:5000\n")
    app.run(debug=True, port=5000)
