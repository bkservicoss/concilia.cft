# -*- coding: utf-8 -*-
"""
Consolida.CFT — conciliação bancária assistida.

Cruza o extrato bancário (OFX, CSV, XLS ou XLSX) com o controle interno,
sugere pares com nota de confiança, mostra as divergências dos dois lados
e gera um relatório em Excel com trilha de auditoria.

Projeto privado · BK Serviços, Produção & Comunicação
"""
import base64
import csv
import html
import io
import numbers
import os
import re
import subprocess
import sys
import unicodedata
from bisect import bisect_left, bisect_right
from datetime import date, datetime, timedelta
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path

import pandas as pd
import streamlit as st

try:
    from zoneinfo import ZoneInfo

    FUSO = ZoneInfo("America/Sao_Paulo")
except Exception:  # sem base de fusos: usa o horário do servidor
    FUSO = None

NOME_APP = "Consolida.CFT"
SUBTITULO = "Conciliação bancária assistida"
PASTA = Path(__file__).resolve().parent
LOGO = PASTA / "assets" / "logo.svg"
ICONE = PASTA / "assets" / "icone.png"
TIPOS_ACEITOS = ["ofx", "csv", "txt", "xlsx", "xls"]
POR_PAGINA = 30
COLUNAS_PADRAO = ["data", "descricao", "centavos", "local"]

# Paleta da identidade visual (ver README · Identidade visual)
COR = {
    "marinho": "#0B2545",
    "azul": "#13315C",
    "verde": "#12A37A",
    "ambar": "#F2B233",
    "vermelho": "#D64545",
    "cinza": "#5B6B7F",
    "borda": "#DCE3EC",
    "fundo": "#F5F7FA",
}


def agora():
    return datetime.now(FUSO) if FUSO else datetime.now()


# ==========================================================
# 1. Normalização de textos, valores e datas
# ==========================================================
def vazio(v):
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    return isinstance(v, str) and not v.strip()


def normalizar(texto):
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return texto.lower().strip()


RE_NUMERO = re.compile(r"^\d+(?:[.,]\d+)*$")
RE_SUFIXO_DC = re.compile(r"^(.*?[\d)\-])(DEBITO|DEB|DB|D|CREDITO|CRED|CR|C)\.?$")


def _padronizar_decimal(s):
    """Decide qual é o separador decimal (padrão brasileiro em caso de dúvida)."""
    tem_virgula, tem_ponto = "," in s, "." in s
    if tem_virgula and tem_ponto:
        decimal = "," if s.rfind(",") > s.rfind(".") else "."
        milhar = "." if decimal == "," else ","
        return s.replace(milhar, "").replace(decimal, ".")
    if tem_virgula:
        return s.replace(",", ".") if s.count(",") == 1 else s.replace(",", "")
    if tem_ponto:
        partes = s.split(".")
        if len(partes) == 2 and len(partes[1]) != 3:
            return s
        return s.replace(".", "")
    return s


def valor_em_centavos(x):
    """Converte um valor monetário em centavos (int). Devolve None se não for valor.

    Aceita: 1500 | 1.500,00 | 1,500.00 | R$ -150,00 | -R$ 150 | (150,00) |
    150,00- | 1.500,00 D | 1.500,00 C | -150.00 (OFX).
    """
    if vazio(x) or isinstance(x, bool):
        return None
    if isinstance(x, numbers.Number):
        try:
            return int((Decimal(str(x)) * 100).quantize(Decimal("1"), ROUND_HALF_UP))
        except InvalidOperation:
            return None
    s = normalizar(x).upper().replace("R$", "").replace("$", "")
    s = re.sub(r"\s+", "", s)
    negativo = False
    m = RE_SUFIXO_DC.match(s)
    if m:
        s, negativo = m.group(1), m.group(2).startswith("D")
    if s.startswith("(") and s.endswith(")"):
        s, negativo = s[1:-1], True
    if s.startswith("-"):
        s, negativo = s[1:], True
    elif s.endswith("-"):
        s, negativo = s[:-1], True
    elif s.startswith("+"):
        s = s[1:]
    if not RE_NUMERO.match(s):
        return None
    try:
        valor = Decimal(_padronizar_decimal(s))
    except InvalidOperation:
        return None
    centavos = int((valor * 100).quantize(Decimal("1"), ROUND_HALF_UP))
    return -centavos if negativo else centavos


def para_data(x):
    """Converte célula em date. Aceita datas do Excel, dd/mm/aaaa, d/m/aa, aaaa-mm-dd e OFX."""
    if vazio(x):
        return None
    if isinstance(x, pd.Timestamp):
        return x.date()
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    if isinstance(x, numbers.Number) and not isinstance(x, bool):
        if 20000 <= float(x) <= 80000:  # número de série do Excel
            return (datetime(1899, 12, 30) + timedelta(days=int(x))).date()
        if not float(x).is_integer():
            return None
        s = str(int(x))
    else:
        s = str(x).strip()
    m = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        ano, mes, dia = map(int, m.groups())
    else:
        m = re.search(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})", s)
        if m:
            dia, mes, ano = map(int, m.groups())
            if ano < 100:
                ano += 2000 if ano < 70 else 1900
        else:
            m = re.match(r"^(\d{4})(\d{2})(\d{2})", s)  # OFX: 20260102120000[-3:BRT]
            if not m:
                return None
            ano, mes, dia = map(int, m.groups())
    try:
        return date(ano, mes, dia)
    except ValueError:
        return None


def eh_data_forte(v):
    """Usado só para adivinhar a coluna de data (não confunde valores com datas)."""
    if isinstance(v, (pd.Timestamp, datetime, date)):
        return True
    return isinstance(v, str) and bool(
        re.search(r"\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}|\d{4}-\d{1,2}-\d{1,2}", v)
    )


def moeda(centavos, sinal=False):
    texto = f"{abs(centavos) / 100:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    prefixo = "-" if centavos < 0 else ("+" if sinal and centavos > 0 else "")
    return f"{prefixo}R$ {texto}"


def data_br(d):
    return d.strftime("%d/%m/%Y") if d else "Sem data"


PALAVRAS_VAZIAS = {
    "pix", "ted", "doc", "transf", "transferencia", "pagamento", "pagto", "pgto",
    "recebido", "recebimento", "enviado", "credito", "debito", "deposito", "boleto",
    "das", "dos", "para", "com", "ltda", "eireli", "sas",
}


def tokens(texto):
    return frozenset(
        p for p in re.findall(r"[a-z0-9]+", normalizar(texto))
        if p not in PALAVRAS_VAZIAS and ((p.isdigit() and len(p) >= 4) or (not p.isdigit() and len(p) >= 3))
    )


def similaridade(a, b):
    if not a or not b:
        return 0.0
    return len(a & b) / min(len(a), len(b))


# ==========================================================
# 2. Leitura dos arquivos
# ==========================================================
def decodificar(conteudo):
    for codificacao in ("utf-8-sig", "cp1252"):
        try:
            return conteudo.decode(codificacao)
        except UnicodeDecodeError:
            continue
    return conteudo.decode("latin-1")


def detectar_separador(texto):
    linhas = [l for l in texto.splitlines()[:60] if l.strip()]
    melhor, melhor_nota = ";", -1
    for sep in (";", "\t", "|", ","):
        contagens = [l.count(sep) for l in linhas]
        positivas = [c for c in contagens if c > 0]
        if not positivas:
            continue
        moda = max(set(positivas), key=positivas.count)
        nota = sum(1 for c in contagens if c == moda)
        if nota > melhor_nota:
            melhor, melhor_nota = sep, nota
    return melhor


def ler_ofx(texto):
    blocos = re.findall(
        r"<STMTTRN>(.*?)(?=</STMTTRN>|<STMTTRN>|</BANKTRANLIST>)", texto, flags=re.S | re.I
    )
    linhas = []
    for i, bloco in enumerate(blocos, 1):
        def tag(nome):
            m = re.search(rf"<{nome}>([^<\r\n]*)", bloco, flags=re.I)
            return m.group(1).strip() if m else ""

        centavos = valor_em_centavos(tag("TRNAMT"))
        if not centavos:
            continue
        partes = []
        for p in (tag("NAME"), tag("MEMO")):
            if p and p not in partes:
                partes.append(p)
        if tag("CHECKNUM"):
            partes.append(f"Doc {tag('CHECKNUM')}")
        linhas.append({
            "data": para_data(tag("DTPOSTED")),
            "descricao": " | ".join(partes) or tag("TRNTYPE") or "Sem descrição",
            "centavos": centavos,
            "local": f"OFX · transação {i}",
        })
    return pd.DataFrame(linhas, columns=COLUNAS_PADRAO)


@st.cache_data(show_spinner=False)
def ler_arquivo(nome, conteudo):
    """Devolve ('ofx', DataFrame padronizado) ou ('tabela', {aba: DataFrame bruto})."""
    extensao = nome.rsplit(".", 1)[-1].lower() if "." in nome else ""
    if extensao == "ofx" or b"<OFX>" in conteudo[:20000].upper():
        return "ofx", ler_ofx(decodificar(conteudo))
    if extensao in ("csv", "txt"):
        texto = decodificar(conteudo)
        linhas = list(csv.reader(io.StringIO(texto), delimiter=detectar_separador(texto)))
        linhas = [l for l in linhas if any(c.strip() for c in l)]
        largura = max((len(l) for l in linhas), default=0)
        df = pd.DataFrame([l + [None] * (largura - len(l)) for l in linhas], dtype=object)
        return "tabela", {"CSV": df}
    abas = pd.read_excel(io.BytesIO(conteudo), sheet_name=None, header=None)
    return "tabela", {n: d.reset_index(drop=True) for n, d in abas.items() if not d.dropna(how="all").empty}


# ==========================================================
# 3. Detecção de cabeçalho e mapeamento de colunas
# ==========================================================
GRUPOS = {
    "saldo": ("saldo", "balance"),
    "data": ("data", "dt.", "dt ", "date", "dia"),
    "credito": ("credito", "entrada", "receita", "credit"),
    "debito": ("debito", "saida", "despesa", "debit"),
    "valor": ("valor", "vlr", "vl.", "montante", "quantia", "amount", "importe"),
    "dc": ("d/c", "c/d", "tipo", "natureza", "sinal"),
    "desc": (
        "historico", "descricao", "lancamento", "complemento", "detalhe", "memo", "nome",
        "favorecido", "pagador", "contribuinte", "profissional", "observ", "beneficiario",
        "cliente", "fornecedor", "documento",
    ),
}


def grupo_do_titulo(titulo):
    t = normalizar(titulo)
    if not t or len(t) > 40 or valor_em_centavos(titulo) is not None:
        return None
    for grupo in ("saldo", "data", "credito", "debito", "valor", "dc", "desc"):
        chaves = GRUPOS[grupo]
        if grupo == "data":
            if t == "dt" or any(t.startswith(k) for k in chaves):
                return grupo
        elif any(k in t for k in chaves):
            return grupo
    return None


def detectar_cabecalho(df):
    melhor, melhor_nota = None, 1
    for i in range(min(30, len(df))):
        grupos = {grupo_do_titulo(v) for v in df.iloc[i] if isinstance(v, str)} - {None}
        if not (grupos & {"data", "valor", "credito", "debito"}):
            continue
        if len(grupos) > melhor_nota:
            melhor, melhor_nota = i, len(grupos)
    return melhor


def letra(i):
    s, i = "", i + 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def rotulos_colunas(df, cab):
    inicio = cab + 1 if cab is not None else 0
    rotulos = {}
    for c in range(df.shape[1]):
        titulo = "" if cab is None or vazio(df.iat[cab, c]) else str(df.iat[cab, c]).strip()
        exemplo = next((str(v).strip() for v in df.iloc[inicio:inicio + 20, c] if not vazio(v)), "")
        if titulo:
            rotulos[c] = f"{letra(c)} · {titulo}"
        elif exemplo:
            rotulos[c] = f"{letra(c)} · ex.: {exemplo[:28]}"
        else:
            rotulos[c] = letra(c)
    return rotulos


def sugerir_mapeamento(df, cab):
    inicio = cab + 1 if cab is not None else 0
    amostra = df.iloc[inicio:inicio + 300]
    n = max(len(amostra), 1)
    colunas = list(range(df.shape[1]))
    grupos = {
        c: (grupo_do_titulo(df.iat[cab, c]) if cab is not None and isinstance(df.iat[cab, c], str) else None)
        for c in colunas
    }
    por_grupo = lambda g: [c for c in colunas if grupos[c] == g]

    def taxa(c, funcao):
        return sum(1 for v in amostra.iloc[:, c] if not vazio(v) and funcao(v)) / n

    col_data = (por_grupo("data") or [None])[0]
    if col_data is None:
        notas = {c: taxa(c, eh_data_forte) for c in colunas}
        if notas and max(notas.values()) >= 0.3:
            col_data = max(notas, key=notas.get)

    proibidas = set(por_grupo("saldo") + por_grupo("dc")) | {col_data}
    modo, col_valor = "unico", (por_grupo("valor") or [None])[0]
    col_cred, col_deb = (por_grupo("credito") or [None])[0], (por_grupo("debito") or [None])[0]
    if col_valor is None and col_cred is not None and col_deb is not None:
        modo = "separadas"
    elif col_valor is None:
        def nota_valor(c):
            vals = [valor_em_centavos(v) for v in amostra.iloc[:, c] if not vazio(v)]
            vals = [v for v in vals if v is not None]
            return len(vals) / n + 0.5 * (sum(1 for v in vals if v % 100) / n)

        candidatas = {c: nota_valor(c) for c in colunas if c not in proibidas}
        if candidatas and max(candidatas.values()) > 0.2:
            col_valor = max(candidatas, key=candidatas.get)

    col_desc = por_grupo("desc")[:2]
    if not col_desc:
        usadas = {col_data, col_valor, col_cred, col_deb} | proibidas
        textos = {}
        for c in colunas:
            if c in usadas:
                continue
            vals = [str(v) for v in amostra.iloc[:, c] if isinstance(v, str) and valor_em_centavos(v) is None]
            if vals:
                textos[c] = sum(len(v) for v in vals) / n
        if textos:
            col_desc = [max(textos, key=textos.get)]

    col_dc = None
    for c in por_grupo("dc"):
        marcas = [normalizar(v) for v in amostra.iloc[:, c] if not vazio(v)]
        if marcas and sum(m[:1] in ("c", "d", "e", "s") and len(m) <= 8 for m in marcas) / len(marcas) > 0.6:
            col_dc = c
            break

    return {"modo": modo, "data": col_data, "valor": col_valor, "credito": col_cred,
            "debito": col_deb, "desc": col_desc, "dc": col_dc}


def aplicar_mapeamento(df, cab, mapa, nome_aba):
    inicio = cab + 1 if cab is not None else 0
    matriz = df.to_numpy(dtype=object)
    registros, descartadas = [], 0
    for i in range(inicio, len(matriz)):
        linha = matriz[i]
        if all(vazio(v) for v in linha):
            continue
        desc = " | ".join(
            re.sub(r"\s+", " ", str(linha[c])).strip() for c in mapa["desc"] if not vazio(linha[c])
        )
        if mapa["modo"] == "unico":
            v = valor_em_centavos(linha[mapa["valor"]]) if mapa["valor"] is not None else None
            if v is not None and mapa.get("dc") is not None and not vazio(linha[mapa["dc"]]):
                indicador = normalizar(linha[mapa["dc"]])
                if indicador.startswith(("d", "s")):
                    v = -abs(v)
                elif indicador.startswith(("c", "e")):
                    v = abs(v)
        else:
            c = valor_em_centavos(linha[mapa["credito"]]) if mapa["credito"] is not None else None
            d = valor_em_centavos(linha[mapa["debito"]]) if mapa["debito"] is not None else None
            v = None if c is None and d is None else abs(c or 0) - abs(d or 0)
        if not v or re.sub(r"\s", "", normalizar(desc)).startswith(("saldo", "total")):
            descartadas += 1
            continue
        if mapa.get("inverter"):
            v = -v
        registros.append({
            "data": para_data(linha[mapa["data"]]) if mapa["data"] is not None else None,
            "descricao": desc or "Sem descrição",
            "centavos": v,
            "local": f"{nome_aba} · linha {i + 1}",
        })
    return pd.DataFrame(registros, columns=COLUNAS_PADRAO), descartadas


# ==========================================================
# 4. Motor de pareamento
# ==========================================================
ORDEM_CONF = {"Alta": 0, "Média": 1, "Baixa": 2, "Manual": 3}


def confianca(dias, sim, ambiguo):
    if dias is None:
        nivel = "Baixa"
    elif dias == 0 or (dias <= 2 and sim >= 0.5):
        nivel = "Alta"
    elif dias <= 3 or sim >= 0.5:
        nivel = "Média"
    else:
        nivel = "Baixa"
    if ambiguo and nivel == "Alta":
        nivel = "Média"
    return nivel


def gerar_sugestoes(banco, sistema, tolerancia, ignorar_sinal):
    """Pareamento em camadas: mesmo valor → data dentro da tolerância → descrição parecida.

    Todas as combinações possíveis recebem uma nota e os pares são escolhidos
    da maior nota para a menor, sem repetir lançamentos.
    """
    chave = abs if ignorar_sinal else (lambda c: c)
    baldes = {}
    for s in sistema.values():
        baldes.setdefault(chave(s["centavos"]), []).append(s)
    indice = {}
    for k, lista in baldes.items():
        com_data = sorted((x for x in lista if x["data"]), key=lambda x: x["data"])
        sem_data = [x for x in lista if not x["data"]]
        indice[k] = ([x["data"].toordinal() for x in com_data], com_data, sem_data)

    arestas, melhor_nota, empates = [], {}, {}
    for b in banco.values():
        item = indice.get(chave(b["centavos"]))
        if not item:
            continue
        ordinais, com_data, sem_data = item
        if b["data"]:
            o = b["data"].toordinal()
            ini, fim = bisect_left(ordinais, o - tolerancia), bisect_right(ordinais, o + tolerancia)
            candidatos = [(com_data[k], abs(ordinais[k] - o)) for k in range(ini, fim)]
            candidatos += [(s, None) for s in sem_data[:200]]
        else:
            candidatos = [(s, None) for s in (com_data + sem_data)[:200]]
        for s, dias in candidatos:
            sim = similaridade(b["tokens"], s["tokens"])
            nota = 100 - (dias * 8 if dias is not None else 40) + round(sim * 30)
            arestas.append((-nota, 999 if dias is None else dias, b["id"], s["id"], nota, dias, sim))
            if nota > melhor_nota.get(b["id"], -10**9):
                melhor_nota[b["id"]], empates[b["id"]] = nota, 1
            elif nota == melhor_nota[b["id"]]:
                empates[b["id"]] += 1

    arestas.sort()
    usados_b, usados_s, pares = set(), set(), {}
    for _, _, bid, sid, nota, dias, sim in arestas:
        if bid in usados_b or sid in usados_s:
            continue
        usados_b.add(bid)
        usados_s.add(sid)
        ambiguo = empates.get(bid, 1) > 1
        pid = f"P{len(pares) + 1}"
        pares[pid] = {
            "id": pid, "b": bid, "s": sid, "nota": nota, "dias": dias, "sim": sim,
            "ambiguo": ambiguo, "conf": confianca(dias, sim, ambiguo),
            "status": "sugestao", "origem": "Automático", "quando": "",
        }
    return pares


def montar_itens(df, prefixo):
    itens = {}
    for i, r in enumerate(df.itertuples(index=False), 1):
        iid = f"{prefixo}{i}"
        itens[iid] = {"id": iid, "data": r.data, "descricao": r.descricao, "centavos": int(r.centavos),
                      "local": r.local, "tokens": tokens(r.descricao)}
    return itens


def iniciar_conciliacao(df_b, df_s, tolerancia, ignorar_sinal, responsavel, nomes):
    banco, sistema = montar_itens(df_b, "B"), montar_itens(df_s, "S")
    conc = {
        "banco": banco, "sistema": sistema,
        "pares": gerar_sugestoes(banco, sistema, tolerancia, ignorar_sinal),
        "justificados": {}, "log": [], "relatorio": None,
        "config": {"tolerancia": tolerancia, "ignorar_sinal": ignorar_sinal,
                   "responsavel": responsavel.strip(), "arquivos": nomes,
                   "inicio": agora().strftime("%d/%m/%Y %H:%M")},
    }
    st.session_state.conc = conc
    registrar("Conciliação iniciada",
              f"{len(banco)} lançamentos do banco, {len(sistema)} do sistema, "
              f"{len(conc['pares'])} sugestões, tolerância de {tolerancia} dia(s)")


# ==========================================================
# 5. Estado e ações (trilha de auditoria)
# ==========================================================
def conc():
    return st.session_state.conc


def item(iid):
    return conc()["banco"].get(iid) or conc()["sistema"].get(iid)


def rotulo_item(x):
    return f"{data_br(x['data'])} · {moeda(x['centavos'])} · {x['descricao'][:60]}"


def registrar(acao, detalhe):
    c = conc()
    c["relatorio"] = None
    c["log"].append({"Data e hora": agora().strftime("%d/%m/%Y %H:%M:%S"),
                     "Responsável": c["config"]["responsavel"], "Ação": acao, "Detalhe": detalhe})


def descrever_par(p):
    b, s = item(p["b"]), item(p["s"])
    return f"{rotulo_item(b)}  ⇄  {rotulo_item(s)}"


def ocupados():
    b, s = set(), set()
    for p in conc()["pares"].values():
        if p["status"] in ("sugestao", "conciliado"):
            b.add(p["b"])
            s.add(p["s"])
    return b, s


def sobras(lado):
    ocup_b, ocup_s = ocupados()
    ocup = ocup_b if lado == "banco" else ocup_s
    return [x for k, x in conc()[lado].items() if k not in ocup and k not in conc()["justificados"]]


def mudar_status(pids, novo):
    nomes = {"conciliado": "Par conciliado", "recusado": "Sugestão recusada", "sugestao": "Par devolvido para revisão"}
    for pid in pids:
        p = conc()["pares"][pid]
        p["status"], p["quando"] = novo, agora().strftime("%d/%m/%Y %H:%M:%S")
        registrar(nomes[novo], descrever_par(p))


def restaurar(pid):
    p = conc()["pares"][pid]
    ocup_b, ocup_s = ocupados()
    livres = p["b"] not in ocup_b and p["s"] not in ocup_s
    livres = livres and p["b"] not in conc()["justificados"] and p["s"] not in conc()["justificados"]
    if livres:
        mudar_status([pid], "sugestao")


def desfazer_conciliacao(pid):
    p = conc()["pares"][pid]
    if p["origem"] == "Manual":
        p["status"] = "desfeito"
        registrar("Conciliação manual desfeita", descrever_par(p))
    else:
        mudar_status([pid], "sugestao")


def conciliar_manual(bid, sid):
    c = conc()
    pid = f"M{len(c['pares']) + 1}"
    b, s = c["banco"][bid], c["sistema"][sid]
    dias = abs((b["data"] - s["data"]).days) if b["data"] and s["data"] else None
    c["pares"][pid] = {"id": pid, "b": bid, "s": sid, "nota": 0, "dias": dias,
                       "sim": similaridade(b["tokens"], s["tokens"]), "ambiguo": False,
                       "conf": "Manual", "status": "conciliado", "origem": "Manual",
                       "quando": agora().strftime("%d/%m/%Y %H:%M:%S")}
    registrar("Conciliação manual", descrever_par(c["pares"][pid]))


def justificar(ids, texto):
    for iid in ids:
        conc()["justificados"][iid] = {"texto": texto.strip(), "quando": agora().strftime("%d/%m/%Y %H:%M:%S")}
        lado = "Banco" if iid.startswith("B") else "Sistema"
        registrar("Pendência justificada", f"{lado}: {rotulo_item(item(iid))} — {texto.strip()}")


def desfazer_justificativa(iid):
    conc()["justificados"].pop(iid, None)
    registrar("Justificativa removida", rotulo_item(item(iid)))


# ==========================================================
# 6. Relatório em Excel
# ==========================================================
def tabela_itens(lista):
    return pd.DataFrame([{"Data": data_br(x["data"]), "Descrição": x["descricao"],
                          "Valor (R$)": x["centavos"] / 100, "Origem no arquivo": x["local"]} for x in lista],
                        columns=["Data", "Descrição", "Valor (R$)", "Origem no arquivo"])


COLUNAS_PARES = ["Data banco", "Descrição banco", "Valor banco (R$)", "Data sistema", "Descrição sistema",
                 "Valor sistema (R$)", "Diferença de dias", "Confiança", "Origem", "Decidido em",
                 "Origem banco", "Origem sistema"]


def tabela_pares(lista):
    linhas = []
    for p in lista:
        b, s = item(p["b"]), item(p["s"])
        linhas.append({
            "Data banco": data_br(b["data"]), "Descrição banco": b["descricao"], "Valor banco (R$)": b["centavos"] / 100,
            "Data sistema": data_br(s["data"]), "Descrição sistema": s["descricao"],
            "Valor sistema (R$)": s["centavos"] / 100,
            "Diferença de dias": p["dias"], "Confiança": p["conf"],
            "Origem": p["origem"], "Decidido em": p["quando"],
            "Origem banco": b["local"], "Origem sistema": s["local"],
        })
    return pd.DataFrame(linhas, columns=COLUNAS_PARES)


def soma(lista):
    return sum(x["centavos"] for x in lista)


def resumo():
    c = conc()
    banco, sistema = list(c["banco"].values()), list(c["sistema"].values())
    pares = c["pares"].values()
    conciliados = [p for p in pares if p["status"] == "conciliado"]
    pendentes = [p for p in pares if p["status"] == "sugestao"]
    so_b, so_s = sobras("banco"), sobras("sistema")
    just = [item(i) for i in c["justificados"]]
    entradas = lambda l: soma([x for x in l if x["centavos"] > 0])
    saidas = lambda l: soma([x for x in l if x["centavos"] < 0])
    linhas = [
        ("Extrato bancário · entradas", len([x for x in banco if x["centavos"] > 0]), entradas(banco)),
        ("Extrato bancário · saídas", len([x for x in banco if x["centavos"] < 0]), saidas(banco)),
        ("Extrato bancário · movimento líquido", len(banco), soma(banco)),
        ("Controle interno · entradas", len([x for x in sistema if x["centavos"] > 0]), entradas(sistema)),
        ("Controle interno · saídas", len([x for x in sistema if x["centavos"] < 0]), saidas(sistema)),
        ("Controle interno · movimento líquido", len(sistema), soma(sistema)),
        ("Diferença de movimento (banco − sistema)", None, soma(banco) - soma(sistema)),
        ("Pares conciliados", len(conciliados), soma([item(p["b"]) for p in conciliados])),
        ("Sugestões ainda não revisadas", len(pendentes), soma([item(p["b"]) for p in pendentes])),
        ("Só no banco (sem registro no sistema)", len(so_b), soma(so_b)),
        ("Só no sistema (sem movimento no banco)", len(so_s), soma(so_s)),
        ("Pendências justificadas", len(just), soma(just)),
    ]
    return pd.DataFrame([{"Indicador": a, "Quantidade": b, "Valor (R$)": v / 100} for a, b, v in linhas])


def gerar_excel():
    c = conc()
    cfg = c["config"]
    pares = list(c["pares"].values())
    cabecalho = pd.DataFrame([
        {"Campo": "Sistema", "Informação": f"{NOME_APP} · {SUBTITULO}"},
        {"Campo": "Responsável", "Informação": cfg["responsavel"]},
        {"Campo": "Conciliação iniciada em", "Informação": cfg["inicio"]},
        {"Campo": "Relatório gerado em", "Informação": agora().strftime("%d/%m/%Y %H:%M")},
        {"Campo": "Extrato bancário", "Informação": cfg["arquivos"][0]},
        {"Campo": "Controle interno", "Informação": cfg["arquivos"][1]},
        {"Campo": "Tolerância de datas", "Informação": f"{cfg['tolerancia']} dia(s)"},
        {"Campo": "Comparação de sinal", "Informação": "Ignorando o sinal" if cfg["ignorar_sinal"] else "Sinal considerado"},
    ])
    just = pd.DataFrame([
        {"Lado": "Banco" if i.startswith("B") else "Sistema", "Data": data_br(item(i)["data"]),
         "Descrição": item(i)["descricao"], "Valor (R$)": item(i)["centavos"] / 100,
         "Justificativa": j["texto"], "Justificado em": j["quando"]}
        for i, j in c["justificados"].items()
    ], columns=["Lado", "Data", "Descrição", "Valor (R$)", "Justificativa", "Justificado em"])
    abas = {
        "Resumo": None,
        "Conciliados": tabela_pares([p for p in pares if p["status"] == "conciliado"]),
        "Só no banco": tabela_itens(sobras("banco")),
        "Só no sistema": tabela_itens(sobras("sistema")),
        "Justificados": just,
        "A revisar": tabela_pares([p for p in pares if p["status"] == "sugestao"]),
        "Recusados": tabela_pares([p for p in pares if p["status"] == "recusado"]),
        "Histórico": pd.DataFrame(c["log"]),
    }
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as escritor:
        cabecalho.to_excel(escritor, sheet_name="Resumo", index=False)
        resumo().to_excel(escritor, sheet_name="Resumo", index=False, startrow=len(cabecalho) + 2)
        for nome, df in abas.items():
            if df is not None:
                df.to_excel(escritor, sheet_name=nome, index=False)
        from openpyxl.styles import Font, PatternFill

        preenchimento = PatternFill("solid", fgColor=COR["marinho"][1:])
        for ws in escritor.book.worksheets:
            for coluna in ws.columns:
                largura = max(len(str(cel.value or "")) for cel in coluna)
                ws.column_dimensions[coluna[0].column_letter].width = min(max(largura + 2, 10), 60)
            for linha in ws.iter_rows():
                for cel in linha:
                    if isinstance(cel.value, str) and cel.value in ("Campo", "Indicador") or cel.row == 1 and cel.value:
                        cel.font, cel.fill = Font(bold=True, color="FFFFFF"), preenchimento
                    titulo = ws.cell(row=1, column=cel.column).value or ""
                    if isinstance(cel.value, float) and ("R$" in str(titulo) or ws.title == "Resumo"):
                        cel.number_format = '"R$" #,##0.00;[Red]-"R$" #,##0.00'
            ws.freeze_panes = "A2"
    return buffer.getvalue()


# ==========================================================
# 7. Interface
# ==========================================================
def svg_base64(caminho):
    try:
        return base64.b64encode(caminho.read_bytes()).decode()
    except OSError:
        return ""


def aplicar_estilo():
    st.markdown(f"""<style>
    .block-container {{ padding-top: 1.6rem; }}
    .cabecalho {{ display:flex; align-items:center; gap:14px; margin-bottom:.2rem; }}
    .cabecalho img {{ height:52px; }}
    .cabecalho h1 {{ font-size:1.75rem; margin:0; color:{COR['marinho']}; font-weight:800; letter-spacing:-.5px; }}
    .cabecalho h1 span {{ color:{COR['verde']}; }}
    .cabecalho p {{ margin:0; color:{COR['cinza']}; font-size:.92rem; }}
    .card-box {{ background:#fff; border:1px solid {COR['borda']}; border-radius:10px; padding:14px 16px;
                 box-shadow:0 1px 3px rgba(11,37,69,.05); min-height:118px; }}
    .card-box .valor {{ float:right; font-weight:700; font-size:1.02rem; }}
    .card-box .verde {{ color:{COR['verde']}; }}
    .card-box .vermelho {{ color:{COR['vermelho']}; }}
    .card-box .lado {{ font-size:.7rem; text-transform:uppercase; letter-spacing:.8px; color:{COR['cinza']}; font-weight:700; }}
    .card-box .quando {{ color:{COR['azul']}; font-size:.9rem; font-weight:600; margin:2px 0 4px; }}
    .card-box .desc {{ color:#3d4a5c; font-size:.86rem; line-height:1.3; }}
    .card-box .local {{ color:#8a97a8; font-size:.74rem; margin-top:4px; }}
    .card-box .sug {{ color:{COR['cinza']}; font-size:.78rem; margin-top:6px; border-top:1px dashed {COR['borda']}; padding-top:6px; }}
    .badge {{ display:inline-block; padding:1px 8px; border-radius:999px; font-size:.72rem; font-weight:700; }}
    .badge.Alta {{ background:#E3F6EF; color:#0B7A5A; }}
    .badge.Média {{ background:#FFF4DA; color:#8A6200; }}
    .badge.Baixa {{ background:#FDE7E7; color:#A32F2F; }}
    .badge.Manual {{ background:#E6ECF5; color:{COR['azul']}; }}
    .rodape {{ text-align:center; color:#8a97a8; font-size:.75rem; margin-top:2.5rem; }}
    </style>""", unsafe_allow_html=True)


def cabecalho():
    logo = svg_base64(LOGO)
    img = f'<img src="data:image/svg+xml;base64,{logo}" alt="{NOME_APP}">' if logo else ""
    st.markdown(f'<div class="cabecalho">{img}<div><h1>Consolida<span>.CFT</span></h1>'
                f"<p>{SUBTITULO}</p></div></div>", unsafe_allow_html=True)


def card_item(x, lado, extra=""):
    cor = "verde" if x["centavos"] > 0 else "vermelho"
    return (f'<div class="card-box"><div class="valor {cor}">{moeda(x["centavos"], sinal=True)}</div>'
            f'<div class="lado">{lado}</div><div class="quando">📅 {data_br(x["data"])}</div>'
            f'<div class="desc">{html.escape(x["descricao"])}</div>'
            f'<div class="local">{html.escape(x["local"])}</div>{extra}</div>')


def mostrar_previa(df):
    previa = df.head(8)
    st.dataframe(pd.DataFrame({"Data": [data_br(d) for d in previa["data"]],
                               "Descrição": previa["descricao"],
                               "Valor": [moeda(v) for v in previa["centavos"]]}),
                 hide_index=True, width="stretch")


def painel_mapeamento(arquivo, chave):
    try:
        tipo, conteudo = ler_arquivo(arquivo.name, arquivo.getvalue())
    except ImportError:
        st.error("Para ler arquivos .xls antigos é preciso o pacote `xlrd` (já listado no requirements.txt).")
        return None
    except Exception as erro:
        st.error(f"Não consegui ler **{arquivo.name}**: {erro}")
        return None

    if tipo == "ofx":
        df = conteudo
        if df.empty:
            st.error("Nenhuma transação encontrada neste OFX.")
            return None
        st.success(f"OFX lido: {len(df)} transações.")
        if st.checkbox("Inverter o sinal deste arquivo", key=f"{chave}_inv_ofx"):
            df = df.assign(centavos=-df["centavos"])
        mostrar_previa(df)
        return df

    abas = conteudo
    if not abas:
        st.error("O arquivo está vazio.")
        return None
    with st.expander("Conferir colunas", expanded=True):
        aba = st.selectbox("Aba", list(abas), key=f"{chave}_aba") if len(abas) > 1 else next(iter(abas))
        bruto = abas[aba]
        cab_auto = detectar_cabecalho(bruto)
        cab = st.number_input("Linha do cabeçalho (0 = sem cabeçalho)", min_value=0, max_value=len(bruto),
                              value=(cab_auto + 1) if cab_auto is not None else 0, key=f"{chave}_cab_{aba}")
        cab_idx = int(cab) - 1 if cab > 0 else None
        sug = sugerir_mapeamento(bruto, cab_idx)
        rot = rotulos_colunas(bruto, cab_idx)
        opcoes = list(range(bruto.shape[1]))
        com_nenhuma = [None] + opcoes
        fmt = lambda c: "— nenhuma —" if c is None else rot[c]
        k = f"{chave}_{aba}_{cab}"
        modo = st.radio("Como o valor aparece", ["Uma coluna de valor", "Colunas separadas de crédito e débito"],
                        index=0 if sug["modo"] == "unico" else 1, key=f"{k}_modo", horizontal=True)
        c1, c2 = st.columns(2)
        col_data = c1.selectbox("Data", com_nenhuma, index=com_nenhuma.index(sug["data"]), format_func=fmt, key=f"{k}_data")
        col_desc = c2.multiselect("Descrição", opcoes, default=sug["desc"], format_func=fmt, key=f"{k}_desc")
        mapa = {"modo": "unico" if modo.startswith("Uma") else "separadas", "data": col_data, "desc": col_desc,
                "valor": None, "credito": None, "debito": None, "dc": None}
        if mapa["modo"] == "unico":
            mapa["valor"] = c1.selectbox("Valor", com_nenhuma, index=com_nenhuma.index(sug["valor"]), format_func=fmt, key=f"{k}_valor")
            mapa["dc"] = c2.selectbox("Indicador D/C (opcional)", com_nenhuma, index=com_nenhuma.index(sug["dc"]),
                                      format_func=fmt, key=f"{k}_dc")
        else:
            mapa["credito"] = c1.selectbox("Crédito", com_nenhuma, index=com_nenhuma.index(sug["credito"]), format_func=fmt, key=f"{k}_cred")
            mapa["debito"] = c2.selectbox("Débito", com_nenhuma, index=com_nenhuma.index(sug["debito"]), format_func=fmt, key=f"{k}_deb")
        mapa["inverter"] = st.checkbox("Inverter o sinal deste arquivo", key=f"{k}_inv",
                                       help="Use quando este arquivo registra saídas como valores positivos.")

    df, descartadas = aplicar_mapeamento(bruto, cab_idx, mapa, aba)
    if df.empty:
        st.warning("Nenhum lançamento reconhecido com essas colunas. Confira o mapeamento acima.")
        return None
    aviso = f" · {descartadas} linha(s) ignorada(s) (saldos, totais ou sem valor)" if descartadas else ""
    st.success(f"{len(df)} lançamentos reconhecidos{aviso}.")
    sem_data = int(df["data"].isna().sum())
    if sem_data:
        st.caption(f"⚠️ {sem_data} lançamento(s) sem data reconhecida: serão pareados só pelo valor, com confiança baixa.")
    mostrar_previa(df)
    return df


def tela_importacao():
    st.markdown("#### 1. Importar arquivos")
    responsavel = st.text_input("Responsável pela conciliação", key="responsavel",
                                placeholder="Nome de quem está conciliando (vai para o relatório)")
    c1, c2 = st.columns(2, gap="large")
    with c1:
        st.info("🏦 **Extrato bancário** · OFX, CSV, XLS ou XLSX")
        arq_b = st.file_uploader("Extrato bancário", type=TIPOS_ACEITOS, key="up_banco", label_visibility="collapsed")
        df_b = painel_mapeamento(arq_b, "banco") if arq_b else None
    with c2:
        st.info("💻 **Controle interno** · CSV, XLS ou XLSX")
        arq_s = st.file_uploader("Controle interno", type=TIPOS_ACEITOS, key="up_sistema", label_visibility="collapsed")
        df_s = painel_mapeamento(arq_s, "sistema") if arq_s else None

    st.markdown("#### 2. Regras de pareamento")
    r1, r2 = st.columns(2, gap="large")
    tolerancia = r1.slider("Tolerância de datas (dias)", 0, 15, 3,
                           help="Diferença máxima entre a data do banco e a do sistema para sugerir um par.")
    ignorar = r2.checkbox("Comparar valores ignorando o sinal",
                          help="Marque quando o controle interno registra as saídas como valores positivos.")

    pronto = df_b is not None and df_s is not None and bool(responsavel.strip())
    if df_b is not None and df_s is not None and not responsavel.strip():
        st.caption("Informe o responsável para liberar a conciliação.")
    if st.button("Iniciar conciliação", type="primary", width="stretch", disabled=not pronto):
        with st.spinner("Cruzando os lançamentos…"):
            iniciar_conciliacao(df_b, df_s, tolerancia, ignorar, responsavel, [arq_b.name, arq_s.name])
        st.rerun()


def paginar(total, chave):
    paginas = max(1, -(-total // POR_PAGINA))
    if st.session_state.get(chave, 1) > paginas:
        st.session_state[chave] = paginas
    if paginas == 1:
        return 0, total
    pagina = st.number_input(f"Página (de {paginas})", 1, paginas, key=chave)
    inicio = (pagina - 1) * POR_PAGINA
    return inicio, min(inicio + POR_PAGINA, total)


def aba_sugestoes(sugestoes):
    if not sugestoes:
        st.success("Nenhuma sugestão aguardando revisão.")
        return
    filtro = st.radio("Confiança", ["Todas", "Alta", "Média", "Baixa"], horizontal=True, key="filtro_conf")
    lista = [p for p in sugestoes if filtro == "Todas" or p["conf"] == filtro]
    lista.sort(key=lambda p: (ORDEM_CONF[p["conf"]], item(p["b"])["data"] or date.max, p["id"]))
    if not lista:
        st.info("Nenhuma sugestão com essa confiança.")
        return
    c1, c2, c3 = st.columns([2, 2, 4], vertical_alignment="center")
    revisado = True
    if filtro != "Alta":
        revisado = c3.checkbox("Revisei as sugestões deste filtro", key=f"revisado_{filtro}",
                               help="Ações em massa fora da confiança alta exigem confirmação.")
    ids = [p["id"] for p in lista]
    c1.button(f"✅ Conciliar {len(lista)}", type="primary", width="stretch", disabled=not revisado,
              on_click=mudar_status, args=(ids, "conciliado"), key=f"massa_ok_{filtro}")
    c2.button(f"❌ Recusar {len(lista)}", width="stretch", disabled=not revisado,
              on_click=mudar_status, args=(ids, "recusado"), key=f"massa_rec_{filtro}")
    st.divider()
    ini, fim = paginar(len(lista), "pag_sug")
    for p in lista[ini:fim]:
        b, s = item(p["b"]), item(p["s"])
        dias = "data sem correspondência" if p["dias"] is None else (
            "mesma data" if p["dias"] == 0 else f"{p['dias']} dia(s) de diferença")
        texto = " · descrição parecida" if p["sim"] >= 0.5 else ""
        alerta = " · ⚠️ há outros lançamentos com o mesmo valor e data" if p["ambiguo"] else ""
        extra = (f'<div class="sug">Confiança <span class="badge {p["conf"]}">{p["conf"]}</span> · '
                 f"{dias}{texto}{alerta}</div>")
        cb, cm, cs = st.columns([4, 1.3, 4], vertical_alignment="center")
        cb.markdown(card_item(b, "Banco"), unsafe_allow_html=True)
        cm.button("✅ Conciliar", key=f"ok_{p['id']}", type="primary", width="stretch",
                  on_click=mudar_status, args=([p["id"]], "conciliado"))
        cm.button("❌ Recusar", key=f"rec_{p['id']}", width="stretch",
                  on_click=mudar_status, args=([p["id"]], "recusado"))
        cs.markdown(card_item(s, "Controle interno", extra), unsafe_allow_html=True)
        st.write("")


def aba_sobras(lista, lado):
    if not lista:
        st.success("Nada pendente deste lado. 🎉")
        return
    st.caption(f"{len(lista)} lançamento(s) · total {moeda(soma(lista))}. "
               "Concilie manualmente ou justifique na aba **Ajustes manuais**.")
    df = tabela_itens(sorted(lista, key=lambda x: (x["data"] or date.max, x["id"])))
    df["Valor (R$)"] = [moeda(round(v * 100)) for v in df["Valor (R$)"]]
    st.dataframe(df, hide_index=True, width="stretch", key=f"tab_{lado}")


def aba_ajustes(so_b, so_s):
    c = conc()
    st.markdown("##### Conciliar manualmente")
    if so_b and so_s:
        busca = normalizar(st.text_input("Filtrar por valor ou texto", key="man_busca",
                                         placeholder="Ex.: 150,00 ou nome do profissional"))
        filtrar = lambda l: [x for x in l if not busca or busca in normalizar(rotulo_item(x))]
        fb, fs = filtrar(so_b), filtrar(so_s)
        c1, c2 = st.columns(2)
        bid = c1.selectbox("Lançamento do banco", [x["id"] for x in fb], format_func=lambda i: rotulo_item(item(i)),
                           key="man_b", index=None, placeholder="Escolha…")
        sid = c2.selectbox("Lançamento do sistema", [x["id"] for x in fs], format_func=lambda i: rotulo_item(item(i)),
                           key="man_s", index=None, placeholder="Escolha…")
        if bid and sid:
            vb, vs = item(bid)["centavos"], item(sid)["centavos"]
            if c["config"]["ignorar_sinal"]:
                vb, vs = abs(vb), abs(vs)
            liberado = True
            if vb != vs:
                st.warning(f"Os valores diferem em {moeda(vb - vs)}.")
                liberado = st.checkbox("Confirmo a conciliação com valores diferentes", key="man_ok")
            st.button("🔗 Conciliar manualmente", type="primary", disabled=not liberado,
                      on_click=conciliar_manual, args=(bid, sid))
    else:
        st.caption("É preciso ter pendências dos dois lados para conciliar manualmente.")

    st.markdown("##### Justificar pendências")
    st.caption("Para tarifas, estornos, rendimentos ou lançamentos que não têm contrapartida.")
    opcoes = [x["id"] for x in so_b + so_s]
    if opcoes:
        with st.form("form_justificar", clear_on_submit=True):
            ids = st.multiselect("Lançamentos", opcoes,
                                 format_func=lambda i: ("🏦 " if i.startswith("B") else "💻 ") + rotulo_item(item(i)))
            texto = st.text_input("Justificativa", placeholder="Ex.: tarifa de manutenção de conta")
            if st.form_submit_button("Marcar como justificado"):
                if ids and texto.strip():
                    justificar(ids, texto)
                    st.rerun()
                st.warning("Escolha ao menos um lançamento e escreva a justificativa.")
    else:
        st.caption("Nenhuma pendência para justificar.")

    if c["justificados"]:
        st.markdown("##### Já justificados")
        for iid, j in list(c["justificados"].items()):
            a, b = st.columns([6, 1], vertical_alignment="center")
            a.markdown(f"{'🏦' if iid.startswith('B') else '💻'} {html.escape(rotulo_item(item(iid)))} — *{html.escape(j['texto'])}*")
            b.button("Desfazer", key=f"desj_{iid}", on_click=desfazer_justificativa, args=(iid,))


def aba_conciliados(conciliados):
    if not conciliados:
        st.info("Nenhum par conciliado ainda.")
        return
    df = tabela_pares(conciliados)
    for col in ("Valor banco (R$)", "Valor sistema (R$)"):
        df[col] = [moeda(round(v * 100)) for v in df[col]]
    st.dataframe(df.drop(columns=["Origem banco", "Origem sistema"]), hide_index=True, width="stretch")
    a, b = st.columns([6, 1.4], vertical_alignment="bottom")
    pid = a.selectbox("Desfazer uma conciliação", [p["id"] for p in conciliados], index=None,
                      format_func=lambda i: descrever_par(conc()["pares"][i]), placeholder="Escolha o par…", key="desf_sel")
    b.button("↩️ Desfazer", disabled=pid is None, on_click=desfazer_conciliacao, args=(pid,), width="stretch")


def aba_recusados(recusados):
    if not recusados:
        st.info("Nenhuma sugestão recusada.")
        return
    st.caption("Ao recusar, os dois lançamentos voltam para as pendências. Restaure se recusou por engano.")
    ocup_b, ocup_s = ocupados()
    ini, fim = paginar(len(recusados), "pag_rec")
    for p in recusados[ini:fim]:
        livre = (p["b"] not in ocup_b and p["s"] not in ocup_s
                 and p["b"] not in conc()["justificados"] and p["s"] not in conc()["justificados"])
        a, b = st.columns([6, 1.2], vertical_alignment="center")
        a.markdown(html.escape(descrever_par(p)))
        b.button("Restaurar", key=f"rest_{p['id']}", disabled=not livre, on_click=restaurar, args=(p["id"],),
                 help=None if livre else "Um dos lançamentos já foi usado em outro par ou justificado.")


def aba_relatorio():
    c = conc()
    df = resumo()
    df["Valor (R$)"] = [moeda(round(v * 100)) for v in df["Valor (R$)"]]
    st.dataframe(df, hide_index=True, width="stretch")
    if c["relatorio"] is None:
        if st.button("📊 Gerar relatório em Excel", type="primary"):
            c["relatorio"] = gerar_excel()
            st.rerun()
    else:
        st.download_button("⬇️ Baixar relatório (.xlsx)", c["relatorio"], type="primary",
                           file_name=f"consolida_cft_{agora():%Y%m%d_%H%M}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.caption("O relatório tem as abas Resumo, Conciliados, Só no banco, Só no sistema, Justificados, "
                   "A revisar, Recusados e Histórico.")
    st.markdown("##### Histórico de ações")
    st.dataframe(pd.DataFrame(c["log"][::-1]), hide_index=True, width="stretch")


def nova_conciliacao():
    for chave in list(st.session_state):
        if chave == "conc" or chave.startswith(("pag_", "man_", "filtro", "revisado_", "desf_", "massa_")):
            del st.session_state[chave]


def tela_conciliacao():
    c = conc()
    pares = list(c["pares"].values())
    sugestoes = [p for p in pares if p["status"] == "sugestao"]
    conciliados = [p for p in pares if p["status"] == "conciliado"]
    recusados = [p for p in pares if p["status"] == "recusado"]
    so_b, so_s = sobras("banco"), sobras("sistema")

    topo, acao = st.columns([5, 1.4], vertical_alignment="center")
    topo.caption(f"Responsável: **{html.escape(c['config']['responsavel'])}** · "
                 f"{html.escape(c['config']['arquivos'][0])} × {html.escape(c['config']['arquivos'][1])} · "
                 f"iniciada em {c['config']['inicio']}")
    with acao.popover("🔄 Nova conciliação", width="stretch"):
        st.write("O trabalho atual será descartado. Gere o relatório antes, se precisar.")
        st.button("Descartar e recomeçar", type="primary", on_click=nova_conciliacao)

    m = st.columns(5)
    m[0].metric("A revisar", len(sugestoes))
    m[1].metric("Conciliados", len(conciliados))
    m[2].metric("Só no banco", len(so_b), help=f"Total {moeda(soma(so_b))}")
    m[3].metric("Só no sistema", len(so_s), help=f"Total {moeda(soma(so_s))}")
    m[4].metric("Justificados", len(c["justificados"]))

    total = len(c["banco"]) + len(c["sistema"])
    resolvidos = 2 * len(conciliados) + len(c["justificados"])
    st.progress(resolvidos / total if total else 0,
                text=f"{resolvidos} de {total} lançamentos resolvidos (conciliados ou justificados)")

    if not sugestoes and not so_b and not so_s:
        st.success("🎉 Conciliação fechada: tudo foi conciliado ou justificado. Gere o relatório na última aba.")
    elif not sugestoes:
        st.warning(f"As sugestões acabaram, mas ainda há {len(so_b)} lançamento(s) só no banco e "
                   f"{len(so_s)} só no sistema. Resolva-os em **Ajustes manuais**.")

    abas = st.tabs([f"Sugestões ({len(sugestoes)})", f"Só no banco ({len(so_b)})", f"Só no sistema ({len(so_s)})",
                    "Ajustes manuais", f"Conciliados ({len(conciliados)})", f"Recusados ({len(recusados)})", "Relatório"])
    with abas[0]:
        aba_sugestoes(sugestoes)
    with abas[1]:
        aba_sobras(so_b, "banco")
    with abas[2]:
        aba_sobras(so_s, "sistema")
    with abas[3]:
        aba_ajustes(so_b, so_s)
    with abas[4]:
        aba_conciliados(conciliados)
    with abas[5]:
        aba_recusados(recusados)
    with abas[6]:
        aba_relatorio()


def main():
    st.set_page_config(page_title=NOME_APP, page_icon=str(ICONE) if ICONE.exists() else "⚖️", layout="wide")
    aplicar_estilo()
    cabecalho()
    st.divider()
    if "conc" in st.session_state:
        tela_conciliacao()
    else:
        tela_importacao()
    st.markdown('<div class="rodape">Consolida.CFT · BK Serviços, Produção &amp; Comunicação</div>',
                unsafe_allow_html=True)


# ==========================================================
# Inicializador: "python consolida_cft.py" abre o Streamlit
# ==========================================================
def _rodando_no_streamlit():
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx

        return get_script_run_ctx(suppress_warning=True) is not None
    except Exception:
        return False


if __name__ == "__main__":
    if _rodando_no_streamlit():
        main()
    else:
        subprocess.run([sys.executable, "-m", "streamlit", "run", os.path.abspath(__file__)])
