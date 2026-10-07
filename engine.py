"""Motor de comparação do Validador Pós-Migração | Agante Tecnologia · Senior HCM.

Sem dependência de Streamlit: lê arquivos, identifica cabeçalhos, monta o DE/PARA,
normaliza valores, classifica cada campo e gera o relatório Excel.
"""
from __future__ import annotations

import csv
import hashlib
import io
import math
import re
import unicodedata
from datetime import date, datetime, timedelta

# --------------------------------------------------------------------------- status
STATUS = {
    "OK":             {"label": "OK", "tone": "ok", "sym": "✓", "sev": 0},
    "CORRIGIDO":      {"label": "CORRIGIDO NO LAYOUT AJUSTADO", "tone": "fix", "sym": "↻", "sev": 1},
    "ALTERADO":       {"label": "ALTERADO NO SENIOR", "tone": "alt", "sym": "≠", "sev": 2},
    "MANUAL":         {"label": "NECESSITA VALIDAÇÃO MANUAL", "tone": "warn", "sym": "?", "sev": 3},
    "NAO_PREENCHIDO": {"label": "CAMPO NÃO PREENCHIDO", "tone": "bad", "sym": "∅", "sev": 4},
    "DIVERGENCIA":    {"label": "DIVERGÊNCIA", "tone": "bad", "sym": "!", "sev": 5},
    "DUPLICADO":      {"label": "REGISTRO DUPLICADO", "tone": "warn", "sym": "⧉", "sev": 6},
    "SOMENTE_SENIOR": {"label": "CRIADO SOMENTE NO SENIOR", "tone": "alt", "sym": "+", "sev": 7},
    "NAO_MIGRADO":    {"label": "NÃO MIGRADO", "tone": "bad", "sym": "✕", "sev": 8},
    "CHAVE_NL":       {"label": "CHAVE NÃO LOCALIZADA", "tone": "warn", "sym": "⌕", "sev": 9},
}
TONE_BG = {"ok": "#E3F3E9", "fix": "#DDEFF6", "alt": "#EBE8F8", "warn": "#FBEFD2", "bad": "#FBE6E3", "mute": "#EEF1EF"}
TONE_FG = {"ok": "#1B7A43", "fix": "#0A6A8A", "alt": "#55459F", "warn": "#865600", "bad": "#B42318", "mute": "#64716B"}
LABEL_TO_CODE = {v["label"]: k for k, v in STATUS.items()}

TIPOS = {"texto": "Texto", "codigo": "Código", "numero": "Número", "moeda": "Valor (R$)", "data": "Data", "cpf": "CPF"}
PAPEIS = {"": "—", "nome": "Nome do colaborador", "matricula": "Matrícula", "cpf": "CPF", "empresa": "Empresa"}
TIPO_ETAPA = {"origem": "Origem", "senior": "Extração Senior", "ajuste": "Layout ajustado"}
DESC_ETAPA = {
    "origem": "Arquivo utilizado como origem para o processo de migração.",
    "senior": "Arquivo extraído do Senior após a execução da migração.",
    "ajuste": "Arquivo corrigido após a identificação de inconsistências.",
}
DEFAULT_RULES = {"trim": True, "case_ins": True, "acentos": True, "data_nula": True,
                 "nulos": "NULL, N/A, NA, NAN, NONE, -, #N/D"}
EXTS = (".xlsx", ".xlsm", ".xls", ".csv", ".txt")

ABSENT = object()   # registro ausente na etapa
NOCOL = object()    # campo não mapeado na etapa

# --------------------------------------------------------------------------- sinônimos Senior
SYN = [
    ["NUMEMP", "CODEMPRESA", "EMPRESA", "CODEMP", "CODIGOEMPRESA", "NUMEROEMPRESA"],
    ["TIPCOL", "TIPOCOLABORADOR", "TIPOCOL", "TIPCOLABORADOR"],
    ["NUMCAD", "MATRICULA", "CADASTRO", "CODCOLABORADOR", "CODIGOCOLABORADOR", "CODFUNCIONARIO"],
    ["NOMFUN", "NOME", "NOMECOLABORADOR", "NOMEFUNCIONARIO"],
    ["NUMCPF", "CPF", "CPFCOLABORADOR", "CPFFUNCIONARIO"],
    ["DATADM", "DATAADMISSAO", "DTADMISSAO", "ADMISSAO"],
    ["DATNAS", "DATANASCIMENTO", "DTNASCIMENTO", "NASCIMENTO"],
    ["CODCAR", "CARGO", "CODCARGO", "CODIGOCARGO"],
    ["CODSIN", "SINDICATO", "CODSINDICATO"],
    ["CODCCU", "CENTROCUSTO", "CENTRODECUSTO", "CCUSTO", "CODCENTROCUSTO"],
    ["VALSAL", "SALARIO", "VALORSALARIO"],
    ["CODFIL", "FILIAL", "CODFILIAL"],
    ["CODDEP", "DEPENDENTE", "CODDEPENDENTE", "CODIGODEPENDENTE"],
    ["NOMDEP", "NOMEDEPENDENTE"],
    ["GRAPAR", "PARENTESCO", "GRAUPARENTESCO"],
    ["INIFER", "INICIOFERIAS", "DATAINICIOFERIAS"],
    ["ESTCIV", "ESTADOCIVIL"],
    ["TIPSEX", "SEXO"],
    ["NUMPIS", "PIS", "PISPASEP", "NIS"],
    ["SIGUFS", "UF", "SIGLAUF", "ESTADO"],
]


def strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def canon(s) -> str:
    return re.sub(r"[^A-Z0-9]", "", strip_accents(str(s if s is not None else "")).upper())


SYN_IDX = {n: gi for gi, g in enumerate(SYN) for n in g}


def syn_name(h) -> str:
    c = canon(h)
    g = SYN_IDX.get(c)
    return SYN[g][0] if g is not None else c


# --------------------------------------------------------------------------- leitura
def decode(buf: bytes) -> str:
    try:
        return buf.decode("utf-8").lstrip("﻿")
    except UnicodeDecodeError:
        return buf.decode("cp1252", errors="replace")


def detect_delim(text: str) -> str:
    sample = "\n".join(text.splitlines()[:8])
    return max([";", "\t", "|", ","], key=lambda d: sample.count(d))


def parse_delimited(text: str) -> list[list]:
    d = detect_delim(text)
    return [row for row in csv.reader(io.StringIO(text), delimiter=d)]


def _cell(v):
    if v is None:
        return ""
    if isinstance(v, float) and math.isnan(v):
        return ""
    return v


def read_source(name: str, data) -> list[tuple[str, list[list]]]:
    """Devolve [(rótulo, matriz)] — uma por arquivo de texto ou por aba de planilha."""
    if isinstance(data, str):
        return [(name, parse_delimited(data.lstrip("﻿")))]
    if isinstance(data, list):
        return [(name, data)]
    low = name.lower()
    if low.endswith((".csv", ".txt")):
        return [(name, parse_delimited(decode(data)))]
    if low.endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        out = []
        for ws in wb.worksheets:
            m = [[_cell(v) for v in row] for row in ws.iter_rows(values_only=True)]
            out.append((f"{name} › {ws.title}" if len(wb.worksheets) > 1 else name, m))
        wb.close()
        return out
    if low.endswith(".xls"):
        import pandas as pd
        sheets = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, engine="xlrd")
        return [(f"{name} › {sh}" if len(sheets) > 1 else name, [[_cell(v) for v in r] for r in df.values.tolist()])
                for sh, df in sheets.items()]
    raise ValueError(f"Formato não suportado: {name}")


def _filled(v) -> bool:
    return str(v).strip() != ""


def data_like(v) -> bool:
    if isinstance(v, (int, float, datetime, date)) and not isinstance(v, bool):
        return True
    s = str(v).strip()
    return s != "" and re.fullmatch(r"[\d\s.,/\-:+()%R$]+", s) is not None


def detect_header(m: list[list]) -> int:
    best, bs = -1, 0
    for i, row in enumerate(m[:30]):
        ne = [c for c in row if _filled(c)]
        if not ne:
            continue
        score = len({str(c).strip().upper() for c in ne if not data_like(c)})
        if score > bs:
            bs, best = score, i
    if best < 0:
        return -1
    ne = [c for c in m[best] if _filled(c)]
    width = max(len(r) for r in m[:50])
    ratio = sum(1 for c in ne if not data_like(c)) / len(ne)
    if ratio < 0.6 or len(ne) < min(2, width):
        return -1
    return best


def matrix_to_table(m: list[list]):
    m = [r for r in m if any(_filled(c) for c in r)]
    if not m:
        return None
    width = max(len(r) for r in m)
    h = detect_header(m)
    if h < 0:
        return {"headers": [f"Coluna {i + 1}" for i in range(width)], "rows": [list(r) + [""] * (width - len(r)) for r in m], "header_row": 0}
    seen: dict[str, int] = {}
    headers = []
    for i in range(max(width, len(m[h]))):
        n = str(m[h][i]).strip() if i < len(m[h]) and _filled(m[h][i]) else f"Coluna {i + 1}"
        if n in seen:
            seen[n] += 1
            n = f"{n} ({seen[n]})"
        else:
            seen[n] = 1
        headers.append(n)
    rows = [list(r) + [""] * (len(headers) - len(r)) for r in m[h + 1:]]
    return {"headers": headers, "rows": rows, "header_row": h + 1}


def build_stage(sources: list[tuple[str, object]]) -> dict:
    """sources: [(nome, bytes | texto colado | matriz)] → headers, rows, info."""
    tables = []
    for name, data in sources:
        try:
            for label, m in read_source(name, data):
                tables.append((label, matrix_to_table(m)))
        except Exception as e:  # arquivo ilegível
            tables.append((name, {"erro": str(e)}))
    headers, rows, info = [], [], []
    base = next((t for _, t in tables if t and "erro" not in t and t["rows"]), None)
    for label, t in tables:
        if t and "erro" in t:
            info.append({"fonte": label, "registros": 0, "usada": False, "situacao": "não foi possível ler: " + t["erro"]})
            continue
        if not t or not t["rows"]:
            info.append({"fonte": label, "registros": 0, "usada": False, "situacao": "sem registros"})
            continue
        if t is base:
            headers, rows = t["headers"], [list(r) for r in t["rows"]]
            info.append({"fonte": label, "registros": len(t["rows"]), "usada": True,
                         "situacao": "cabeçalho detectado" if t["header_row"] else "sem cabeçalho (colunas numeradas)"})
            continue
        cmap = [next((j for j, o in enumerate(t["headers"]) if canon(o) == canon(h)), -1) for h in headers]
        hits = sum(1 for j in cmap if j >= 0)
        if hits >= math.ceil(len(headers) * 0.5):
            rows.extend([[r[j] if j >= 0 else "" for j in cmap] for r in t["rows"]])
            info.append({"fonte": label, "registros": len(t["rows"]), "usada": True,
                         "situacao": "somada" + (" (colunas parciais)" if hits < len(headers) else "")})
        else:
            info.append({"fonte": label, "registros": len(t["rows"]), "usada": False, "situacao": "ignorada: colunas diferentes da primeira fonte"})
    return {"headers": headers, "rows": rows, "info": info}


def files_from_path(path: str) -> list[tuple[str, bytes]]:
    """Lê um arquivo ou todos os arquivos suportados de uma pasta (sem subpastas)."""
    import os
    p = path.strip().strip('"').strip("'")
    if not p:
        return []
    if os.path.isdir(p):
        names = sorted(f for f in os.listdir(p) if f.lower().endswith(EXTS) and not f.startswith("~$"))
        return [(f, open(os.path.join(p, f), "rb").read()) for f in names]
    if os.path.isfile(p):
        return [(os.path.basename(p), open(p, "rb").read())]
    raise FileNotFoundError(p)


# --------------------------------------------------------------------------- vários layouts
LAYOUT_STOP = {"LAYOUT", "INICIAL", "ORIGEM", "ORIGINAL", "SENIOR", "EXTRACAO", "EXTRAIDO", "EXTRAIDA", "EXPORT", "EXPORTACAO",
               "AJUSTADO", "AJUSTADA", "AJUSTE", "CORRIGIDO", "CORRIGIDA", "CORRECAO", "POS", "PRE", "MIGRACAO", "CARGA",
               "FINAL", "EXEMPLO", "DADOS", "ARQUIVO", "BASE", "IMPORTACAO", "IMPORT", "1A", "2A", "3A", "V1", "V2", "V3",
               "PRIMEIRA", "SEGUNDA", "TERCEIRA", "DE", "DA", "DO", "DOS", "DAS", "E", "CSV", "TXT", "XLSX", "XLS", "PLAN", "PLANILHA",
               "SHEET", "SHEET1", "PLAN1", "ABA", "COLADO", "CONTEUDO"}


def layout_label(fonte: str) -> str:
    """Nome legível do layout a partir do nome do arquivo/aba (ex.: '1031_dependentes_inicial.xlsx' → '1031 Dependentes')."""
    parts = [p.strip() for p in fonte.split("›")]
    cand = parts[-1] if len(parts) > 1 and canon(parts[-1]) not in LAYOUT_STOP else parts[0]
    cand = re.sub(r"\.(xlsx|xlsm|xls|csv|txt)$", "", cand, flags=re.I)
    toks = [t for t in re.split(r"[\s_\-.()\[\]]+", strip_accents(cand)) if t]
    keep = [t for t in toks if canon(t) not in LAYOUT_STOP and not re.fullmatch(r"\d{6,8}", t)]
    if not keep:
        return ""
    return " ".join(t if t.isdigit() or t.isupper() and len(t) <= 4 else t.capitalize() for t in keep)


def read_tables(sources: list[tuple[str, object]]) -> list[dict]:
    """Lê cada arquivo/aba como tabela separada (sem somar)."""
    out = []
    for name, data in sources:
        try:
            for label, m in read_source(name, data):
                t = matrix_to_table(m)
                out.append({"fonte": label, "arquivo": name, **(t or {"headers": [], "rows": [], "header_row": 0})})
        except Exception as e:
            out.append({"fonte": name, "arquivo": name, "headers": [], "rows": [], "header_row": 0, "erro": str(e)})
    return out


def group_stage(tables: list[dict]) -> list[dict]:
    """Agrupa as tabelas de uma etapa por estrutura de colunas; tabelas da mesma estrutura são somadas."""
    groups: list[dict] = []
    info = []
    for t in tables:
        if t.get("erro"):
            info.append({"fonte": t["fonte"], "registros": 0, "grupo": None, "situacao": "não foi possível ler: " + t["erro"]})
            continue
        if not t["rows"]:
            info.append({"fonte": t["fonte"], "registros": 0, "grupo": None, "situacao": "sem registros"})
            continue
        tc = [canon(h) for h in t["headers"]]
        target = None
        for g in groups:
            gc = [canon(h) for h in g["headers"]]
            hits = len(set(gc) & set(tc))
            if hits >= math.ceil(len(gc) * 0.5) and hits >= math.ceil(len(tc) * 0.5):
                target = g
                break
        if target is None:
            groups.append({"headers": list(t["headers"]), "rows": [list(r) for r in t["rows"]], "fontes": [t["fonte"]],
                           "arquivos": [t["arquivo"]], "header_row": t["header_row"]})
            info.append({"fonte": t["fonte"], "registros": len(t["rows"]), "grupo": len(groups) - 1,
                         "situacao": "cabeçalho detectado" if t["header_row"] else "sem cabeçalho (colunas numeradas)"})
        else:
            cmap = [next((j for j, o in enumerate(t["headers"]) if canon(o) == canon(h)), -1) for h in target["headers"]]
            target["rows"].extend([[r[j] if 0 <= j < len(r) else "" for j in cmap] for r in t["rows"]])
            target["fontes"].append(t["fonte"])
            if t["arquivo"] not in target["arquivos"]:
                target["arquivos"].append(t["arquivo"])
            info.append({"fonte": t["fonte"], "registros": len(t["rows"]), "grupo": groups.index(target), "situacao": "somada ao mesmo layout"})
    for i, g in enumerate(groups):
        g["gid"] = hashlib.md5("|".join(g["fontes"]).encode()).hexdigest()[:10]
        g["nome"] = next((n for n in (layout_label(f) for f in g["fontes"]) if n), f"Layout {i + 1}")
    return {"groups": groups, "info": info}


def _syn_set(headers) -> set:
    return {syn_name(h) for h in headers if canon(h)}


def _name_tokens(nome: str) -> set:
    return {canon(t) for t in re.split(r"\s+", nome) if canon(t)}


def match_score(a: dict, b: dict) -> float:
    A, B = _syn_set(a["headers"]), _syn_set(b["headers"])
    if not A or not B:
        return 0.0
    score = len(A & B) / min(len(A), len(B))
    ta, tb = _name_tokens(a["nome"]), _name_tokens(b["nome"])
    if ta and tb and ta & tb:
        score += 0.35
    return score


def assign_layouts(stage_groups: list[list[dict]], overrides: dict | None = None) -> dict:
    """Cada grupo da origem vira um layout; grupos das demais etapas vão para o layout de maior afinidade.
    overrides: {gid: lid | '__ignorar'} definido pelo usuário."""
    overrides = overrides or {}
    origin = stage_groups[0] if stage_groups else []
    layouts = [{"lid": g["gid"], "nome": g["nome"], "groups": [g] + [None] * (len(stage_groups) - 1)} for g in origin]
    by_lid = {L["lid"]: L for L in layouts}
    unassigned = []
    for si in range(1, len(stage_groups)):
        scored = []
        for g in stage_groups[si]:
            ov = overrides.get(g["gid"])
            if ov == "__ignorar":
                unassigned.append((si, g))
                continue
            if ov in by_lid:
                scored.append((99.0, g, by_lid[ov]))
                continue
            best = max(((match_score(L["groups"][0], g), L) for L in layouts), key=lambda x: x[0], default=(0, None))
            if best[1] is not None and best[0] >= 0.4:
                scored.append((best[0], g, best[1]))
            else:
                unassigned.append((si, g))
        for sc, g, L in sorted(scored, key=lambda x: -x[0]):
            if L["groups"][si] is None:
                L["groups"][si] = g
            else:
                unassigned.append((si, g))
    return {"layouts": layouts, "unassigned": unassigned}


def layout_stages(L: dict, stage_defs: list[dict]) -> list[dict]:
    out = []
    for g, sd in zip(L["groups"], stage_defs):
        if g is None:
            out.append({"headers": [], "rows": [], "nome": sd["nome"], "tipo": sd["tipo"], "arquivos": ""})
        else:
            out.append({"headers": g["headers"], "rows": g["rows"], "nome": sd["nome"], "tipo": sd["tipo"], "arquivos": ", ".join(g["fontes"])})
    return out


# --------------------------------------------------------------------------- DE/PARA
def guess_type(h: str, headers: list, rows: list) -> str:
    c, g = canon(h), syn_name(h)
    if g in ("DATADM", "DATNAS", "INIFER") or re.match(r"^(DAT|DATA|DT)", c):
        return "data"
    if g == "NUMCPF" or "CPF" in c:
        return "cpf"
    if g == "VALSAL" or re.search(r"SALARIO|^VAL|VALOR", c):
        return "moeda"
    if g in ("NUMEMP", "NUMCAD", "CODCAR", "CODSIN", "CODCCU", "TIPCOL", "CODFIL", "CODDEP", "GRAPAR", "NUMPIS") or re.match(r"^(NUM|COD|TIP)", c):
        return "codigo"
    if h in headers:
        ix = headers.index(h)
        vals = [r[ix] for r in rows[:60] if _filled(r[ix])]
        if len(vals) >= 3:
            hit = sum(1 for v in vals if isinstance(v, (datetime, date)) or
                      (isinstance(v, str) and re.match(r"^\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}$|^\d{4}-\d{2}-\d{2}", v.strip())))
            if hit / len(vals) >= 0.8:
                return "data"
    return "texto"


def find_match(name: str, headers: list) -> str:
    c = canon(name)
    for h in headers:
        if canon(h) == c:
            return h
    g = SYN_IDX.get(c)
    if g is not None:
        for h in headers:
            if SYN_IDX.get(canon(h)) == g:
                return h
    # nome composto (ex.: CPF_DEPENDENTE): procura o sinônimo mais longo contido no nome
    best = None  # prioriza o sinônimo no início do nome (CPF_DEPENDENTE → CPF), depois o mais longo
    for alias, gi in SYN_IDX.items():
        if len(alias) >= 3 and alias in c:
            rank = (c.startswith(alias), len(alias))
            if best is None or rank > best[2]:
                best = (alias, gi, rank)
    if best:
        cands = [h for h in headers if SYN_IDX.get(canon(h)) == best[1]]
        if len(cands) == 1:
            return cands[0]
    return ""


def new_field(h: str, base: dict) -> dict:
    tipo = guess_type(h, base["headers"], base["rows"])
    papel = {"NUMCAD": "matricula", "NUMCPF": "cpf", "NOMFUN": "nome", "NUMEMP": "empresa"}.get(syn_name(h), "")
    return {"nome": h, "cols": [], "tipo": tipo, "comparar": True, "obrigatorio": False,
            "tolerancia": 0.01 if tipo == "moeda" else 0.0, "zeros": "remover" if tipo == "codigo" else "manter", "papel": papel}


def auto_map(stages: list[dict], fields: list[dict]) -> list[dict]:
    """stages: [{'headers': [...], 'rows': [...]}]; preserva ajustes manuais existentes."""
    base = stages[0]
    if not base["headers"]:
        return []
    old = {f["nome"]: f for f in fields}
    out = []
    for h in base["headers"]:
        o = old.get(h)
        f = dict(o) if o else new_field(h, base)
        ocols = list(o["cols"]) if o else []
        cols = []
        for si, st in enumerate(stages):
            if si == 0:
                cols.append(h)
            elif si < len(ocols) and ocols[si] and ocols[si] in st["headers"]:
                cols.append(ocols[si])
            else:
                cols.append(find_match(h, st["headers"]))
        f["cols"] = cols
        out.append(f)
    if not any(f.get("papel") == "nome" for f in out):
        for f in out:
            if syn_name(f["nome"]) == "NOMDEP":
                f["papel"] = "nome"
                break
    return out


def _dup_ratio(base: dict | None, fields: list[dict], key: list[str]) -> float:
    if not base or not base.get("rows") or not key:
        return 0.0
    idx = [base["headers"].index(k) for k in key if k in base["headers"]]
    if not idx:
        return 1.0
    seen, dup = set(), 0
    for r in base["rows"]:
        t = tuple(canon(r[i]) if i < len(r) else "" for i in idx)
        if t in seen:
            dup += 1
        seen.add(t)
    return dup / len(base["rows"])


def guess_key(fields: list[dict], base: dict | None = None) -> list[str]:
    """Sugere a chave pelos nomes Senior e completa com colunas até os registros ficarem únicos na origem."""
    key = _guess_key_names(fields)
    if not base or not key or _dup_ratio(base, fields, key) <= 0.01:
        return key
    prio = ["CODDEP", "INIFER", "DATNAS", "DATADM"]
    cands = [f["nome"] for f in fields if f["nome"] not in key]
    cands.sort(key=lambda n: (prio.index(syn_name(n)) if syn_name(n) in prio else len(prio),
                              0 if re.match(r"^(COD|NUM|SEQ|DAT|DATA|TIP|INI)", canon(n)) else 1))
    ratio = _dup_ratio(base, fields, key)
    for _ in range(3):
        best = min(cands, key=lambda n: _dup_ratio(base, fields, key + [n]), default=None)
        if best is None:
            break
        new = _dup_ratio(base, fields, key + [best])
        if new > ratio * 0.5:  # a coluna não resolve a repetição (ex.: linha realmente duplicada)
            break
        key, ratio = key + [best], new
        cands.remove(best)
        if ratio <= 0.01:
            break
    return key


def _guess_key_names(fields: list[dict]) -> list[str]:
    by = {syn_name(f["nome"]): f["nome"] for f in reversed(fields)}
    e, t, c, cpf = by.get("NUMEMP"), by.get("TIPCOL"), by.get("NUMCAD"), by.get("NUMCPF")
    if e and t and c:
        return [e, t, c]
    if c:
        return [e, c] if e else [c]
    if cpf:
        return [cpf]
    return [fields[0]["nome"]] if fields else []


# --------------------------------------------------------------------------- normalização
class Norm:
    __slots__ = ("e", "k", "bad", "num")

    def __init__(self, k="", e=False, bad=False, num=False):
        self.k, self.e, self.bad, self.num = k, e, bad, num


class Normalizer:
    def __init__(self, rules: dict):
        self.r = {**DEFAULT_RULES, **(rules or {})}
        self.nulls = {s.strip().upper() for s in self.r["nulos"].split(",") if s.strip()}

    def is_null(self, v) -> bool:
        if v is None or v is ABSENT or v is NOCOL:
            return True
        if isinstance(v, float) and math.isnan(v):
            return True
        s = str(v).strip()
        return s == "" or s.upper() in self.nulls

    def txt(self, v) -> str:
        s = str(v)
        if self.r["trim"]:
            s = re.sub(r"\s+", " ", s.strip())
        if self.r["acentos"]:
            s = strip_accents(s)
        if self.r["case_ins"]:
            s = s.upper()
        return s

    def norm(self, v, f: dict) -> Norm:
        if self.is_null(v):
            return Norm(e=True)
        t = f["tipo"]
        if t == "data":
            d = parse_date(v)
            if d is None:
                return Norm(str(v), bad=True)
            if self.r["data_nula"] and d == "1900-12-31":
                return Norm(e=True)
            return Norm(d)
        if t in ("numero", "moeda"):
            n = parse_num(v)
            return Norm(str(v), bad=True) if n is None else Norm(n, num=True)
        if t == "cpf":
            d = re.sub(r"\D", "", str(int(v)) if isinstance(v, float) and v.is_integer() else str(v))
            if not d or len(d) > 11:
                return Norm(str(v), bad=True)
            return Norm(d.zfill(11))
        if t == "codigo":
            if isinstance(v, float) and v.is_integer():
                s = str(int(v))
            elif isinstance(v, (int, float)):
                s = str(v)
            else:
                s = self.txt(v)
            s = re.sub(r"^(-?\d+)[.,]0+$", r"\1", s)
            if f.get("zeros") == "remover":
                s = re.sub(r"^0+(?=.)", "", s)
            return Norm(s)
        return Norm(self.txt(v))

    @staticmethod
    def eq(a: Norm, b: Norm, f: dict) -> bool:
        if a.e and b.e:
            return True
        if a.e != b.e:
            return False
        if a.num and b.num:
            return abs(a.k - b.k) <= float(f.get("tolerancia") or 0) + 1e-9
        return a.k == b.k


def parse_num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v) if math.isfinite(v) else None
    s = re.sub(r"R\$|\s", "", str(v).strip())
    if not s:
        return None
    neg = bool(re.match(r"^\(.*\)$", s) or s.endswith("-"))
    s = s.replace("(", "").replace(")", "").rstrip("-")
    lc, ld = s.rfind(","), s.rfind(".")
    if lc >= 0 and ld >= 0:
        s = s.replace(".", "").replace(",", ".") if lc > ld else s.replace(",", "")
    elif lc >= 0:
        s = s.replace(",", ".")
    elif ld >= 0 and re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")
    if not re.fullmatch(r"-?\d*\.?\d+", s):
        return None
    n = float(s)
    return -n if neg else n


def _ymd(y, m, d):
    if y < 100:
        y += 1900 if y > 50 else 2000
    try:
        return date(y, m, d).isoformat()
    except ValueError:
        return None


def parse_date(v):
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return (date(1899, 12, 30) + timedelta(days=round(v))).isoformat() if 0 < v < 200000 else None
    s = str(v).strip()
    m = re.fullmatch(r"(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})(?:\s.*)?", s)
    if m:
        return _ymd(int(m[3]), int(m[2]), int(m[1]))
    m = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?", s)
    if m:
        return _ymd(int(m[1]), int(m[2]), int(m[3]))
    m = re.fullmatch(r"(\d{2})(\d{2})(\d{4})", s)
    if m:
        return _ymd(int(m[3]), int(m[2]), int(m[1]))
    if re.fullmatch(r"\d{4,6}", s):
        return (date(1899, 12, 30) + timedelta(days=int(s))).isoformat()
    return None


def br_date(iso: str) -> str:
    return "/".join(reversed(iso.split("-"))) if iso else ""


def show(v, f=None, nz: Normalizer | None = None) -> str:
    if v is ABSENT:
        return "(ausente)"
    if v is NOCOL:
        return "(não mapeado)"
    if v == "presente":
        return "presente"
    if (nz and nz.is_null(v)) or v is None or str(v).strip() == "":
        return "(vazio)"
    if f and f["tipo"] == "data" and (isinstance(v, (datetime, date, int, float)) and not isinstance(v, bool)):
        d = parse_date(v)
        if d:
            return br_date(d)
    if isinstance(v, datetime):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


# --------------------------------------------------------------------------- comparação
def classify(f, I, S, A, S2, nm, eq):
    tipo = TIPOS[f["tipo"]]
    if I.bad or S.bad or (A and A.bad) or (S2 and S2.bad):
        return "MANUAL", f"Valor em formato não reconhecido para o tipo {tipo}.", nm["s"]
    if eq(I, S, f):
        if I.e and f.get("obrigatorio"):
            return "NAO_PREENCHIDO", "Campo obrigatório sem valor na origem e no Senior.", nm["s"]
        if not A:
            return "OK", ("Campo vazio na origem e no Senior." if I.e else "Valor mantido do Layout Inicial para o Senior."), ""
        if eq(A, S, f):
            if S2 and not eq(S2, A, f):
                return "DIVERGENCIA", "Valor alterado na carga posterior ao ajuste.", nm["s2"]
            return "OK", "Valor mantido em todas as etapas.", ""
        return "MANUAL", "O layout ajustado alterou um valor que já estava aderente.", nm["a"]
    if A and not A.e and not eq(A, S, f):
        if S2:
            if eq(S2, A, f):
                return "CORRIGIDO", "Senior divergiu da origem; valor corrigido no layout ajustado e confirmado na carga seguinte.", nm["s"]
            return "DIVERGENCIA", "Correção do layout ajustado não foi refletida na carga seguinte.", nm["s2"]
        if I.e:
            return "CORRIGIDO", "Origem vazia; valor definido no layout ajustado.", nm["s"]
        if eq(A, I, f):
            return "CORRIGIDO", "Senior apresentou valor divergente da origem e precisou ser corrigido.", nm["s"]
        return "CORRIGIDO", "Senior divergiu da origem; o layout ajustado definiu um novo valor.", nm["s"]
    if A and not S.e and eq(A, S, f):
        return "ALTERADO", ("Informação complementada durante a migração." if I.e else "Valor alterado no Senior e confirmado no layout ajustado."), nm["s"]
    if I.e and not S.e:
        return "ALTERADO", "Informação complementada durante a migração.", nm["s"]
    if S.e:
        return "NAO_PREENCHIDO", "Campo preenchido na origem e vazio no Senior.", nm["s"]
    return "DIVERGENCIA", "Valor do Senior difere da origem, sem correção no layout ajustado.", nm["s"]


def validate(stages: list[dict], fields: list[dict], key: list[str]) -> list[str]:
    e = []
    if not stages or not stages[0]["rows"]:
        e.append("Carregue o arquivo do Layout Inicial (primeira etapa).")
    if not any(s["tipo"] == "senior" and s["rows"] for s in stages):
        e.append("Carregue ao menos um arquivo extraído do Senior.")
    if not key:
        e.append("Defina a chave de comparação (uma ou mais colunas).")
    fmap = {f["nome"]: f for f in fields}
    for si, s in enumerate(stages):
        if not s["rows"]:
            continue
        for k in key:
            f = fmap.get(k)
            if f and (si >= len(f["cols"]) or not f["cols"][si]):
                e.append(f"O campo-chave {k} não está mapeado em “{s['nome']}”. Ajuste o DE/PARA.")
    return e


def run_comparison(stages: list[dict], fields: list[dict], key: list[str], rules: dict, meta: dict) -> dict:
    nz = Normalizer(rules)
    eq = nz.eq
    loaded = [bool(s["rows"]) for s in stages]
    idx = {f["nome"]: [s["headers"].index(f["cols"][si]) if si < len(f["cols"]) and f["cols"][si] in s["headers"] else -1
                       for si, s in enumerate(stages)] for f in fields}
    fmap = {f["nome"]: f for f in fields}
    key_f = [fmap[k] for k in key]
    cmp_f = [f for f in fields if f.get("comparar") and f["nome"] not in key]
    role = {r: next((f for f in fields if f.get("papel") == r), None) for r in ("nome", "matricula", "cpf", "empresa")}

    def val(si, ri, f):
        i = idx[f["nome"]][si]
        if i < 0:
            return NOCOL
        row = stages[si]["rows"][ri]
        return row[i] if i < len(row) else ""

    maps = []
    for si, s in enumerate(stages):
        m, dup, bad = {}, {}, []
        if loaded[si]:
            for ri in range(len(s["rows"])):
                parts = [nz.norm(val(si, ri, f), f) for f in key_f]
                if any(p.e or p.bad for p in parts):
                    bad.append(ri)
                    continue
                k = " · ".join(str(int(p.k) if isinstance(p.k, float) and p.k.is_integer() else p.k) for p in parts)
                if k in m:
                    dup[k] = dup.get(k, 1) + 1
                else:
                    m[k] = ri
        maps.append({"m": m, "dup": dup, "bad": bad})

    keys, seen = [], set()
    for x in maps:
        for k in x["m"]:
            if k not in seen:
                seen.add(k)
                keys.append(k)
    sen = [i for i, s in enumerate(stages) if s["tipo"] == "senior" and loaded[i]]
    aju = [i for i, s in enumerate(stages) if s["tipo"] == "ajuste" and loaded[i]]
    when = br_date(meta.get("data") or date.today().isoformat())
    names = [s["nome"] for s in stages]

    def role_val(rec, r):
        f = role[r]
        if not f:
            return ""
        for si in range(len(stages)):
            if rec["present"][si]:
                v = val(si, rec["rows"][si], f)
                if v is not NOCOL and not nz.is_null(v):
                    return show(v, f)
        return ""

    records = []
    for k in keys:
        present = [k in x["m"] for x in maps]
        rows = [x["m"].get(k) for x in maps]
        rec = {"key": k, "present": present, "rows": rows, "fields": [], "notes": [], "trail": [], "level": None, "late": False, "etapa": ""}
        for r in role:
            rec[r] = role_val(rec, r)
        dup_at = [i for i, x in enumerate(maps) if k in x["dup"]]
        if dup_at:
            rec["notes"].append("Chave repetida em: " + ", ".join(f"{names[i]} ({maps[i]['dup'][k]}×)" for i in dup_at)
                                + ". Comparação feita com a primeira ocorrência.")
        in_i = present[0]
        s_f = next((i for i in sen if present[i]), None)
        if not in_i and s_f is None:
            rec["level"] = "CHAVE_NL"
            rec["notes"].append("Chave encontrada somente em: " + ", ".join(n for n, p in zip(names, present) if p) + ".")
            rec["etapa"] = names[present.index(True)]
        elif not in_i:
            rec["level"] = "SOMENTE_SENIOR"
            rec["notes"].append(f"Registro inexistente no Layout Inicial e encontrado em {names[s_f]}.")
            rec["etapa"] = names[s_f]
        elif s_f is None:
            rec["level"] = "NAO_MIGRADO"
            rec["notes"].append("Registro existente no Layout Inicial e não localizado em nenhuma extração do Senior.")
            rec["etapa"] = names[sen[0]]
        else:
            late = s_f != sen[0]
            a_list = [i for i in aju if i > s_f and present[i]]
            a_i = a_list[-1] if a_list else None
            s2_list = [i for i in sen if a_i is not None and i > a_i and present[i]]
            s2 = s2_list[-1] if s2_list else None
            nm = {"s": names[s_f], "a": names[a_i] if a_i is not None else "", "s2": names[s2] if s2 is not None else ""}
            for f in cmp_f:
                raws = [ABSENT if not present[si] else val(si, rows[si], f) for si in range(len(stages))]
                if raws[s_f] is NOCOL:
                    code, nota, etapa = "MANUAL", f"Campo não mapeado no arquivo “{names[s_f]}”.", names[s_f]
                else:
                    A = nz.norm(raws[a_i], f) if a_i is not None and raws[a_i] is not NOCOL else None
                    S2 = nz.norm(raws[s2], f) if s2 is not None and raws[s2] is not NOCOL else None
                    code, nota, etapa = classify(f, nz.norm(raws[0], f), nz.norm(raws[s_f], f), A, S2, nm, eq)
                rec["fields"].append({"f": f, "raws": raws, "code": code, "nota": nota, "etapa": etapa})
                prev = prev_raw = None
                for si, rv in enumerate(raws):
                    if rv is ABSENT or rv is NOCOL:
                        continue
                    cur = nz.norm(rv, f)
                    if prev is not None and not eq(prev, cur, f):
                        rec["trail"].append({"campo": f["nome"], "de": show(prev_raw, f, nz), "para": show(rv, f, nz),
                                             "etapa": names[si], "arquivo": stages[si].get("arquivos", ""), "data": when})
                    prev, prev_raw = cur, rv
            if late:
                rec["late"] = True
                rec["notes"].append(f"Registro não migrado em “{names[sen[0]]}” e incluído em “{names[s_f]}”.")
        code = rec["level"] or "OK"
        if not rec["level"]:
            for x in rec["fields"]:
                if STATUS[x["code"]]["sev"] > STATUS[code]["sev"]:
                    code = x["code"]
        if dup_at and STATUS["DUPLICADO"]["sev"] > STATUS[code]["sev"]:
            code = "DUPLICADO"
        if rec["late"] and code == "OK":
            code = "CORRIGIDO"
        rec["status"] = code
        records.append(rec)

    for si, x in enumerate(maps):
        for ri in x["bad"]:
            desc = ", ".join(f"{f['nome']}={show(val(si, ri, f), f, nz)}" for f in key_f)
            present = [i == si for i in range(len(stages))]
            rec = {"key": f"(chave incompleta) {names[si]} · registro {ri + 1}", "present": present,
                   "rows": [ri if i == si else None for i in range(len(stages))], "fields": [], "trail": [],
                   "level": "CHAVE_NL", "status": "CHAVE_NL", "etapa": names[si], "late": False,
                   "notes": [f"Chave de comparação vazia ou inválida ({desc})."]}
            for r in role:
                rec[r] = role_val(rec, r)
            records.append(rec)

    lines = []
    for rec in records:
        if rec["level"] or rec["late"] or rec["status"] == "DUPLICADO":
            lines.append({"rec": rec, "f": None, "raws": ["presente" if p else ABSENT for p in rec["present"]],
                          "code": rec["level"] or ("CORRIGIDO" if rec["late"] else "DUPLICADO"),
                          "nota": " ".join(rec["notes"]), "etapa": rec["etapa"]})
        for x in rec["fields"]:
            lines.append({"rec": rec, "f": x["f"], "raws": x["raws"], "code": x["code"], "nota": x["nota"], "etapa": x["etapa"] or ""})

    both = [r for r in records if not r["level"]]
    has = lambda r, codes: any(x["code"] in codes for x in r["fields"])  # noqa: E731
    cnt = lambda c: sum(1 for r in records if r["status"] == c)  # noqa: E731
    field_stats = []
    for f in cmp_f:
        reg = div = 0
        for r in both:
            x = next((y for y in r["fields"] if y["f"] is f), None)
            if not x:
                continue
            reg += 1
            if x["code"] != "OK":
                div += 1
        field_stats.append({"campo": f["nome"], "registros": reg, "divergencias": div, "pct": (reg - div) / reg * 100 if reg else 100.0})
    dist: dict[str, int] = {}
    for r in both:
        for x in r["fields"]:
            dist[x["code"]] = dist.get(x["code"], 0) + 1
    fl = [l for l in lines if l["f"] is not None]
    k = {
        "total": len(records), "corretos": cnt("OK"),
        "divergentes": sum(1 for r in both if any(x["code"] != "OK" for x in r["fields"])),
        "nao_migrados": cnt("NAO_MIGRADO"), "somente_senior": cnt("SOMENTE_SENIOR"), "chave_nl": cnt("CHAVE_NL"),
        "duplicados": sum(1 for r in records if any(n.startswith("Chave repetida") for n in r["notes"])),
        "alterados": sum(1 for r in both if has(r, ["ALTERADO"])),
        "corrigidos": sum(1 for r in both if has(r, ["CORRIGIDO"]) or r["late"]),
        "manual": sum(1 for r in records if has(r, ["MANUAL"]) or r["status"] in ("DUPLICADO", "CHAVE_NL")),
        "pendentes": sum(1 for r in both if has(r, ["DIVERGENCIA", "NAO_PREENCHIDO"])),
        "div_campos": sum(1 for l in fl if l["code"] != "OK"),
        "corr_campos": sum(1 for l in fl if l["code"] == "CORRIGIDO"),
        "alt_campos": sum(1 for l in fl if l["code"] == "ALTERADO"),
        "pend_campos": sum(1 for l in fl if l["code"] in ("DIVERGENCIA", "NAO_PREENCHIDO", "MANUAL")),
    }
    k["aderencia"] = k["corretos"] / k["total"] * 100 if k["total"] else 0.0
    res = {"records": records, "lines": lines, "k": k, "field_stats": field_stats, "dist": dist, "key": list(key),
           "cmp_f": cmp_f, "stages": [{"nome": s["nome"], "tipo": s["tipo"], "arquivos": s.get("arquivos", ""), "n": len(s["rows"])} for s in stages],
           "meta": dict(meta), "at": datetime.now(), "nz": nz}
    res["summary"] = build_summary(res)
    return res


def fmt_int(n) -> str:
    return f"{int(n):,}".replace(",", ".")


def fmt_pct(n) -> str:
    return f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + "%"


def _list(a):
    return "".join(a) if len(a) <= 1 else ", ".join(a[:-1]) + " e " + a[-1]


def build_summary(R) -> str:
    k = R["k"]
    top = [x["campo"] for x in sorted([x for x in R["field_stats"] if x["divergencias"]], key=lambda x: -x["divergencias"])[:3]]
    p = [f"Foram analisados {fmt_int(k['total'])} registros, comparados pela chave {' + '.join(R['key'])}.",
         f"{fmt_int(k['corretos'])} registros apresentaram aderência integral entre o Layout Inicial e o Senior, "
         f"o que representa {fmt_pct(k['aderencia'])} de aderência da migração."]
    if k["div_campos"]:
        conc = f", concentradas principalmente {'nos campos' if len(top) > 1 else 'no campo'} {_list(top)}" if top else ""
        p.append(f"Foram encontradas {fmt_int(k['div_campos'])} divergências de campo em {fmt_int(k['divergentes'])} registros{conc}.")
        alt = f", {fmt_int(k['alt_campos'])} correspondem a valores alterados ou complementados no Senior" if k["alt_campos"] else ""
        p.append(f"Dessas divergências, {fmt_int(k['corr_campos'])} foram corrigidas no Layout Ajustado{alt} e "
                 f"{fmt_int(k['pend_campos'])} permanecem pendentes de validação.")
    else:
        p.append("Não foram encontradas divergências de campo entre os registros migrados.")
    extra = []
    if k["nao_migrados"]:
        extra.append(f"{fmt_int(k['nao_migrados'])} " + ("registros não foram migrados" if k["nao_migrados"] > 1 else "registro não foi migrado"))
    if k["somente_senior"]:
        extra.append(f"{fmt_int(k['somente_senior'])} " + ("existem" if k["somente_senior"] > 1 else "existe") + " somente no Senior")
    if k["duplicados"]:
        extra.append(f"{fmt_int(k['duplicados'])} " + ("chaves aparecem duplicadas" if k["duplicados"] > 1 else "chave aparece duplicada"))
    if k["chave_nl"]:
        extra.append(f"{fmt_int(k['chave_nl'])} " + ("registros têm" if k["chave_nl"] > 1 else "registro tem") + " chave não localizada")
    if extra:
        p.append("Além disso, " + _list(extra) + ".")
    return " ".join(p)


# --------------------------------------------------------------------------- tabelas e Excel
def lines_frame(R, lines, with_layout=False):
    import pandas as pd
    names = [s["nome"] for s in R["stages"]]
    cols = (["Layout"] if with_layout else []) + ["Chave", "Colaborador", "Matrícula", "CPF", "Empresa", "Campo"] + names + ["Resultado", "Etapa", "Análise"]
    data = []
    nz = R["nz"]
    for l in lines:
        rec = l["rec"]
        data.append(([R["meta"].get("layout", "")] if with_layout else []) + [rec["key"], rec.get("nome", ""), rec.get("matricula", ""), rec.get("cpf", ""), rec.get("empresa", ""),
                     l["f"]["nome"] if l["f"] else "(registro)", *[show(v, l["f"], nz) for v in l["raws"]],
                     STATUS[l["code"]]["label"], l["etapa"], l["nota"]])
    # nomes de etapa repetidos não podem virar colunas duplicadas
    seen: dict[str, int] = {}
    uniq = []
    for c in cols:
        if c in seen:
            seen[c] += 1
            uniq.append(f"{c} ({seen[c]})")
        else:
            seen[c] = 1
            uniq.append(c)
    return pd.DataFrame(data, columns=uniq)


def overall(results: list[dict]) -> dict:
    """Indicadores somados de vários layouts."""
    keys = ["total", "corretos", "divergentes", "nao_migrados", "somente_senior", "chave_nl", "duplicados", "alterados",
            "corrigidos", "manual", "pendentes", "div_campos", "corr_campos", "alt_campos", "pend_campos"]
    k = {x: sum(R["k"][x] for R in results) for x in keys}
    k["aderencia"] = k["corretos"] / k["total"] * 100 if k["total"] else 0.0
    k["layouts"] = len(results)
    return k


def top_fields(R, n=3) -> list[str]:
    return [x["campo"] for x in sorted([x for x in R["field_stats"] if x["divergencias"]], key=lambda x: -x["divergencias"])[:n]]


def layout_status(R) -> str:
    k = R["k"]
    if not k["div_campos"] and not k["nao_migrados"] and not k["somente_senior"] and not k["manual"]:
        return "Sem divergências"
    if k["pend_campos"] or k["nao_migrados"] or k["manual"]:
        return "Com pendências"
    return "Divergências corrigidas"


def divergence_rows(results: list[dict]) -> list[dict]:
    """Uma linha por layout com o que precisa de atenção."""
    return [{"Layout": R["meta"].get("layout", ""), "Situação": layout_status(R),
             "Divergências de campo": R["k"]["div_campos"], "Pendentes": R["k"]["pend_campos"],
             "Corrigidas": R["k"]["corr_campos"], "Não migrados": R["k"]["nao_migrados"],
             "Somente Senior": R["k"]["somente_senior"], "Validação manual": R["k"]["manual"],
             "Campos mais afetados": ", ".join(top_fields(R)) or "—",
             "Aderência (%)": round(R["k"]["aderencia"], 2)} for R in results]


def overview_rows(results: list[dict]) -> list[dict]:
    return [{"Layout": R["meta"].get("layout", ""), "Chave": " + ".join(R["key"]), "Registros": R["k"]["total"],
             "Corretos": R["k"]["corretos"], "Com divergência": R["k"]["divergentes"], "Não migrados": R["k"]["nao_migrados"],
             "Somente Senior": R["k"]["somente_senior"], "Corrigidos": R["k"]["corrigidos"], "Validação manual": R["k"]["manual"],
             "Divergências de campo": R["k"]["div_campos"], "Aderência (%)": round(R["k"]["aderencia"], 2)} for R in results]


def build_summary_multi(results: list[dict]) -> str:
    if len(results) == 1:
        return results[0]["summary"]
    k = overall(results)
    nomes = [R["meta"].get("layout") or "sem nome" for R in results]
    p = [f"Foram validados {len(results)} layouts ({_list(nomes)}), com {fmt_int(k['total'])} registros no total e "
         f"aderência geral de {fmt_pct(k['aderencia'])}."]
    pior = min(results, key=lambda R: R["k"]["aderencia"])
    melhor = max(results, key=lambda R: R["k"]["aderencia"])
    if pior is not melhor:
        p.append(f"A menor aderência está em {pior['meta'].get('layout')} ({fmt_pct(pior['k']['aderencia'])}) e a maior em "
                 f"{melhor['meta'].get('layout')} ({fmt_pct(melhor['k']['aderencia'])}).")
    if k["div_campos"]:
        p.append(f"Ao todo, foram encontradas {fmt_int(k['div_campos'])} divergências de campo: {fmt_int(k['corr_campos'])} corrigidas no "
                 f"Layout Ajustado e {fmt_int(k['pend_campos'])} pendentes de validação.")
    if k["nao_migrados"]:
        p.append(f"{fmt_int(k['nao_migrados'])} " + ("registros não foram migrados." if k["nao_migrados"] > 1 else "registro não foi migrado."))
    partes = []
    for R in results:
        kk = R["k"]
        if not (kk["div_campos"] or kk["nao_migrados"] or kk["somente_senior"]):
            partes.append(f"{R['meta'].get('layout')}: sem divergências")
            continue
        itens = []
        if kk["div_campos"]:
            tf = top_fields(R)
            itens.append(f"{fmt_int(kk['div_campos'])} divergências de campo" + (f" ({', '.join(tf)})" if tf else ""))
        if kk["nao_migrados"]:
            itens.append(f"{fmt_int(kk['nao_migrados'])} não migrado(s)")
        if kk["somente_senior"]:
            itens.append(f"{fmt_int(kk['somente_senior'])} somente no Senior")
        partes.append(f"{R['meta'].get('layout')}: " + ", ".join(itens))
    p.append("Por layout — " + "; ".join(partes) + ".")
    return " ".join(p)


def export_excel(R, stages: list[dict]) -> bytes:
    return export_excel_multi([(R, stages)])


def export_excel_multi(items: list[tuple[dict, list[dict]]]) -> bytes:
    """items: [(resultado, etapas do layout)] — um único Excel para todos os layouts."""
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    results = [R for R, _ in items]
    multi = len(results) > 1
    wb = Workbook()
    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="12804A")
    band_fill = PatternFill("solid", fgColor="E2F2E9")

    def style_head(ws, row):
        for c in ws[row]:
            if c.value is not None:
                c.font, c.fill = head_font, head_fill

    def sheet(title, df, widths=None):
        ws = wb.create_sheet(title)
        ws.append(list(df.columns))
        style_head(ws, 1)
        res_col = list(df.columns).index("Resultado") + 1 if "Resultado" in df.columns else None
        for row in df.itertuples(index=False):
            ws.append([("" if v is None else v) for v in row])
            if res_col:
                cell = ws.cell(row=ws.max_row, column=res_col)
                tone = STATUS[LABEL_TO_CODE.get(cell.value, "OK")]["tone"]
                cell.fill = PatternFill("solid", fgColor=TONE_BG[tone][1:])
                cell.font = Font(bold=True, color=TONE_FG[tone][1:])
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for i, col in enumerate(df.columns, start=1):
            w = (widths or {}).get(col) or min(60, max(10, len(str(col)) + 2, *(len(str(v)) + 2 for v in df[col].head(200))))
            ws.column_dimensions[get_column_letter(i)].width = w
        return ws

    m = results[0]["meta"]
    ws = wb.active
    ws.title = "RESUMO"
    ws.append(["Validador Pós-Migração — Agante Tecnologia | Senior HCM"])
    ws["A1"].font = Font(bold=True, size=14, color="12804A")
    ws.append([])
    for lbl, v in [("Cliente", m.get("cliente", "")), ("Projeto", m.get("projeto", "")), ("Módulo", m.get("modulo", "")),
                   ("Layouts", ", ".join(R["meta"].get("layout", "") for R in results)), ("Data da validação", br_date(m.get("data", ""))),
                   ("Responsável", m.get("responsavel", "")), ("Processado em", results[0]["at"].strftime("%d/%m/%Y %H:%M"))]:
        ws.append([lbl, v])
    ws.append([])
    ws.append(["Divergências por layout"])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True, size=12)
    dr = divergence_rows(results)
    ws.append(list(dr[0].keys()))
    style_head(ws, ws.max_row)
    for r in dr:
        ws.append(list(r.values()))
        sit = ws.cell(row=ws.max_row, column=2)
        tone = {"Sem divergências": "ok", "Divergências corrigidas": "fix"}.get(sit.value, "bad")
        sit.fill = PatternFill("solid", fgColor=TONE_BG[tone][1:])
        sit.font = Font(bold=True, color=TONE_FG[tone][1:])
    if multi:
        k = overall(results)
        ws.append([])
        ws.append(["Visão geral"])
        ws.cell(row=ws.max_row, column=1).font = Font(bold=True, size=12)
        ov = overview_rows(results)
        ws.append(list(ov[0].keys()))
        style_head(ws, ws.max_row)
        for r in ov:
            ws.append(list(r.values()))
        ws.append(["TOTAL", "", k["total"], k["corretos"], k["divergentes"], k["nao_migrados"], k["somente_senior"], k["corrigidos"],
                   k["manual"], k["div_campos"], round(k["aderencia"], 2)])
        for c in ws[ws.max_row]:
            c.font = Font(bold=True)
        ws.append([])
        ws.append(["Resumo geral"])
        ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
        ws.append([build_summary_multi(results)])
        ws.cell(row=ws.max_row, column=1).alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=8)
        ws.row_dimensions[ws.max_row].height = 75
    for R in results:
        k = R["k"]
        ws.append([])
        ws.append([f"Layout: {R['meta'].get('layout') or 'sem nome'}"])
        for c in ws[ws.max_row]:
            c.font, c.fill = Font(bold=True, size=12, color="12804A"), band_fill
        rows = [["Chave de comparação", " + ".join(R["key"])], [],
                ["Etapa", "Tipo", "Arquivo(s)", "Registros"],
                *[[s["nome"], TIPO_ETAPA[s["tipo"]], s["arquivos"] or "(não carregado)", s["n"]] for s in R["stages"]], [],
                ["Indicador", "Valor"],
                ["Total de registros analisados", k["total"]], ["Registros corretos", k["corretos"]],
                ["Registros com divergência", k["divergentes"]], ["Registros não migrados", k["nao_migrados"]],
                ["Registros criados somente no Senior", k["somente_senior"]], ["Registros alterados", k["alterados"]],
                ["Registros corrigidos", k["corrigidos"]], ["Registros que necessitam validação manual", k["manual"]],
                ["Divergências de campo", k["div_campos"]], ["Aderência da migração (%)", round(k["aderencia"], 2)], [],
                ["Campo", "Registros", "Divergências", "% Correto"],
                *[[x["campo"], x["registros"], x["divergencias"], round(x["pct"], 2)] for x in R["field_stats"]], [],
                ["Resumo da validação"], [R["summary"]]]
        for r in rows:
            ws.append(r)
            if r and r[0] in ("Etapa", "Indicador", "Campo", "Resumo da validação"):
                for c in ws[ws.max_row]:
                    if c.value is not None:
                        c.font = Font(bold=True)
        ws.cell(row=ws.max_row, column=1).alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=8)
        ws.row_dimensions[ws.max_row].height = 95
    ws.column_dimensions["A"].width = 42
    for col in "BCDEFGHIJK":
        ws.column_dimensions[col].width = 18 if multi else 30

    w = {"Análise": 60, "Resultado": 30}

    def lines_all(pred):
        frames = [lines_frame(R, [l for l in R["lines"] if pred(l)], with_layout=True) for R in results]
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

    sheet("COMPARATIVO COMPLETO", lines_all(lambda l: True), w)
    sheet("DIVERGÊNCIAS", lines_all(lambda l: l["code"] != "OK"), w)

    def rec_sheet(title, code, pick_stage):
        ws2 = wb.create_sheet(title)
        first = True
        for R, stages in items:
            si = pick_stage(stages)
            recs = [r for r in R["records"] if si is not None and r["status"] == code and r["present"][si]]
            if not recs:
                continue
            if not first:
                ws2.append([])
            ws2.append([f"Layout: {R['meta'].get('layout')}"])
            for c in ws2[ws2.max_row]:
                c.font, c.fill = Font(bold=True, color="12804A"), band_fill
            first = False
            st_ = stages[si]
            ws2.append(["Layout", "Chave", *st_["headers"]])
            style_head(ws2, ws2.max_row)
            for r in recs:
                ws2.append([R["meta"].get("layout", ""), r["key"], *[show(v) if v is not None else "" for v in st_["rows"][r["rows"][si]]]])
        if first:
            ws2.append(["Nenhum registro."])
        ws2.column_dimensions["A"].width = 20
        ws2.column_dimensions["B"].width = 24
        for i in range(3, 40):
            ws2.column_dimensions[get_column_letter(i)].width = 16
        return ws2

    rec_sheet("NÃO MIGRADOS", "NAO_MIGRADO", lambda st_: 0)
    rec_sheet("SOMENTE SENIOR", "SOMENTE_SENIOR",
              lambda st_: next((i for i, x in enumerate(st_) if x["tipo"] == "senior" and x["rows"]), None))
    sheet("CORRIGIDOS", lines_all(lambda l: l["code"] == "CORRIGIDO"), w)
    sheet("VALIDAÇÃO MANUAL", lines_all(lambda l: l["code"] in ("MANUAL", "DUPLICADO", "CHAVE_NL")), w)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# --------------------------------------------------------------------------- exemplo
def make_example():
    """Três arquivos fictícios (matrizes) para demonstração."""
    import random
    rnd = random.Random(7)
    nomes = ["Ana Beatriz Souza", "Bruno Henrique Lima", "Camila Rodrigues", "Diego Martins", "Eduarda Ferreira", "Fábio Nogueira",
             "Gabriela Ribeiro", "Henrique Alves", "Isabela Carvalho", "João da Silva", "Juliana Barbosa", "Lucas Pereira",
             "Mariana Gonçalves", "Nicolas Araújo", "Otávio Mendes", "Patrícia Cardoso", "Rafael Teixeira", "Renata Moreira",
             "Sérgio Batista", "Tatiane Correia", "Vinícius Rocha", "Yasmin Freitas", "André Monteiro", "Beatriz Castro",
             "Caio Fernandes", "Débora Pinto", "Elias Moura", "Fernanda Lopes", "Gustavo Ramos", "Helena Duarte", "Igor Santana",
             "Jéssica Vieira", "Leonardo Cunha", "Luana Machado", "Marcelo Dias", "Natália Campos", "Paulo Azevedo",
             "Priscila Nunes", "Rodrigo Farias", "Simone Reis", "Thiago Melo", "Valéria Prado", "Wagner Cruz", "Aline Tavares",
             "Bruno Costa", "Carla Siqueira", "Daniel Brito", "Érica Pacheco"]
    brl = lambda n: f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")  # noqa: E731
    base = []
    for i, n in enumerate(nomes):
        base.append({"cad": str(1001 + i), "nome": n, "cpf": "".join(str(rnd.randint(0, 9)) for _ in range(11)),
                     "y": 2012 + rnd.randint(0, 11), "m": rnd.randint(1, 12), "d": rnd.randint(1, 28),
                     "car": rnd.choice(["1001", "1002", "1003", "2001", "2002", "3001"]), "sin": rnd.choice(["10", "12", "15"]),
                     "cc": rnd.choice(["1.01.001", "1.01.002", "1.02.001", "2.01.003"]),
                     "sal": rnd.choice([2850, 3400, 4200, 5000, 6100, 7800, 9500, 12300]), "uf": "SP" if rnd.random() < .8 else rnd.choice(["MG", "PR", "RJ"])})
    base[3]["car"], base[11]["car"], base[21]["car"], base[27]["car"] = "1001", "2001", "1001", "2002"
    h0 = ["COD_EMPRESA", "TIPO_COLABORADOR", "MATRICULA", "NOME", "CPF_COLABORADOR", "DATA_ADMISSAO", "CARGO", "SINDICATO", "CENTRO_CUSTO", "SALARIO", "UF"]
    h1 = ["NumEmp", "TipCol", "NumCad", "NomFun", "NumCpf", "DatAdm", "CodCar", "CodSin", "CodCcu", "ValSal", "SigUfs"]
    mask = lambda c: f"{c[:3]}.{c[3:6]}.{c[6:9]}-{c[9:]}"  # noqa: E731

    def init_row(b, i):
        return ["1", "1", b["cad"], b["nome"], mask(b["cpf"]), "31/02/2019" if i == 44 else f"{b['d']:02d}/{b['m']:02d}/{b['y']}",
                b["car"], b["sin"], b["cc"], brl(b["sal"]), "" if i in (12, 36) else b["uf"]]

    ini = [h0] + [init_row(b, i) for i, b in enumerate(base)] + [init_row(base[8], 8)]
    sen = [h1]
    for i, b in enumerate(base):
        if i in (5, 17, 33):
            continue
        car, sin, cc, sal, d, m, y = b["car"], "0" + b["sin"], b["cc"], b["sal"], b["d"], b["m"], b["y"]
        car = {3: "1002", 11: "2002", 21: "1003", 27: "3001"}.get(i, car)
        if i in (6, 14, 40):
            sin = ""
        cc = {9: "1.01.009", 30: "2.01.004"}.get(i, cc)
        if i == 19:
            sal -= 0.004
        if i == 24:
            sal += 100
        if i == 15:
            d += 1
        if i == 44:
            y, m, d = 2019, 2, 28
        sen.append(["1", "1", b["cad"], strip_accents(b["nome"]).upper(), b["cpf"], f"{y}-{m:02d}-{d:02d}", car, sin, cc, f"{sal:.2f}", b["uf"]])
    sen.append(["1", "1", "1990", "MARCOS VIEIRA LIMA", "40512378901", "2024-08-05", "1002", "012", "1.01.002", "3400.00", "SP"])
    aj = [h0] + [init_row(base[i], i) for i in (3, 11, 6, 14, 24, 15)]
    r21 = init_row(base[21], 21)
    r21[6] = "1003"
    aj.append(r21)
    return [("exemplo_layout_inicial.xlsx", ini), ("exemplo_extracao_senior.csv", sen), ("exemplo_layout_ajustado.xlsx", aj)]



def make_example_multi():
    """Exemplo com dois layouts no mesmo lote: Colaboradores e 1031 Dependentes."""
    import random
    col = make_example()
    ini_c, sen_c, aj_c = col[0][1], col[1][1], col[2][1]
    rnd = random.Random(11)
    nomes = ["Lucas", "Sofia", "Miguel", "Helena", "Arthur", "Laura", "Davi", "Alice", "Gael", "Valentina", "Theo", "Cecília",
             "Heitor", "Maria Clara", "Bernardo", "Isadora", "Samuel", "Lorena", "Pedro", "Manuela", "Gabriel", "Luiza", "Rafael", "Beatriz"]
    deps = []
    for i, r in enumerate(ini_c[1:41:2]):
        sobrenome = r[3].split()[-1]
        for j in range(1 + (i % 3 == 0)):
            nm = f"{nomes[(i * 2 + j) % len(nomes)]} {sobrenome}"
            deps.append({"emp": r[0], "tip": r[1], "cad": r[2], "dep": str(j + 1), "nome": nm,
                         "par": rnd.choice(["1", "1", "4", "2"]), "nas": f"{rnd.randint(1, 28):02d}/{rnd.randint(1, 12):02d}/{rnd.randint(2008, 2022)}",
                         "cpf": "".join(str(rnd.randint(0, 9)) for _ in range(11))})
    h0 = ["COD_EMPRESA", "TIPO_COLABORADOR", "MATRICULA", "COD_DEPENDENTE", "NOME_DEPENDENTE", "GRAU_PARENTESCO", "DATA_NASCIMENTO", "CPF_DEPENDENTE"]
    h1 = ["NumEmp", "TipCol", "NumCad", "CodDep", "NomDep", "GraPar", "DatNas", "NumCpf"]
    ini_d = [h0] + [[d["emp"], d["tip"], d["cad"], d["dep"], d["nome"], d["par"], d["nas"], d["cpf"]] for d in deps]
    sen_d = [h1]
    for i, d in enumerate(deps):
        if i in (4, 17):
            continue
        par = "3" if i == 7 else d["par"]
        dd, mm, yy = d["nas"].split("/")
        nas = f"{yy}-{mm}-{'01' if i == 11 else dd}"
        sen_d.append([d["emp"], d["tip"], d["cad"], d["dep"], strip_accents(d["nome"]).upper(), par, nas, "" if i == 14 else d["cpf"]])
    aj_d = [h0, [deps[7]["emp"], deps[7]["tip"], deps[7]["cad"], deps[7]["dep"], deps[7]["nome"], deps[7]["par"], deps[7]["nas"], deps[7]["cpf"]]]
    return {
        "origem": [("colaboradores_layout_inicial.xlsx", ini_c), ("1031_dependentes_layout_inicial.xlsx", ini_d)],
        "senior": [("colaboradores_extracao_senior.csv", sen_c), ("1031_dependentes_extracao_senior.csv", sen_d)],
        "ajuste": [("colaboradores_layout_ajustado.xlsx", aj_c), ("1031_dependentes_layout_ajustado.xlsx", aj_d)],
    }
