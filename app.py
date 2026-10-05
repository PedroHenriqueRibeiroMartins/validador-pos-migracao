"""Validador Pós-Migração | Agante Tecnologia · Senior HCM — versão Streamlit.

Executar:  streamlit run app.py
"""
from __future__ import annotations

import hashlib
import json
import os
import uuid
from datetime import date

import pandas as pd
import streamlit as st

import engine as E

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(APP_DIR, "configs", "layouts.json")
ASSETS = os.path.join(APP_DIR, "assets")

st.set_page_config(page_title="Validador Pós-Migração | Agante", page_icon="✅", layout="wide")

# --------------------------------------------------------------------------- estilo
st.markdown("""
<style>
.block-container{padding-top:1.4rem;max-width:1400px}
.vpm-head{display:flex;align-items:center;gap:16px;flex-wrap:wrap;padding-bottom:12px;border-bottom:1px solid #dce3df;margin-bottom:8px}
.vpm-wm{font-weight:700;letter-spacing:.06em;font-size:13px;padding:6px 10px;border:1.5px solid currentColor;border-radius:4px;line-height:1}
.vpm-ag{color:#12804a}.vpm-sr{color:#0c9399}
.vpm-sep{width:1px;height:30px;background:#dce3df}
.vpm-title h1{font-size:1.35rem;margin:0;padding:0;line-height:1.2}
.vpm-title p{margin:0;color:#4a5952;font-size:.85rem}
.vpm-eyebrow{font-size:11px;text-transform:uppercase;letter-spacing:.09em;color:#73817a;font-weight:600}
.vpm-flow{display:flex;gap:6px;align-items:center;flex-wrap:wrap;font-size:13px}
.vpm-node{padding:4px 10px;border-radius:6px;border:1px solid #dce3df;font-weight:600}
.vpm-node.res{background:#e2f2e9;border-color:transparent;color:#12804a}
.vpm-pill{display:inline-block;padding:2px 9px;border-radius:999px;font-size:11px;font-weight:700;letter-spacing:.03em;white-space:nowrap}
.vpm-ader{font-size:2.6rem;font-weight:700;color:#12804a;line-height:1.1;font-variant-numeric:tabular-nums}
.vpm-meter{height:10px;border-radius:5px;background:#eef2f0;overflow:hidden;display:flex;margin:8px 0}
.vpm-meter span{display:block;height:100%}
.vpm-legend{display:flex;flex-wrap:wrap;gap:6px 14px;font-size:12px}
.vpm-legend i{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:5px}
.vpm-evo{border:1px solid #dce3df;border-radius:9px;padding:10px 12px;margin-bottom:10px}
.vpm-evo small{color:#73817a;display:block}
.vpm-evo code{font-size:12.5px}
</style>
""", unsafe_allow_html=True)


def pill(code: str) -> str:
    s = E.STATUS[code]
    return f'<span class="vpm-pill" style="background:{E.TONE_BG[s["tone"]]};color:{E.TONE_FG[s["tone"]]}">{s["sym"]} {s["label"]}</span>'


def flow_html(names, with_res=True) -> str:
    parts = []
    for i, n in enumerate(names):
        if i:
            parts.append("<span>→</span>")
        parts.append(f'<span class="vpm-node">{n}</span>')
    if with_res:
        parts.append('<span>→</span><span class="vpm-node res">Situação final</span>')
    return f'<div class="vpm-flow">{"".join(parts)}</div>'


# --------------------------------------------------------------------------- estado
def new_stage(nome, tipo):
    return {"id": uuid.uuid4().hex[:8], "nome": nome, "tipo": tipo, "files": [], "nonce": 0}


def default_stages():
    return [new_stage("Layout Inicial", "origem"), new_stage("Senior", "senior"), new_stage("Layout Ajustado", "ajuste")]


def default_meta(resp=""):
    return {"cliente": "", "projeto": "Migração de Dados HCM", "modulo": "Administração de Pessoal", "layout": "",
            "data": date.today().isoformat(), "responsavel": resp}


ss = st.session_state
if "stages" not in ss:
    ss.stages = default_stages()
    ss.meta = default_meta()
    ss.fields, ss.key, ss.rules = [], [], dict(E.DEFAULT_RULES)
    ss.result, ss.xlsx, ss.is_example, ss.sig, ss.errors = None, None, False, None, []
    ss.page = "Início"
    ss.map_ver = 0


@st.cache_data(show_spinner=False, max_entries=64)
def build_cached(sources: tuple):
    return E.build_stage([(n, d) for n, d in sources])


def built_stages():
    out = []
    for s in ss.stages:
        b = build_cached(tuple((n, d) for n, d in s["files"])) if s["files"] else {"headers": [], "rows": [], "info": []}
        out.append({**b, "nome": s["nome"], "tipo": s["tipo"], "arquivos": ", ".join(n for n, _ in s["files"])})
    return out


def sync_fields(bs):
    sig = tuple(tuple(b["headers"]) for b in bs)
    if sig != ss.sig:
        ss.sig = sig
        ss.fields = E.auto_map(bs, ss.fields)
        ss.key = [k for k in ss.key if any(f["nome"] == k for f in ss.fields)]
        if not ss.key:
            ss.key = E.guess_key(ss.fields)
        ss.map_ver += 1


def load_example():
    ss.stages = default_stages()
    for st_, (name, m) in zip(ss.stages, E.make_example()):
        st_["files"] = [(name, m)]
    ss.meta = {"cliente": "Cliente Exemplo S.A.", "projeto": "Migração de Dados HCM", "modulo": "Administração de Pessoal",
               "layout": "Colaboradores (exemplo)", "data": date.today().isoformat(), "responsavel": "Equipe de Inteligência"}
    ss.fields, ss.key, ss.sig = [], [], None
    bs = built_stages()
    sync_fields(bs)
    for f in ss.fields:
        if f["nome"] == "UF":
            f["tipo"] = "texto"
    ss.is_example = True
    execute(bs)


def reset_validation():
    resp = ss.meta.get("responsavel", "") if not ss.is_example else ""
    ss.stages, ss.meta = default_stages(), default_meta(resp)
    ss.fields, ss.key, ss.sig, ss.result, ss.xlsx, ss.is_example, ss.errors = [], [], None, None, None, False, []


def execute(bs=None):
    bs = bs or built_stages()
    ss.errors = E.validate(bs, ss.fields, ss.key)
    if ss.errors:
        return False
    with st.spinner("Comparando registros…"):
        ss.result = E.run_comparison(bs, ss.fields, ss.key, ss.rules, ss.meta)
        ss.xlsx = E.export_excel(ss.result, bs)
    return True


def file_name_for_export():
    import re
    m = ss.result["meta"]
    slug = lambda s: re.sub(r"[^A-Za-z0-9]+", "_", E.strip_accents(str(s or ""))).strip("_")[:40]  # noqa: E731
    return f"Validacao_PosMigracao_{slug(m.get('cliente')) or 'cliente'}_{slug(m.get('layout')) or 'layout'}_{(m.get('data') or '').replace('-', '')}.xlsx"


def load_configs() -> list:
    try:
        with open(CONFIG_FILE, encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return []


def save_configs(cfgs: list):
    os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as fh:
        json.dump(cfgs, fh, ensure_ascii=False, indent=2)


def apply_config(c: dict):
    ss.meta["layout"] = c.get("nome", ss.meta["layout"])
    if c.get("modulo"):
        ss.meta["modulo"] = c["modulo"]
    ss.rules = {**ss.rules, **c.get("rules", {})}
    for i, e in enumerate(c.get("etapas", [])):
        if i >= len(ss.stages):
            ss.stages.append(new_stage(e["nome"], e["tipo"]))
        elif not ss.stages[i]["files"]:
            ss.stages[i]["nome"], ss.stages[i]["tipo"] = e["nome"], e["tipo"]
    bs = built_stages()
    sync_fields(bs)
    by = {f["nome"]: f for f in c.get("fields", [])}
    for f in ss.fields:
        o = by.get(f["nome"])
        if not o:
            continue
        for k in ("tipo", "comparar", "obrigatorio", "tolerancia", "zeros", "papel"):
            f[k] = o.get(k, f[k])
        f["cols"] = [f["nome"] if si == 0 else (o["cols"][si] if si < len(o["cols"]) and o["cols"][si] in b["headers"] else f["cols"][si])
                     for si, b in enumerate(bs)]
    ss.key = [k for k in c.get("key", []) if any(f["nome"] == k for f in ss.fields)] or ss.key
    ss.map_ver += 1


if ss.result is None and not ss.get("_example_done"):
    ss._example_done = True
    load_example()

# --------------------------------------------------------------------------- cabeçalho e menu
def logo(fname, cls, txt):
    p = os.path.join(ASSETS, fname)
    if os.path.exists(p):
        import base64
        b64 = base64.b64encode(open(p, "rb").read()).decode()
        return f'<img src="data:image/png;base64,{b64}" style="height:34px;width:auto" alt="{txt}">'
    return f'<span class="vpm-wm {cls}">{txt}</span>'


st.markdown(f"""<div class="vpm-head">{logo("logo_agante.png", "vpm-ag", "AGANTE")}{logo("logo_senior.png", "vpm-sr", "SENIOR")}
<div class="vpm-sep"></div><div class="vpm-title"><h1>Validador Pós-Migração</h1><p>Agante Tecnologia | Senior HCM</p></div></div>""",
            unsafe_allow_html=True)

PAGES = ["Início", "Nova Validação", "Comparativo", "Divergências", "Dashboard", "Configuração de Layouts", "DE/PARA de Campos"]
if "nav_target" in ss:
    ss.page = ss.pop("nav_target")
with st.sidebar:
    st.radio("Menu", PAGES, key="page")
    if ss.result is not None:
        st.divider()
        st.caption(f"{ss.result['meta'].get('cliente') or 'Sem cliente'} · {ss.result['meta'].get('layout') or 'Sem layout'}")
        st.download_button("Exportar Excel", ss.xlsx, file_name=file_name_for_export(), width="stretch",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")


def go(page):
    ss.nav_target = page


if ss.is_example and ss.page in ("Início", "Comparativo", "Divergências", "Dashboard"):
    c1, c2 = st.columns([5, 1.4])
    c1.info("**Dados de exemplo.** Os resultados abaixo vêm de arquivos fictícios gerados para demonstração.")
    if c2.button("Iniciar com meus arquivos", type="primary", width="stretch"):
        reset_validation()
        go("Nova Validação")
        st.rerun()

BS = built_stages()
sync_fields(BS)


# --------------------------------------------------------------------------- componentes
def mapping_editor(bs, key_prefix):
    if not ss.fields:
        st.caption("Carregue o Layout Inicial para montar o DE/PARA.")
        return
    tipo_lbl = {v: k for k, v in E.TIPOS.items()}
    papel_lbl = {v: k for k, v in E.PAPEIS.items()}
    zeros_lbl = {"Ignorar": "remover", "Considerar": "manter"}
    rows = []
    for f in ss.fields:
        r = {"Campo": f["nome"]}
        for si, b in enumerate(bs[1:], start=1):
            r[f"{si + 1}. {b['nome']}"] = f["cols"][si] if si < len(f["cols"]) else ""
        r.update({"Tipo": E.TIPOS[f["tipo"]], "Comparar": bool(f["comparar"]), "Obrigatório": bool(f["obrigatorio"]),
                  "Tolerância": float(f["tolerancia"] or 0), "Zeros à esquerda": "Ignorar" if f["zeros"] == "remover" else "Considerar",
                  "Papel": E.PAPEIS[f.get("papel", "")]})
        rows.append(r)
    df = pd.DataFrame(rows)
    cfg = {"Campo": st.column_config.TextColumn(disabled=True, help=f"Coluna em {bs[0]['nome']}")}
    for si, b in enumerate(bs[1:], start=1):
        cfg[f"{si + 1}. {b['nome']}"] = st.column_config.SelectboxColumn(options=[""] + b["headers"], help="Coluna correspondente nesta etapa")
    cfg.update({"Tipo": st.column_config.SelectboxColumn(options=list(E.TIPOS.values()), required=True),
                "Tolerância": st.column_config.NumberColumn(min_value=0.0, step=0.01, format="%.2f", help="Diferença aceita em Número/Valor"),
                "Zeros à esquerda": st.column_config.SelectboxColumn(options=list(zeros_lbl), help="Para campos do tipo Código"),
                "Papel": st.column_config.SelectboxColumn(options=list(E.PAPEIS.values()), help="Usado nos filtros e no detalhe do registro")})
    unmapped = sum(1 for f in ss.fields for si, b in enumerate(bs[1:], start=1) if b["headers"] and not f["cols"][si])
    if unmapped:
        st.warning(f"{unmapped} pareamento(s) sem coluna. Campos sem par em uma extração do Senior ficam como validação manual.")
    ed = st.data_editor(df, column_config=cfg, hide_index=True, width="stretch", num_rows="fixed",
                        key=f"{key_prefix}_{ss.map_ver}_{hashlib.md5('|'.join(b['nome'] for b in bs).encode()).hexdigest()[:6]}")
    for f, (_, r) in zip(ss.fields, ed.iterrows()):
        for si, b in enumerate(bs[1:], start=1):
            v = r.get(f"{si + 1}. {b['nome']}") or ""
            f["cols"][si] = v if v in b["headers"] else ""
        f["tipo"] = tipo_lbl.get(r["Tipo"], f["tipo"])
        f["comparar"], f["obrigatorio"] = bool(r["Comparar"]), bool(r["Obrigatório"])
        f["tolerancia"] = float(r["Tolerância"] or 0)
        f["zeros"] = zeros_lbl.get(r["Zeros à esquerda"], f["zeros"])
        f["papel"] = papel_lbl.get(r["Papel"], "")


def rules_editor(prefix):
    r = ss.rules
    c = st.columns(4)
    r["trim"] = c[0].checkbox("Remover espaços antes e depois", r["trim"], key=f"{prefix}_trim")
    r["case_ins"] = c[1].checkbox("Ignorar maiúsculas/minúsculas", r["case_ins"], key=f"{prefix}_case")
    r["acentos"] = c[2].checkbox("Ignorar acentos", r["acentos"], key=f"{prefix}_ac")
    r["data_nula"] = c[3].checkbox("Tratar 31/12/1900 como data vazia", r["data_nula"], key=f"{prefix}_dn",
                                   help="Convenção do Senior para data não informada.")
    r["nulos"] = st.text_input("Valores tratados como vazio (separados por vírgula)", r["nulos"], key=f"{prefix}_nulos",
                               help="Valores como 10, 10.0 e 10,00 são equivalentes nos tipos Número, Valor e Código.")


def save_current_config():
    if not ss.fields:
        st.toast("Carregue o Layout Inicial antes de salvar a configuração.")
        return
    c = {"nome": ss.meta["layout"] or "Layout sem nome", "modulo": ss.meta["modulo"], "key": list(ss.key), "rules": dict(ss.rules),
         "etapas": [{"nome": s["nome"], "tipo": s["tipo"]} for s in ss.stages],
         "fields": [{k: f[k] for k in ("nome", "cols", "tipo", "comparar", "obrigatorio", "tolerancia", "zeros", "papel")} for f in ss.fields],
         "atualizado_em": pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")}
    cfgs = [x for x in load_configs() if x["nome"] != c["nome"]]
    save_configs([c] + cfgs)
    st.toast(f"Configuração “{c['nome']}” salva.")


def styled(df: pd.DataFrame):
    def color(v):
        code = E.LABEL_TO_CODE.get(v)
        if not code:
            return ""
        t = E.STATUS[code]["tone"]
        return f"background-color:{E.TONE_BG[t]};color:{E.TONE_FG[t]};font-weight:700"
    return df.style.map(color, subset=["Resultado"]) if "Resultado" in df.columns and len(df) else df


def paged_table(df: pd.DataFrame, key: str, per=500):
    st.caption(f"{E.fmt_int(len(df))} itens")
    if len(df) > per:
        pages = (len(df) - 1) // per + 1
        p = st.number_input(f"Página (1–{pages})", 1, pages, 1, key=key)
        df = df.iloc[(p - 1) * per: p * per]
    st.dataframe(styled(df), hide_index=True, width="stretch", height=min(640, 38 + 35 * max(1, len(df))))


def record_detail(R, rec_key: str):
    rec = next((r for r in R["records"] if r["key"] == rec_key), None)
    if not rec:
        return
    nz, names = R["nz"], [s["nome"] for s in R["stages"]]
    st.markdown(f"#### {rec.get('nome') or rec['key']}")
    ids = " · ".join(x for x in [rec.get("matricula") and "Matrícula " + rec["matricula"], rec.get("cpf") and "CPF " + rec["cpf"],
                                  rec.get("empresa") and "Empresa " + rec["empresa"]] if x)
    st.markdown(f"{pill(rec['status'])} &nbsp; <code>Chave {rec['key']}</code> &nbsp; <span style='color:#73817a'>{ids}</span>", unsafe_allow_html=True)
    for n in rec["notes"]:
        st.warning(n)
    if rec["fields"]:
        rows = [{"Campo": x["f"]["nome"], **{n: E.show(v, x["f"], nz) for n, v in zip(names, x["raws"])},
                 "Resultado": E.STATUS[x["code"]]["label"], "Análise": x["nota"]} for x in rec["fields"]]
        st.dataframe(styled(pd.DataFrame(rows)), hide_index=True, width="stretch")
    else:
        st.markdown(flow_html([f"{n}: {'presente' if p else 'ausente'}" for n, p in zip(names, rec["present"])], False), unsafe_allow_html=True)
    diffs = [x for x in rec["fields"] if x["code"] != "OK"]
    if diffs:
        st.markdown("**Evolução dos campos com diferença**")
        cols = st.columns(min(3, len(diffs)))
        for i, x in enumerate(diffs):
            chain = "<div>↓</div>".join(f"<small>{n}</small><code>{E.show(v, x['f'], nz)}</code>"
                                        for n, v in zip(names, x["raws"]) if v is not E.ABSENT)
            cols[i % len(cols)].markdown(f'<div class="vpm-evo"><b>{x["f"]["nome"]}</b><br>{chain}<hr style="margin:8px 0">{pill(x["code"])}'
                                         f'<div style="font-size:12.5px;margin-top:6px">{x["nota"]}</div></div>', unsafe_allow_html=True)
    if rec["trail"]:
        st.markdown("**Trilha de alterações** · " + " → ".join(names))
        st.dataframe(pd.DataFrame([{"Etapa": t["etapa"], "Campo": t["campo"], "Valor anterior": t["de"], "Valor novo": t["para"],
                                    "Arquivo": t["arquivo"], "Data da análise": t["data"]} for t in rec["trail"]]),
                     hide_index=True, width="stretch")


def need_result():
    st.info("Nenhuma comparação executada ainda.")
    if st.button("Nova validação", type="primary"):
        go("Nova Validação")
        st.rerun()


# --------------------------------------------------------------------------- páginas
def page_inicio():
    st.markdown('<div class="vpm-eyebrow">Início</div>', unsafe_allow_html=True)
    st.subheader("O que foi enviado, o que ficou no Senior e o que precisou ser ajustado")
    st.caption("Compare o arquivo de origem com a extração do Senior e com os layouts ajustados, campo a campo, e veja a situação final de cada registro.")
    st.markdown(flow_html([s["nome"] for s in ss.stages]), unsafe_allow_html=True)
    R = ss.result
    if R:
        k = R["k"]
        with st.container(border=True):
            a, b = st.columns([1, 1.3])
            with a:
                st.markdown('<div class="vpm-eyebrow">Validação atual</div>', unsafe_allow_html=True)
                st.markdown(f"**{R['meta'].get('cliente') or 'Cliente não informado'}**  \n{R['meta'].get('projeto', '')} · "
                            f"{R['meta'].get('modulo', '')} · {R['meta'].get('layout') or 'Layout não informado'}")
                st.markdown(f'<div class="vpm-ader">{E.fmt_pct(k["aderencia"])}</div><div style="color:#73817a">de aderência da migração · '
                            f'{E.fmt_int(k["corretos"])} de {E.fmt_int(k["total"])} registros sem divergência</div>', unsafe_allow_html=True)
                b1, b2 = st.columns(2)
                if b1.button("Abrir comparativo", type="primary", width="stretch"):
                    go("Comparativo"); st.rerun()
                if b2.button("Dashboard", width="stretch"):
                    go("Dashboard"); st.rerun()
            with b:
                st.markdown('<div class="vpm-eyebrow">Resumo da validação</div>', unsafe_allow_html=True)
                st.write(R["summary"])
        c = st.columns(4)
        c[0].metric("Com divergência", E.fmt_int(k["divergentes"]))
        c[1].metric("Corrigidos", E.fmt_int(k["corrigidos"]))
        c[2].metric("Não migrados", E.fmt_int(k["nao_migrados"]))
        c[3].metric("Validação manual", E.fmt_int(k["manual"]))
    with st.container(border=True):
        st.markdown("**Três formas de carregar os arquivos** (escolha em cada etapa da Nova Validação)")
        c = st.columns(3)
        c[0].markdown("**Upload**  \nArraste um ou vários arquivos XLSX, XLS, CSV ou TXT para a etapa.")
        c[1].markdown("**Leitura (colar)**  \nCopie o conteúdo do CSV ou as células do Excel, com o cabeçalho, e cole.")
        c[2].markdown("**Caminho na máquina**  \nCole o caminho do arquivo ou da pasta (ex.: `C:\\Migracao\\Cliente`). "
                      "Funciona quando o validador roda no seu computador ou tem acesso à pasta de rede.")


def stage_block(si: int, s: dict, b: dict):
    with st.container(border=True):
        c1, c2, c3 = st.columns([3, 2, 1])
        s["nome"] = c1.text_input("Nome da etapa", s["nome"], key=f"nm_{s['id']}", label_visibility="collapsed")
        if si == 0:
            c2.markdown("<span class='vpm-pill' style='background:#eef1ef;color:#4a5952'>ORIGEM</span>", unsafe_allow_html=True)
        else:
            opts = ["senior", "ajuste"]
            s["tipo"] = c2.selectbox("Tipo", opts, index=opts.index(s["tipo"]), format_func=lambda x: E.TIPO_ETAPA[x],
                                     key=f"tp_{s['id']}", label_visibility="collapsed")
        if si >= 3 and c3.button("Excluir etapa", key=f"del_{s['id']}"):
            ss.stages.pop(si)
            for f in ss.fields:
                if si < len(f["cols"]):
                    f["cols"].pop(si)
            st.rerun()
        st.caption(E.DESC_ETAPA[s["tipo"]] + (" Opcional; pode conter só os registros corrigidos." if s["tipo"] == "ajuste" else ""))
        t_up, t_paste, t_path = st.tabs(["Upload", "Leitura (colar)", "Caminho na máquina"])
        with t_up:
            def on_upload(sid=s["id"]):
                stg = next(x for x in ss.stages if x["id"] == sid)
                ups = ss.get(f"up_{sid}_{stg['nonce']}") or []
                stg["files"] += [(u.name, u.getvalue()) for u in ups]
                stg["nonce"] += 1
            st.file_uploader("Arraste um ou vários arquivos", type=["xlsx", "xlsm", "xls", "csv", "txt"], accept_multiple_files=True,
                             key=f"up_{s['id']}_{s['nonce']}", on_change=on_upload, label_visibility="collapsed")
        with t_paste:
            txt = st.text_area("Cole o conteúdo", key=f"ps_{s['id']}_{s['nonce']}", height=120, label_visibility="collapsed",
                               placeholder="Cole aqui o conteúdo do CSV ou as células copiadas do Excel, incluindo o cabeçalho.")
            if st.button("Ler conteúdo", key=f"psb_{s['id']}"):
                if txt.strip():
                    n = sum(1 for nm, _ in s["files"] if nm.startswith("Conteúdo colado")) + 1
                    s["files"].append((f"Conteúdo colado {n}" if n > 1 else "Conteúdo colado", txt))
                    s["nonce"] += 1
                    st.rerun()
                else:
                    st.warning("Nada para ler. Cole o conteúdo no campo acima.")
        with t_path:
            paths = st.text_area("Caminhos", key=f"pt_{s['id']}", height=80, label_visibility="collapsed",
                                 placeholder="Um caminho por linha: arquivo (C:\\Migracao\\layout.csv) ou pasta inteira (C:\\Migracao\\Cliente)")
            if st.button("Ler caminho", key=f"ptb_{s['id']}"):
                read, errs = [], []
                for p in [x for x in paths.splitlines() if x.strip()]:
                    try:
                        got = E.files_from_path(p)
                        read += got
                        if not got:
                            errs.append(f"Nenhum arquivo XLSX, XLS, CSV ou TXT em: {p}")
                    except FileNotFoundError:
                        errs.append(f"Caminho não encontrado neste computador: {p}")
                    except PermissionError:
                        errs.append(f"Sem permissão para ler: {p}")
                for e in errs:
                    st.error(e)
                if read:
                    s["files"] += read
                    st.rerun()
            st.caption("Lê do computador onde o validador está rodando. No servidor Streamlit Cloud, use Upload ou Leitura.")
        if s["files"]:
            info = pd.DataFrame(b["info"]).rename(columns={"fonte": "Fonte", "registros": "Registros", "usada": "Usada", "situacao": "Situação"})
            st.dataframe(info, hide_index=True, width="stretch")
            c1, c2 = st.columns([4, 1])
            c1.markdown(f"**{E.fmt_int(len(b['rows']))}** registros · **{len(b['headers'])}** colunas")
            if c2.button("Remover arquivos", key=f"clr_{s['id']}", width="stretch"):
                s["files"] = []
                s["nonce"] += 1
                if si == 0:
                    ss.fields, ss.key = [], []
                st.rerun()


def page_nova():
    st.markdown('<div class="vpm-eyebrow">Nova validação</div>', unsafe_allow_html=True)
    h1, h2 = st.columns([4, 1.2])
    h1.subheader("Configurar e executar a comparação")
    h1.caption("Os arquivos são processados somente nesta sessão; nada é gravado.")
    if h2.button("Limpar tudo", width="stretch"):
        reset_validation(); st.rerun()
    if ss.errors:
        st.error("**Revise antes de comparar:**\n\n" + "\n".join(f"- {e}" for e in ss.errors))

    st.markdown("##### 1 · Identificação")
    cfgs = load_configs()
    if cfgs:
        sel = st.selectbox("Aplicar configuração salva", [""] + [c["nome"] for c in cfgs], format_func=lambda x: x or "Selecione um layout salvo…")
        if sel and st.button("Aplicar configuração"):
            apply_config(next(c for c in cfgs if c["nome"] == sel)); st.toast(f"Configuração “{sel}” aplicada."); st.rerun()
    m = ss.meta
    c = st.columns(3)
    m["cliente"] = c[0].text_input("Cliente", m["cliente"], placeholder="Razão social ou nome do cliente")
    m["projeto"] = c[1].text_input("Projeto", m["projeto"], placeholder="Ex.: Migração de Dados HCM")
    m["modulo"] = c[2].text_input("Módulo", m["modulo"], placeholder="Ex.: Administração de Pessoal, Benefícios, Ponto")
    c = st.columns(3)
    m["layout"] = c[0].text_input("Layout", m["layout"], placeholder="Ex.: 1031 – Dependentes")
    m["data"] = c[1].date_input("Data da validação", date.fromisoformat(m["data"]), format="DD/MM/YYYY").isoformat()
    m["responsavel"] = c[2].text_input("Responsável pela validação", m["responsavel"], placeholder="Nome do analista")

    st.markdown("##### 2 · Arquivos")
    st.caption("XLSX, XLS, CSV ou TXT. Cabeçalho, separador e abas são identificados automaticamente; vários arquivos na mesma etapa são somados.")
    for si, (s, b) in enumerate(zip(ss.stages, BS)):
        stage_block(si, s, b)
    c = st.columns([1, 1, 1, 3])
    for col, (nm, tp) in zip(c, [("Senior 2ª Carga", "senior"), ("Ajuste Final", "ajuste"), ("Nova etapa", "senior")]):
        if col.button(f"+ {nm}", width="stretch"):
            ss.stages.append(new_stage(nm, tp))
            for f in ss.fields:
                f["cols"].append("")
            st.rerun()
    st.markdown(flow_html([s["nome"] for s in ss.stages]), unsafe_allow_html=True)

    st.markdown("##### 3 · Chave de comparação")
    if ss.fields:
        names = [f["nome"] for f in ss.fields]
        c1, c2 = st.columns([5, 1])
        ss.key = c1.multiselect("Colunas que identificam o mesmo registro em todos os arquivos", names,
                                default=[k for k in ss.key if k in names], key=f"key_{ss.map_ver}")
        if c2.button("Sugerir chave", width="stretch"):
            ss.key = E.guess_key(ss.fields); ss.map_ver += 1; st.rerun()
        st.caption("Exemplos: CPF · Matrícula · Empresa + Tipo de Colaborador + Cadastro (NumEmp + TipCol + NumCad) · Código do dependente.")
    else:
        st.caption("Carregue o Layout Inicial para escolher as colunas da chave.")

    st.markdown("##### 4 · DE/PARA de campos")
    st.caption("Colunas pareadas automaticamente por nome e por sinônimos Senior (ex.: MATRICULA → NumCad). Revise o que ficar sem par.")
    mapping_editor(BS, "map_nova")
    if st.button("Salvar configuração do layout"):
        save_current_config()

    st.markdown("##### 5 · Regras de normalização")
    rules_editor("nova")
    st.divider()
    if st.button("Executar comparação", type="primary"):
        ss.is_example = False
        if execute():
            go("Dashboard"); st.rerun()
        else:
            st.rerun()


def lines_filtered(R, flt):
    out = []
    for l in R["lines"]:
        rec = l["rec"]
        if flt.get("only_div") and l["code"] == "OK":
            continue
        if flt.get("status") and E.STATUS[l["code"]]["label"] != flt["status"]:
            continue
        campo = flt.get("campo")
        if campo and ((l["f"] is not None) if campo == "(registro)" else (l["f"] is None or l["f"]["nome"] != campo)):
            continue
        if flt.get("etapa") and l["etapa"] != flt["etapa"]:
            continue
        q = (flt.get("q") or "").strip().upper()
        if q and not any(q in str(rec.get(x, "")).upper() for x in ("key", "nome", "matricula", "cpf")):
            continue
        for fld in ("nome", "matricula", "empresa"):
            v = (flt.get(fld) or "").strip().upper()
            if v and v not in str(rec.get(fld, "")).upper():
                break
        else:
            cpf = "".join(ch for ch in (flt.get("cpf") or "") if ch.isdigit())
            if cpf and cpf not in "".join(ch for ch in str(rec.get("cpf", "")) if ch.isdigit()):
                continue
            out.append(l)
    return out


def page_comparativo():
    R = ss.result
    if not R:
        return need_result()
    st.markdown('<div class="vpm-eyebrow">Comparativo</div>', unsafe_allow_html=True)
    st.subheader("Comparação campo a campo")
    st.caption("Cada linha mostra o valor do campo em todas as etapas e a situação final.")
    st.markdown(flow_html([s["nome"] for s in R["stages"]]), unsafe_allow_html=True)
    c = st.columns([2.2, 1.4, 1.6, 1])
    flt = {"q": c[0].text_input("Buscar", placeholder="Chave, nome, matrícula ou CPF"),
           "campo": c[1].selectbox("Campo", ["", "(registro)"] + [f["nome"] for f in R["cmp_f"]], format_func=lambda x: x or "Todos"),
           "status": c[2].selectbox("Resultado", [""] + [v["label"] for v in E.STATUS.values()], format_func=lambda x: x or "Todos")}
    c[3].markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
    flt["only_div"] = c[3].checkbox("Somente diferenças")
    lines = lines_filtered(R, flt)
    paged_table(E.lines_frame(R, lines), "pg_comp")
    keys = list(dict.fromkeys(l["rec"]["key"] for l in lines))
    with st.expander("Ver registro completo", expanded=False):
        sel = st.selectbox("Registro", keys, format_func=lambda k: next((f"{k} · {r.get('nome')}" for r in R["records"] if r["key"] == k), k),
                           key="rec_comp") if keys else None
        if sel:
            record_detail(R, sel)


def page_divergencias():
    R = ss.result
    if not R:
        return need_result()
    st.markdown('<div class="vpm-eyebrow">Divergências</div>', unsafe_allow_html=True)
    st.subheader("Análise de Divergências")
    st.caption(f"Cliente: **{R['meta'].get('cliente') or '—'}** · Layout: **{R['meta'].get('layout') or '—'}** · "
               f"processado em {R['at'].strftime('%d/%m/%Y %H:%M')}")
    div = [l for l in R["lines"] if l["code"] != "OK"]
    codes = [v["label"] for k, v in E.STATUS.items() if k != "OK" and any(l["code"] == k for l in div)]
    etapas = sorted({l["etapa"] for l in div if l["etapa"]})
    c = st.columns(3)
    flt = {"only_div": True,
           "campo": c[0].selectbox("Campo", ["", "(registro)"] + [f["nome"] for f in R["cmp_f"]], format_func=lambda x: x or "Todos", key="d_campo"),
           "status": c[1].selectbox("Tipo de divergência", [""] + codes, format_func=lambda x: x or "Todos", key="d_status"),
           "etapa": c[2].selectbox("Etapa da migração", [""] + etapas, format_func=lambda x: x or "Todas", key="d_etapa")}
    c = st.columns(4)
    flt["nome"] = c[0].text_input("Colaborador", key="d_nome")
    flt["matricula"] = c[1].text_input("Matrícula", key="d_mat")
    flt["cpf"] = c[2].text_input("CPF", key="d_cpf")
    flt["empresa"] = c[3].text_input("Empresa", key="d_emp")
    lines = lines_filtered(R, flt)
    df = E.lines_frame(R, lines).drop(columns=["Chave"])
    paged_table(df, "pg_div")
    keys = list(dict.fromkeys(l["rec"]["key"] for l in lines))
    with st.expander("Ver registro completo", expanded=False):
        sel = st.selectbox("Registro", keys, format_func=lambda k: next((f"{k} · {r.get('nome')}" for r in R["records"] if r["key"] == k), k),
                           key="rec_div") if keys else None
        if sel:
            record_detail(R, sel)


def page_dashboard():
    R = ss.result
    if not R:
        return need_result()
    k = R["k"]
    st.markdown('<div class="vpm-eyebrow">Dashboard</div>', unsafe_allow_html=True)
    st.subheader("Painel da validação")
    st.caption(f"{R['meta'].get('cliente') or 'Cliente não informado'} · {R['meta'].get('layout') or 'Layout não informado'} · "
               f"processado em {R['at'].strftime('%d/%m/%Y %H:%M')}")
    with st.container(border=True):
        a, b = st.columns([1, 1.4])
        a.markdown('<div class="vpm-eyebrow">Aderência da migração</div>'
                   f'<div class="vpm-ader">{E.fmt_pct(k["aderencia"])}</div>'
                   f'<div class="vpm-meter"><span style="width:{k["aderencia"]}%;background:#1b7a43"></span></div>'
                   f'<div style="color:#73817a;font-size:12px">Registros sem divergência ÷ total analisado × 100 = '
                   f'{E.fmt_int(k["corretos"])} ÷ {E.fmt_int(k["total"])}</div>', unsafe_allow_html=True)
        b.markdown('<div class="vpm-eyebrow">Resumo da validação</div>', unsafe_allow_html=True)
        b.write(R["summary"])
    tiles = [("Registros analisados", k["total"]), ("Corretos", k["corretos"]), ("Com divergência", k["divergentes"]),
             ("Não migrados", k["nao_migrados"]), ("Alterados", k["alterados"]), ("Corrigidos", k["corrigidos"]),
             ("Validação manual", k["manual"]), ("Somente no Senior", k["somente_senior"])]
    for row in (tiles[:4], tiles[4:]):
        cols = st.columns(4)
        for col, (lbl, v) in zip(cols, row):
            col.metric(lbl, E.fmt_int(v))
    left, right = st.columns([1.2, 1])
    with left:
        st.markdown("**Indicadores por campo** · do menor para o maior % correto")
        fs = pd.DataFrame(R["field_stats"]).sort_values("pct") if R["field_stats"] else pd.DataFrame(columns=["campo", "registros", "divergencias", "pct"])
        fs = fs.rename(columns={"campo": "Campo", "registros": "Registros", "divergencias": "Divergências", "pct": "% Correto"})
        st.dataframe(fs, hide_index=True, width="stretch",
                     column_config={"% Correto": st.column_config.ProgressColumn(format="%.2f%%", min_value=0, max_value=100)})
    with right:
        total = sum(R["dist"].values()) or 1
        order = ["OK", "CORRIGIDO", "ALTERADO", "MANUAL", "NAO_PREENCHIDO", "DIVERGENCIA"]
        bar = "".join(f'<span style="width:{R["dist"][c] / total * 100}%;background:{E.TONE_FG[E.STATUS[c]["tone"]]}"></span>'
                      for c in order if R["dist"].get(c))
        leg = "".join(f'<span><i style="background:{E.TONE_FG[E.STATUS[c]["tone"]]}"></i>{E.STATUS[c]["label"]} · <b>{E.fmt_int(R["dist"][c])}</b></span>'
                      for c in order if R["dist"].get(c))
        st.markdown(f"**Resultado das comparações de campo** · {E.fmt_int(total)} comparações"
                    f'<div class="vpm-meter">{bar}</div><div class="vpm-legend">{leg}</div>', unsafe_allow_html=True)
        st.markdown("**Auditoria**")
        m = R["meta"]
        aud = [("Cliente", m.get("cliente")), ("Projeto", m.get("projeto")), ("Módulo", m.get("modulo")), ("Layout", m.get("layout")),
               ("Data da validação", E.br_date(m.get("data", ""))), ("Responsável", m.get("responsavel")),
               ("Processado em", R["at"].strftime("%d/%m/%Y %H:%M")), ("Chave", " + ".join(R["key"]))]
        aud += [(s["nome"], f"{s['arquivos'] or 'não carregado'} ({E.fmt_int(s['n'])})") for s in R["stages"]]
        aud += [("Registros", E.fmt_int(k["total"])), ("Divergências", E.fmt_int(k["div_campos"])), ("Aderência", E.fmt_pct(k["aderencia"]))]
        st.dataframe(pd.DataFrame(aud, columns=["Item", "Valor"]).fillna("—"), hide_index=True, width="stretch")


def page_layouts():
    st.markdown('<div class="vpm-eyebrow">Configuração de Layouts</div>', unsafe_allow_html=True)
    st.subheader("Parametrização reutilizável por layout")
    st.caption("Salve chave, DE/PARA, tipos, obrigatoriedade, tolerâncias e regras de um layout e reaplique em novas validações, para qualquer cliente.")
    if ss.fields:
        with st.container(border=True):
            st.markdown(f"**Configuração atual:** {ss.meta['layout'] or 'Layout sem nome'} · chave `{' + '.join(ss.key) or '—'}` · "
                        f"{len(ss.fields)} campos, {sum(1 for f in ss.fields if f['comparar'])} comparados")
            if st.button("Salvar como configuração do layout", type="primary"):
                save_current_config(); st.rerun()
    cfgs = load_configs()
    st.markdown(f"**Layouts salvos** ({len(cfgs)}) · arquivo `configs/layouts.json`")
    if not cfgs:
        st.caption("Nenhum layout salvo. Monte o DE/PARA em Nova Validação e use “Salvar configuração do layout”.")
    cols = st.columns(3)
    for i, c in enumerate(cfgs):
        with cols[i % 3].container(border=True):
            st.markdown(f"**{c['nome']}**  \n{c.get('modulo') or 'Módulo não informado'}")
            st.caption(f"Chave: {' + '.join(c.get('key', []))} · {len(c.get('fields', []))} campos · "
                       f"etapas: {' → '.join(e['nome'] for e in c.get('etapas', []))} · atualizado em {c.get('atualizado_em', '—')}")
            b1, b2 = st.columns(2)
            if b1.button("Aplicar", key=f"ap_{i}", type="primary", width="stretch"):
                apply_config(c); go("Nova Validação"); st.rerun()
            confirm = b2.checkbox("Confirmar exclusão", key=f"cf_{i}")
            if confirm and b2.button("Excluir", key=f"dl_{i}", width="stretch"):
                save_configs([x for x in cfgs if x["nome"] != c["nome"]]); st.rerun()
    st.divider()
    c1, c2 = st.columns(2)
    c1.download_button("Exportar layouts (JSON)", json.dumps(cfgs, ensure_ascii=False, indent=2).encode("utf-8"),
                       file_name="layouts_validador.json", mime="application/json", disabled=not cfgs)
    up = c2.file_uploader("Importar layouts (JSON)", type=["json"], key="imp_cfg")
    if up is not None and st.button("Importar"):
        try:
            novos = json.loads(up.getvalue().decode("utf-8"))
            nomes = {x["nome"] for x in novos}
            save_configs(novos + [x for x in cfgs if x["nome"] not in nomes])
            st.toast(f"{len(novos)} layout(s) importado(s)."); st.rerun()
        except Exception as e:
            st.error(f"Arquivo inválido: {e}")


def page_depara():
    st.markdown('<div class="vpm-eyebrow">DE/PARA de Campos</div>', unsafe_allow_html=True)
    st.subheader("Mapeamento entre arquivos")
    st.caption("Relacione cada coluna do Layout Inicial com a coluna correspondente em cada etapa. "
               "O pareamento automático reconhece nomes iguais e sinônimos usuais do Senior HCM.")
    if ss.fields:
        st.markdown(flow_html([s["nome"] for s in ss.stages], False), unsafe_allow_html=True)
    mapping_editor(BS, "map_dp")
    st.markdown("**Regras de normalização**")
    rules_editor("dp")
    c1, c2, _ = st.columns([1.3, 1.3, 4])
    if c1.button("Salvar configuração do layout"):
        save_current_config()
    if c2.button("Executar comparação", type="primary"):
        ss.is_example = False
        if execute():
            go("Dashboard")
        else:
            go("Nova Validação")
        st.rerun()


{"Início": page_inicio, "Nova Validação": page_nova, "Comparativo": page_comparativo, "Divergências": page_divergencias,
 "Dashboard": page_dashboard, "Configuração de Layouts": page_layouts, "DE/PARA de Campos": page_depara}[ss.page]()
