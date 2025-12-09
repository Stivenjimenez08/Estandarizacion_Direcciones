import pandas as pd
import re
import os

# ============================================================
# REGEX ESPECÍFICOS: PORTAL PRADERA / ZAGUANES / CHAMBRANAS
# ============================================================
regex_portal_pradera = re.compile(
    r"""(?:URB|CJT)?\s*PORTAL\s+PRADERA
        .*?(?:BLQ|BL|BLOQUE)\s*(?P<bloque>[A-Z])
        (?:\s*(?:APTO?|APARTAMENTO|AP)\s*)?(?P<apt>\d+)
        (?:\s+(?:8000|NIU\s*\#?\s*\d*))*
    """,
    re.IGNORECASE | re.VERBOSE,
)

regex_zaguanes = re.compile(
    r"""(?:URB|BRR)?\s*ZAGUANES
        .*?(?:MZ|MZA|MNZ|MZN|MANZANA)\s*(?P<manzana>[0-9]+[A-Z]?)
        .*?(?P<tipo_casa>CS|CASA|C|LC)\s*(?P<casa>[0-9]+[A-Z]?)
        (?:\s+(?:8000|NIU\s*\#?\s*\d*))* 
    """,
    re.IGNORECASE | re.VERBOSE,
)

regex_chambranas = re.compile(
    r"""BRR\s*CHAMBRANAS
        .*?(?:MZ|MZA|MNZ|MZN|MANZANA)\s*(?P<manzana>[0-9]+[A-Z]?)
        .*?(?P<tipo_casa>CS|CASA|C|LC)\s*(?P<casa>[0-9]+[A-Z]?)
        (?:\s*(?:PI|PISO)\s*(?P<piso>\d+))?
        (?:\s*(?:APTO?|APARTAMENTO|AP)\s*(?P<ap>\d+))?
        (?:\s+(?:8000|NIU\s*\#?\s*\d*))* 
    """,
    re.IGNORECASE | re.VERBOSE,
)

# ============================================================
# REGEX CRA / CL y CLL / CR  (basados en tu script de C29)
# ============================================================

# ---------- CRA / CL ----------
regex_cra_cl_ap_then_to = re.compile(
    r"""CRA\s*(?P<cra>\d+)\s*
        CL\s*(?P<cl>\d+[A-Z]?)\s*
        (?:(?:-|–|—)\s*|\s+)(?P<guion>0*\d+)\s*
        (?:APTO?|APARTAMENTO|AP)\s*(?P<ap>\d+)\s+
        (?:(?:\bTO\b|\bTORRE\b|\bBQ\b|\bBLQ\b|\bBL\b|\bBLOQUE\b))\s*(?P<to>[A-Z0-9]+)
    """,
    re.IGNORECASE | re.VERBOSE,
)

regex_cra_cl_ap_textnum = re.compile(
    r"""CRA\s*(?P<cra>\d+)\s*
        CL\s*(?P<cl>\d+[A-Z]?)\s*
        (?:(?:-|–|—)\s*|\s+)(?P<guion>0*\d+)\s*
        (?:APTO?|APARTAMENTO|AP)\s*(?:[A-ZÁÉÍÓÚÜÑ]+(?:\s+[A-ZÁÉÍÓÚÜÑ]+)*)\s*(?P<ap>\d+)
    """,
    re.IGNORECASE | re.VERBOSE,
)

regex_cra_cl_bq_to_ap = re.compile(
    r"""CRA\s*(?P<cra>\d+)\s*
        CL\s*(?P<cl>\d+[A-Z]?)\s*
        (?:(?:-|–|—)\s*|\s+)(?P<guion>0*\d+)\s*
        (?:ET(?:APA)?\s*(?P<et>\d+))?\s*
        .*?(?P<tipo>(?:\bTO\b|\bTORRE\b|\bT\b|\bBL\b|\bBQ\b|\bBLQ\b|\bBLOQUE\b))\s*
        (?P<numtipo>(?:[A-Z0-9]+))
        (?:\s*(?:APTO?|APARTAMENTO|AP)\s*(?P<ap>\d+))?
        (?:\s*(?:PI|PISO)\s*(?P<piso>\d+))?
        (?:\s*LC\s+(?P<lc_raw>(?:\d+[A-Z]?|[A-Z]{1,3}\s*\d{1,4}|[A-Z]{1,3}\d{1,4}|\d+\s*-\s*\d+)))?
        (?:\s*(?P<tail>(?:PU\s+VIGILANCIA|MOTOBOMBA|(?:OFI(?:CINA)?|OF(?:ICINA)?)\s*\d+(?:\s*-\s*\d+)?|MACROMEDIDOR\s*\d+|ECR\s+[A-ZÁÉÍÓÚÜÑ\s]+)))?
    """,
    re.IGNORECASE | re.VERBOSE,
)

regex_cra_cl_cs_lc = re.compile(
    r"""CRA\s*(?P<cra>\d+)\s*
        CL\s*(?P<cl>\d+[A-Z]?)\s*
        (?:(?:-|–|—)\s*|\s+)?(?P<guion>0*\d+)?      # guion opcional
        (?:\s*N\s*\d+)?\s*
        (?P<tipo2>CS|LC)\s*
        (?P<num2>(?:\d+[A-Z]?|[A-Z]{1,3}\s*\d{1,4}|[A-Z]{1,3}\d{1,4}|\d+\s*-\s*\d+))?
        (?:\s*(?:PI|PISO)\s*(?P<piso>\d+))?
    """,
    re.IGNORECASE | re.VERBOSE,
)

regex_cra_cl_basico = re.compile(
    r"""CRA\s*(?P<cra>\d+)\s*
        CL\s*(?P<cl>\d+[A-Z]?)\s*
        (?:(?:-|–|—)\s*|\s+)(?P<guion>0*\d+)\b
        (?:\s*(?P<tail>(?:PU\s+VIGILANCIA|MOTOBOMBA|MACROMEDIDOR\s*\d+|(?:OFI(?:CINA)?|OF(?:ICINA)?)\s*\d+(?:\s*-\s*\d+)?)))?
        (?:\s*(?:PI|PISO)\s*(?P<piso>\d+))?
        (?:\s+(?:8000|NIU\s*\#?\s*\d+))*
    """,
    re.IGNORECASE | re.VERBOSE,
)

regex_cra_cl_ap = re.compile(
    r"""CRA\s*(?P<cra>\d+)\s*
        CL\s*(?P<cl>\d+[A-Z]?)\s*
        (?:(?:-|–|—)\s*|\s+)(?P<guion>0*\d+)\s*
        (?:APTO?|APARTAMENTO|AP)\s*(?P<ap>\d+)
        (?:\s*(?:PI|PISO)\s*(?P<piso>\d+))?
        (?:\s+(?:8000|NIU\s*\#?\s*\d+))*
    """,
    re.IGNORECASE | re.VERBOSE,
)

regex_cra_cl_ap_sin_num = re.compile(
    r"""CRA\s*(?P<cra>\d+)\s*
        CL\s*(?P<cl>\d+[A-Z]?)\s*
        (?:(?:-|–|—)\s*|\s+)(?P<guion>0*\d+)\s*
        (?:APTO?|APARTAMENTO|AP)\b(?!\s*\d)
    """,
    re.IGNORECASE | re.VERBOSE,
)

regex_cra_cl_macro = re.compile(
    r"""CRA\s*(?P<cra>\d+)\s*
        CL\s*(?P<cl>\d+[A-Z]?)\s*
        (?:(?:-|–|—)\s*|\s+)(?P<guion>0*\d+)\s*
        MACRO\s*(?P<macro>\d+)
    """,
    re.IGNORECASE | re.VERBOSE,
)

regex_cra_cl_mz_cs = re.compile(
    r"""CRA\s*(?P<cra>\d+)\s*
        CL\s*(?P<cl>\d+[A-Z]?)\s*
        (?:(?:-|–|—)\s*|\s+)(?P<guion>0*\d+)
        (?:\s+[A-ZÁÉÍÓÚÜÑ0-9]+){0,6}?\s*
        MZ\s*(?P<mz>[A-Z0-9]+)\s*
        CS\s*(?P<cs>[A-Z0-9]+)
    """,
    re.IGNORECASE | re.VERBOSE,
)

regex_cra_cl_to_ap = re.compile(
    r"""CRA\s*(?P<cra>\d+)\s*
        CL\s*(?P<cl>\d+[A-Z]?)\s*
        -\s*(?P<guion>0*\d+)\s*
        (?:TO|TORRE|T|BQ)\s*(?P<to>[A-Z0-9]+)\s*
        (?:APTO?|APARTAMENTO|AP)\s*(?P<ap1>\d+)(?:\s*(?:-|–|—)\s*(?P<ap2>\d+))?
        (?:\s+MONTEAZUL)?
        (?:\s+(?:8000|NIU\s*\#?\s*\d+))* 
    """,
    re.IGNORECASE | re.VERBOSE,
)

# MONTEAZUL (BQ/BL ≡ TO)
regex_monteazul = re.compile(
    r"""URB\s+MONTEAZUL
        (?:
            .*?(?:APTO?|APARTAMENTO|AP)\s*(?P<ap>\d+).*?(?:BQ|BL|BLQ|BLOQUE|TO|TORRE|T)\s*(?P<to>\d+)
            |
            .*?(?:TO|TORRE|T|BQ|BL|BLQ|BLOQUE)\s*(?P<to2>\d+).*?(?:APTO?|APARTAMENTO|AP)\s*(?P<ap2>\d+)
        )
        (?:\s+(?:8000|NIU\s*\#?\s*\d+))*   
    """,
    re.IGNORECASE | re.VERBOSE,
)

# ---------- CLL / CR GENERAL ----------
regex_cll_cr_general = re.compile(
    r"""CLL\s*(?P<cll>\d+(?:[A-Z]{1,2})?(?:\s*BIS(?:\s*[A-Z])?)?)
        (?:\s*(?:NORTE|SUR|ESTE|OESTE|ORIENTE|OCCIDENTE|N|S|E|O|NTE|STE|OTE|OCC))?
        \s*(?:CR|CRA|KR|KRA|K|CL)\s*(?P<cr>\d+(?:[A-Z]{1,2})?(?:\s*BIS(?:\s*[A-Z])?)?)
        (?:\s*(?:NORTE|SUR|ESTE|OESTE|ORIENTE|OCCIDENTE|N|S|E|O|NTE|STE|OTE|OCC))?
        (?:(?:(?:-|–|—)\s*|\#\s*|\s+)(?P<guion>0*\d+))?
        (?:\s*(?:APTO?|APARTAMENTO|AP)\s*(?P<ap>\d+))?
        (?:\s*(?:PI|PISO)\s*(?P<piso>\d+))?
        (?:\s*(?:OFI(?:CINA)?|OF(?:ICINA)?)\s*(?P<of>\d+(?:\s*-\s*\d+)?))?
        (?:\s*LC\s+(?P<lc>(?:\d+[A-Z]?|[A-Z]{1,3}\s*\d{1,4}|[A-Z]{1,3}\d{1,4}|\d+\s*-\s*\d+)))?
        (?:\s*CS\s+(?P<cs>[A-Z0-9]+))?
        (?:\s*AC\b)?
    """,
    re.IGNORECASE | re.VERBOSE,
)

# ============================================================
# HELPERS / LIMPIEZAS  (mismos que en C29)
# ============================================================

ROTULOS_A_REMOVER = [
    "CLUB HOUSE", "CLUBHOUSE",
    "CENTENARIO MALL", "MALL CENTENARIO", "FLORIDA BAJA",
    "BALEARES", "AV BOLIVAR", "ED EL PILAR", "AMANECER",
    "MALL ZN ORO", "LUXOR",
]

def limpiar_rotulos_finales(texto: str) -> str:
    t = texto
    for r in ROTULOS_A_REMOVER:
        t = re.sub(rf"\s+{re.escape(r)}\s*[\.\-]*\s*$", "", t, flags=re.IGNORECASE)
    return t

_RE_SOTANO = re.compile(r"\bSOTANO\s*(\d+)\b", re.IGNORECASE)
_RE_PLLIBRE = re.compile(r"\bPL\s+LIBRE\b", re.IGNORECASE)
_RE_CAJERO  = re.compile(r"\bCAJERO\s*(\d+)\b", re.IGNORECASE)
_RE_GRUPO   = re.compile(r"\bGRUPO\s+[A-ZÁÉÍÓÚÜÑ0-9 ]+\b", re.IGNORECASE)

def normalizar_cola_lc(seg: str) -> str:
    s = seg.strip()
    s = limpiar_rotulos_finales(s)
    s = re.sub(r'\s+AV(?:ENIDA)?\s+[A-ZÁÉÍÓÚÜÑ0-9\s]+$', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\s+ED(?:IFICIO)?\s+[A-ZÁÉÍÓÚÜÑ0-9\s]+$', '', s, flags=re.IGNORECASE)

    m = re.search(r'\bLC\b\s*(\S+)?', s, flags=re.IGNORECASE)
    if not m:
        return s

    base = "LC"
    lc_id = m.group(1)
    if lc_id:
        m2 = re.match(r'[A-Z0-9\-]+', lc_id, flags=re.IGNORECASE)
        lc_id = m2.group(0) if m2 else None
    if lc_id:
        base += f" {lc_id}"

    tail = s[m.end():].strip()
    tail = _RE_SOTANO.sub(lambda x: f"SOT {int(x.group(1))}", tail)
    tail = _RE_PLLIBRE.sub("", tail)
    tail = _RE_CAJERO.sub(lambda x: f"CAJ {int(x.group(1))}", tail)
    tail = _RE_GRUPO.sub("", tail)

    mpi = re.search(r'\b(?:PI|PISO)\s*(\d+)\b', tail, flags=re.IGNORECASE)
    pi_txt = f" PI {int(mpi.group(1))}" if mpi else ""

    out = (base + pi_txt).strip()
    return out

LC_PERMITIDO_COMPLETO = re.compile(
    r'^LC\s+(?:\d+[A-Z]?|[A-Z]{1,3}\s*\d{1,4}|[A-Z]{1,3}\d{1,4}|\d+\s*-\s*\d+)(?:\s+PI\s*\d+)?$',
    re.IGNORECASE,
)

def lc_es_complejo(texto: str) -> bool:
    m = re.search(r'\bLC\b.*$', texto, flags=re.IGNORECASE)
    if not m:
        return False
    seg = m.group(0)
    seg_norm = normalizar_cola_lc(seg)
    return LC_PERMITIDO_COMPLETO.match(seg_norm) is None

ORIENT_MAP = {
    "NORTE": "N", "SUR": "S", "ESTE": "E", "OESTE": "O",
    "ORIENTE": "E", "OCCIDENTE": "O",
    "NTE": "N", "STE": "S", "OTE": "E", "OCC": "O",
    "N": "N", "S": "S", "E": "E", "O": "O",
}

# ============================================================
# NORMALIZADOR PRINCIPAL
# ============================================================

def normalizar_direccion(direccion: str):
    if pd.isna(direccion):
        return direccion, "0"

    d = direccion.upper().strip()

    # Espacios/guiones unicode
    d = re.sub(r"[\u00A0\u2000-\u200B\u202F\u205F\u3000]", " ", d)
    d = re.sub(r"[‐‒–—―-]", "-", d)

    # Limpiezas base
    d = re.sub(r"\s*-\s*ARMENIA\b", "", d)
    d = re.sub(r"-?\s*NIU\s*\#?\s*-?\s*\d+\b", "", d)
    d = re.sub(r"\b(APTO?|APARTAMENTO|AP)\s*-\s*", r"\1 ", d)
    d = re.sub(r"\bCONS\b", "CN", d)
    d = re.sub(r"(?<!MACROMEDIDOR)(?<!MACRO)\s+8000\b", "", d)

    d = re.sub(r"\s*(?:-|–|—)\s*", " - ", d)
    d = re.sub(r"\s+", " ", d).strip()

    d = limpiar_rotulos_finales(d)
    d = re.sub(r"\s+AV(?:ENIDA)?\s+[A-ZÁÉÍÓÚÜÑ0-9\s]+$", "", d, flags=re.IGNORECASE)
    d = re.sub(r"\s+ED(?:IFICIO)?\s+[A-ZÁÉÍÓÚÜÑ0-9\s]+$", "", d, flags=re.IGNORECASE)

    d = re.sub(r"\bNIVEL\s*\d+\b", "", d, flags=re.IGNORECASE)
    d = re.sub(
        r"\s*-\s+(?!\d+\b)(?!ECR\b)(?!MACRO\b)(?!MACROMEDIDOR\b)[A-ZÁÉÍÓÚÜÑ]{3,}(?:\s+[A-ZÁÉÍÓÚÜÑ0-9]{2,})*$",
        "",
        d,
        flags=re.IGNORECASE,
    )
    d = re.sub(r"\bIN\b(?=\s*\d)", "AP", d, flags=re.IGNORECASE)
    d = re.sub(r"\s+", " ", d).strip()
    d = re.sub(r"\s*-\s*$", "", d)

    # Guardas
    if re.search(r"\bLT\s*\d+\b", d):
        return d, "0"
    if lc_es_complejo(d):
        return d, "0"

    # OF global
    ofm = re.search(r"\b(?:OFI(?:CINA)?|OF(?:ICINA)?)\s*(\d+)\s*-\s*(\d+)\b", d)
    if ofm:
        of_text = f"OF {int(ofm.group(1))}-{int(ofm.group(2))}"
    else:
        ofs = re.search(r"\b(?:OFI(?:CINA)?|OF(?:ICINA)?)\s*(\d+)\b", d)
        of_text = f"OF {int(ofs.group(1))}" if ofs else None

    cnm = re.search(r"\bCN\s*([A-Z0-9]+)\b", d)
    cn_text = f"CN {cnm.group(1)}" if cnm else None

    # ================= CLL / CR =================
    m = regex_cll_cr_general.search(d)
    if m:
        cll, cr, guion = m.group("cll"), m.group("cr"), m.group("guion")
        ap, piso, ofn, lc, cs = m.group("ap"), m.group("piso"), m.group("of"), m.group("lc"), m.group("cs")

        if not guion:
            mnum = re.search(
                rf"(?:CR|CRA|KR|KRA|K|CL)\s*{re.escape(cr)}\s*(?:\#\s*|(?:-|–|—)\s*|\s+)(\d+)",
                d,
            )
            if mnum:
                guion = mnum.group(1)

        if not guion:
            return d, "0"

        out = f"CLL {cll} CR {cr} - {guion}"
        if ap:   out += f" AP {int(ap)}"
        if piso: out += f" PI {int(piso)}"
        if lc:
            out_lc = normalizar_cola_lc(f"LC {lc.strip()}")
            out += f" {out_lc}"
        elif re.search(r'\bLC\b', d, flags=re.IGNORECASE):
            seg = re.search(r'\bLC\b.*$', d, flags=re.IGNORECASE).group(0)
            out_lc = normalizar_cola_lc(seg)
            out += f" {out_lc}"
        if cs:   out += f" CS {cs.strip()}"
        if ofn:
            out += f" OF {ofn.replace(' ', '')}"
        elif of_text and "OF" not in out:
            out += f" {of_text}"
        if cn_text and "CN" not in out:
            out += f" {cn_text}"
        return out, "1"

    # ================= CRA / CL =================
    m = regex_cra_cl_ap_then_to.search(d)
    if m:
        cra, cl, guion = m.group("cra"), m.group("cl"), m.group("guion")
        ap, to = int(m.group("ap")), m.group("to")
        out = f"CRA {cra} CL {cl} - {guion} TO {to} AP {ap}"
        if of_text and "OF" not in out: out += f" {of_text}"
        if cn_text and "CN" not in out: out += f" {cn_text}"
        return out, "1"

    m = regex_cra_cl_ap_textnum.search(d)
    if m:
        cra, cl, guion = m.group("cra"), m.group("cl"), m.group("guion")
        ap = int(m.group("ap"))
        out = f"CRA {cra} CL {cl} - {guion} AP {ap}"
        if of_text and "OF" not in out: out += f" {of_text}"
        if cn_text and "CN" not in out: out += f" {cn_text}"
        return out, "1"

    m = regex_cra_cl_bq_to_ap.search(d)
    if m:
        cra, cl, guion = m.group("cra"), m.group("cl"), m.group("guion")
        tipo, numt = m.group("tipo").upper(), m.group("numtipo")
        ap, piso, et = m.group("ap"), m.group("piso"), m.group("et")
        lc_raw, tail = m.group("lc_raw"), m.group("tail")
        etiqueta = "BQ" if tipo in ("BL", "BLQ", "BLOQUE", "BQ") else "TO"
        out = f"CRA {cra} CL {cl} - {guion} {etiqueta} {numt}"
        if ap:   out += f" AP {int(ap)}"
        if piso: out += f" PI {int(piso)}"
        if et:   out += f" ET {int(et)}"
        if lc_raw:
            out_lc = normalizar_cola_lc(f"LC {lc_raw.strip()}")
            out += f" {out_lc}"
        elif re.search(r'\bLC\b', d, flags=re.IGNORECASE):
            seg = re.search(r'\bLC\b.*$', d, flags=re.IGNORECASE).group(0)
            out_lc = normalizar_cola_lc(seg)
            out += f" {out_lc}"
        if tail: out += f" {tail.strip()}"
        if of_text and "OF" not in out: out += f" {of_text}"
        if cn_text and "CN" not in out: out += f" {cn_text}"
        return out, "1"

    m = regex_cra_cl_ap.search(d)
    if m:
        cra, cl, guion = m.group("cra"), m.group("cl"), m.group("guion")
        ap, piso = int(m.group("ap")), m.group("piso")
        out = f"CRA {cra} CL {cl} - {guion} AP {ap}"
        if piso: out += f" PI {int(piso)}"
        if of_text and "OF" not in out: out += f" {of_text}"
        if cn_text and "CN" not in out: out += f" {cn_text}"
        return out, "1"

    m = regex_cra_cl_ap_sin_num.search(d)
    if m:
        cra, cl, guion = m.group("cra"), m.group("cl"), m.group("guion")
        out = f"CRA {cra} CL {cl} - {guion} AP"
        if of_text and "OF" not in out: out += f" {of_text}"
        if cn_text and "CN" not in out: out += f" {cn_text}"
        return out, "1"

    m = regex_cra_cl_macro.search(d)
    if m:
        cra, cl, guion = m.group("cra"), m.group("cl"), m.group("guion")
        macro = int(m.group("macro"))
        out = f"CRA {cra} CL {cl} - {guion} MACRO {macro}"
        if of_text and "OF" not in out: out += f" {of_text}"
        if cn_text and "CN" not in out: out += f" {cn_text}"
        return out, "1"

    m = regex_cra_cl_mz_cs.search(d)
    if m:
        cra, cl, guion = m.group("cra"), m.group("cl"), m.group("guion")
        mz, cs = m.group("mz"), m.group("cs")
        out = f"CRA {cra} CL {cl} - {guion} MZ {mz} CS {cs}"
        if of_text and "OF" not in out: out += f" {of_text}"
        if cn_text and "CN" not in out: out += f" {cn_text}"
        return out, "1"

    m = regex_cra_cl_cs_lc.search(d)
    if m:
        cra, cl, guion = m.group("cra"), m.group("cl"), m.group("guion")
        tipo2, num2, piso = m.group("tipo2"), m.group("num2"), m.group("piso")
        out = f"CRA {cra} CL {cl}"
        if guion: out += f" - {guion}"
        out += f" {tipo2.upper()}"
        if num2:  out += f" {num2}"
        if piso:  out += f" PI {int(piso)}"
        if of_text and "OF" not in out: out += f" {of_text}"
        if cn_text and "CN" not in out: out += f" {cn_text}"
        return out, "1"

    m = regex_cra_cl_basico.search(d)
    if m:
        cra, cl, guion = m.group("cra"), m.group("cl"), m.group("guion")
        tail, piso = m.group("tail"), m.group("piso")
        out = f"CRA {cra} CL {cl} - {guion}"
        if tail: out += f" {tail.strip()}"
        if piso: out += f" PI {int(piso)}"
        if of_text and "OF" not in out: out += f" {of_text}"
        if cn_text and "CN" not in out: out += f" {cn_text}"
        return out, "1"

    m = regex_cra_cl_to_ap.search(d)
    if m:
        cra, cl, guion = m.group("cra"), m.group("cl"), m.group("guion")
        to = m.group("to")
        ap1 = int(m.group("ap1"))
        ap2 = m.group("ap2")
        out = f"CRA {cra} CL {cl} - {guion} TO {to} AP {ap1}"
        if ap2:
            out += f"-{int(ap2)}"
        if of_text and "OF" not in out: out += f" {of_text}"
        if cn_text and "CN" not in out: out += f" {cn_text}"
        return out, "1"

    if "MONTEAZUL" in d:
        m = regex_monteazul.search(d)
        if m:
            ap = m.group("ap") or m.group("ap2")
            to = m.group("to") or m.group("to2")
            out = f"CRA 6 CL 51N - 25 TO {int(to)} AP {int(ap)}"
            if of_text and "OF" not in out: out += f" {of_text}"
            if cn_text and "CN" not in out: out += f" {cn_text}"
            return out, "1"

    # Portal Pradera
    if "PRADERA" in d:
        m = regex_portal_pradera.search(d)
        if m:
            bloque = m.group("bloque")
            apt = int(m.group("apt"))
            out = f"CJT PORTAL PRADERA BQ {bloque} AP {apt}"
            if of_text and "OF" not in out: out += f" {of_text}"
            if cn_text and "CN" not in out: out += f" {cn_text}"
            return out, "1"

    # ZAGUANES (se corrige el bug del script original)
    if "ZAGUANES" in d:
        m = regex_zaguanes.search(d)
        if m:
            manzana = m.group("manzana")
            tipo = m.group("tipo_casa").upper()
            casa = m.group("casa")
            out = f"BRR ZAGUANES MZ {manzana} {tipo} {casa}"
            if of_text and "OF" not in out: out += f" {of_text}"
            if cn_text and "CN" not in out: out += f" {cn_text}"
            return out, "1"

    # CHAMBRANAS
    if "CHAMBRANAS" in d:
        m = regex_chambranas.search(d)
        if m:
            manzana = m.group("manzana")
            tipo = m.group("tipo_casa").upper()
            casa = m.group("casa")
            piso = m.group("piso")
            ap = m.group("ap")
            out = f"BRR CHAMBRANAS MZ {manzana} {tipo} {casa}"
            if ap:   out += f" AP {int(ap)}"
            if piso: out += f" PI {int(piso)}"
            if of_text and "OF" not in out: out += f" {of_text}"
            if cn_text and "CN" not in out: out += f" {cn_text}"
            return out, "1"

    # Sin normalizar
    return d, "0"

# ============================================================
# FUNCIÓN PÚBLICA: procesar(df_in)
# ============================================================

def procesar(df_in: pd.DataFrame) -> pd.DataFrame:
    """
    Normaliza direcciones para PRADERA / ZAGUANES / CHAMBRANAS / MONTEAZUL
    y todas las estructuras CRA/CL y CLL/CR que encajen con los patrones.

    Devuelve un DataFrame con columnas:
    NIU, DIRECCION, DIRECCION_NORMALIZADA, VALIDACION
    """
    if "DIRECCION" not in df_in.columns:
        raise ValueError("El DataFrame de entrada debe contener la columna 'DIRECCION'.")

    # Detectar columna de identificador y renombrarla a NIU
    if "NIU" in df_in.columns:
        id_col = "NIU"
    elif "CLIENTE_ID" in df_in.columns:
        id_col = "CLIENTE_ID"
    else:
        raise ValueError("No se encontró columna de NIU (ni 'NIU' ni 'CLIENTE_ID').")

    df = df_in[[id_col, "DIRECCION"]].copy()
    df.rename(columns={id_col: "NIU"}, inplace=True)

    df[["DIRECCION_NORMALIZADA", "VALIDACION"]] = df["DIRECCION"].apply(
        lambda x: pd.Series(normalizar_direccion(x))
    )

    # Filtro de interés (igual al original, pero sobre este df)
    filtro = df["DIRECCION"].astype(str).str.contains(
        r"PRADERA|ZAGUANES|CHAMBRANAS|MONTEAZUL|CRA\s*\d+\s*CL\s*\d+|CLL\s*\d+\s*(?:CR|CRA|CL)\s*\d+",
        case=False,
        na=False,
    )
    df_filtrado = df[filtro].copy()

    # Regla de duplicados (bases genéricas)
    def es_base_generica(s: str) -> bool:
        if not isinstance(s, str):
            return False
        if re.search(
            r"\b(AP|PI|BQ|TO|ET|CS|LC|MZ|MACRO|MACROMEDIDOR|PU|MOTOBOMBA|OF|ECR|CN)\b", s
        ):
            return False
        if re.match(r"^CRA\s+\d+\s+CL\s+\S+\s+-\s+\S+\s*$", s):
            return True
        if re.match(r"^CLL\s+\S+\s+CR\s+\S+\s+-\s+\S+\s*$", s):
            return True
        return False

    cont = df_filtrado["DIRECCION_NORMALIZADA"].value_counts()
    candidatas = {k for k, v in cont.items() if v > 2 and es_base_generica(k)}
    mask_dup = df_filtrado["DIRECCION_NORMALIZADA"].isin(candidatas)
    df_filtrado.loc[mask_dup, "DIRECCION_NORMALIZADA"] = df_filtrado.loc[mask_dup, "DIRECCION"]
    df_filtrado.loc[mask_dup, "VALIDACION"] = "0"

    return df_filtrado[["NIU", "DIRECCION", "DIRECCION_NORMALIZADA", "VALIDACION"]]

# ============================================================
# MODO SCRIPT: reproduce el comportamiento original
# ============================================================

if __name__ == "__main__":
    ruta_entrada = "CICLO 25_PDIRECCION.xlsx"
    ruta_salida = "CICLOS_PROCESADOS.xlsx"
    hoja_salida = "PRADERA_ZAGUANES_CHAMBRANAS"

    df_in = pd.read_excel(ruta_entrada, dtype=str)
    df_out = procesar(df_in)

    total = df_out.shape[0]
    normalizadas = (df_out["VALIDACION"] == "1").sum()
    efectividad = (normalizadas / total * 100) if total > 0 else 0

    modo = "a" if os.path.exists(ruta_salida) else "w"
    with pd.ExcelWriter(ruta_salida, engine="openpyxl", mode=modo) as writer:
        df_out.to_excel(writer, sheet_name=hoja_salida, index=False)

    print(f"✅ Archivo procesado y guardado como {ruta_salida}")
    print(f"📄 Hoja: {hoja_salida}")
    print(f"🔍 Total direcciones filtradas: {total}")
    print(f"📊 Direcciones normalizadas: {normalizadas}")
    print(f"📈 Efectividad: {round(efectividad, 2)}%")
