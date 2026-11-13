#!/usr/bin/env python3
"""
Gera vinhetas clínicas padronizadas a partir da planilha codificada do TCC.

Uso:
    python3 gerar_vinhetas.py TCC_dicionario_reorganizado_v6.1.xlsx
    python3 gerar_vinhetas.py planilha.xlsx --sem-radiografia=""          # ignora a lista abaixo
    python3 gerar_vinhetas.py planilha.xlsx --sem-radiografia=P04,P06     # outra lista

Saída (na mesma pasta da planilha):
    vinhetas.xlsx  – ID + comando + vinheta, uma linha por caso
    vinhetas.txt   – as mesmas vinhetas em texto corrido

O texto é montado por regras fixas a partir do dicionário (aba com Letra/Código/
Significado/Subgrupo), sem nenhuma IA: o mesmo código gera sempre o mesmo texto.
Novos códigos entram automaticamente, desde que recebam um dos subgrupos já usados.
Requer: pip install openpyxl
"""
import sys, re
from pathlib import Path
import openpyxl

COMANDO = "Preciso que, baseado nesses dados, você elabore as principais hipóteses diagnósticas, em ordem de probabilidade:"
INCLUIR_RADIOGRAFIA = True   # acrescenta "Radiografia da região em anexo." quando o caso conta como "com radiografia"

# Casos com H = 2 na planilha cuja radiografia foi considerada INSATISFATÓRIA: contam como
# "sem radiografia" na vinheta (nenhuma menção a imagem em anexo e nenhuma imagem enviada).
# A planilha não é alterada; a regra vale só para a geração do texto.
SEM_RADIOGRAFIA_QUALIDADE = {"P04", "P20"}

# Concordância dos termos de coloração (D37–D41) com a palavra "coloração".
COLORACAO_FEM = {"arroxeado": "arroxeada", "azulado": "azulada",
                 "normocrômico": "normocrômica", "amarelado": "amarelada",
                 "eritematoso/avermelhado": "eritematosa/avermelhada"}

FEMININO = {"diabético": "diabética", "diabético descompensado": "diabética descompensada",
            "hipertenso": "hipertensa", "imunossuprimido": "imunossuprimida",
            "etilista crônico": "etilista crônica"}
ADJETIVOS = {"tabagista", "ex-tabagista", "etilista crônico", "gestante", "diabético",
             "diabético descompensado", "hipertenso", "imunossuprimido"}
LOCAL_TEXTO = {"ATM/região pré-auricular": "ATM/pré-auricular"}


def low(s):
    s = " ".join(str(s).split())   # minúsculas, preservando siglas (HPV, ISRS, ATM, TGI, LP)
    return re.sub(r"\w+", lambda m: m.group() if (m.group().isupper() and len(m.group()) > 1) else m.group().lower(), s)


def lista(itens):
    itens = [i for i in itens if i]
    return itens[0] if len(itens) == 1 else ", ".join(itens[:-1]) + " e " + itens[-1]


def achar_abas(wb):
    dic = dados = None
    for ws in wb:
        h = [str(c.value or "").strip() for c in ws[1]]
        if dic is None and {"Letra", "Código", "Significado"} <= set(h):
            dic = ws
        if dados is None and "ID" in h and any(x.startswith("D ") for x in h):   # 1ª aba de dados
            dados = ws
    if not dic or not dados:
        sys.exit("Não encontrei a aba do dicionário e/ou a aba de dados.")
    return dic, dados


def ler_dicionario(ws):
    h = [str(c.value or "").strip() for c in ws[1]]
    iL, iC, iS = h.index("Letra"), h.index("Código"), h.index("Significado")
    iG = h.index("Subgrupo") if "Subgrupo" in h else None
    d, letra = {}, None
    for r in ws.iter_rows(min_row=2, values_only=True):
        letra = r[iL] or letra
        try:
            cod = int(float(r[iC]))
        except (TypeError, ValueError):
            continue
        d[(letra, cod)] = (low(r[iS]), (r[iG] or "") if iG is not None else "")
    return d


def codigos(v):
    if v is None:
        return []
    return sorted({int(float(x)) for x in re.split(r"[;,]", str(v)) if x.strip()})


def com_radiografia(pid, caso, sem_rad=SEM_RADIOGRAFIA_QUALIDADE):
    """True se o caso tem radiografia (H = 2) E ela não foi descartada por qualidade."""
    return caso.get("H") == [2] and pid not in sem_rad


def vinheta(caso, D, radiografia=None):
    """radiografia: True/False força a frase; None usa H = 2 (comportamento antigo)."""
    t = lambda L: [D[(L, c)] for c in caso[L]]          # [(termo, subgrupo)]
    fem = caso["A"] == [1]
    frases = []

    # Localização (E)
    locais = [LOCAL_TEXTO.get(s, s) for s, _ in t("E")]
    difuso = "difuso" in locais
    locais = [l for l in locais if l != "difuso"]
    onde = ("em região de " + lista(locais)) if locais else ""
    if difuso:
        onde = (onde + ", de forma difusa") if onde else "de forma difusa"

    # 1. Identificação + queixa (com o local da queixa, quando houver sintoma)
    sexo = "feminino" if fem else "masculino"
    sint = [s for s, _ in t("C")]
    assintomatico = not sint or caso["C"] == [1]
    # sintoma que já traz o próprio local (ex.: "dor em face") mantém o local no exame
    local_na_queixa = (not assintomatico) and bool(onde) and not any(" em " in s for s in sint)
    if assintomatico:
        queixa = "sem sintomatologia"
    else:
        queixa = "relata " + lista(sint) + (f" {onde}" if local_na_queixa else "")
    frases.append(f"Paciente do sexo {sexo}, {caso['B']} anos, {queixa}.")

    # 2. Exame clínico: achado principal + localização + sinais associados
    exame = "Ao exame clínico da região" if local_na_queixa else "Ao exame clínico"
    onde_ex = "" if local_na_queixa else (f" {onde}" if onde else "")

    d = t("D")
    lesoes = [s for s, g in d if g == "Lesão fundamental"]
    outros = [s for s, g in d if g.startswith(("Alteração", "ATM"))]
    sinais = [s for s, g in d if g.startswith("Sinais")]
    carac = {}                      # subgrupo -> lista de termos (um subgrupo pode ter vários)
    for s, g in d:
        if g.startswith("Característica"):
            carac.setdefault(g.split("–")[-1].strip(), []).append(s)

    if caso["D"] == [1]:
        if local_na_queixa or not locais:
            frases.append(f"{exame}, não foram observadas alterações.")
        else:
            frases.append(f"Ao exame clínico da região de {lista(locais)}, não foram observadas alterações.")
    elif lesoes or outros:
        frase = f"{exame}, observa-se presença de {lista(lesoes + outros)}{onde_ex}"
        frase += f", com {lista(sinais)}" if sinais else ""
        frases.append(frase + ".")
    elif sinais:
        frases.append(f"{exame}, observa-se presença de {lista(sinais)}{onde_ex}.")

    # 3. Características da lesão
    if carac:
        plural = "múltiplos" in carac.get("número", [])
        if lesoes:
            sujeito = "Lesões" if plural else "Lesão"
            if "número" in carac:
                sujeito += " múltiplas" if plural else " única"
        else:
            sujeito = "Quadro"

        def agrupado(chave, substantivo):
            """'superfície estriada' + 'superfície rugosa' -> 'superfície estriada e rugosa'."""
            termos = carac.get(chave, [])
            if not termos:
                return None
            adj = [t[len(substantivo):].strip() if t.lower().startswith(substantivo) else t
                   for t in termos]
            return f"{substantivo} {lista(adj)}"

        meio = []
        for k in ("limites", "base", "consistência"):
            if k in carac:
                meio.append(lista(carac[k]))
        for k, subst in (("superfície", "superfície"), ("bordas", "bordas")):
            p = agrupado(k, subst)
            if p:
                meio.append(p)
        if "coloração" in carac:
            meio.append("coloração " + lista([COLORACAO_FEM.get(c, c)
                                              for c in carac["coloração"]]))

        # a vírgula só cabe quando o sujeito tem qualificador ("Lesão única, de ...");
        # em "Lesão" ou "Quadro" sozinhos ela fica solta
        liga = ", de " if (" " in sujeito and meio) else (" de " if meio else "")
        frase = sujeito + (liga + lista(meio) if meio else "")
        if "tempo de evolução" in carac:
            tempo = carac["tempo de evolução"][0]
            if plural and sujeito.startswith("Lesões"):
                tempo = tempo.replace("presente", "presentes", 1)
            frase += ("," if meio else "") + " " + tempo
        frases.append(frase + ".")

    # 4. Fatores locais
    if caso["F"] and caso["F"] != [1]:
        frases.append("Como fatores locais, apresenta " + lista([s for s, _ in t("F")]) + ".")
    elif caso["F"] == [1]:
        frases.append("Sem fatores locais modificadores.")

    # 5. Fatores sistêmicos: condições → histórico → medicamentos
    if caso["G"] == [1]:
        frases.append("Sem fatores sistêmicos associados.")
    elif caso["G"]:
        g = t("G")
        adj = [FEMININO.get(s, s) if fem else s for s, _ in g if s in ADJETIVOS]
        hist = [s for s, grp in g if s not in ADJETIVOS and not grp.startswith("Medicamento")]
        meds = [re.sub(r"^medicamentos? (da classe d[oa]s |para )", "", s)
                for s, grp in g if grp.startswith("Medicamento")]
        if adj:
            frases.append("Sistemicamente, é " + lista(adj) + ".")
        if hist:
            frases.append(("Apresenta" if not adj else "Apresenta também") + " histórico de " + lista(hist) + ".")
        if meds:
            frases.append("Faz uso de " + lista(meds) + ".")

    # 6. Radiografia (só se o caso conta como "com radiografia" — ver SEM_RADIOGRAFIA_QUALIDADE)
    tem_rad = (caso.get("H") == [2]) if radiografia is None else radiografia
    if INCLUIR_RADIOGRAFIA and tem_rad:
        frases.append("Radiografia da região em anexo.")

    return " ".join(frases)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    opts = [a for a in sys.argv[1:] if a.startswith("--")]
    if not args:
        sys.exit(__doc__)
    sem_rad = set(SEM_RADIOGRAFIA_QUALIDADE)
    for o in opts:
        if o.startswith("--sem-radiografia="):
            sem_rad = {x.strip().upper() for x in o.split("=", 1)[1].split(",") if x.strip()}
    src = Path(args[0])
    wb = openpyxl.load_workbook(src, data_only=True)
    wdic, wdat = achar_abas(wb)
    D = ler_dicionario(wdic)
    h = [str(c.value or "").strip() for c in wdat[1]]
    col = {L: next(i for i, x in enumerate(h) if x.startswith(L + " ")) for L in "ABCDEFGH"}
    iid = h.index("ID")

    saida, descartadas = [], []
    for r in wdat.iter_rows(min_row=2, values_only=True):
        if not r[iid]:
            continue
        caso = {L: codigos(r[col[L]]) for L in "ACDEFGH"}
        caso["B"] = int(float(r[col["B"]])) if r[col["B"]] is not None else "?"
        faltando = [f"{L}{c}" for L in "ACDEFGH" for c in caso[L] if (L, c) not in D]
        if faltando:
            sys.exit(f"{r[iid]}: códigos sem correspondência no dicionário: {faltando}")
        rad = com_radiografia(r[iid], caso, sem_rad)
        if caso["H"] == [2] and not rad:
            descartadas.append(r[iid])
        saida.append((r[iid], vinheta(caso, D, radiografia=rad)))

    out = openpyxl.Workbook(); ws = out.active; ws.title = "Vinhetas"
    ws.append(["ID", "Comando", "Vinheta"])
    for pid, v in saida:
        ws.append([pid, COMANDO, v])
    ws.column_dimensions["B"].width = 30; ws.column_dimensions["C"].width = 120
    out.save(src.with_name("vinhetas.xlsx"))
    with open(src.with_name("vinhetas.txt"), "w", encoding="utf-8") as f:
        for pid, v in saida:
            f.write(f"{pid}\n{COMANDO}\n{v}\n\n")
    print(f"{len(saida)} vinhetas geradas em {src.with_name('vinhetas.xlsx')} e vinhetas.txt")
    if descartadas:
        print("Radiografia descartada por qualidade (tratados como sem radiografia):", ", ".join(descartadas))


if __name__ == "__main__":
    main()
