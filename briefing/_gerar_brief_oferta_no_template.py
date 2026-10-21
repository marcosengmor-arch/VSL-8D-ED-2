# -*- coding: utf-8 -*-
"""Preenche o template 'BRIEF DE OFERTA' (docx do coordenador) com o conteúdo do brief em markdown,
preservando o layout: Tab 1/Tab 2, títulos, bullets, destaque verde nos nomes, caixa da Pesquisa de Mercado,
barras pretas (Personagens, VSLs, Provas) e tabelas. Edita o document.xml por clonagem de protótipos.
Uso: python3 _gerar_brief_oferta_no_template.py template.docx brief.md saida.docx
"""
import re, sys, zipfile, shutil, os, tempfile, html

TPL, MD, OUT = sys.argv[1], sys.argv[2], sys.argv[3]

# ------------------------------------------------------------------ template
tmp = tempfile.mkdtemp()
with zipfile.ZipFile(TPL) as z: z.extractall(tmp)
docx_xml_path = os.path.join(tmp, 'word', 'document.xml')
X = open(docx_xml_path, encoding='utf8').read()
head, body, tail = X[:X.index('<w:body>') + 8], X[X.index('<w:body>') + 8:X.index('</w:body>')], X[X.index('</w:body>'):]

def tbl_end(s, i):
    d = 0; k = i
    while True:
        a = s.find('<w:tbl>', k); b = s.find('</w:tbl>', k)
        if a != -1 and a < b: d += 1; k = a + 7
        else:
            d -= 1; k = b + 8
            if d == 0: return k

els = []; i = 0; sect = ''
while i < len(body):
    if body.startswith('<w:p ', i) or body.startswith('<w:p>', i):
        j = body.index('</w:p>', i) + 6; els.append(body[i:j]); i = j
    elif body.startswith('<w:tbl>', i):
        e = tbl_end(body, i); els.append(body[i:e]); i = e
    elif body.startswith('<w:sectPr', i):
        sect = body[i:]; break
    else:
        n = body.find('<', i + 1); i = n if n != -1 else len(body)

def rows_of(t):
    out = []; pos = 0
    while True:
        c = [v for v in (t.find('<w:tr ', pos), t.find('<w:tr>', pos)) if v != -1]
        if not c: break
        a = min(c); d = 0; k = a
        while True:
            o = [v for v in (t.find('<w:tr ', k), t.find('<w:tr>', k)) if v != -1]; o = min(o) if o else -1
            cl = t.find('</w:tr>', k)
            if o != -1 and o < cl: d += 1; k = o + 5
            else:
                d -= 1; k = cl + 7
                if d == 0: break
        out.append(t[a:k]); pos = k
    return out

def cells_of(r):
    out = []; pos = 0
    while True:
        a = r.find('<w:tc>', pos)
        if a == -1: break
        d = 0; k = a
        while True:
            o = r.find('<w:tc>', k); cl = r.find('</w:tc>', k)
            if o != -1 and o < cl: d += 1; k = o + 6
            else:
                d -= 1; k = cl + 7
                if d == 0: break
        out.append(r[a:k]); pos = k
    return out

def paras_of(xml):
    return re.findall(r'<w:p[ >].*?</w:p>', xml, re.S)

# ------------------------------------------------------------------ paragraph builder
def esc(s): return html.escape(s, quote=False)

def rpr_of(p):
    m = re.search(r'<w:r [^>]*>\s*<w:rPr>(.*?)</w:rPr>', p, re.S)
    if m: return m.group(1)
    m = re.search(r'<w:pPr>.*?<w:rPr>(.*?)</w:rPr>\s*</w:pPr>', p, re.S)
    return m.group(1) if m else '<w:rtl w:val="0"/>'

def mk_run(text, rpr, bold=None, hl=None):
    r = rpr
    r = re.sub(r'<w:highlight [^>]*/>', '', r)
    if bold is True and '<w:b ' not in r and '<w:b/>' not in r: r = '<w:b w:val="1"/><w:bCs w:val="1"/>' + r
    if bold is False: r = re.sub(r'<w:bCs? [^>]*/>', '', r)
    if hl: r = r + f'<w:highlight w:val="{hl}"/>'
    return f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r>'

def inline_runs(md, rpr, base_bold=None, hl=None):
    md = md.replace('`', '')
    md = re.sub(r'\[([^\]]+)\]\((https?://[^)]+)\)', r'\1 (\2)', md)
    out = []
    for part in re.split(r'(\*\*[^*]+\*\*)', md):
        if not part: continue
        if part.startswith('**') and part.endswith('**'): out.append(mk_run(part[2:-2], rpr, bold=True, hl=hl))
        else: out.append(mk_run(part, rpr, bold=base_bold, hl=hl))
    return ''.join(out)

def P(proto, md_text, bold=None, hl=None, strip_hl=True):
    """Clone proto paragraph, keep pPr, replace content."""
    m = re.match(r'(<w:p[ >].*?<w:pPr>.*?</w:pPr>)', proto, re.S)
    if m: pre = m.group(1)
    else: pre = re.match(r'(<w:p[^>]*>)', proto).group(1)
    pre = re.sub(r'<w:bookmark(Start|End)[^>]*/>', '', pre)
    if strip_hl: pre = re.sub(r'<w:highlight [^>]*/>', '', pre)
    rpr = rpr_of(proto)
    if strip_hl: rpr = re.sub(r'<w:highlight [^>]*/>', '', rpr)
    return pre + inline_runs(md_text, rpr, bold, hl) + '</w:p>'

def set_level(p, lvl):
    p = re.sub(r'<w:ilvl w:val="\d+"/>', f'<w:ilvl w:val="{lvl}"/>', p)
    return re.sub(r'<w:ind w:left="\d+"', f'<w:ind w:left="{720 + 720 * lvl}"', p)

def set_num(p, numid):
    return re.sub(r'<w:numId w:val="\d+"/>', f'<w:numId w:val="{numid}"/>', p)

# prototypes (indexes from the template dump)
TAB1_TITLE, BRIEF_TITLE, PESQ = els[0], els[1], els[2]
H1, H2, NOTE, H3G, H3 = els[3], els[26], els[27], els[28], els[31]
B0, B1 = els[16], els[8]
H1_MARKET, EMPTY, T_MARKET = els[50], els[51], els[52]
TAB2 = els[54:61]
H3BLACK, T_PERS = els[61], els[62]
NUMBOLD, PLAIN, BOLDT, NUM5, NUM12, T_PROV = els[65], els[69], els[87], els[90], els[98], els[109]
TAILS = els[110:]

# extra numbered lists that restart: add numIds 15,16,17 on abstractNum 5
numbering_path = os.path.join(tmp, 'word', 'numbering.xml')
N = open(numbering_path, encoding='utf8').read()
extra = ''.join(f'<w:num w:numId="{n}"><w:abstractNumId w:val="5"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>' for n in (15, 16, 17, 18))
N = N.replace('</w:numbering>', extra + '</w:numbering>')
open(numbering_path, 'w', encoding='utf8').write(N)

# ------------------------------------------------------------------ markdown parsing
lines = open(MD, encoding='utf8').read().split('\n')
sections = {}; order = []; cur = None
def new_sec(title):
    global cur
    cur = title; sections[cur] = []; order.append(cur)
i = 0
while i < len(lines):
    ln = lines[i]; s = ln.strip()
    m = re.match(r'^(#{1,4})\s+(.*)$', s)
    if m: new_sec(m.group(2).replace('`', '').strip()); i += 1; continue
    if cur is None: i += 1; continue
    if s.startswith('|'):
        rows = []
        while i < len(lines) and lines[i].strip().startswith('|'):
            r = lines[i].strip()
            if not re.match(r'^\|?\s*:?-{2,}', r): rows.append([c.strip() for c in r.strip('|').split('|')])
            i += 1
        sections[cur].append(('t', rows)); continue
    mb = re.match(r'^(\s*)[-•]\s+(.*)$', ln)
    if mb: sections[cur].append(('b', len(mb.group(1)) // 2, mb.group(2))); i += 1; continue
    mn = re.match(r'^(\s*)(\d+)\.\s+(.*)$', ln)
    if mn: sections[cur].append(('n', mn.group(2), mn.group(3))); i += 1; continue
    if s and s != '---': sections[cur].append(('p', s))
    i += 1

def sec(prefix):
    for t in order:
        if t.startswith(prefix): return sections[t]
    raise KeyError(prefix)

def items(prefix, kinds=('b',)): return [x for x in sec(prefix) if x[0] in kinds]
def table(prefix): return [x for x in sec(prefix) if x[0] == 't'][0][1]
def paras(prefix): return [x[1] for x in sec(prefix) if x[0] == 'p']

def bullets(blocks, l0=B0, l1=B1):
    out = []
    for b in blocks:
        if b[0] == 'b':
            lvl = b[1]
            out.append(P(l0, b[2]) if lvl == 0 else set_level(P(l1, b[2]), lvl))
        elif b[0] == 'p': out.append(P(PLAIN, b[1]))
        elif b[0] == 'n': out.append(set_level(P(l1, f'{b[1]}. {b[2]}'), 1))
    return out

def numbered(blocks, numid):
    return [set_num(P(NUM5, b[2]), numid) for b in blocks if b[0] == 'n']

def label_text(md):
    """split '**Label:** rest' or 'Label: rest'"""
    m = re.match(r'^\*\*([^*]+?)\*\*:?\s*(.*)$', md)
    if m: return m.group(1).rstrip(':'), m.group(2)
    return None, md

# ------------------------------------------------------------------ tables
def fill_table(t_xml, data, header=None):
    rows = rows_of(t_xml); head_r, proto_r = rows[0], rows[1]
    pre = t_xml[:t_xml.index(rows[0])]; post = t_xml[t_xml.index(rows[-1]) + len(rows[-1]):]
    def fill_row(r, vals, bold=None):
        cs = cells_of(r); out = r[:r.index(cs[0])]
        for ci, c in enumerate(cs):
            v = vals[ci] if ci < len(vals) else ''
            p0 = paras_of(c)[0]
            p0r = re.sub(r'<w:rFonts [^>]*/>', '<w:rFonts w:ascii="Roboto" w:cs="Roboto" w:eastAsia="Roboto" w:hAnsi="Roboto"/>', p0)
            newc = c[:c.index(p0)] + P(p0r, v, bold=bold) + c[c.index(p0) + len(p0):]
            # drop any extra paragraphs beyond first
            for extra_p in paras_of(c)[1:]: newc = newc.replace(extra_p, '', 1)
            out += newc
        return out + r[r.rindex('</w:tc>') + 7:]
    out = pre + fill_row(head_r, header if header else [c for c in data[0]], bold=True)
    for vals in (data[1:] if header is None else data): out += fill_row(proto_r, vals)
    return out + post

# ------------------------------------------------------------------ build Tab 1
out = [TAB1_TITLE, BRIEF_TITLE, PESQ]
first = [x[1] for x in sections.get(order[0], []) if x[0] == 'p'][:0]  # (title line ignored)
# version / legend lines (first section before 'Premissas')
intro = [x[1] for x in sections[order[0]] if x[0] == 'p']
for ptxt in intro: out.append(P(NOTE, ptxt))
out.append(P(H1, 'Premissas'))
out.append(P(els[4], 'Contexto Geral da Oferta'))

# Produto
prod = items('Produto')
out.append(P(els[5], 'Produto:'))
for b in prod: out.append(set_level(P(B1, b[2]), 1))
# Promessa
out.append(P(B0, 'Promessa principal:'))
for b in items('Promessa principal'): out.append(set_level(P(B1, b[2]), 1))
# Transformação
out.append(P(B0, 'Transformação esperada (marcos, as duas opções):'))
for b in items('Transformação esperada'): out.append(set_level(P(B1, b[2]), 1))
# Público
out.append(P(B0, 'Público-alvo (detalhado, as duas opções):'))
for b in items('Público-alvo'): out.append(set_level(P(B1, b[2]), 1))

# Três ideias
out.append(P(H2, 'Três Ideias de Ângulo Principal (Top 3 Big Ideas)'))
out.append(P(NOTE, '(Cada uma deve ser clara, única e com potencial de escalabilidade. Tudo sem marca é igual nas duas opções; [OP1] e [OP2] marcam o que difere.)'))
ideias = [t for t in order if t.startswith('Ideia ')]
for n, t in enumerate(ideias):
    out.append(P(H3G if n == 0 else H3, t, hl='green' if n == 0 else None, strip_hl=(n != 0)))
    for b in items(t): out.append(P(els[29], b[2]))
out.append(P(H3, 'Leads para teste (A/B, como na VSL 1)'))
for r in table('Leads para teste')[1:]:
    out.append(P(els[29], f'**Lead {r[0]}:** {r[1]} — Opção {r[2]}'))

# Mecanismos
out.append(P(H2, 'Mecanismos – Problema e Solução (uma por opção + reserva)'))
out.append(P(NOTE, '(Cada mecanismo deve explicar a causa real do problema + o PORQUÊ o produto resolve)'))
for t in [t for t in order if t.startswith('Mecanismo ')]:
    out.append(P(H3, t))
    for b in items(t): out.append(P(els[40], b[2]))

# Nome chiclete
out.append(P(H2, 'Nome "Chiclete" (Hipóteses de Nome para Oferta ou Produto)'))
out.append(P(NOTE, '(Opções fáceis de lembrar e comunicar; o animal vem antes do veículo)'))
nome_rows = table('Nome')[1:]
funcs = {r[0]: r for r in nome_rows}
def nr(key): return [r for r in nome_rows if r[0].startswith(key)][0]
r_truque, r_prod, r_gar = nr('Truque'), nr('Produto'), nr('Garantia')
out.append(P(H3G, f'Nome 1 [OP1]: {r_truque[1].replace("**","")} / {r_prod[1].replace("**","")} / {r_gar[1].replace("**","")}', hl='green', strip_hl=False))
out.append(P(els[47], f'Justificativa: {r_truque[3]}. Produto: {r_prod[3]}. Garantia: {r_gar[3]}.'))
out.append(P(H3, f'Nome 2 [OP2]: {r_truque[2].replace("**","")} / {r_prod[2].replace("**","")} / {r_gar[2].replace("**","")}'))
out.append(P(els[47], f'Justificativa: {r_truque[3]}. Produto: {r_prod[3]}.'))
others = [r for r in nome_rows if r not in (r_truque, r_prod, r_gar)]
out.append(P(H3, 'Nome 3: variações de lugar, hook e nome de boca'))
for r in others: out.append(P(els[47], f'{r[0]}: [OP1] {r[1]} · [OP2] {r[2]} — {r[3]}'))
for ptxt in paras('Nome'): out.append(P(PLAIN, ptxt))

# Pesquisa de mercado (box table)
out.append(P(H1_MARKET, 'Pesquisa de Mercado'))
out.append(EMPTY)
mrows = rows_of(T_MARKET)
def replace_cell_content(row, new_paras_xml):
    c = cells_of(row)[0]; ps = paras_of(c)
    start = c.index(ps[0]); end = c.rindex(ps[-1]) + len(ps[-1])
    newc = c[:start] + new_paras_xml + c[end:]
    return row.replace(c, newc)
# r2 content prototypes
r2c = cells_of(mrows[2])[0]; r2p = paras_of(r2c)
LABEL, EMPTYC, PLAINC = r2p[0], r2p[1], r2p[2]
PLAINC = re.sub(r'<w:hyperlink [^>]*>.*?</w:hyperlink>', '', PLAINC, flags=re.S)
cont = []
for b in items('Tendências de Conteúdo'):
    lab, rest = label_text(b[2])
    if not lab:
        mm = re.match(r'^([^:*]{3,30}):\s+(.*)$', b[2])
        if mm: lab, rest = mm.group(1), mm.group(2)
    if lab and len(lab) < 40:
        cont.append(P(LABEL, lab + ':')); cont.append(P(PLAINC, rest)); cont.append(P(EMPTYC, ''))
    else:
        cont.append(P(PLAINC, b[2])); cont.append(P(EMPTYC, ''))
r2 = replace_cell_content(mrows[2], ''.join(cont))
# r6 marketing bullets
r6c = cells_of(mrows[6])[0]; MB = paras_of(r6c)[0]
mk = []
for n, b in enumerate(items('Tendências de Marketing')):
    lab, rest = label_text(b[2])
    hl = 'green' if n < 4 else ('yellow' if 'Não fazer' in b[2] else None)
    if lab: mk.append(P(MB, '', ) .replace('</w:p>', inline_runs(lab + '. ', rpr_of(MB), base_bold=True, hl=hl) + inline_runs(rest, rpr_of(MB)) + '</w:p>'))
    else: mk.append(P(MB, b[2]))
r6 = replace_cell_content(mrows[6], ''.join(mk))
# r8 explicação das VSLs: black bar + nested table
r8c = cells_of(mrows[8])[0]; r8p = paras_of(re.sub(r'<w:tbl>.*</w:tbl>', '', r8c, flags=re.S))
H3B_IN = r8p[0]
nested0 = r8c[r8c.index('<w:tbl>'):tbl_end(r8c, r8c.index('<w:tbl>'))]
vsl_rows = table('Explicação das VSLs')
nested_new = fill_table(nested0, vsl_rows)
r8_content = P(H3B_IN, 'VSLs de referência (o que cada uma empresta)') + nested_new + P(EMPTYC, '') + P(PLAINC, paras('Explicação das VSLs')[0] if paras('Explicação das VSLs') else '') + P(EMPTYC, '')
start = r8c.index(r8p[0]); c_new = r8c[:start] + r8_content + '</w:tc>'
r8 = mrows[8].replace(r8c, c_new)
tm = T_MARKET[:T_MARKET.index(mrows[0])] + mrows[0] + mrows[1] + r2 + mrows[3] + mrows[4] + mrows[5] + r6 + mrows[7] + r8 + T_MARKET[T_MARKET.index(mrows[-1]) + len(mrows[-1]):]
out.append(tm)
out.append(EMPTY)

# ------------------------------------------------------------------ Tab 2
out += TAB2
out.append(P(H3BLACK, 'Personagens (as duas opções)'))
out.append(fill_table(T_PERS, table('Personagens')))
out += [EMPTY, EMPTY]
out.append(P(NUMBOLD, 'Resumo de como o problema aconteceu até o fundo do poço (Pacino, pegada do Costner)'))
out.append(EMPTY)
for blk in sec('Resumo de como'):
    if blk[0] == 'p': out.append(P(PLAIN, blk[1])); out.append(EMPTY)
    elif blk[0] == 'b': out.append(P(B0, blk[2]))
out.append(P(NUMBOLD, 'Fundo do poço (a virada)'))
out.append(EMPTY)
for blk in sec('Fundo do poço'):
    if blk[0] == 'p': out.append(P(PLAIN, blk[1])); out.append(EMPTY)
out.append(P(BOLDT, 'PONTOS LÓGICOS (Problema e Solução):'))
out.append(EMPTY)
out.append(P(H3, 'Explicação por A + B porque o problema existe'))
for ptxt in paras('Explicação por A + B porque o problema'): out.append(P(NOTE, ptxt))
out += numbered(sec('Explicação por A + B porque o problema'), 5)
out.append(P(H3, 'Explicação por A + B porque a solução funciona (e é a única solução que funciona)'))
out += numbered(sec('Explicação por A + B porque a solução'), 12)
out += [EMPTY, EMPTY]
out.append(P(H3BLACK, 'Provas porque isso funciona'))
out.append(fill_table(T_PROV, table('Provas porque isso funciona')))
out += [EMPTY, EMPTY]

# Oferta e close
out.append(P(BOLDT, 'OFERTA E CLOSE (resumo operacional, as duas opções):'))
out.append(EMPTY)
out += bullets(items('Oferta e Close'))
out += [EMPTY, EMPTY]
# Desqualificação
out.append(P(BOLDT, 'DESQUALIFICAÇÃO (cavalo, prateleira e truques concorrentes):'))
out.append(EMPTY)
for ptxt in paras('Desqualificação'): out.append(P(PLAIN, ptxt)); out.append(EMPTY)
out.append(P(H3BLACK, 'Munição'))
mun = table('Munição')
out.append(fill_table(T_PROV, [[r[2], r[1], r[3]] for r in mun[1:]], header=['Status', 'Fato', 'Uso na copy']))
out.append(EMPTY)
for t in ('Onde entra na lead', 'Onde entra no MUS', 'Edição (vai para o briefing)'):
    out.append(P(H3, t)); out += bullets(sec(t))
out += [EMPTY, EMPTY]
out.append(P(BOLDT, 'AJUSTES PARA A COPY (vindos do briefing de edição, valem para as duas):'))
out.append(EMPTY)
out += numbered(sec('Ajustes para a copy'), 15)
out += [EMPTY, EMPTY]
out.append(P(BOLDT, 'REGRA DE ESTILO DA COPY:'))
out.append(EMPTY)
out += bullets(items('Regra de estilo'))
out += [EMPTY, EMPTY]
out.append(P(BOLDT, 'PENDÊNCIAS E DECISÕES EM ABERTO:'))
out.append(EMPTY)
out += numbered(sec('Pendências'), 16)
out += TAILS

# ------------------------------------------------------------------ write
new_body = ''.join(out)
new_body = re.sub(r'<w:bookmark(Start|End)[^>]*/>', '', new_body)
# unique paraIds
counter = [0x10000000]
def pid(m):
    counter[0] += 1; return f'w14:paraId="{counter[0]:08X}"'
new_body = re.sub(r'w14:paraId="[0-9A-Fa-f]+"', pid, new_body)
X2 = head + new_body + sect + tail
open(docx_xml_path, 'w', encoding='utf8').write(X2)
if os.path.exists(OUT): os.remove(OUT)
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    for root, _, files in os.walk(tmp):
        for f in files:
            full = os.path.join(root, f); z.write(full, os.path.relpath(full, tmp))
shutil.rmtree(tmp)
print('ok', OUT, 'elements:', len(out))
