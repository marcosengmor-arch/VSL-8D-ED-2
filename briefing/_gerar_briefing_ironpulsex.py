# -*- coding: utf-8 -*-
# Gera BRIEFING_Edicao_IronPulseX.docx no layout exato do template do coordenador
# (usa o docx do briefing da Bull Soda Rise como base e troca o conteúdo), e acrescenta as abas da copy EN-US
# com tags em negrito, orientações em PT em itálico e caixas de Referência Visual (com frames das VSLs modelo).
import copy, re, sys, os
import docx
from docx.shared import Inches, Pt
from docx.oxml.ns import qn
from docx.table import Table, _Row
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _conteudo_briefing import *
from _plano_copy import PLANO

ROOT='/home/user/vsl-8d-ed-2'
TEMPLATE='/home/user/Teste-ED-Project/analise-ed/briefing/BRIEFING_Edicao_BullSodaRise.docx'
OUT=f'{ROOT}/briefing/BRIEFING_Edicao_IronPulseX.docx'
HF=f'{ROOT}/00_base/referencia_honeyfil_male/frames/'
RH=f'{ROOT}/00_base/referencia_red_horse/frames_lead/'
IMG={'hf_news':HF+'b1-4_apresentador-de-telejornal-elon-musk-d_00-00-54.jpg','hf_fire':HF+'b1-9_cybertruck-em-chamas-a-noite_00-04-24.jpg',
 'hf_soda':HF+'b1-10_caixa-de-bicarbonato-arm-hammer_00-05-21.jpg','hf_bee':HF+'b13-1_coletor-de-mel-na-arvore-puxando-favo_00-30-51.jpg',
 'hf_docbottle':HF+'b9-2_dr-michael-carter-segurando-o-frasco-h_00-24-31.jpg','hf_doc':HF+'b10-6_dr-michael-carter-no-consultorio_00-26-43.jpg',
 'hf_3dplaque':HF+'b6-1_animacao-3d-de-tecido-fibroso-com-plac_00-18-47.jpg','hf_3dsponge':HF+'b3-2_animacao-de-tecido-esponjoso-vasos-com_00-11-26.jpg',
 'hf_3dvessel':HF+'b12-1_animacao-3d-de-vaso-com-placas-escuras_00-28-38.jpg','hf_bottles':HF+'b10-6_seis-frascos-honeyfil-male-numa-bancad_00-26-34.jpg',
 'hf_empty':HF+'b17-3_deposito-com-prateleiras-quase-vazias_00-40-09.jpg','hf_gift':HF+'b17-3_caixa-de-presente-vermelha-com-laco-do_00-40-22.jpg',
 'hf_duo':HF+'b2-5_elon-musk-e-medico-de-scrubs-pretos-dr_00-09-46.jpg','hf_armchair':HF+'b22-1_elon-musk-na-poltrona_00-49-28.jpg',
 'hf_lab':HF+'b4-1_dois-cientistas-de-macacao-branco-e-ma_00-12-48.jpg','hf_micro':HF+'b4-3_cientista-de-jaleco-olhando-no-microsc_00-14-53.jpg',
 'hf_hacker':HF+'b19-1_hacker-digitando-em-laptop-com-codigo_00-45-21.jpg','hf_box':HF+'b5-2_homem-abrindo-uma-caixa-de-papelao-na_00-17-33.jpg',
 'hf_ugcman':HF+'b7-1_depoimento-homem-idoso-de-cabelo-branc_00-20-16.jpg','hf_ugcwoman':HF+'b13-3_depoimento-mulher-morena-com-toalha-na_00-32-21.jpg',
 'hf_study':HF+'b3-1_capa-da-revista-nature-24-april-2023_00-10-51.jpg','hf_field':HF+'b9-1_elon-musk-vestido-de-campo-segurando-f_00-23-38.jpg',
 'hf_bedphone':HF+'b18-3_elon-musk-na-cama-a-noite-olhando-o-ce_00-44-16.jpg',
 'rh_kitchen':RH+'f_00-00.jpg','rh_horse':RH+'f_01-42.jpg','rh_tv':RH+'f_02-49.jpg'}

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
d=docx.Document(TEMPLATE)
T=d.tables
assert len(T)==49, len(T)

def _base_run(p):
    for r in p.findall(W+'r'):
        if r.find(W+'t') is not None: return r
    rs=p.findall(W+'r')
    return rs[0] if rs else None

def _mk_run(base, text, bold=False, italic=False):
    r=copy.deepcopy(base) if base is not None else docx.oxml.OxmlElement('w:r')
    for ch in list(r):
        if ch.tag!=W+'rPr': r.remove(ch)
    rpr=r.find(W+'rPr')
    if rpr is None:
        rpr=docx.oxml.OxmlElement('w:rPr'); r.insert(0,rpr)
    for tag in ('b','bCs','i','iCs','color','u'):
        for e in rpr.findall(W+tag): rpr.remove(e)
    if bold:
        for tag in ('bCs','b'):
            e=docx.oxml.OxmlElement('w:'+tag); e.set(qn('w:val'),'1'); rpr.insert(1 if rpr.find(W+'rFonts') is not None else 0,e)
    if italic:
        e=docx.oxml.OxmlElement('w:i'); e.set(qn('w:val'),'1'); rpr.append(e)
    t=docx.oxml.OxmlElement('w:t'); t.set(qn('xml:space'),'preserve'); t.text=text; r.append(t)
    return r

def fill_par(p, text):
    base=_base_run(p)
    for r in p.findall(W+'r')+p.findall(W+'hyperlink'): p.remove(r)
    for seg in re.split(r'(\*\*.*?\*\*)', text):
        if not seg: continue
        b=seg.startswith('**') and seg.endswith('**')
        p.append(_mk_run(base, seg[2:-2] if b else seg, bold=b))

def fill_cell(cell, text):
    ps=cell._tc.findall(W+'p')
    tpl=ps[0]
    for p in ps[1:]: cell._tc.remove(p)
    lines=text.split('\n')
    fill_par(tpl, lines[0])
    last=tpl
    for line in lines[1:]:
        np_=copy.deepcopy(tpl); fill_par(np_, line); last.addnext(np_); last=np_

def rebuild(table, rows, header_rows=1, drop_dup_header=False):
    trs=table._tbl.findall(W+'tr')
    if drop_dup_header and len(trs)>1: table._tbl.remove(trs[1]); trs=table._tbl.findall(W+'tr')
    tpl=copy.deepcopy(trs[header_rows])
    for tr in trs[header_rows:]: table._tbl.remove(tr)
    for vals in rows:
        tr=copy.deepcopy(tpl); table._tbl.append(tr)
        for c,v in zip(_Row(tr,table).cells, vals): fill_cell(c, v)

# ---------- Seção 1 ----------
fill_cell(T[2].cell(0,0), TITULO)
for p in d.paragraphs:
    if p.text.startswith('Misto — especificar'):
        fill_par(p._p, FORMATO_MISTO); break
fill_cell(T[3].cell(0,0), REF_PRINCIPAL)
fill_cell(T[4].cell(0,0), REF_SECUNDARIA)
fill_cell(T[5].cell(0,0), PERS1); fill_cell(T[6].cell(0,0), PERS2); fill_cell(T[7].cell(0,0), PERS3)
fill_cell(T[8].cell(0,0), APOIO); fill_cell(T[9].cell(0,0), CENARIO); fill_cell(T[10].cell(0,0), RITMO)
# ---------- Seção 2 ----------
rebuild(T[12], AVATARES, header_rows=1, drop_dup_header=True)
fill_cell(T[13].cell(0,0), REDES)
# ---------- Seção 3 ----------
fill_cell(T[15].cell(0,0), ESTILO_BROLL)
rebuild(T[16], BROLL_TIPOS)
fill_cell(T[17].cell(0,0), BENCH); fill_cell(T[18].cell(0,0), GRAFICOS); fill_cell(T[19].cell(0,0), AJUSTES)
# ---------- Seção 4 ----------
for k,(nome,obj,fontes) in enumerate(BLOCOS):
    fill_cell(T[21+3*k].cell(0,0), nome); fill_cell(T[22+3*k].cell(0,0), obj); fill_cell(T[23+3*k].cell(0,0), fontes)
# ---------- Seção 5 ----------
rebuild(T[43], LEADS)

# ---------- Seção 6: tempos a 190 palavras/min sobre a copy EN-US ----------
EN=f'{ROOT}/copy/EN-US/'
def body(f): return open(EN+f).read().split('\n',2)[2]
def wc(t): return sum(len(p.split()) for p in t.split('\n\n') if p.strip() and not re.fullmatch(r'\[.*\]',p.strip()))
def parts(t, anchors):
    ps=[p for p in t.split('\n\n') if p.strip()]; idx=[0]
    for a in anchors:
        i=next(k for k,p in enumerate(ps) if p.startswith(a))
        if i>0 and re.fullmatch(r'\[.*\]',ps[i-1].strip()): i-=1
        idx.append(i)
    idx.append(len(ps)); return [wc('\n\n'.join(ps[idx[j]:idx[j+1]])) for j in range(len(idx)-1)]
WPM=190
def mmss(words): s=round(words*60/WPM); return f'{s//60}:{s%60:02d}'
L=[wc(body(f'EN_0{i}_Lead_0{i}.md')) for i in (1,2,3)]
BG=parts(body('EN_04_Background_Story.md'),['Al. You called me.','I went after those 2 answers'])
MT=parts(body('EN_05_Marketing_Thesis.md'),["And here's what almost nobody sees",'To dissolve the cement and never'])
BU=wc(body('EN_06_Product_Build_Up.md'))
BO=parts(body('EN_07_Big_Offer.md'),["That's why I put together this special offer",'And because I want you to walk','As soon as you lock in'])
_bo=body('EN_07_Big_Offer.md'); TO49=wc(_bo[:_bo.find('$49')])
CL=wc(body('EN_08_Close.md')); FQ=wc(body('EN_09_FAQ.md'))
acc=[L[0]]
def nxt(n): acc.append(acc[-1]+n); return acc[-1]
t_bg=acc[-1]; t_emo=nxt(BG[0]); t_disc=nxt(BG[1]); t_mt=nxt(BG[2]); t_mtf=t_mt+MT[0]; t_mts=t_mtf+MT[1]; t_bu=nxt(sum(MT)); t_bo=nxt(BU)
t_49=t_bo+TO49; t_pitch=t_bo+BO[0]; t_cl=nxt(sum(BO)); t_faq=nxt(CL)
fill_cell(T[45].cell(0,0), '[inserir link: HONEYFIL MALE — vídeo .mp4 + transcrição EN + tradução PT-BR + 281 prints] — tempos abaixo estimados a 190 palavras/minuto sobre a versão EN-US (copy/EN-US/), '
 f'com a Lead 01 ({L[0]} palavras); com a Lead 02 ({L[1]}) ou a Lead 03 ({L[2]}), deslocar o início dos blocos em {"+" if L[1]>=L[0] else "−"}{mmss(abs(L[1]-L[0]))} / {"+" if L[2]>=L[0] else "−"}{mmss(abs(L[2]-L[0]))}.')
TIMELINE=[
 ('Lead', f'0:00 – ~{mmss(L[0])} (Lead 01 Abella + Kendra; Lead 02 ~{mmss(L[1])}; Lead 03 ~{mmss(L[2])})'),
 ('Background Story', f'~{mmss(t_bg)} – ~{mmss(t_emo)} (credenciais do Hyman, depoimentos 01–03, a ligação do Al, por que ele aceitou gravar)'),
 ('Emotional Story', f'~{mmss(t_emo)} – ~{mmss(t_disc)} (o Al na sala: Viagra, Cialis, a noite da terça, "Tudo bem, Al", TRT, personal, câmera da cozinha; o lado da Noor)'),
 ('Discovery Story', f'~{mmss(t_disc)} – ~{mmss(t_mt)} (73 estudos, glifosato/Roundup, RFK, memorando Bayer, FDA, $32 bi, $47 mi, o Gupta, o bode no cocho)'),
 ('Marketing Thesis', f'~{mmss(t_mt)} – ~{mmss(t_bu)} (Problem {mmss(t_mt)}–{mmss(t_mtf)} · Function {mmss(t_mtf)}–{mmss(t_mts)} · Solution {mmss(t_mts)}–{mmss(t_bu)}, os 4 verbos + cronograma)'),
 ('Product Build Up', f'~{mmss(t_bu)} – ~{mmss(t_bo)} (pote do Gupta, 14 de janeiro, dias 4/11/21, RM, namorada grávida, bodes irmãos, 117 homens, Brandon + selo, 14 labs, Caldwell, trial 3.000, reveal)'),
 ('Offer', f'~{mmss(t_bo)} – ~{mmss(t_cl)} (apresentação, Oz, atriz/Mick Blue, Ford/Van Damme, Diane, 67 frascos, Whitman; pitch a partir de ~{mmss(t_pitch)} — $49 em ~{mmss(t_49)} —, 15 minutos, o Al e o iate, bônus, GOAT Guarantee)'),
 ('Close', f'~{mmss(t_cl)} – ~{mmss(t_faq)} (o Al abre + dois caminhos + recap + 3 motivos + CTA "Today.")'),
 ('FAQ (se houver)', f'Após ~{mmss(t_faq)} — segmento anexo (~{mmss(FQ)}), 6 perguntas de clientes lidas pelo Hyman'),
]
rebuild(T[46], TIMELINE)
# ---------- Seção 7 ----------
fill_cell(T[48].cell(0,0), CRITICOS)

# ---------- Abas da copy ----------
SEC_HDR=T[1]._tbl; BLK_HDR=T[21]._tbl; BOX=T[2]._tbl
FONT='Calibri'
def add_par(text='', bold=False, italic=False, size=10.5):
    p=d.add_paragraph()
    if text:
        r=p.add_run(text); r.bold=bold; r.italic=italic; r.font.name=FONT; r.font.size=Pt(size)
    return p
def add_tbl(tpl, text):
    anchor=add_par()
    tbl=copy.deepcopy(tpl); anchor._p.addnext(tbl)
    t=Table(tbl, d._body); fill_cell(t.cell(0,0), text)
    return t
def add_ref(label, imgs):
    t=add_tbl(BOX, f'**[{label}]**')
    c=t.cell(0,0)
    for k in imgs:
        r=c.add_paragraph().add_run(); r.add_picture(IMG[k], height=Inches(2.4))
    add_par()

TABS=[('01_Lead_01','LEAD 01'),('02_Lead_02','LEAD 02'),('03_Lead_03','LEAD 03'),('04_Background_Story','BACKGROUND STORY'),
      ('05_Marketing_Thesis','MARKETING THESIS'),('06_Product_Build_Up','PRODUCT BUILD UP'),('07_Big_Offer','BIG OFFER'),('08_Close','CLOSE'),('09_FAQ','FAQ')]
md=[]; nrefs=0; nimgs=0
for key,title in TABS:
    d.add_page_break()
    add_tbl(SEC_HDR, f'**COPY EN-US — {title}**'); add_par()
    md.append(f'\n# COPY EN-US — {title}\n')
    paras=[p.strip() for p in body(f'EN_{key}.md').split('\n\n') if p.strip()]
    plan=PLANO.get(key,[]); ins={}; pos=0
    for anchor,items,pre in plan:
        i=next((k for k in range(pos,len(paras)) if paras[k].startswith(anchor)),None)
        if i is None: raise SystemExit(f'âncora não achada em {key}: {anchor}')
        j=i-1 if pre and i>0 and re.fullmatch(r'\[.*\]',paras[i-1]) else i
        ins.setdefault(j,[]).extend(items); pos=i
    for k,ptxt in enumerate(paras):
        for it in ins.get(k,[]):
            if it[0]=='h': add_tbl(BLK_HDR, f'**{it[1]}**'); add_par(); md.append(f'\n## {it[1]}\n')
            elif it[0]=='tag': add_par(it[1], bold=True); add_par(); md.append(f'**{it[1]}**\n')
            elif it[0]=='dir': add_par(it[1], italic=True); add_par(); md.append(f'*{it[1]}*\n')
            elif it[0]=='ref': add_ref(it[1], it[2]); nrefs+=1; nimgs+=len(it[2]); md.append(f'> **[{it[1]}]**' + (f'  ·  imagens: {", ".join(it[2])}' if it[2] else '') + '\n')
        if re.fullmatch(r'\[.*\]', ptxt): add_par(ptxt, bold=True); md.append(f'**{ptxt}**\n')
        else: add_par(ptxt); md.append(ptxt+'\n')
        add_par()
d.save(OUT)
open(f'{ROOT}/briefing/BRIEFING_Edicao_IronPulseX_copy.md','w').write('\n'.join(md))
print('ok', OUT, 'caixas de referência:', nrefs, 'imagens:', nimgs)
print('timeline:', [x[1][:40] for x in TIMELINE])
