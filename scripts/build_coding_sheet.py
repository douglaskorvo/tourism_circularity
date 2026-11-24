# Builds the double-coding workbook (Codificador_1, Codificador_2, Kappa, Resumo) from data/reviews_raw.csv.
# Usage: python scripts/build_coding_sheet.py data/reviews_raw.csv coding_sheet.xlsx
# The output contains review text and must NOT be committed (see .gitignore).
import pandas as pd, re, sys, numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as L_
SRC, OUT = sys.argv[1], sys.argv[2]
df=pd.read_csv(SRC).reset_index().rename(columns={'index':'row_id'})
N_TOTAL_ROWS=len(df)
df=df[df.text.notna()&(df.text.str.strip()!='')].copy()
N_TEXT=len(df)
LEX=[  # (codigo, categoria, regex)  ordem = prioridade
('R8','R8 Reciclar',r'\brecicl\w*'),
('R7','R7 Repropor / upcycling',r'upcycl\w*|segunda vida|\bsucata\w*|reuso adaptativo|\btransformad[oa]s? em\b'),
('R5-6','R5-R6 Restaurar / remanufaturar',r'\brestaur(?:ad[oa]s?|ação|ações|ar|ou)\b'),
('R4','R4 Reparar',r'\bconsert\w*|\brepar(?:o|os|ar|ou|ação|ações|ado|ada)\b'),
('R3','R3 Reutilizar',r'\breaprovei\w*|\breutiliz\w*|\breuso\b|segunda mão|\bbrechó\w*|\b(?:roupas?|móveis|livros?|peças?|máquinas?|objetos?|itens|produtos?|artigos?)\s+(?:\w+\s+)?usad[oa]s\b|\bantiguidades?\b|\bantiquári\w*|\bvintage\b'),
('R0-2','R0-R2 Recusar / repensar / reduzir',r'desperd[ií]\w*|\bresíduos?\b|\blixo\b|lixo zero|zero waste|\bsobras?\b|\bplástico\w*|\bdescartáve\w*|\bcanudos?\b|\bembalage\w*'),
('SD','Sustentabilidade declarada',r'sustentáve\w*|sustentabilidade|ecológic\w*|meio ambiente|ambientalmente|eco[- ]?friendly|\becoturism\w*'),
('CL','Cadeia local',r'\b(?:produt\w*|ingredientes?|matérias?[- ]primas?|insumos?|fornecedor\w*|comida|alimentos?)\s+(?:\w+\s+)?(?:loca(?:l|is)|da região|regiona(?:l|is))\b|comércio justo|fair trade|\bcooperativas? (?:de |das |dos )(?:artes|mulher|produtor|agricult|tecel|bordad|oleir|ceramist|costur)\w*|\borgânic\w*'),
('AR','Artesania',r'feit[oa]s? (?:à|a) mão|\bartesana(?:l|is)\b|\bartes(?:ão|ãos|ã|ãs)\b'),
]
NOISE=[r'desperd[ií]\w*\s+(?:de\s+|do\s+|o\s+|meu\s+|seu\s+|nosso\s+|total\s+de\s+)?(?:dinheiro|tempo|dia|viagem|energia)',
 r'\bum lixo\b|\blixo de\b|\bé lixo\b|\bpuro lixo\b|\bque lixo\b',r'\bsobras?\s+(?:de\s+)?(?:funcionári|tempo)',
 r'precis\w* de (?:alguns |uns |muitos )?reparos|necessita\w* de (?:alguns )?reparos']
recs=[]
for _,r in df.iterrows():
    low=r.text.lower(); hits=[]
    for code,cat,p in LEX:
        ms=[m for m in re.finditer(p,low) if not any(re.search(n,low[max(0,m.start()-60):m.end()+60]) for n in NOISE)]
        if ms: hits.append((code,cat,ms))
    allm=[(code,m) for code,cat,p in LEX for m in re.finditer(p,low)]
    if not hits:
        if allm: recs.append(dict(r, grupo='ruido', codes='; '.join(sorted({c for c,_ in allm})), termos='; '.join(sorted({r.text[m.start():m.end()] for _,m in allm})), i=allm[0][1].start(), j=allm[0][1].end()))
        continue
    code,cat,ms=hits[0]
    recs.append(dict(r, grupo='candidato', principal=cat, codes='; '.join(h[0] for h in hits),
        termos='; '.join(sorted({r.text[m.start():m.end()] for h in hits for m in h[2]})), i=ms[0].start(), j=ms[0].end()))
c=pd.DataFrame(recs)
def ctx(r): 
    t=r.text.replace('\n',' '); i,j=r.i,r.j
    return ('… ' if i>160 else '')+t[max(0,i-160):j+160]+(' …' if j+160<len(t) else '')
c['contexto']=c.apply(ctx,axis=1)
cand=c[c.grupo=='candidato'].copy(); noise=c[c.grupo=='ruido'].copy()
ar_only=cand[cand.codes=='AR']; rest=cand[cand.codes!='AR']
AR_N=200
ar_s=pd.concat([g.sample(int(round(AR_N*len(g)/len(ar_only))),random_state=20261119) for _,g in ar_only.groupby('segment')])
code_set=pd.concat([rest,ar_s]).sample(frac=1,random_state=20261119).reset_index(drop=True)
code_set['ID']=[f'C{n:04d}' for n in range(1,len(code_set)+1)]
# subamostra dupla ~25% estratificada pela categoria principal
rng=np.random.default_rng(20261119)
code_set['dupla']='não'
for cat,g in code_set.groupby('principal'):
    k=max(5,int(round(0.25*len(g)))); k=min(k,len(g))
    code_set.loc[rng.choice(g.index,k,replace=False),'dupla']='sim'
print('codificar',len(code_set),'dupla',(code_set.dupla=='sim').sum(),'AR universo',len(ar_only),'ruido',len(noise))
print(code_set.principal.value_counts())

# ---------- workbook ----------
F='Arial'; H=Font(name=F,bold=True,color='FFFFFF'); HF=PatternFill('solid',fgColor='1F4E78')
B=Font(name=F); BB=Font(name=F,bold=True); Y=PatternFill('solid',fgColor='FFF2CC'); G=PatternFill('solid',fgColor='E2EFDA')
thin=Side(style='thin',color='BFBFBF'); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
WR=Alignment(wrap_text=True,vertical='top'); TOP=Alignment(vertical='top')
wb=Workbook()
STRAT=['R0-R2 Recusar / repensar / reduzir','R3 Reutilizar','R4 Reparar','R5-R6 Restaurar / remanufaturar','R7 Repropor / upcycling','R8 Reciclar','R9 Recuperar (energia)','Cadeia local','Artesania','Sustentabilidade declarada','Não se aplica']
QUEM=['Fornecedor','Visitante','Destino / comunidade','Indeterminado']
VAL=['Positiva','Negativa','Neutra']
# Listas
ls=wb.active; ls.title='00 Leia-me'
lst=wb.create_sheet('Listas')
for col,(name,vals) in enumerate([('Estrategia',STRAT),('Quem',QUEM),('Valencia',VAL),('SN',['S','N'])],1):
    lst.cell(1,col,name).font=BB
    for k,v in enumerate(vals,2): lst.cell(k,col,v).font=B
    lst.column_dimensions[L_(col)].width=38
def dv(ws,col,ref,n):
    d=DataValidation(type='list',formula1=ref,allow_blank=True,showErrorMessage=True,errorTitle='Valor inválido',error='Escolha um valor da lista.')
    ws.add_data_validation(d); d.add(f'{col}2:{col}{n+1}')
FIX=['ID','row_id','País','Cidade','Segmento','Nota','lang','Categoria principal (auto)','Categorias (auto)','Termos casados','Contexto (±160 caracteres)']
def coder_sheet(name,pref,data):
    ws=wb.create_sheet(name)
    hdr=FIX+[f'{pref}_relevante',f'{pref}_estrategia',f'{pref}_estrategia_2',f'{pref}_quem',f'{pref}_valencia',f'{pref}_obs','Texto completo']
    ws.append(hdr)
    for c_ in ws[1]: c_.font=H; c_.fill=HF; c_.alignment=WR; c_.border=BOX
    for _,r in data.iterrows():
        ws.append([r.ID,int(r.row_id),r.country,r.city,r.segment,int(r.rating),r.lang,r.principal,r.codes,r.termos,r.contexto,None,None,None,None,None,None,r.text[:32000]])
    n=len(data)
    widths=[8,8,13,14,12,6,7,24,12,20,70,11,26,22,16,11,24,60]
    for k,w in enumerate(widths,1): ws.column_dimensions[L_(k)].width=w
    for row in ws.iter_rows(min_row=2,max_row=n+1):
        for c_ in row:
            c_.font=B; c_.border=BOX; c_.alignment=WR if c_.column in (11,18) else TOP
            if 12<=c_.column<=17: c_.fill=Y
    ws.freeze_panes='L2'; ws.auto_filter.ref=f'A1:{L_(len(hdr))}{n+1}'
    dv(ws,'L','=Listas!$D$2:$D$3',n); dv(ws,'M','=Listas!$A$2:$A$12',n); dv(ws,'N','=Listas!$A$2:$A$12',n)
    dv(ws,'O','=Listas!$B$2:$B$5',n); dv(ws,'P','=Listas!$C$2:$C$4',n)
    return ws,n
w1,n1=coder_sheet('Codificador_1','C1',code_set)
dup=code_set[code_set.dupla=='sim']
w2,n2=coder_sheet('Codificador_2','C2',dup)
# colunas de comparação no Codificador_2
base=len(FIX)+7
heads=['C1_relevante (puxado)','C1_estrategia (puxado)','Concorda relevância','Concorda estratégia']
for k,h in enumerate(heads):
    c_=w2.cell(1,base+1+k,h); c_.font=H; c_.fill=PatternFill('solid',fgColor='548235'); c_.alignment=WR; c_.border=BOX
    w2.column_dimensions[L_(base+1+k)].width=16
R1=f"Codificador_1!$A$2:$A${n1+1}"
for i in range(2,n2+2):
    S,T,U,V=[L_(base+1+k) for k in range(4)]
    w2[f'{S}{i}']=f'=IFERROR(INDEX(Codificador_1!$L$2:$L${n1+1},MATCH($A{i},{R1},0))&"","")'
    w2[f'{T}{i}']=f'=IFERROR(INDEX(Codificador_1!$M$2:$M${n1+1},MATCH($A{i},{R1},0))&"","")'
    w2[f'{U}{i}']=f'=IF(OR($L{i}="",{S}{i}=""),"",IF($L{i}={S}{i},1,0))'
    w2[f'{V}{i}']=f'=IF(OR($M{i}="",{T}{i}=""),"",IF($M{i}={T}{i},1,0))'
    for col in (S,T,U,V): w2[f'{col}{i}'].font=B; w2[f'{col}{i}'].border=BOX; w2[f'{col}{i}'].fill=G
# Kappa
kp=wb.create_sheet('Kappa')
kp['A1']='Concordância entre codificadores (Cohen\'s kappa), calculada sobre as linhas da subamostra dupla com os dois códigos preenchidos'; kp['A1'].font=BB
S,T,U,V=[L_(base+1+k) for k in range(4)]
rngU=f'Codificador_2!${U}$2:${U}${n2+1}'; rngV=f'Codificador_2!${V}$2:${V}${n2+1}'
kp['A3']='Relevância (S/N)'; kp['A3'].font=BB
kp['A4']='Pares válidos (n)'; kp['B4']=f'=COUNT({rngU})'
kp['A5']='Concordância observada (po)'; kp['B5']=f'=IF(B4=0,"",SUM({rngU})/B4)'
kp['A7']='Categoria'; kp['B7']='Codificador 1'; kp['C7']='Codificador 2'; kp['D7']='Produto das proporções'
for k,cat in enumerate(['S','N']):
    r=8+k; kp[f'A{r}']=cat
    kp[f'B{r}']=f'=COUNTIFS({rngU},">=0",Codificador_2!${S}$2:${S}${n2+1},A{r})'
    kp[f'C{r}']=f'=COUNTIFS({rngU},">=0",Codificador_2!$L$2:$L${n2+1},A{r})'
    kp[f'D{r}']=f'=IF($B$4=0,"",(B{r}/$B$4)*(C{r}/$B$4))'
kp['A10']='Concordância esperada (pe)'; kp['B10']='=IF(B4=0,"",SUM(D8:D9))'
kp['A11']='Kappa'; kp['B11']='=IF(OR(B4=0,B10=1),"",(B5-B10)/(1-B10))'; kp['A11'].font=BB; kp['B11'].font=BB
kp['A14']='Estratégia (código principal)'; kp['A14'].font=BB
kp['A15']='Pares válidos (n)'; kp['B15']=f'=COUNT({rngV})'
kp['A16']='Concordância observada (po)'; kp['B16']=f'=IF(B15=0,"",SUM({rngV})/B15)'
kp['A18']='Categoria'; kp['B18']='Codificador 1'; kp['C18']='Codificador 2'; kp['D18']='Produto das proporções'
for k,cat in enumerate(STRAT):
    r=19+k; kp[f'A{r}']=cat
    kp[f'B{r}']=f'=COUNTIFS({rngV},">=0",Codificador_2!${T}$2:${T}${n2+1},A{r})'
    kp[f'C{r}']=f'=COUNTIFS({rngV},">=0",Codificador_2!$M$2:$M${n2+1},A{r})'
    kp[f'D{r}']=f'=IF($B$15=0,"",(B{r}/$B$15)*(C{r}/$B$15))'
e=19+len(STRAT)-1
kp[f'A{e+1}']='Concordância esperada (pe)'; kp[f'B{e+1}']=f'=IF(B15=0,"",SUM(D19:D{e}))'
kp[f'A{e+2}']='Kappa'; kp[f'B{e+2}']=f'=IF(OR(B15=0,B{e+1}=1),"",(B16-B{e+1})/(1-B{e+1}))'; kp[f'A{e+2}'].font=BB; kp[f'B{e+2}'].font=BB
kp[f'A{e+4}']='Interpretação usual (Landis & Koch, 1977): 0,41–0,60 moderada; 0,61–0,80 substancial; >0,80 quase perfeita.'
for row in kp.iter_rows():
    for c_ in row:
        if c_.font!=BB and c_.font!=H: c_.font=Font(name=F,bold=c_.font.bold)
for r in (7,18):
    for col in 'ABCD': kp[f'{col}{r}'].font=H; kp[f'{col}{r}'].fill=HF
for addr in ['B5','B10','B11','B16',f'B{e+1}',f'B{e+2}']: kp[addr].number_format='0.000'
for col in 'BCD':
    for r in range(8,e+1):
        if kp[f'{col}{r}'].value and col=='D': kp[f'{col}{r}'].number_format='0.0000'
kp.column_dimensions['A'].width=40; [kp.column_dimensions.__setitem__(x,kp.column_dimensions[x]) for x in 'BCD']
for x in 'BCD': kp.column_dimensions[x].width=16
# Resumo
rs=wb.create_sheet('Resumo')
rs['A1']='Prevalência das estratégias no corpus, a partir da codificação do Codificador 1'; rs['A1'].font=BB
rs['A3']='Avaliações com texto no corpus (insumo do pipeline)'; rs['B3']=N_TEXT
rs['A4']='Universo de avaliações só de Artesania (insumo do pipeline)'; rs['B4']=len(ar_only)
rs['A5']='Amostra de Artesania codificada (insumo do pipeline)'; rs['B5']=len(ar_s)
rs['A6']='Fator de expansão da Artesania'; rs['B6']='=IF(B5=0,"",B4/B5)'; rs['B6'].number_format='0.00'
rs['C3']='Fonte: reviews_raw.csv, linhas com texto não vazio'; rs['C4']='Avaliações cujo único sinal casado é AR; as demais categorias são censo'
hdr=['Estratégia (C1_estrategia)','Codificadas','Relevantes (S)','Creative (S)','Mass market (S)','Nota média (S)','% do corpus (S)']
for k,h in enumerate(hdr,1):
    c_=rs.cell(8,k,h); c_.font=H; c_.fill=HF; c_.alignment=WR
M1=f'Codificador_1!$M$2:$M${n1+1}'; Lr=f'Codificador_1!$L$2:$L${n1+1}'; Sg=f'Codificador_1!$E$2:$E${n1+1}'; Nt=f'Codificador_1!$F$2:$F${n1+1}'; Pc=f'Codificador_1!$H$2:$H${n1+1}'
for k,cat in enumerate(STRAT[:-1]):
    r=9+k; rs[f'A{r}']=cat
    rs[f'B{r}']=f'=COUNTIFS({M1},A{r})'
    rs[f'C{r}']=f'=COUNTIFS({M1},A{r},{Lr},"S")'
    rs[f'D{r}']=f'=COUNTIFS({M1},A{r},{Lr},"S",{Sg},"creative")'
    rs[f'E{r}']=f'=COUNTIFS({M1},A{r},{Lr},"S",{Sg},"mass_market")'
    rs[f'F{r}']=f'=IFERROR(AVERAGEIFS({Nt},{M1},A{r},{Lr},"S"),"")'; rs[f'F{r}'].number_format='0.00'
    rs[f'G{r}']=f'=IF(A{r}="Artesania",IFERROR(COUNTIFS({M1},A{r},{Lr},"S",{Pc},"Artesania")*$B$6+COUNTIFS({M1},A{r},{Lr},"S",{Pc},"<>Artesania"),"")/$B$3,C{r}/$B$3)'
    rs[f'G{r}'].number_format='0.000%'
last=9+len(STRAT)-2
rs[f'A{last+1}']='Total codificado (todas as linhas)'; rs[f'B{last+1}']=f'=COUNTA(Codificador_1!$A$2:$A${n1+1})'
rs[f'A{last+2}']='Linhas já codificadas pelo C1'; rs[f'B{last+2}']=f'=COUNTA({Lr})'
rs[f'A{last+3}']='Precisão do léxico (S / codificadas)'; rs[f'B{last+3}']=f'=IF(B{last+2}=0,"",COUNTIF({Lr},"S")/B{last+2})'; rs[f'B{last+3}'].number_format='0.0%'
rs[f'A{last+5}']='Nota: a % de Artesania expande pela amostra apenas as linhas cuja categoria principal automática é Artesania.'
for row in rs.iter_rows():
    for c_ in row:
        if c_.font.color is None or c_.font.color.rgb!='FFFFFFFF': c_.font=Font(name=F,bold=c_.font.bold)
for x,w in zip('ABCDEFG',[44,12,13,12,14,12,14]): rs.column_dimensions[x].width=w
for addr in ('B3','B4','B5'): rs[addr].font=Font(name=F,color='000000')
# Ruido
rz=wb.create_sheet('Excluidos_ruido')
rz.append(['row_id','País','Cidade','Segmento','Nota','Categorias casadas','Termos','Contexto'])
for c_ in rz[1]: c_.font=H; c_.fill=HF
for _,r in noise.iterrows(): rz.append([int(r.row_id),r.country,r.city,r.segment,int(r.rating),r.codes,r.termos,r.contexto])
for row in rz.iter_rows(min_row=2):
    for c_ in row: c_.font=B; c_.alignment=WR if c_.column==8 else TOP
for x,w in zip('ABCDEFGH',[8,13,14,12,6,14,20,90]): rz.column_dimensions[x].width=w
rz.freeze_panes='A2'
# Codebook
cb=wb.create_sheet('Codebook',1)
rows=[('Campo','Valores','Regra de decisão'),
('_relevante','S / N','S quando o trecho descreve, no sentido do construto, uma prática ou atributo de produção ou consumo responsável ligado ao fornecedor, ao visitante ou ao destino. N quando o termo aparece em outro sentido (ex.: "desperdício de dinheiro", "um lixo de lugar", "precisa de reparos" como manutenção, "embalagem" só como reclamação de entrega, "artesanal" como mero adjetivo de qualidade sem referência ao modo de produção).'),
('_estrategia','Lista (Kirchherr et al., 2017)','Estratégia principal descrita no trecho, que pode diferir da categoria automática. R0-R2: evitar/reduzir consumo de material, embalagem, descartáveis, desperdício de alimento. R3: reuso, segunda mão, antiquários, reaproveitamento. R4: serviço de conserto ou reparo. R5-R6: restauração ou recondicionamento de bem ou edifício, incluindo reuso adaptativo de patrimônio. R7: dar nova função ao material (upcycling). R8: reciclagem de material. R9: recuperação energética. Cadeia local: insumo, alimento ou produto de produtor local, comércio justo, cooperativa produtiva, orgânico. Artesania: produção manual, artesão como produtor. Sustentabilidade declarada: menção genérica a sustentável/ecológico sem estratégia identificável. Não se aplica: quando _relevante = N.'),
('_estrategia_2','Lista (opcional)','Segunda estratégia, se o trecho descrever duas de forma clara. Deixar vazio caso contrário.'),
('_quem','Fornecedor / Visitante / Destino / comunidade / Indeterminado','Quem realiza a prática: o estabelecimento avaliado, o próprio visitante (escolha de consumo), a comunidade ou o destino em geral.'),
('_valencia','Positiva / Negativa / Neutra','Atitude do autor da avaliação em relação à prática (não a nota geral). Ex.: elogia o material reciclado = Positiva; reclama do excesso de plástico = Negativa; só descreve = Neutra.'),
('_obs','Texto livre','Dúvidas, casos limítrofes, sugestão de novo termo para o léxico.'),
('Regra geral','','Codificar pelo contexto (±160 caracteres); abrir o texto completo (última coluna) quando o contexto não bastar. Não consultar o código do outro codificador antes de terminar.'),
]
for r in rows: cb.append(list(r))
for c_ in cb[1]: c_.font=H; c_.fill=HF
for row in cb.iter_rows(min_row=2):
    for c_ in row: c_.font=B; c_.alignment=WR
cb.column_dimensions['A'].width=16; cb.column_dimensions['B'].width=30; cb.column_dimensions['C'].width=110
cb.append([]); cb.append(['Exemplo de linha codificada (não faz parte dos dados)']); cb.cell(cb.max_row,1).font=BB
cb.append(['Contexto','“Os artesanatos são muito criativos e sempre feitos com material de reaproveitamento, gerando um impacto positivo para o meio ambiente”'])
cb.append(['_relevante','S']); cb.append(['_estrategia','R3 Reutilizar']); cb.append(['_estrategia_2','Artesania']); cb.append(['_quem','Fornecedor']); cb.append(['_valencia','Positiva']); cb.append(['_obs','Material de reaproveitamento = reuso de insumo; poderia ser R7 se descrevesse nova função.'])
for row in cb.iter_rows(min_row=cb.max_row-6):
    for c_ in row: c_.font=B; c_.alignment=WR
# Leia-me
txt=[('Planilha de codificação — sinais de circularidade no reviews_raw.csv (artigo para a Production, Thematic Section de circularidade)',True),
('',False),
('O que preencher',True),
('Somente as células amarelas das abas Codificador_1 e Codificador_2 (colunas _relevante a _obs), escolhendo os valores nas listas suspensas. Regras na aba Codebook, com uma linha de exemplo.',False),
('Codificador_1 codifica todas as linhas. Codificador_2 codifica só a subamostra dupla (cerca de 25%, estratificada pela categoria principal), sem ver os códigos do C1. As colunas verdes do Codificador_2 puxam o código do C1 por fórmula e a aba Kappa calcula a concordância.',False),
('',False),('Como o conjunto foi montado',True),
(f'Corpus: {N_TOTAL_ROWS} linhas, {N_TEXT} com texto. Léxico de sementes em português (o texto é original pt ou traduzido para pt-BR), agrupado pelas estratégias R de Kirchherr, Reike & Hekkert (2017), mais Cadeia local, Artesania e Sustentabilidade declarada.',False),
('Censo: todas as avaliações com pelo menos um termo fora de Artesania. Amostra: 200 avaliações cujo único sinal é Artesania (aleatória, proporcional ao segmento, semente 20261119), porque Artesania é a categoria de comparação e é muito mais frequente.',False),
(f'Filtro de ruído: termos em expressões idiomáticas ("desperdício de dinheiro/tempo", "um lixo de lugar", "precisa de reparos") foram retirados. As {len(noise)} avaliações em que só sobrou ruído estão na aba Excluidos_ruido para auditoria.',False),
('A categoria principal automática é a primeira na ordem R8 > R7 > R5-6 > R4 > R3 > R0-2 > SD > CL > AR; a coluna "Categorias (auto)" lista todas as que casaram.',False),
('',False),('Abas',True),
('Codebook · Codificador_1 · Codificador_2 · Kappa · Resumo (prevalência por estratégia, segmento e nota, por fórmula) · Excluidos_ruido · Listas (valores das listas suspensas)',False),
('',False),('Cuidados',True),
('Não redistribuir: os textos são conteúdo da plataforma (Google). Nome do autor da avaliação foi omitido. Não usar o léxico do aspecto "environment" do artigo massa vs. criativo.',False)]
for k,(t,b) in enumerate(txt,1):
    c_=ls.cell(k,1,t); c_.font=Font(name=F,bold=b,size=12 if k==1 else 10); c_.alignment=Alignment(wrap_text=True,vertical='top')
ls.column_dimensions['A'].width=130
wb.move_sheet('Listas',offset=10)
wb.save(OUT)
code_set.drop(columns=['text','author_name','i','j','grupo']).to_csv(OUT.replace('.xlsx','_ids.csv'),index=False)
