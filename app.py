from io import BytesIO
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Giolitti 48", page_icon="🏠", layout="wide", initial_sidebar_state="collapsed")
BROWN="#3B2A22"; COCOA="#644B3C"; BEIGE="#E8DDD1"; CREAM="#F7F4EF"; SAGE="#7D9278"; TERRA="#B56E4A"; GOLD="#C29A61"; RED="#B44C43"; GREY="#766E69"
ICONS={"Cucina":"🍳","Bagno":"🛁","Soggiorno":"🛋️","Camera":"🛏️","Studio":"💻","Lavanderia":"🧺","Casa":"🏠","Ingresso":"🚪","Ripostiglio":"📦","Esterno":"🌿"}
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Playfair+Display:wght@700&display=swap');
html,body,[class*="css"]{{font-family:Inter,sans-serif}} .stApp{{background:{CREAM};color:{BROWN}}}
.block-container{{max-width:1480px;padding:1.25rem 2.2rem 4rem}} header[data-testid="stHeader"]{{background:transparent}}
.hero{{background:linear-gradient(120deg,{BROWN},{COCOA});padding:30px 34px;border-radius:24px;color:#fff;box-shadow:0 14px 34px rgba(59,42,34,.16);margin-bottom:18px}}
.hero h1{{font-family:'Playfair Display',serif;font-size:38px;margin:0 0 8px}}.hero p{{margin:0;color:#eadfd6}}
.title{{font-size:20px;font-weight:800;color:{BROWN};margin:24px 0 11px}}.subtle{{color:{GREY};font-size:12px}}
.kpi,.room,.priority{{background:#fff;border:1px solid #e6ddd6;border-radius:18px;box-shadow:0 5px 18px rgba(59,42,34,.055)}}
.kpi{{padding:17px 18px;min-height:118px;border-top:4px solid var(--a)}}.kl{{font-size:11px;font-weight:800;color:{GREY};letter-spacing:.5px}}.kv{{font-size:27px;font-weight:800;color:{BROWN};margin:9px 0 5px}}.kn{{font-size:12px;color:{GREY}}}
.room{{padding:15px;min-height:122px;border-bottom:4px solid {BEIGE}}}.ri{{font-size:25px}}.rt{{font-size:16px;font-weight:800;margin:6px 0;color:{BROWN}}}.rv{{font-size:20px;font-weight:800;color:{COCOA}}}.rm{{font-size:12px;color:{GREY}}}
.room-selected{{border:2px solid {TERRA};background:#fffaf6}}.priority{{padding:16px;min-height:105px}}.pn{{font-size:27px;font-weight:800}}.pt{{font-size:11px;font-weight:800;color:{GREY}}}
.detail{{background:#fff;border:1px solid #e6ddd6;border-radius:19px;padding:20px;margin-top:10px}}.tag{{display:inline-block;background:{BEIGE};color:{BROWN};padding:5px 9px;border-radius:999px;font-size:11px;font-weight:700;margin-right:6px}}
.package{{background:#f1e9df;border-left:5px solid {GOLD};padding:14px 16px;border-radius:12px;color:{BROWN};margin:10px 0 14px}}
.stButton>button{{width:100%;border-radius:12px;background:white;color:{BROWN};border:1px solid #d7c9bd;font-weight:700}}.stButton>button:hover{{border-color:{TERRA};color:{TERRA}}}
[data-testid="stDataFrame"]{{border:1px solid #e6ddd6;border-radius:14px;overflow:hidden}} #MainMenu,footer{{visibility:hidden}}
</style>""",unsafe_allow_html=True)

@st.cache_data(show_spinner=False)
def load(raw):
    d=pd.read_excel(BytesIO(raw),sheet_name="Lista",header=3,engine="openpyxl").iloc[:,:16].copy()
    d.columns=[str(x).strip() for x in d.columns]; d=d[d.get("Voce").notna()|d.get("Categoria").notna()].copy()
    for c in ["Q.tà","Prezzo €","Totale €"]:
        if c not in d:d[c]=0
        d[c]=pd.to_numeric(d[c],errors="coerce").fillna(0)
    m=d["Totale €"].eq(0)&d["Prezzo €"].ne(0); d.loc[m,"Totale €"]=d.loc[m,"Q.tà"].replace(0,1)*d.loc[m,"Prezzo €"]
    for c in ["✔","Categoria","Voce","Stato","Stanza","Quando","Note","Brand","Modello"]:
        if c not in d:d[c]=""
        d[c]=d[c].fillna("").astype(str).str.strip()
    return d

def euro(v):return f"€ {v:,.0f}".replace(",",".")
def ring(value,total,label,color):
    pct=max(0,min(100,(value/total*100 if total else 0)))
    fig=go.Figure(go.Pie(values=[pct,100-pct],hole=.77,sort=False,direction="clockwise",marker=dict(colors=[color,"#EDE6E0"],line=dict(color=CREAM,width=3)),textinfo="none",hoverinfo="skip"))
    fig.add_annotation(text=f"<b>{pct:.0f}%</b><br><span style='font-size:11px'>{label}</span>",showarrow=False,font=dict(color=BROWN,size=19))
    fig.update_layout(height=245,margin=dict(l=8,r=8,t=8,b=8),showlegend=False,paper_bgcolor=CREAM)
    st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False})

def table_block(data,title):
    st.markdown(f"**{title}**")
    if data.empty: st.caption("Nessuna voce")
    else:
        x=data[["Voce","Categoria","Stato","Quando","Totale €"]].copy()
        x["Totale €"]=x["Totale €"].map(euro)
        st.dataframe(x,use_container_width=True,hide_index=True,height=min(350,38+35*len(x)))

with open("Casa_Giolitti_Acquisti.xlsx","rb") as f:default=f.read()
st.markdown('<div class="hero"><h1>Giolitti 48</h1><p>Acquisti · Budget · Avanzamento lavori</p></div>',unsafe_allow_html=True)
with st.expander("Aggiorna dati da Excel"):
    up=st.file_uploader("Carica il file aggiornato",type=["xlsx"])
raw=up.getvalue() if up else default
df=load(raw)

f1,f2,f3,f4=st.columns(4)
with f1:fc=st.multiselect("Categoria",sorted(x for x in df.Categoria.unique() if x),placeholder="Tutte")
with f2:fr=st.multiselect("Stanza",sorted(x for x in df.Stanza.unique() if x),placeholder="Tutte")
with f3:fs=st.multiselect("Stato",sorted(x for x in df.Stato.unique() if x),placeholder="Tutti")
with f4:fp=st.multiselect("Priorità / fase",sorted(x for x in df.Quando.unique() if x),placeholder="Tutte")
v=df.copy()
if fc:v=v[v.Categoria.isin(fc)]
if fr:v=v[v.Stanza.isin(fr)]
if fs:v=v[v.Stato.isin(fs)]
if fp:v=v[v.Quando.isin(fp)]
sel=v[v["✔"].str.contains("✔",regex=False)].copy(); missing=v[v.Stato.eq("Da valutare")].copy()
total=sel["Totale €"].sum(); paid=sel.loc[sel.Stato.isin(["Pagato","Consegnato"]),"Totale €"].sum(); ordered=sel.loc[sel.Stato.eq("Ordinato"),"Totale €"].sum(); quote=sel.loc[sel.Stato.eq("Preventivo"),"Totale €"].sum(); due=max(total-paid,0)

st.markdown('<div class="title">Quadro generale</div>',unsafe_allow_html=True)
def kpi(t,val,note,color):st.markdown(f'<div class="kpi" style="--a:{color}"><div class="kl">{t}</div><div class="kv">{val}</div><div class="kn">{note}</div></div>',unsafe_allow_html=True)
cols=st.columns(5)
for c,args in zip(cols,[("TOTALE IMPEGNATO",euro(total),f"{len(sel)} voci selezionate",COCOA),("GIÀ PAGATO",euro(paid),"pagato o consegnato",SAGE),("ORDINATO",euro(ordered),"da saldare",TERRA),("PREVENTIVI",euro(quote),"da confermare",GOLD),("DA PAGARE",euro(due),"residuo sul totale",RED)]):
    with c:kpi(*args)

st.markdown('<div class="title">Speso e ancora da coprire</div>',unsafe_allow_html=True)
r1,r2,r3=st.columns(3)
with r1:ring(paid,total,"già pagato",SAGE)
with r2:ring(due,total,"ancora da pagare",TERRA)
with r3:ring(len(sel),len(v),"voci definite",COCOA)

st.markdown('<div class="title">Stanze</div><div class="subtle">Clicca una stanza per vedere cosa è fatto e cosa manca.</div>',unsafe_allow_html=True)
if "room" not in st.session_state:st.session_state.room=None
order=["Cucina","Bagno","Soggiorno","Camera","Studio","Lavanderia","Casa","Ingresso","Ripostiglio","Esterno"]
rooms=[x for x in order if x in v.Stanza.unique()]
for start in range(0,len(rooms),5):
    cc=st.columns(5)
    for c,room in zip(cc,rooms[start:start+5]):
        sub=v[v.Stanza.eq(room)]; chosen=sub[sub["✔"].str.contains("✔",regex=False)]; val=chosen["Totale €"].sum(); open_n=len(sub[sub.Stato.eq("Da valutare")])
        with c:
            cls="room room-selected" if st.session_state.room==room else "room"
            st.markdown(f'<div class="{cls}"><div class="ri">{ICONS.get(room,"🏠")}</div><div class="rt">{room}</div><div class="rv">{euro(val)}</div><div class="rm">{len(sub)} voci · {open_n} da valutare</div></div>',unsafe_allow_html=True)
            if st.button(f"Apri {room}",key=f"b_{room}"):st.session_state.room=room;st.rerun()

if st.session_state.room:
    room=st.session_state.room; d=v[v.Stanza.eq(room)].copy(); done=d[d.Stato.isin(["Pagato","Consegnato","Ordinato","Preventivo"])]; todo=d[d.Stato.eq("Da valutare")]
    st.markdown(f'<div class="detail"><span class="tag">{ICONS.get(room,"🏠")} {room}</span><span class="tag">{len(done)} fatte / definite</span><span class="tag">{len(todo)} mancanti</span></div>',unsafe_allow_html=True)
    if room=="Cucina":
        pack=d[(d.Categoria.eq("Cucina: arredo & accessori"))&(d["Totale €"].gt(0))]
        included=d[(d.Categoria.eq("Cucina: arredo & accessori"))&(d["Totale €"].eq(0))&(d.Stato.isin(["Pagato","Consegnato"]))]
        if not pack.empty:
            pack_value=pack["Totale €"].max(); names=", ".join(included.Voce.tolist())
            st.markdown(f'<div class="package"><b>Pacchetto cucina: {euro(pack_value)}</b><br>La cifra principale comprende anche le voci già definite a €0 nel file: {names}. Le voci sono mostrate come incluse e non sommate una seconda volta.</div>',unsafe_allow_html=True)
    a,b=st.columns(2)
    with a:table_block(done,"✅ Fatto / definito")
    with b:table_block(todo,"○ Da decidere / acquistare")

st.markdown('<div class="title">Priorità operative</div>',unsafe_allow_html=True)
pcols=st.columns(3)
for c,phase,color,sub in zip(pcols,["Prima delle tracce","Fine lavori","Dopo trasloco"],[RED,GOLD,SAGE],["urgenti per i lavori","prima della consegna","acquisti differibili"]):
    n=len(missing[missing.Quando.eq(phase)])
    with c:st.markdown(f'<div class="priority" style="border-top:4px solid {color}"><div class="pt">{phase.upper()}</div><div class="pn" style="color:{color}">{n}</div><div class="subtle">{sub}</div></div>',unsafe_allow_html=True)

st.markdown('<div class="title">Lista acquisti</div>',unsafe_allow_html=True)
q=st.text_input("Cerca",placeholder="Voce, stanza, brand...")
t=v.copy()
if q:
    hay=t[["Voce","Brand","Modello","Note","Categoria","Stanza"]].astype(str).agg(" ".join,axis=1);t=t[hay.str.contains(q,case=False,na=False)]
st.dataframe(t[["✔","Voce","Categoria","Stanza","Stato","Quando","Totale €","Note"]],use_container_width=True,hide_index=True,height=420,column_config={"Totale €":st.column_config.NumberColumn(format="€ %.2f")})
