#!/usr/bin/env python3
"""Build the independent Royal Bank of Canada equity research presentation.
All quantities and assumptions are documented in SOURCES.md and METHOD.md.
"""
import os, csv, math
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR

ROOT=Path(__file__).resolve().parent
ASSETS=ROOT/'assets';DATA=ROOT/'data'; OUT=ROOT/'deliverables'
ASSETS.mkdir(exist_ok=True); DATA.mkdir(exist_ok=True);OUT.mkdir(exist_ok=True)
Navy='#102A43'; Navy2='#173F5F'; Teal='#06AAB6'; Cyan='#DDF8FA'; Ink='#142D45'; Slate='#60758A'; Light='#F3F7FA'; Mid='#E4ECF2'; Red='#BF5261'; Green='#287F73'; Gold='#D59B41'; White='#FFFFFF'
FONT='Aptos'

def C(c):return RGBColor.from_string(c.lstrip('#'))

def fill(sh,color):
 sh.fill.solid();sh.fill.fore_color.rgb=C(color);sh.line.fill.background();return sh

def rect(s,x,y,w,h,color=White,r=.1,line=None):
 sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if r else MSO_SHAPE.RECTANGLE,Inches(x),Inches(y),Inches(w),Inches(h))
 fill(sh,color)
 if line: sh.line.color.rgb=C(line);sh.line.width=Pt(.6)
 if r:
  try: sh.adjustments[0]=.08
  except:pass
 return sh

def ln(s,x1,y1,x2,y2,color=Mid,width=1.0):
 sh=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Inches(x1),Inches(y1),Inches(x2),Inches(y2))
 sh.line.color.rgb=C(color); sh.line.width=Pt(width);return sh

def txt(s,text,x,y,w,h,size=16,color=Ink,bold=False,align='left',font=None,margin=.02,vertical='middle',breaks=False):
 shape=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h))
 tf=shape.text_frame;tf.clear();tf.word_wrap=True
 tf.margin_left=tf.margin_right=Inches(margin);tf.margin_top=tf.margin_bottom=Inches(margin)
 tf.vertical_anchor={'middle':MSO_ANCHOR.MIDDLE,'top':MSO_ANCHOR.TOP,'bottom':MSO_ANCHOR.BOTTOM}[vertical]
 # Line breaks via text as given; do not use stray manual breaks unless needed.
 for i,line in enumerate(text.split('\n')):
  p=tf.paragraphs[0] if i==0 else tf.add_paragraph();p.text=line;p.alignment={'left':PP_ALIGN.LEFT,'center':PP_ALIGN.CENTER,'right':PP_ALIGN.RIGHT}[align]
  p.space_before=Pt(0);p.space_after=Pt(0);p.line_spacing=1.05
  for run in p.runs:
   run.font.name=font or FONT;run.font.size=Pt(size);run.font.bold=bold;run.font.color.rgb=C(color)
 return shape

def label(s,text,x,y,w,color=Teal,size=10):return txt(s,text.upper(),x,y,w,.22,size,color,True,vertical='middle')

def pill(s,text,x,y,w,color=Light,fore=Ink,size=11):
 rect(s,x,y,w,.34,color,r=.18);txt(s,text,x+.09,y+.02,w-.18,.30,size,fore,True,'center')

def box_title(s,title,body,x,y,w,h,accent=Teal):
 rect(s,x,y,w,h,Light,r=.16)
 rect(s,x+.15,y+.17,.06,h-.34,accent,r=0)
 txt(s,title,x+.32,y+.16,w-.5,.32,15,Ink,True)
 txt(s,body,x+.32,y+.56,w-.52,h-.69,12.5,Slate,vertical='top')

def metric(s,val,name,x,y,w,h=1.04,detail=None,accent=Teal):
 rect(s,x,y,w,h,Light,r=.17)
 txt(s,val,x+.18,y+.12,w-.36,.40,25,Navy,True)
 txt(s,name,x+.18,y+.54,w-.36,.20,10.5,Slate,True)
 if detail: txt(s,detail,x+.18,y+.78,w-.36,.18,9.5,accent)

def head(s,n,title,sub='',tag='FINANCIALS',foot=''):
 rect(s,0,0,13.333,.095,Teal,r=0)
 label(s,'INDEPENDENT EQUITY RESEARCH  /  '+tag,.58,.32,9.5,Slate,9)
 txt(s,f'{n:02d} / {TOTAL:02d}',11.93,.32,.8,.24,9,Slate,True,'right')
 txt(s,title,.62,.85,12.1,.56,26,Navy,True)
 if sub:txt(s,sub,.63,1.49,11.95,.38,12.5,Slate)
 ln(s,.62,7.055,12.75,7.055,Mid,.8)
 txt(s,'AAMIR SOHAIB  •  05 NOV 2026',.65,7.13,4.2,.19,8.4,Slate,True)
 if foot:txt(s,foot,4.4,7.11,8.32,.27,8.2,Slate,False,'right')

def slide(p,title,sub='',tag='FINANCIALS',foot=''):
 global IDX
 IDX+=1; s=p.slides.add_slide(p.slide_layouts[6]); head(s,IDX,title,sub,tag,foot);return s

def bar(s,name,val,maximum,y,left=1.2,right=11.55,color=Teal,value_text=None,small=False):
 h=.35 if small else .47
 txt(s,name,left,y,2.9,h,11 if small else 12,Ink,True)
 start=left+3.10;maxw=right-start
 rect(s,start,y+.04,maxw,h-.08,Mid,r=.08)
 rect(s,start,y+.04,maxw*max(0,val)/maximum,h-.08,color,r=.08)
 if value_text:txt(s,value_text,right+.13,y,1.03,h,11,Ink,True,'right')

def chart_segment(s,x,y,w,h,data,foot=''):
 total=sum(v for _,v,_ in data)
 maxv=max(v for _,v,_ in data)
 for j,(name,value,col) in enumerate(data):
  yy=y+j*h/len(data)
  txt(s,name,x,yy,w*.39,.35,11.3,Ink,True)
  rect(s,x+w*.41,yy+.04,w*.42,.23,Mid,r=.05)
  rect(s,x+w*.41,yy+.04,w*.42*value/maxv,.23,col,r=.05)
  txt(s,f'{value/1000:.2f}B',x+w*.85,yy,w*.15,.35,11.4,Navy,True,'right')
 if foot:txt(s,foot,x,y+h+.05,w,.25,10,Slate)

def presentation():
 global IDX,TOTAL
 TOTAL=23;IDX=0
 p=Presentation();p.slide_width=Inches(13.333);p.slide_height=Inches(7.5)
 # 1 Cover
 s=p.slides.add_slide(p.slide_layouts[6]);IDX=1
 rect(s,0,0,13.333,7.5,Navy,r=0);rect(s,.0,0,.14,7.5,Teal,r=0)
 rect(s,8.33,0,5,7.5,Navy2,r=0)
 for k in range(9):
  ln(s,8.55+k*.52,1.1,8.55+k*.52,6.55,'#315C76',1)
 for k in range(5):
  ln(s,8.55,1.15+k*1.0,12.85,1.15+k*1.0,'#315C76',.8)
 # architectural motif that does not reuse licensed photo
 for xx,hh in [(8.75,3.1),(9.65,4.2),(10.55,5.1),(11.52,3.8)]:
  rect(s,xx,6.38-hh,.56,hh,'#26556B',r=0)
  rect(s,xx+.14,6.40-hh,.12,hh,'#3D8192',r=0)
 txt(s,'INDEPENDENT EQUITY RESEARCH',.78,.73,7.3,.35,14,Teal,True)
 txt(s,'Royal Bank\nof Canada',.77,1.68,7.35,1.67,40,White,True,vertical='top')
 rect(s,.8,3.74,5.8,.06,Teal,r=0)
 txt(s,'BUSINESS QUALITY  /  CREDIT RISK  /  VALUATION',.8,4.17,7.1,.40,13,'#C9E3EB',True)
 txt(s,'NYSE: RY  •  FY2025 actuals  •  FY2026–27 outlook',.8,5.05,6.9,.45,13,White)
 txt(s,'Aamir Sohaib',.8,6.12,5.5,.40,19,White,True)
 txt(s,'Research cut-off: 05 Nov 2026  •  Stock reference close: 05 Nov 2026',.8,6.68,7.4,.35,10,'#B6D3DE')
 # 2 Summary
 s=slide(p,'The thesis | quality business, demanding price','Provisional stance: HOLD / WATCHLIST, based on an illustrative scenario model','INVESTMENT VIEW','Sources: RBC FY2025, Q3 FY2026; Yahoo Finance; model assumptions')
 metric(s,'16.3%','FY2025 ROE',.7,2.05,2.8)
 metric(s,'13.5%','Q3 FY2026 CET1',3.62,2.05,2.8)
 metric(s,'C$20.37B','FY2025 NET INCOME',6.54,2.05,2.8)
 metric(s,'15.7×','FY2027E P/E*',9.46,2.05,2.8,detail='At 05 Nov close, FX assumption')
 box_title(s,'Why own it','Diversified fee businesses, scale advantages, strong capital and FY2025 ROE support an earnings-quality premium.',.74,3.47,5.74,1.55)
 box_title(s,'Why be selective','Credit stress, a high starting multiple, currency effects and a narrow modeled base-case upside constrain expected returns.',6.69,3.47,5.74,1.55,Red)
 rect(s,.74,5.36,11.69,.82,Cyan)
 txt(s,'Watch for FY2026 results on 03 Dec 2026; reassess FY2027 EPS, credit costs and entry valuation.',.99,5.50,11.17,.43,16,Navy,True)
 txt(s,'* P/E = US$203.76 / (C$17.91 FY2027 consensus EPS × US$0.723/C$). FX assumption is illustrative.',.79,6.43,11.9,.25,10,Slate)
 # 3 Company and mix
 s=slide(p,'RBC’s business model | diversified sources of revenue','FY2025 reported segment revenue; C$ billions','COMPANY PROFILE','Source: RBC FY2025 Annual Report, Table 14')
 chart_segment(s,.8,2.11,11.35,3.6,[('Wealth Management',22378,Teal),('Personal Banking',19854,Navy2),('Capital Markets',14426,'#4589AD'),('Commercial Banking',8562,Gold),('Insurance',1321,'#7BB3AE')])
 rect(s,.8,5.9,11.71,.65,Light)
 txt(s,'Corporate Support revenue: C$0.064B; five operating segments + Corporate Support reconcile to C$66.605B.',1.03,5.99,11.20,.35,12,Ink)
 # 4 Geo
 s=slide(p,'Canada anchors the franchise','FY2025 geographic revenue shares; borrower/client footprint differs from office locations','GEOGRAPHY','Source: RBC FY2025 Annual Report, geographic revenue table')
 for j,(name,v,col,note) in enumerate([('Canada',62.8,Navy2,'Home-market concentration'),('United States',25.8,Teal,'City National + capital markets'),('Other international',11.3,Gold,'Cross-border and institutional businesses')]):
  y=2.15+j*1.21
  txt(s,name,.84,y,3.23,.35,16,Ink,True)
  rect(s,4.0,y+.02,6.6,.37,Mid)
  rect(s,4.0,y+.02,6.6*v/63,.37,col)
  txt(s,f'{v:.1f}%',10.85,y-.01,1.2,.41,19,Navy,True,'right')
  txt(s,note,.84,y+.43,8.5,.26,11,Slate)
 rect(s,.82,6.02,11.5,.62,Cyan)
 txt(s,'Interpretation: domestic concentration supports franchise density but increases sensitivity to Canadian credit and rates.',1.02,6.10,11.1,.4,13,Navy,True)
 # 5 Snapshot
 s=slide(p,'FY2025 | profitability plus balance-sheet scale','CAD figures; fiscal year ended 31 October 2025','FINANCIAL SNAPSHOT','Source: RBC FY2025 Annual Report')
 for val,lab,x,y in [('C$66.61B','REVENUE',.85,2.17),('C$20.37B','NET INCOME',4.72,2.17),('C$2.33T','ASSETS',8.59,2.17),('C$1.52T','DEPOSITS',.85,3.73),('C$14.07','DILUTED EPS',4.72,3.73),('16.3%','ROE',8.59,3.73)]:metric(s,val,lab,x,y,3.53,1.32)
 rect(s,.85,5.45,11.26,.70,Cyan)
 txt(s,'Revenue grew 16.1% and reported net income grew 25.4% versus FY2024.',1.11,5.59,10.88,.38,17,Navy,True)
 # 6 historical
 s=slide(p,'Underlying earnings trajectory | FY2023–FY2025','Reported C$ amounts; FY2023 onward uses IFRS 17 comparable reporting basis','HISTORICAL PERFORMANCE','Source: RBC FY2025 Annual Report, selected annual data')
 for x,title,vals,maxv,color in [(1.05,'REVENUE · C$B',[('2023',51.464),('2024',57.344),('2025',66.605)],72,Teal),(7.0,'NET INCOME · C$B',[('2023',14.612),('2024',16.240),('2025',20.369)],22,Navy2)]:
  label(s,title,x,2.09,4.9,Ink,12)
  for j,(yr,v) in enumerate(vals):
   y=2.83+j*1.00
   txt(s,yr,x,y,1.05,.34,13,Ink,True)
   rect(s,x+1.0,y+.05,3.7,.31,Mid)
   rect(s,x+1.0,y+.05,3.7*v/maxv,.31,color)
   txt(s,f'{v:.2f}',x+4.8,y,1.02,.35,13,Navy,True,'right')
 rect(s,.9,6.18,11.57,.59,Light)
 txt(s,'Avoid naïve FY2022 → FY2023 comparisons: the FY2023 insurance accounting basis changed under IFRS 17.',1.07,6.24,11.12,.46,12,Slate,True)
 # 7 bridge
 s=slide(p,'What drove net income higher in FY2025?','A fully reconciled reported earnings bridge, C$ billions','EARNINGS DRIVERS','Source: RBC FY2025 Annual Report, Table 16')
 values=[('FY2024 NI',16.24,Navy2),('Revenue',9.261,Green),('Expenses',-2.342,Red),('Credit losses',-1.130,Red),('Taxes',-1.660,Red),('FY2025 NI',20.369,Teal)]
 # waterfall bars anchored at common horizontal baseline with connector levels
 xmin=.75; width=11.85; chart_y=2.32;chart_h=3.65; ymin=0;ymax=27
 def yy(v):return chart_y+chart_h-(v-ymin)/(ymax-ymin)*chart_h
 cur=0
 for i,(nm,v,col) in enumerate(values):
  bx=xmin+i*2.02; bw=1.49
  if i==0:bottom=0;top=v;cur=v
  elif i==len(values)-1:bottom=0;top=v
  else:bottom=min(cur,cur+v);top=max(cur,cur+v);cur+=v
  rect(s,bx,yy(top),bw,max(.09,yy(bottom)-yy(top)),col,r=.04)
  txt(s,nm,bx-.14,6.12,1.8,.46,10.2,Ink,True,'center')
  lab=f'{v:+.2f}' if 0<i<5 else f'{v:.2f}'
  txt(s,lab,bx-.13,yy(top)-.33,1.8,.27,12,Navy,True,'center')
  if i not in (0,5):ln(s,bx-.36,yy(cur),bx+.11,yy(cur),Slate,.8)
 txt(s,'C$16.24B + 9.26 revenue – 2.34 operating expense – 1.13 PCL – 1.66 tax = C$20.37B.',.88,6.66,11.8,.25,10.8,Slate)
 # 8 profitability
 s=slide(p,'Profitability | stronger operating efficiency','Reported results; adjusted metrics are separately labeled','PROFITABILITY','Source: RBC FY2025 Annual Report; FY2025 earnings release')
 metric(s,'16.3%','FY2025 ROE',.88,2.16,3.72,detail='+190 bps vs. 2024')
 metric(s,'54.9%','EFFICIENCY RATIO',4.82,2.16,3.72,detail='59.7% in 2024')
 metric(s,'C$20.87B','ADJUSTED NET INCOME',8.76,2.16,3.72,detail='Reported: C$20.37B')
 box_title(s,'Earnings quality','Profit growth exceeded revenue growth as efficiency improved; provisioning and taxes still absorbed a material part of incremental revenue.',.88,3.62,5.50,1.6)
 box_title(s,'Critical distinction','Adjusted earnings exclude specified non-GAAP items. Valuation calculations must consistently distinguish adjusted versus reported EPS.',6.61,3.62,5.50,1.6,Gold)
 rect(s,.90,5.55,11.19,.72,Light)
 txt(s,'Bank metrics to emphasize: ROE, efficiency, capital, PCL, impaired loans and forward P/E / P-B.',1.09,5.71,10.9,.35,14,Navy,True)
 # 9 capital & credit
 s=slide(p,'Capital protects | credit quality requires monitoring','FY2025 annual credit measures vs FY2024; Q3 FY2026 is a separate reporting period','CREDIT & CAPITAL','Sources: RBC FY2025 Annual Report, Table 46; Q3 FY2026 results')
 for j,(t,a,b,note,col) in enumerate([
 ('Gross impaired loan ratio','0.83%','0.59%','Worsened by 24 bps YoY',Red),
 ('PCL / avg. net loans','0.43%','0.35%','Credit costs increased 8 bps YoY',Red),
 ('CET1 ratio, Q3 FY2026','13.5%','—','Capital remains above requirements',Green),
 ('PCL / loans, Q3 FY2026','36 bps','—','Annualized quarterly ratio; not FY2025',Gold)]):
  y=2.05+j*1.0
  rect(s,.86,y,11.45,.84,Light)
  txt(s,t,1.12,y+.15,4.85,.43,14,Ink,True)
  txt(s,a,6.22,y+.13,1.68,.40,19,Navy,True)
  txt(s,'FY24: '+b if b!='—' else 'Latest Q3',7.95,y+.18,1.43,.30,10.4,Slate)
  txt(s,note,9.37,y+.12,2.73,.48,10.2,col,True)
 txt(s,'Rising impaired loans are a key counterweight to RBC’s strong capital and earnings momentum.',.89,6.38,11.53,.38,13,Navy,True)
 # 10 segment profitability
 s=slide(p,'Not all revenue converts equally into profit','FY2025 C$ net income by operating segment; reported basis','SEGMENT ECONOMICS','Source: RBC FY2025 Annual Report, Table 14')
 chart_segment(s,1.0,2.07,10.8,3.98,[('Personal Banking',7105,Navy2),('Capital Markets',5393,'#4589AD'),('Wealth Management',4289,Teal),('Commercial Banking',3020,Gold),('Insurance',828,'#7BB3AE')])
 rect(s,.86,6.19,11.39,.54,Cyan)
 txt(s,'Corporate Support: –C$0.266B. Five segments plus Corporate Support = C$20.369B consolidated.',1.06,6.25,10.96,.40,12,Navy,True)
 # 11 distributions
 s=slide(p,'Capital return | dividend plus buybacks','FY2025 payouts; dividends reported in C$, not USD','CAPITAL RETURN','Source: RBC FY2025 Annual Report')
 for val,nam,x in [('C$6.04','DIVIDEND / SHARE',.95),('43%','DIVIDEND PAYOUT',4.86),('15.24M','SHARES REPURCHASED',8.77)]: metric(s,val,nam,x,2.36,3.53,1.38)
 txt(s,'Capital return is supported by profitability, but must remain sensitive to credit costs and capital rules.',1.05,4.33,11.19,.53,18,Navy,True,'center')
 rect(s,1.01,5.28,11.24,.74,Cyan)
 txt(s,'FY2025 reported combined common dividends and repurchases: C$11.3B; total payout ratio 57%.',1.27,5.44,10.7,.36,13.2,Ink,True,'center')
 # 12 peer logic
 s=slide(p,'Peer lens | use comparisons carefully','Two Canadian peers and four international banks; stock performance uses NYSE USD listings','RELATIVE ANALYSIS','Sources: original slide deck, RBC + peer FY2025 filings')
 group=[('CANADA','BMO','Bank of Montreal'),('CANADA','BNS','Scotiabank'),('GLOBAL','BCS','Barclays'),('GLOBAL','UBS','UBS Group'),('GLOBAL','HSBC','HSBC Holdings'),('GLOBAL','SAN','Banco Santander')]
 for i,(region,ticker,name) in enumerate(group):
  col=i%3;row=i//3;x=.82+col*4.2;y=2.26+row*1.56
  rect(s,x,y,3.98,1.31,Light)
  pill(s,region,x+.16,y+.16,1.06,Cyan,Green,9.5)
  txt(s,ticker,x+.17,y+.59,1.25,.42,22,Navy,True)
  txt(s,name,x+1.47,y+.61,2.32,.36,12.3,Ink)
 rect(s,.82,5.71,12.34,.77,Cyan)
 txt(s,'Warning: fiscal calendars, bank structures, revenue recognition and adjusted-vs-reported EPS can differ.',1.03,5.91,11.84,.33,13,Navy,True)
 # 13 valuation comparables simple
 s=slide(p,'Relative valuation | RY commands a premium','Approximate public forward P/E figures (NYSE, 05 Nov 2026)','MARKET VALUATION','Source: StockAnalysis ticker valuation pages; values can change intraday')
 pes=[('RY',16.36,Teal),('BMO',15.02,Navy2),('BNS',13.83,Navy2),('UBS',11.67,Gold),('HSBC',9.77,Gold),('SAN',9.05,Gold),('BCS',7.54,Gold)]
 for j,(symbol,v,col) in enumerate(pes):
  y=2.10+j*.57
  txt(s,symbol,.94,y,1.0,.34,12.7,Ink,True)
  rect(s,2.00,y+.08,8.68,.22,Mid)
  rect(s,2.00,y+.08,8.68*v/18,.22,col)
  txt(s,f'{v:.2f}×',10.87,y,1.20,.34,12.7,Navy,True,'right')
 rect(s,.94,6.43,11.46,.36,Cyan)
 txt(s,'Premium is not automatically mispricing: relate multiple to ROE, stability, growth and credit risk.',1.11,6.47,11.10,.28,11,Ink,True)
 # 14 investment premium
 s=slide(p,'Is the valuation premium justified?','Evidence supporting premium pricing, and the variables that could undermine it','VALUATION FRAMEWORK','Analytical interpretation of RBC financial disclosures')
 box_title(s,'Support for a premium','16.3% FY2025 ROE, 13.5% latest CET1, scale advantages, recurring wealth fees and earnings diversification.',.91,2.12,5.44,2.0,Green)
 box_title(s,'Reasons to apply a discount','FY2025 impaired loan ratio increased to 0.83%; Canadian geographic concentration and higher interest rates create sensitivity.',6.63,2.12,5.44,2.0,Red)
 rect(s,.9,4.45,11.18,1.38,Cyan)
 txt(s,'A multiple deserves support only if excess profitability survives the credit cycle.',1.12,4.71,10.71,.79,21,Navy,True,'center')
 # 15 10 yr return
 s=slide(p,'10-year stock-price performance','NYSE USD prices; split-adjusted; dividends excluded; 31 Oct 2016–30 Oct 2026','PRICE HISTORY','Source: Alpaca SIP stock bars, supplied source chart')
 pic=str(ASSETS/'original_chart_13.png')
 s.shapes.add_picture(pic,Inches(.7),Inches(1.91),width=Inches(11.98),height=Inches(3.80))
 returns=[('RY','+224.9%'),('UBS','+238.9%'),('BMO','+169.1%'),('BNS','+74.4%'),('BCS','+141.9%'),('HSBC','+127.0%'),('SAN','+131.2%')]
 for j,(a,b) in enumerate(returns):
  x=.83+j*1.75
  rect(s,x,5.89,1.65,.59,Cyan if a=='RY' else Light)
  txt(s,a,x+.09,5.99,.50,.26,10.6,Navy,True)
  txt(s,b,x+.55,5.99,1.05,.26,10.5,Navy,True,'right')
 # 16 one yr return
 s=slide(p,'1-year stock-price performance','NYSE USD prices; split-adjusted; dividends excluded; 31 Oct 2025–30 Oct 2026','PRICE HISTORY','Source: Alpaca SIP stock bars, supplied source chart')
 s.shapes.add_picture(str(ASSETS/'original_chart_14.png'),Inches(.7),Inches(1.91),width=Inches(11.98),height=Inches(3.80))
 returns=[('RY','+38.6%'),('BNS','+42.9%'),('BMO','+37.8%'),('UBS','+24.3%'),('HSBC','+21.9%'),('SAN','+10.2%'),('BCS','+3.8%')]
 for j,(a,b) in enumerate(returns):
  x=.83+j*1.75
  rect(s,x,5.89,1.65,.59,Cyan if a=='RY' else Light)
  txt(s,a,x+.09,5.99,.50,.26,10.6,Navy,True)
  txt(s,b,x+.56,5.99,1.03,.26,10.5,Navy,True,'right')
 # 17 correlation
 s=slide(p,'Diversification | bank shares remain positively correlated','Pearson correlation of 260 weekly split-adjusted USD log returns, Nov 2021–Oct 2026','PORTFOLIO RISK','Source: Alpaca SIP; reproduced against source deck')
 syms=['RY','BMO','BNS','BCS','UBS','HSBC','SAN']
 matrix=[[1,.77,.76,.60,.57,.45,.48],[.77,1,.73,.58,.58,.45,.48],[.76,.73,1,.53,.50,.41,.41],[.60,.58,.53,1,.70,.67,.69],[.57,.58,.50,.70,1,.57,.65],[.45,.45,.41,.67,.57,1,.67],[.48,.48,.41,.69,.65,.67,1]]
 xx=1.28;yy=2.28;cw=.68;ch=.53
 for j,a in enumerate(syms):txt(s,a,xx+.63+j*cw,yy-.50,cw,.29,10.5,Navy,True,'center');txt(s,a,xx-.65,yy+j*ch+.10,.62,.27,10.5,Navy,True,'right')
 for i,row in enumerate(matrix):
  for j,v in enumerate(row):
   col= '#DCECEF' if v<.55 else '#B2D8DC' if v<.7 else '#7FBDC5' if v<.9 else Navy2
   rect(s,xx+.63+j*cw,yy+i*ch,cw-.05,ch-.05,col,r=.03)
   txt(s,f'{v:.2f}',xx+.63+j*cw,yy+i*ch+.07,cw-.05,ch-.16,10.2,White if v==1 else Ink,True,'center')
 box_title(s,'Portfolio implication','RY–BMO correlation is about 0.77. Pairing Canadian banks does not eliminate shared macroeconomic and credit exposure.',7.45,2.45,4.65,1.74)
 box_title(s,'Relatively lower pair','RY–HSBC correlation is about 0.45 in this sample, but correlations change and cannot guarantee a hedge.',7.45,4.37,4.65,1.56,Gold)
 # 18 eps earnings
 s=slide(p,'Quarterly EPS | four consecutive beats','RY adjusted diluted EPS, C$; one consistent consensus data source','EARNINGS TRACK RECORD','Source: Yahoo Finance RY earnings-history table, accessed 05 Nov 2026')
 quarters=[('Q4 FY25',3.55,3.85),('Q1 FY26',3.84,4.08),('Q2 FY26',3.79,3.90),('Q3 FY26',4.07,4.28)]
 for i,(q,e,a) in enumerate(quarters):
  x=.88+i*3.03
  rect(s,x,2.29,2.79,3.17,Light)
  label(s,q,x+.2,2.58,2.28,Slate,11)
  txt(s,f'C${a:.2f}',x+.2,3.0,2.25,.66,30,Navy,True)
  txt(s,'Actual adjusted EPS',x+.2,3.68,2.2,.28,11,Slate)
  txt(s,f'C${e:.2f} consensus',x+.2,4.08,2.25,.32,11.8,Ink,True)
  ppct=(a/e-1)*100
  pill(s,f'+{ppct:.1f}% beat',x+.19,4.66,1.76,Cyan,Green,12)
 txt(s,'Surprise percentages are recalculated from rounded displayed actual/estimate amounts; vendor surprises may differ slightly.',.96,6.13,11.6,.52,11,Slate)
 # 19 forward
 s=slide(p,'Consensus outlook | FY2026 and FY2027','C$ estimates, not yet reported results; as displayed 05 Nov 2026','EARNINGS OUTLOOK','Sources: Yahoo Finance RY analysis; RBC FY2025 Annual Report')
 for j,(name,a,b,c,fmt) in enumerate([('REVENUE',66.605,72.42,76.44,'C$B'),('ADJUSTED EPS',14.43,16.42,17.91,'C$')]):
  y=2.24+j*2.16
  label(s,name,.94,y,4.8,Navy,13)
  for k,(n,v) in enumerate([('FY2025 actual',a),('FY2026E',b),('FY2027E',c)]):
   x=1.00+k*4.08
   rect(s,x,y+.47,3.84,1.12,Light if k==0 else Cyan)
   txt(s,f'{v:.2f}{"B" if fmt=="C$B" else ""}',x+.18,y+.60,3.30,.50,24,Navy,True)
   txt(s,n,x+.18,y+1.14,3.44,.26,11,Slate,True)
  txt(s,f'FY26E growth: {(b/a-1)*100:+.1f}%     FY27E growth: {(c/b-1)*100:+.1f}%',1.01,y+1.73,10.9,.27,12,Green,True)
 # 20 governance
 s=slide(p,'Governance and management | continuity matters','Biographies as of Nov 2026; no third-party executive portraits reused','LEADERSHIP','Sources: RBC leadership page; RBC 2026 Proxy Circular')
 leaders=[('Jacynthe Côté','Independent Board Chair','Director since 2014; chair since Apr 2023'),('David McKay','President & Chief Executive Officer','CEO since Aug 2014'),('Katherine Gibson','Chief Financial Officer','CFO since Sep 2024')]
 for j,(a,b,c) in enumerate(leaders):
  x=.85+j*4.17
  rect(s,x,2.32,3.86,2.30,Light)
  txt(s,a,x+.22,2.63,3.42,.40,21,Navy,True)
  txt(s,b,x+.22,3.18,3.42,.29,12,Green,True)
  ln(s,x+.22,3.73,x+3.62,3.73,Mid,1)
  txt(s,c,x+.22,3.92,3.4,.40,13,Ink)
 rect(s,.85,5.05,12.18,.95,Cyan)
 txt(s,'What to monitor: succession planning, capital allocation discipline, credit controls and execution on the HSBC Canada integration.',1.10,5.22,11.71,.55,14,Navy,True)
 # 21 catalysts/risks
 s=slide(p,'Catalysts and risks | a balanced forward view','Management disclosures through Q3 FY2026 and Bank of Canada decision of 28 Oct 2026','DRIVERS & RISKS','Sources: RBC Q3 FY2026; Bank of Canada 28 Oct 2026')
 box_title(s,'Tailwinds','Wealth-management fees and capital-markets activity; Personal and Commercial Banking balance growth; HSBC Canada cross-sell; CET1 at 13.5%.',.78,2.17,5.84,2.45,Green)
 box_title(s,'Headwinds','FY2025 impaired loans rose; Q3 FY2026 PCL was ~C$1.0B; a 2.50% policy rate can constrain borrower demand and credit performance.',6.79,2.17,5.84,2.45,Red)
 rect(s,.80,5.10,11.80,.92,Cyan)
 txt(s,'Key asymmetry: stable credit costs support premium valuation; credit deterioration can compress both earnings and the P/E multiple.',1.05,5.28,11.25,.54,15,Navy,True)
 # 22 valuation scenario
 spot=203.76;fx=.723;scens=[('BEAR',17.00,14.0,Red),('BASE',17.91,16.5,Teal),('BULL',19.00,19.0,Green)]
 s=slide(p,'Illustrative 12-month valuation | scenario not a price target','FY2027 EPS assumptions × terminal P/E × illustrative USD/CAD factor','SCENARIO VALUATION','Stock benchmark US$203.76 (05 Nov close); USD/CAD conversion 0.723')
 for i,(name,eps,pe,col) in enumerate(scens):
  x=.81+i*4.19
  tgt=eps*pe*fx;ret=(tgt/spot-1)*100
  rect(s,x,2.23,3.96,3.47,Light)
  label(s,name,x+.23,2.50,3.4,col,13)
  txt(s,f'US${tgt:.0f}',x+.23,2.90,3.43,.65,32,Navy,True)
  txt(s,f'{ret:+.1f}% price return',x+.23,3.69,3.37,.39,17,Green if ret>=0 else Red,True)
  ln(s,x+.23,4.29,x+3.69,4.29,Mid,1)
  txt(s,f'C${eps:.2f} EPS × {pe:.1f}× P/E',x+.23,4.49,3.45,.36,12,Ink)
  txt(s,'FX: US$0.723/C$',x+.23,4.96,3.37,.29,10.6,Slate)
 txt(s,'FY2027 consensus EPS is ~C$17.91; bear/bull EPS and all terminal multiples are author assumptions. No dividend, fees or FX changes in price return.',.89,6.13,11.60,.57,10.6,Slate)
 # 23 conclusion
 s=slide(p,'Investment conclusion | HOLD / WATCHLIST','An independent analytical judgment, not personalized investment advice','INVESTMENT DECISION','For educational research; assumptions and sources in repository')
 rect(s,.87,2.13,11.56,1.04,Navy)
 txt(s,'High-quality bank  ≠  attractive entry at every price',1.14,2.30,11.08,.70,22,White,True,'center')
 for i,(title,body) in enumerate([
 ('What I like','Durable franchise, diversified earnings, strong capital and improving FY2025 efficiency.'),
 ('Why I would wait','The model’s base case has limited price upside while credit and multiple risks remain meaningful.'),
 ('What changes my mind','A lower entry multiple, upside to FY2027 EPS, or evidence of stabilizing impaired loans and PCL.')]):
  y=3.52+i*.83
  rect(s,.96,y,11.33,.72,Light)
  txt(s,title,1.18,y+.14,3.00,.37,13.3,Navy,True)
  txt(s,body,4.13,y+.13,7.86,.46,11.8,Ink)
 txt(s,'Next catalyst: fiscal FY2026 results — 03 Dec 2026.',1.08,6.37,10.81,.34,13,Green,True)
 assert IDX==TOTAL, (IDX,TOTAL)
 out=OUT/'RBC_Independent_Equity_Research_Full.pptx';p.save(out);print('full',out,len(p.slides))

# LinkedIn 4:5 portrait document slides

def li_base(p,n,kicker,title,subtitle=''):
 s=p.slides.add_slide(p.slide_layouts[6])
 rect(s,0,0,8.0,10.0,White,r=0);rect(s,0,0,8.0,.13,Teal,r=0)
 label(s,kicker,.55,.43,7.10,Teal,10.1)
 txt(s,title,.56,.95,6.88,1.13,29,Navy,True,vertical='top')
 if subtitle:txt(s,subtitle,.58,2.13,6.78,.67,13,Slate,vertical='top')
 ln(s,.57,9.37,7.48,9.37,Mid,.8)
 txt(s,'AAMIR SOHAIB  /  INDEPENDENT EQUITY RESEARCH',.60,9.48,6.24,.25,9,Slate,True)
 txt(s,f'{n:02d}/08',6.94,9.46,.48,.30,10,Teal,True,'right')
 return s

def li_card(s,x,y,w,h,headline,body,accent=Teal):
 rect(s,x,y,w,h,Light)
 rect(s,x+.17,y+.20,.06,h-.40,accent,r=0)
 txt(s,headline,x+.40,y+.28,w-.56,.53,17,Navy,True,vertical='top')
 txt(s,body,x+.39,y+.94,w-.57,h-1.14,13,Ink,vertical='top')

def linkedin():
 p=Presentation();p.slide_width=Inches(8);p.slide_height=Inches(10)
 s=li_base(p,1,'ROYAL BANK OF CANADA  /  NYSE: RY','A quality franchise. Is the stock priced for it?','Independent equity research | FY2025 financials + FY2026–27 outlook | 05 Nov 2026')
 rect(s,.56,3.24,6.85,3.22,Navy)
 txt(s,'HOLD  /  WATCHLIST',.88,3.66,6.18,.70,29,White,True)
 txt(s,'Capital strength and earnings quality support a premium — but my valuation base case leaves limited upside.',.88,4.55,6.05,1.19,20,'#DBF4F7',vertical='top')
 txt(s,'8-slide research summary →',.58,7.53,6.8,.55,20,Teal,True)
 txt(s,'Sources and complete methodology are in the full GitHub research pack.',.58,8.45,6.58,.47,10.5,Slate)
 s=li_base(p,2,'01  /  INVESTMENT THESIS','Three numbers define the setup','FY2025 actuals unless labeled')
 for i,(v,lab,description) in enumerate([('16.3%','ROE','Profitable franchise'),('13.5%','CET1 (Q3 FY26)','Capital buffer'),('0.83%','Impaired loans (FY25)','Credit stress rising')]):
  y=3.20+i*1.60;rect(s,.65,y,6.62,1.39,Light)
  txt(s,v,.93,y+.22,2.48,.66,29,Navy,True)
  txt(s,lab,3.51,y+.20,3.24,.34,14,Ink,True)
  txt(s,description,3.51,y+.62,3.35,.33,12,Slate)
 txt(s,'Takeaway: quality does not remove credit-cycle risk.',.71,8.33,6.5,.56,16,Teal,True)
 s=li_base(p,3,'02  /  FINANCIAL PERFORMANCE','FY2025 net income grew 25.4%','RBC annual results (C$ billions)')
 vals=[('FY2024 earnings',16.24,Navy2),('FY2025 earnings',20.369,Teal)]
 for i,(name,v,col) in enumerate(vals):
  y=3.16+i*1.17
  txt(s,name,.76,y,3.08,.43,14,Ink,True)
  rect(s,.80,y+.58,5.73,.39,Mid)
  rect(s,.80,y+.58,5.73*v/21,.39,col)
  txt(s,f'{v:.2f}',6.59,y+.52,.62,.49,15,Navy,True,'right')
 li_card(s,.68,5.88,6.59,2.12,'The bridge','+C$9.26B revenue – C$2.34B expense – C$1.13B provisions – C$1.66B taxes = +C$4.13B net income.',Green)
 s=li_base(p,4,'03  /  BUSINESS MODEL','Five engines. One consolidated bank.','FY2025 segment revenue; C$ billions')
 for i,(name,v) in enumerate([('Wealth',22.38),('Personal',19.85),('Capital Markets',14.43),('Commercial',8.56),('Insurance',1.32)]):
  y=3.0+i*.89
  txt(s,name,.68,y,2.58,.37,12.3,Ink,True)
  rect(s,3.15,y+.05,3.12,.27,Mid)
  rect(s,3.15,y+.05,3.12*v/23,.27,Teal if i<2 else Navy2)
  txt(s,f'{v:.1f}',6.45,y-.01,.72,.39,12,Navy,True,'right')
 rect(s,.69,7.76,6.57,.77,Cyan);txt(s,'Diversification supports earnings resilience.',.97,7.94,5.98,.39,15,Navy,True)
 s=li_base(p,5,'04  /  RISK FACTORS','Credit quality is the warning sign','FY2025 vs FY2024; RBC annual report')
 for i,(a,b,c) in enumerate([('0.83%','0.59%','Gross impaired loans / related loans'),('0.43%','0.35%','PCL / average net loans')]):
  y=3.1+i*2.10
  rect(s,.68,y,6.55,1.81,Light)
  txt(s,a,.94,y+.24,2.23,.66,29,Red,True)
  txt(s,'FY2025',.94,y+.95,2.3,.29,12,Slate)
  txt(s,'vs '+b,3.48,y+.36,3.1,.50,23,Navy,True)
  txt(s,c,3.48,y+1.02,3.13,.38,12,Ink)
 txt(s,'Strong capital and higher impaired-loan ratios can coexist.',.72,8.03,6.54,.82,16,Navy,True)
 s=li_base(p,6,'05  /  EARNINGS OUTLOOK','Analysts expect more growth','C$ estimates, not yet reported results')
 for j,(a,b,c,growth) in enumerate([('FY2026E revenue','72.42B','+8.7%','vs FY2025'),('FY2027E revenue','76.44B','+5.6%','vs FY2026E'),('FY2026E adj. EPS','16.42','+13.8%','vs FY2025'),('FY2027E adj. EPS','17.91','+9.1%','vs FY2026E')]):
  y=3.00+j*1.25
  rect(s,.73,y,6.49,1.10,Light)
  txt(s,a,.91,y+.16,3.15,.32,12.4,Ink,True)
  txt(s,'C$'+b,.92,y+.52,3.15,.42,20,Navy,True)
  txt(s,c,4.42,y+.20,2.19,.43,19,Green,True,'right')
  txt(s,growth,4.42,y+.68,2.18,.21,10.1,Slate,False,'right')
 txt(s,'Consensus source: Yahoo Finance RY analysis, checked 05 Nov 2026.',.76,8.54,6.39,.43,10.5,Slate)
 s=li_base(p,7,'06  /  VALUATION','The payoff depends on your assumptions','Illustrative 12-month USD scenarios; not analyst price targets')
 for i,(name,eps,pe,col) in enumerate([('BEAR',17,14,Red),('BASE',17.91,16.5,Teal),('BULL',19,19,Green)]):
  tgt=eps*pe*.723;ret=(tgt/203.76-1)*100;y=3.08+i*1.56
  rect(s,.76,y,6.39,1.38,Light)
  txt(s,name,1.00,y+.19,1.33,.31,12,col,True)
  txt(s,f'${tgt:.0f}',2.27,y+.20,2.13,.58,26,Navy,True)
  txt(s,f'{ret:+.1f}%',4.70,y+.18,1.81,.60,24,Green if ret>0 else Red,True,'right')
  txt(s,f'C${eps:.2f} EPS  ·  {pe:g}× P/E',.98,y+.94,5.85,.27,10.2,Slate)
 txt(s,'Benchmarked to US$203.76 on Nov 5; uses US$0.723/C$ illustrative FX.',.77,8.27,6.34,.55,10.3,Slate)
 s=li_base(p,8,'07  /  CONCLUSION','My view: HOLD / WATCHLIST','Independent research, not individualized investment advice')
 li_card(s,.74,3.02,6.43,1.60,'The positive','Strong ROE, capital resilience and multiple engines of growth.',Green)
 li_card(s,.74,4.87,6.43,1.60,'The constraint','The current premium leaves a modest base-case return and real credit downside.',Red)
 li_card(s,.74,6.72,6.43,1.61,'What I will watch','FY2026 results (Dec 3), FY2027 estimates, credit trends and the entry multiple.',Gold)
 out=OUT/'RBC_LinkedIn_8_Slide_Carousel.pptx';p.save(out);print('linkedin',out,len(p.slides))

if __name__=='__main__':
 presentation();linkedin()
