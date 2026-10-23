"""Recalculates the Excel model with the `formulas` engine, checks for errors, and compares Base/Bull/Bear
revenue and operating income against an independent pandas re-implementation of the forecast logic."""
import sys, os, sqlite3, numpy as np, pandas as pd, formulas
ROOT=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
X=ROOT+'/excel/alphabet_pnl_model.xlsx'
xl=formulas.ExcelModel().loads(X).finish()
sol=xl.calculate()
def v(sheet,cell):
    k=f"'[alphabet_pnl_model.xlsx]{sheet.upper()}'!{cell}"
    return sol[k].value[0,0] if k in sol else ''
# 1. errors anywhere?
errs=[k for k,val in sol.items() if hasattr(val,'value') and isinstance(val.value,np.ndarray) and any(isinstance(x,(str,)) and str(x).startswith('#') or (hasattr(x,'__class__') and 'XlError' in type(x).__name__) for x in val.value.ravel())]
print('error cells:',len(errs), errs[:5])
# 2. independent re-implementation
con=sqlite3.connect(ROOT+'/data/alphabet.db')
df=pd.read_sql("select * from v_pnl where quarter_end>='2019-03-31'",con)
segs=['rev_search','rev_youtube','rev_network','rev_subscriptions','rev_cloud','rev_other_bets']
fade=0.9
deltas={'Bear':[-.02,-.02,0,0,-.05,0],'Base':[0]*6,'Bull':[.02,.02,0,0,.05,0]}
def run(sc):
    H={s:list(df[s]) for s in segs}
    for j,s in enumerate(segs):
        g=H[s][-1]/H[s][-5]-1
        for k in range(1,7):
            H[s].append(H[s][-4]*(1+g*fade**k+deltas[sc][j]))
    tot=[sum(H[s][i] for s in segs)+ (df.rev_hedging.iloc[i] if i<30 else 0) for i in range(36)]
    ratios={}
    for key,col in (('cor','cost_of_revenues'),('rd','rd'),('sm','sales_marketing'),('ga','g_and_a')):
        r=[df[col].iloc[i]/tot[i] for i in range(30)]
        for i in range(30,36): r.append(np.mean(r[i-4:i]))
        ratios[key]=r
    base_tot=None
    return H,tot,ratios
Hb,totb,rb=run('Base')
def oi(sc):
    H,tot,r=run(sc)
    out=[]
    for i in range(30,36):
        rd=totb[i]*rb['rd'][i]; ga=totb[i]*rb['ga'][i]
        if sc=='Base': rd=tot[i]*r['rd'][i]; ga=tot[i]*r['ga'][i]
        out.append(tot[i]*(1-r['cor'][i]-r['sm'][i])-rd-ga)
    return tot,out
cols=['AG','AH','AI','AJ','AK','AL']
for sc in ('Base','Bull','Bear'):
    tot,o=oi(sc)
    xt=[v('Model_'+sc,f'{c}15') for c in cols]; xo=[v('Model_'+sc,f'{c}35') for c in cols]
    print(sc,'max |rev diff|',max(abs(a-b) for a,b in zip(xt,tot[30:])),'max |OI diff|',max(abs(a-b) for a,b in zip(xo,o)))
# history ties to DB
print('Q4 25 rev/OI in sheet:',v('Model_Base','AE15'),v('Model_Base','AE35'),' DB:',df.revenue.iloc[28],df.operating_income.iloc[28])
print('Scenarios sheet:')
for r in range(4,18):
    print([ (round(v('Scenarios',f'{c}{r}'),3) if isinstance(v('Scenarios',f'{c}{r}'),(int,float,np.floating)) else v('Scenarios',f'{c}{r}')) for c in 'ACDEFGHI'])
