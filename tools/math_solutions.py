"""Independent calculations and plots for variant 6; no Mathcad is invoked."""
from pathlib import Path
import sys,math,json
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'tmp/python-packages'))
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

plt.rcParams.update({'font.family':'DejaVu Sans','font.size':11,'axes.spines.top':False,'axes.spines.right':False})
L1=ROOT/'02_Mathcad_выражения_и_графики'
L2=ROOT/'03_Mathcad_массивы_и_системы'
for d in [L1,L2]:(d/'Графики').mkdir(parents=True,exist_ok=True)

def save(fig,d,name):
    fig.tight_layout();fig.savefig(d/'Графики'/name,dpi=190);plt.close(fig)

def curve(x,y,xlabel,ylabel):
    fig,ax=plt.subplots(figsize=(7,3.7));ax.plot(x,y,color='#174e85',lw=1.8)
    ax.axhline(0,color='gray',lw=.7);ax.axvline(0,color='gray',lw=.7);ax.grid(alpha=.2)
    ax.set(xlabel=xlabel,ylabel=ylabel)
    return fig,ax

I=2*(math.sqrt(1+math.log(10))-math.sqrt(1+math.log(3)))
S=math.fsum(7**k*math.sin(k*math.pi/4) for k in range(1,11))
# An independent closed form groups the eight-period trigonometric values.
Scheck=7**10-7**6+7**2+(7+7**3-7**5-7**7+7**9)/math.sqrt(2)
assert abs(S-Scheck)<1e-5
x=np.r_[np.linspace(-10,-.03,2600),np.nan,np.linspace(.03,10,2600)]
fig,ax=curve(x,x*np.sin(1/x),'x','y = x sin(1/x)');ax.plot([0],[0],marker='o',mfc='white',mec='#174e85');save(fig,L1,'01_функция.png')
x=np.linspace(-3,3,80);X,Y=np.meshgrid(x,x);Z=X**3-X*Y+Y**2-2*X+3*Y-4
fig=plt.figure(figsize=(7,4.5));ax=fig.add_subplot(111,projection='3d');ax.plot_surface(X,Y,Z,cmap='viridis',linewidth=0,antialiased=True);ax.set(xlabel='x',ylabel='y',zlabel='z');ax.view_init(24,-57);save(fig,L1,'02_поверхность.png')
x=np.linspace(-3,1,500);fig,ax=curve(x,x**3+3*x*x+6*x+4,'x','f(x)');ax.plot(-1,0,'o',color='#b63232');ax.annotate('(-1; 0)',(-1,0),xytext=(-.8,7));save(fig,L1,'03_корень.png')

B=np.array([[2,1,0,1],[0,3,1,0],[1,0,4,1],[0,1,0,5]],float)
C=np.diag([1.,2.,1.,2.])
D=np.array([[4,0,1,0],[1,5,0,1],[0,1,6,0],[1,0,1,7]],float)
A=(B-C)@(D-C)+(D-B)@np.linalg.inv(C)
assert np.allclose(C@np.linalg.inv(C),np.eye(4))
f=lambda x:2*x**3-8*x*x+3*x+50
lo,hi=-2.,-1.
for _ in range(60):
 mid=(lo+hi)/2
 if f(lo)*f(mid)<=0:hi=mid
 else:lo=mid
root=(lo+hi)/2
M=np.array([[1,1,-1],[1,1,2],[1,1,1]],float);v=np.array([0,9,6],float)
assert np.linalg.matrix_rank(M)==np.linalg.matrix_rank(np.column_stack([M,v]))==2
for t in [-20,0,1.5,3,100]:assert np.allclose(M@np.array([t,3-t,3]),v)
x=np.linspace(-3,4,700);fig,ax=curve(x,f(x),'x','f(x)');ax.plot(root,0,'o',color='#b63232');ax.annotate(f'x ≈ {root:.6f}',(root,0),xytext=(-1.5,-15));save(fig,L2,'01_кубическое_уравнение.png')
x=np.linspace(.2,7,700);fig,ax=plt.subplots(figsize=(7,3.7));ax.plot(x,x/math.e,label='y = x/e');ax.plot(x,np.log(x),label='y = ln(x)');ax.plot(math.e,1,'o',color='#b63232');ax.annotate('(e; 1)',(math.e,1),xytext=(3.2,.7));ax.set(xlabel='x',ylabel='y');ax.grid(alpha=.2);ax.legend();save(fig,L2,'02_нелинейная_система.png')
data={'lab1':{'integral':I,'sum':Scheck,'root':-1},'lab2':{'B':B.tolist(),'C':C.tolist(),'D':D.tolist(),'A':A.tolist(),'maximum':float(A.max()),'cubic_root':root,'residual':f(root),'linear_rank':2,'linear_solution':'(t, 3-t, 3), t in R','nonlinear_solution':[math.e,1.]}}
(ROOT/'tmp/math_results.json').write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(data,ensure_ascii=False,indent=2))
