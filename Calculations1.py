#import xlwings as xw
import pandas as pd 
import numpy as np
s=r'‪C:\Users\aayushi\Desktop\Book2d.xlsx'
s = s.lstrip('\u202a')
xlsx = pd.ExcelFile(s)
df1=pd.read_excel(xlsx,'Sheet1')
df10=pd.read_excel(xlsx,'Sheet2')

col_mapping = [f"{c[0]}:{c[1]}" for c in enumerate(df1.columns)]
#TABLE V CALCULATIONS-the calculations are down to up.

#BLOCK                          PARAMETER                                                  PYTHON FUNCTION
#(e)                      Effective Principal  Quantun no n'                                 EffectQuantNo
#                             n'=  (Sum of  orbit exponents for each occupied atomic orbital
#                 including screening charges(1s+2s+..+6s))/z(Atomic Number)

#(d)                         q=(1.98ξn)/n'                                                     qi
#                      ξ is the overall atomic orbital exponent(Values given in col.4 Table V)
#                        n is the principal quantum number(Values given in col.5 Table V)
#                             n' taken from the calculation of block (e)   

#(b)                        P_o is called the spatial energy parameter                          Po
#                                P_0=  (q^2 η_g)/(q^2+ η_g )
#                          q as calculated in block(d),Values given in col.6 Table V
#                     η_g  Values given in col.7 Table V

#(a)                         P_E=  P_o/r_i  effective Polarisation Energy Parameter in eV       PE
#                           P_o , block(b) Values given in col.7 Table V
#                      r_i, Atomic radii in Ang. ,Values given in col.3 Table V 


#First Sheet of the worbook contains calculations for the effective principal Quantum Number n' or n1.
#  The effective principal quantum number depends on the effective nucleus charge ,which is a result of 
#of the screening of the nucleus charge by the inner orbital electrons. 
#  The table consists the value of orbital exponents 𝝃 ,which is given by 𝝃 = (Z- σ)/n ,with z being the 
#nuclear charge and  𝝈 being the screening constant. Effective principle quantum number is calculated by dividing
#the sum of orbital exponent with z, including some conditions. when the p orbitals are included, there is an 
#addition of correction factor , multiple of 0.5 ,as shown below.


cols = ['1s','2s','2p','3s','3p','4s','4p','3d','5s','4d','4f','4f1', '5d', '6s','2s','2p','5p']  # We don't want to convert the Final grade column.
for col in cols:  # Iterate over chosen columns
  df1[col] = [float(str(val).replace('. ','.')) for val in df1[col].values]
df1['Sum']=df1.iloc[:,2:17].astype(float).sum(axis=1)

#x is the nuclear charge z , y is the sum of orbital exponents 𝝃 
#n1 is the effective principal quantum number denoted as n'


def EffectQuantNo(x,y):
    if(x<5):
        return round((y/x),1)
    elif(x>=5 and x<=10):
        return round((y/x) + 0.5,1)
    elif(x>=13 and x<19):
        return round((y/x) + 1.5,1) 
    elif(x>=31 and x<=36):
        return round((y/x) + 1.5,1)
    elif(x>49 and x<=54):
        return round((y/x) + 2.0,1)
    else :
        return round((y/x),1)

#Below step applies the function to calcualte n'

df1.insert(18,'n1',np.nan)
for i in range(0,78):
  df2=df1.iloc[i,1].astype(float)
  df3=df1.iloc[i,17].astype(float)
 
  

  df4 = EffectQuantNo(df2,df3)
  df1.iloc[i,18]=df4

print(df1)
p=r'‪C:\Users\aayushi\Desktop\output2dd.xlsx'
p = p.lstrip('\u202a')
writer = pd.ExcelWriter(p, engine = 'xlsxwriter')
df1.to_excel(writer, sheet_name = 'SheetA')
df5 = df1[['Element', 'n1']].copy()

#ri is the atomic radii in Angstrong
df5.loc[:,'ri']=df10.iloc[:,0]

#OrbExp is the orbital exponent value of the respective elements.
df5.loc[:,'OrbExp']=df10.iloc[:,1]

#n gives the principal quantum number of the elements
df5.loc[:,'n']=df10.iloc[:,2]

#In second sheet of the workbook , the table calculates the value of P_o and P_E given below. The other columns used 
#are q , η  (GlobHard), orbital exponent as OrbExp, atomic radii as ri.
# 1/P_o =  1/q^2 +1/η  
# P_E=  P_o/r_i 
#P_o is called the spatial energy parameter, P_E is the called the effective energy parameter , averarge polaristaion
#energy of the valence electrons
#η  is the Global Hardness Factor used as GlobHard variable here.

#q=z_eff/n'
#z_(eff )= ξ*n
#q=(ξ*n)/n' 

def qi(n1,OrbExp,n):
    return round((OrbExp*n*1.98)/n1,3)

df5.insert(5,'q',np.nan)
for i in range(0,78):
  df2=df5.iloc[i,1].astype(float)
  df3=df5.iloc[i,3].astype(float)
  df4=df5.iloc[i,4].astype(float)
 
  

  df7 = qi(df2,df3,df4)
  df5.iloc[i,5]=df7


df5.insert(6,'GlobHard',np.nan)


#GlobHard gives the Global Hardness Factor which is a measure of the polarisation capacity of an element.
df5.loc[:,'GlobHard']=df10.iloc[:,3]

#Below function calculates 1/P_o =  1/q^2 +1/η  
def  Po(q,GlobHard):
      return round((q*q*GlobHard)/(q*q+GlobHard),3)
#Po is the Spatial Energy Paramter in eV-Angs.
df5.insert(7,'Po',np.nan)
for i in range(0,78):
  df2=df5.iloc[i,5].astype(float)
  df3=df5.loc[i,'GlobHard'].astype(float)
  df7 =  Po(df2,df3)
  df5.iloc[i,7]=df7

#Below function calculates P_E
def  PE(ri,Po):
      return round((Po/ri),3)
#PE is the effective Polarisation Energy Parameter in eV
df5.insert(8,'PE',np.nan)
for i in range(0,78):
  df2=df5.iloc[i,2].astype(float)
  df3=df5.iloc[i,7].astype(float)
  df7 = PE(df2,df3)
  df5.iloc[i,8]=df7
df5.to_excel(writer, sheet_name = 'SheetB')


writer.save()