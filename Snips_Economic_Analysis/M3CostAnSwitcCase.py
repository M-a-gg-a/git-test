# -*- coding: utf-8 -*-
"""
Created on Thu Jan  7 16:27:28 2021

@author: Kolovos
"""
"""
Economic Analysis Program M3

Created on Wed Dec 23 08:53:35 2020
@author: Kolovos
"""
'''
# --------------------------Assumptions------------------------------
1. GIVES A YEAR FORCAST (main reason missing gas frwd curve in farther future)
 . Also the 'Market values' & 'weather patterns' in M2 will have to be for multiple years 
2. values  given in 'Input from M2' section MUST be initialized in prior modules
 . preferably M1 (all var = 0)
'''
'''
Using Switch/Case mecanism --- doesn't need to indicate 'Existing Resources'
------'Key' for numbering/labeling additional resources and combinations-------

Additional resources
1 = Wind 
2 = Solar
3 = Battery Wind
4 = Battery Solar
5 = Combine Cycle (CC)
6 = Combination Tutbine (CT)
7 = RIC
8 = Aeroderivative (A)
2 conmbinations= double #
3 conmbinations= triple #    (not used in senarios here)
4 conmbinations= quadruple # (not used in senarios here)
'''
# PURPOSE: Read only from excel (not Write)

#-----------------------------Input from M2------------------------------------
# Create Test Variables (outcomes of M2)
# these results should arive from M2 (here are just for testing)
    
#MWh in Year for Solar 
MWhS = 500
# curtailment additions [MWh] Solar in a Year
MWhS_crt = 45   # used by battery when available - Solar
MWhS_Xcrt = 5   # Extra curtailment not stored - Solar
MWhS_Tcrt = 50  # Total (used & stored) - Solar
 
#MWh in Year for Wind
MWhW = 550
# curtailment additions [MWh] Wind in a Year
MWhW_crt = 50   # used by battery when available - Wind
MWhW_Xcrt = 10  # Extra curtailment not stored - Wind
MWhW_Tcrt = 60  # Total (used & stored) - Wind
    
# Values from M2 independed of Existing Resource    
# CC MWh produced each month
JrCC = 30; FbCC= 15; MrCC = 40; ApCC =70; MyCC = 35;JnCC = 20;
JlCC = 70; AgCC=60; SpCC = 35;OcCC = 60; NvCC = 10; DcCC = 50
# CT MWh produced each month
JrCT = 0; FbCT= 0; MrCT = 0; ApCT =0; MyCT = 0;JnCT = 0;
JlCT = 0; AgCT=0; SpCT = 0;OcCT = 0; NvCT = 0; DcCT = 0
# RIC MWh produced each month
JrRIC = 30; FbRIC= 15; MrRIC = 0; ApRIC =70; MyRIC = 35;JnRIC = 20;
JlRIC = 70; AgRIC=60; SpRIC = 35;OcRIC = 60; NvRIC = 0; DcRIC = 50
# A MWh produced each month
JrA = 0; FbA= 0; MrA = 0; ApA = 0; MyA = 0;JnA = 0;
JlA = 0; AgA= 0; SpA = 0;OcA = 00; NvA = 00; DcA = 0

# small Cc, Ct, Ric _a are used when COMBINED additional resource   
# Cc MWh produced each month
JrCc = 15; FbCc= 10; MrCc = 20; ApCc =35; MyCc = 20;JnCc = 10;
JlCc = 35; AgCc=30; SpCc = 20;OcCc = 10; NvCc = 5; DcCc = 25
# Ct MWh produced each month
JrCt = 0; FbCt= 0; MrCt = 0; ApCt =0; MyCt = 0;JnCt = 0;
JlCt = 0; AgCt=0; SpCt = 0;OcCt = 0; NvCt = 0; DcCt = 0
# Ric MWh produced each month
JrRic = 15; FbRic= 10; MrRic = 0; ApRic =35; MyRic = 20;JnRic = 10;
JlRic = 35; AgRic=15; SpRic = 20;OcRic = 30; NvRic = 0; DcRic = 25
# _a MWh produced each month
Jr_a = 0; Fb_a= 0; Mr_a = 0; Ap_a = 0; My_a = 0;Jn_a = 0;
Jl_a = 0; Ag_a= 0; Sp_a = 0;Oc_a = 00; Nv_a = 00; Dc_a = 0

#*************************Actual Code Starts here***************************
#------------------List yearMWh =[monthMWh,monthMWh... ]--------------------
# CC MWh in Year
MWhCC =[JrCC,FbCC,MrCC,ApCC,MyCC,JnCC,         # this is the val. of the list
        JlCC,AgCC,SpCC,OcCC,NvCC,DcCC]
# CT MWh in Year
MWhCT =[JrCT,FbCT,MrCT,ApCT,MyCT,JnCT,         # this is the val. of the list
        JlCT,AgCT,SpCT,OcCT,NvCT,DcCT]
# RIC MWh in Year
MWhRIC =[JrRIC,FbRIC,MrRIC,ApRIC,MyRIC,JnRIC,  # this is the val. of the list
        JlRIC,AgRIC,SpRIC,OcRIC,NvRIC,DcRIC]
# A MWh in Year
MWhA =[JrA,FbA,MrA,ApA,MyA,JnA,                # this is the val. of the list
        JlA,AgA,SpA,OcA,NvA,DcA]

# ---------------------When gas COMBINED additional resource-------------------
# Cc MWh in Year
MWhCc =[JrCc,FbCc,MrCc,ApCc,MyCc,JnCc,         # this is the val. of the list
        JlCc,AgCc,SpCc,OcCc,NvCc,DcCc]
# Ct MWh in Year
MWhCt =[JrCt,FbCt,MrCt,ApCt,MyCt,JnCt,         # this is the val. of the list
        JlCt,AgCt,SpCt,OcCt,NvCt,DcCt]
# Ric MWh in Year
MWhRic =[JrRic,FbRic,MrRic,ApRic,MyRic,JnRic,  # this is the val. of the list
        JlRic,AgRic,SpRic,OcRic,NvRic,DcRic]
# _a MWh in Year
MWh_a =[Jr_a,Fb_a,Mr_a,Ap_a,My_a,Jn_a,         # this is the val. of the list
        Jl_a,Ag_a,Sp_a,Oc_a,Nv_a,Dc_a]

#--------------------------Formulas------------------------------------------
# Reading data from 'EconomicAnalysisM3.xlsx' in 'project program' folder
# pandas can read formulas directly from excel sheet
import pandas as pd

sheet = pd.read_excel('EconomicAnalysisM3.xlsx') 

# starting from 0
C26 = sheet.iloc[24, 2]
D27 = sheet.iloc[25, 3]
E26 = sheet.iloc[24, 4]
E27= sheet.iloc[25, 4]
#--------------------------------------------------------------
#RESOURCE---RESOURCE---RESOURCE---RESOURCE---RESOURCE---RESOURCE------------- 

#-----------Based upon Existing Resources-----MWhW or MWhS = 0 --Short Itself-
# 1a) Wind (included  Wind_curtail) - NO BATTERY (in use)
Wind_D = C26*MWhW                          # $
WindCurt_D = C26*MWhW_Tcrt                 # $
if MWhW > 0:
    W_DpE = (Wind_D + WindCurt_D)/ MWhW    # $/MWh
else:
     W_DpE =0
Sen1 = [MWhW,Wind_D,MWhW_Tcrt,WindCurt_D,W_DpE]

# 1b) wind (included  Wind_curtail) + BATTERY (in use)
Wind_D = C26*MWhW                          # $
windCurt_D = C26*MWhW_Xcrt                 # $
if MWhW > 0:
    w_DpE = (Wind_D + windCurt_D)/ MWhW    # $/MWh
else:
  w_DpE =0  
  
 # 3 Batteries Wind (included  *curtail taken at wind)
BattW_D = E26*MWhW_crt                    # $
if MWhW_crt >0:                           
    BatW_DpE = BattW_D/ MWhW_crt          # $/MWh
else:
  BatW_DpE =0

# 2a) Solar (included  Solar_curtail) - NO BATTERY (in use)
Solar_D = D27*MWhS                          # $
SolarCurt_D = D27*MWhS_Tcrt                 # $
if MWhS > 0:
    Sol_DpE = (Solar_D + SolarCurt_D)/ MWhS # $/MWh
else:
    Sol_DpE = 0 
Sen2 = [MWhS,Solar_D,MWhS_Tcrt,SolarCurt_D,Sol_DpE]

# 2b) solar (included  Solar_curtail) + BATTERY (in use)
Solar_D = D27*MWhS                           # $
solarCurt_D = D27*MWhS_Xcrt                  # $
if MWhS > 0:
    sol_DpE = (Solar_D + solarCurt_D)/ MWhS  # $/MWh
else:
    sol_DpE = 0 

 # 4 Batteries Solar  (included  *curtail taken at Solar)
BattS_D = E27*MWhS_crt                    # $
if MWhS_crt >0:                           
    BatS_DpE = BattS_D/ MWhS_crt          # $/MWh
else:
    BatS_DpE =0

#----------------------Share Senarios (All Existing Resources)---------------
#-----------Calculating gas values when gas ONLY additional resource------
# 5 Combine Cycle
# initialization for the 'for loops'
CC_DpMWh = 0   # from source excel (fromula) = $/MWh
CCmonth_MWh= 0 # MWh for each month -LIST
CC_monthD= 0   # $ for each month 
CC_D=0         # Total $ CC
SumMWhCC= 0    # Total MWh CC
# formula =(B29*$F$9)/1000+$F$7
F9= sheet.iloc[7,5]  # from 'EconomicAnalysisM3.xlsx'
F7= sheet.iloc[5,5]  # from 'EconomicAnalysisM3.xlsx'
# Total [MWH] & Total [$] CC
for i in range (27,39):  
    CC_DpMWh= (sheet.iloc[i,1]* F9/1000) + F7
    CCmonth_MWh= MWhCC[i-27]  
    CC_monthD = CC_DpMWh * CCmonth_MWh 
    SumMWhCC+= CCmonth_MWh    # Total [MWH]
    CC_D+=CC_monthD           # Total [$]
# Total $/MWh for all months
if SumMWhCC > 0:
    CC_DpE=CC_D/SumMWhCC
else:
     CC_DpE=0  
Sen5 = [SumMWhCC,CC_D,0,0,CC_DpE]

# 6 Combination Turbine 
# initialization for the 'for loops'
CT_DpMWh = 0   # from source excel (fromula) = $/MWh
CTmonth_MWh= 0 # MWh for each month
CT_monthD= 0   # $ for each month ($)
CT_D=0         # Total $ CT
SumMWhCT= 0    # Total MWh CT
# formula =(B29*$G$9)/1000+$G$7
G9= sheet.iloc[7,6]  # from 'EconomicAnalysisM3.xlsx'
G7= sheet.iloc[5,6]  # from 'EconomicAnalysisM3.xlsx'
# Total [MWH] & Total [$] CT
for i in range (27,39):  
    CT_DpMWh= (sheet.iloc[i,1]* G9/1000) + G7
    CTmonth_MWh= MWhCT[i-27]  
    CT_monthD = CT_DpMWh * CTmonth_MWh
    SumMWhCT+= CTmonth_MWh    # Total [MWH]
    CT_D+=CT_monthD           # Total [$]
if SumMWhCT > 0:
    CT_DpE=CT_D/SumMWhCT
else:
     CT_DpE=0 
Sen6 = [SumMWhCT,CT_D,0,0,CT_DpE]
      
# 7 Reciprocating IC
# initialization for the 'for loops'
RIC_DpMWh = 0   # from source excel (fromula) = $/MWh
RICmonth_MWh= 0 # MWh for each month
RIC_monthD= 0   # $ for each month ($)
RIC_D=0           # Total $ RIC
SumMWhRIC= 0    # Total MWh RIC
# formula =(B29*$H$9)/1000+$H$7
H9= sheet.iloc[7,7]  # from 'EconomicAnalysisM3.xlsx'
H7= sheet.iloc[5,7]  # from 'EconomicAnalysisM3.xlsx'
# Total [MWH] & Total [$] RIC
for i in range (27,39):  
    RIC_DpMWh= (sheet.iloc[i,1]* H9/1000) + H7
    RICmonth_MWh= MWhRIC[i-27]  
    RIC_monthD = RIC_DpMWh * RICmonth_MWh
    SumMWhRIC+= RICmonth_MWh    # Total [MWH]
    RIC_D+=RIC_monthD           # Total [$]
if SumMWhRIC > 0:
    RIC_DpE=RIC_D/SumMWhRIC
else:
     RIC_DpE=0 
Sen7 = [SumMWhRIC,RIC_D,0,0,RIC_DpE]
 
# 8 Aeroderivative
# initialization for the 'for loops'
A_DpMWh = 0   # from source excel (fromula) = $/MWh
Amonth_MWh= 0 # MWh for each month
A_monthD= 0   # $ for each month ($)
A_D=0         # Total $ A
SumMWhA= 0    # Total MWh A
# formula =(B29*$I$9)/1000+$I$7 
I9= sheet.iloc[7,8]  # from 'EconomicAnalysisM3.xlsx'
I7= sheet.iloc[5,8]  # from 'EconomicAnalysisM3.xlsx'
# Total [MWH] & Total [$] A
for i in range (27,39):  
    A_DpMWh= (sheet.iloc[i,1]* I9/1000) + I7
    Amonth_MWh= MWhA[i-27]  
    A_monthD = A_DpMWh * Amonth_MWh
    SumMWhA+= Amonth_MWh    # Total [MWH]
    A_D+= A_monthD          # Total [$]
if SumMWhA > 0:
    A_DpE=A_D/SumMWhA
else:
     A_DpE=0 
Sen8 = [SumMWhA,A_D,0,0,A_DpE]

#-----------Calculating gas values when gas COMBINED additional resource------
# 5b. Combine Cycle
# initialization for the 'for loops'
Cc_DpMWh = 0   # from source excel (fromula) = $/MWh
Ccmonth_MWh= 0 # MWh for each month -LIST
Cc_monthD= 0   # $ for each month 
Cc_D=0         # Total $ Cc
SumMWhCc= 0    # Total MWh Cc
# formula =(B29*$F$9)/1000+$F$7
F9= sheet.iloc[7,5]  # from 'EconomicAnalysisM3.xlsx'
F7= sheet.iloc[5,5]  # from 'EconomicAnalysisM3.xlsx'
# Total [MWH] & Total [$] Cc
for i in range (27,39):  
    Cc_DpMWh= (sheet.iloc[i,1]* F9/1000) + F7
    Ccmonth_MWh= MWhCc[i-27]  
    Cc_monthD = Cc_DpMWh * Ccmonth_MWh 
    SumMWhCc+= Ccmonth_MWh    # Total [MWH]
    Cc_D+=Cc_monthD           # Total [$]
# Total $/MWh for all months
if SumMWhCc > 0:
    Cc_DpE=Cc_D/SumMWhCc
else:
     Cc_DpE=0  

# 6b. Combination Turbine 
# initialization for the 'for loops'
Ct_DpMWh = 0   # from source excel (fromula) = $/MWh
Ctmonth_MWh= 0 # MWh for each month
Ct_monthD= 0   # $ for each month ($)
Ct_D=0         # Total $ Ct
SumMWhCt= 0    # Total MWh Ct
# formula =(B29*$G$9)/1000+$G$7
G9= sheet.iloc[7,6]  # from 'EconomicAnalysisM3.xlsx'
G7= sheet.iloc[5,6]  # from 'EconomicAnalysisM3.xlsx'
# Total [MWH] & Total [$] Ct
for i in range (27,39):  
    Ct_DpMWh= (sheet.iloc[i,1]* G9/1000) + G7
    Ctmonth_MWh= MWhCt[i-27]  
    Ct_monthD = Ct_DpMWh * Ctmonth_MWh
    SumMWhCt+= Ctmonth_MWh    # Total [MWH]
    Ct_D+=Ct_monthD           # Total [$]
if SumMWhCt > 0:
    Ct_DpE=Ct_D/SumMWhCt
else:
     Ct_DpE=0  
      
# 7b. Reciprocating ic
# initialization for the 'for loops'
Ric_DpMWh = 0   # from source excel (fromula) = $/MWh
Ricmonth_MWh= 0 # MWh for each month
Ric_monthD= 0   # $ for each month ($)
Ric_D=0           # Total $ Ric
SumMWhRic= 0    # Total MWh Ric
# formula =(B29*$H$9)/1000+$H$7
H9= sheet.iloc[7,7]  # from 'EconomicAnalysisM3.xlsx'
H7= sheet.iloc[5,7]  # from 'EconomicAnalysisM3.xlsx'
# Total [MWH] & Total [$] Ric
for i in range (27,39):  
    Ric_DpMWh= (sheet.iloc[i,1]* H9/1000) + H7
    Ricmonth_MWh= MWhRic[i-27]  
    Ric_monthD = Ric_DpMWh * Ricmonth_MWh
    SumMWhRic+= Ricmonth_MWh    # Total [MWH]
    Ric_D+=Ric_monthD           # Total [$]
if SumMWhRic > 0:
    Ric_DpE=Ric_D/SumMWhRic
else:
     Ric_DpE=0  
 
# 8b. Aeroderivative (_a)
# initialization for the 'for loops'
_a_DpMWh = 0   # from source excel (fromula) = $/MWh
_amonth_MWh= 0 # MWh for each month
_a_monthD= 0   # $ for each month ($)
_a_D=0         # Total $ _a
SumMWh_a= 0    # Total MWh _a
# formula =(B29*$I$9)/1000+$I$7 
I9= sheet.iloc[7,8]  # from 'EconomicAnalysisM3.xlsx'
I7= sheet.iloc[5,8]  # from 'EconomicAnalysisM3.xlsx'
# Total [MWH] & Total [$] _a
for i in range (27,39):  
    _a_DpMWh= (sheet.iloc[i,1]* I9/1000) + I7
    _amonth_MWh= MWh_a[i-27]  
    _a_monthD = _a_DpMWh * _amonth_MWh
    SumMWh_a += _amonth_MWh   # Total [MWH]
    _a_D+= _a_monthD          # Total [$]
if SumMWh_a > 0:
    _a_DpE= _a_D/SumMWh_a
else:
     _a_DpE=0 
     
#-----------------Senarior differer [Based on 'Existing Resources']---------
# From Existing Solar
# 13 Battery with Wind
WBE= MWhW+MWhW_crt
WBD=Wind_D+BattW_D
WBCrE=MWhW_Xcrt
WBCrD=windCurt_D
WBDpE =(WBD+ WBCrD) /WBE
if WBE >0:
    WBDpE =(WBD+ WBCrD) /WBE
else:
   WBDpE =0
Sen13 = [WBE,WBD,WBCrE,WBCrD,WBDpE] 

# 15 Wind & Cc            
WCcE =MWhW+SumMWhCc           # Energy Resource
WCcD=Wind_D+Cc_D              # $ Resource
WCcCrE=MWhW_Tcrt              # Energy Curt
WCcCrD = WindCurt_D           # $ Curt
if MWhW > 0:
    WCcDpE=(WCcD+WCcCrD)/WCcE # $/MWh Resource
else:
    WCcDpE= 0
Sen15 = [WCcE,WCcD,WCcCrE,WCcCrD,WCcDpE]
     
# 16 Wind & Ct            
WCtE=MWhW+SumMWhCt            # Energy Resource
WCtD=Wind_D+Ct_D              # $ Resource
WCtCrE=MWhW_Tcrt              # Energy Curt
WCtCrD=WindCurt_D             # $ Curt    
if MWhW > 0:
    WCtDpE=(WCtD+WCtCrD)/WCtE # $/MWh Resource
else:
    WCtDpE=0
Sen16 = [WCtE,WCtD,WCtCrE,WCtCrD,WCtDpE]
      
# 17 Wind & Ric  
WRicE=MWhW+SumMWhRic              # Energy Resource
WRicD=Wind_D+Ric_D                # $ Resource
WRicCrE=MWhW_Tcrt                 # Energy Curt
WRicCrD=WindCurt_D                # $ Curt
if MWhW > 0:
    WRicDpE=(WRicD+WRicCrD)/WRicE # $/MWh Resource
else:
    WRicDpE=0
Sen17 = [WRicE,WRicD,WRicCrE,WRicCrD,WRicDpE]
      
# 18 Wind & _a 
W_aE=MWhW+SumMWh_a                # Energy Resource
W_aD=Wind_D+ _a_D                 # $ Resource
W_aCrE= MWhW_Tcrt                 # Energy Curt
W_aCrD= WindCurt_D                # $ Curt
if MWhW > 0:
    W_aDpE= (W_aD+W_aCrD)/W_aE    # $/MWh Resource    
else:
  W_aDpE=0  
Sen18 = [W_aE,W_aD,W_aCrE,W_aCrD,W_aDpE]
#------------------------------------------------------------------------
# From Existing Wind
# 24 Battery with Solar
SBE= MWhS+MWhS_crt
SBD=Solar_D+BattS_D
SBCrE=MWhS_Xcrt
SBCrD=solarCurt_D
if WBE >0:
    SBDpE =(SBD+ SBCrD) /SBE
else:
    SBDpE =0
Sen24 = [SBE,SBD,SBCrE,SBCrD,SBDpE]

# 15 Solar & Cc            
SCcE =MWhS+SumMWhCc           # Energy Resource
SCcD=Solar_D+ Cc_D            # $ Resource
SCcCrE=MWhS_Tcrt              # Energy Curt         
SCcCrD = SolarCurt_D          # $ Curt  
if MWhW > 0:
    SCcDpE=(SCcD+SCcCrD)/SCcE # $/MWh Resource
else:
    SCcDpE= 0
Sen25 = [SCcE,SCcD,SCcCrE,SCcCrD,SCcDpE]
     
# 16 Solar & Ct            
SCtE=MWhS+SumMWhCt           # Energy Resource
SCtD=Solar_D+Ct_D            # $ Resource
SCtCrE=MWhS_Tcrt             # Energy Curt 
SCtCrD=SolarCurt_D           # $ Curt  
if MWhW > 0:
    SCtDpE=(SCtD+SCtCrD)/SCtE    # $/MWh Resource 
else:
    SCtDpE=0
Sen26 = [SCtE,SCtD,SCtCrE,SCtCrD,SCtDpE]
      
# 17 Solar & Ric  
SRicE=MWhS+SumMWhRic               # Energy Resource
SRicD=Solar_D+Ric_D                # $ Resource
SRicCrE=MWhS_Tcrt                  # Energy Curt 
SRicCrD=SolarCurt_D                # $ Curt
if MWhW > 0:
    SRicDpE=(SRicD+SRicCrD)/SRicE  # $/MWh Resource
else:
     SRicDpE=0
Sen27 = [SRicE,SRicD,SRicCrE,SRicCrD,SRicDpE]
      
# 18 Solar & _a 
S_aE=MWhS+SumMWh_a               # Energy Resource
S_aD=Solar_D+ _a_D               # $ Resource
S_aCrE= MWhS_Tcrt                # Energy Curt
S_aCrD= SolarCurt_D              # $ Curt
if MWhW > 0:
     S_aDpE= (S_aD+S_aCrD)/S_aE  # $/MWh Resource   
else:
  S_aDpE=0  
Sen28 = [S_aE,S_aD,S_aCrE,S_aCrD,S_aDpE]

#-----------------Shorted by the 'Switch Case' Mechanism----------------
#+++++++++++++++Switch Case (assigh values to resources)+++++++++++

# lists used by the Switch/Case Mechanism
KeyList=[1,2,5,6,7,8,13,15,16,17,18,24,25,26,26,27,28] # 'Key' of all senarios
SenTrue = [] # empty list (to receive [value]/[00000] values from TrueSen fn)
Zr = [0,0,0,0,0] # [00000] replaces Sen# in SenTrue list when senarion is F

# loop for determining if Sen# are T/F
    # & (T= last T), (F = 1st F)
    # | (T = 1st T), (F = Last F)
for i in KeyList:
    # TrueSen fn places the T_values at SenTrue [], or [00000]
    def TrueSen (i):    # each loop (i) determins T or F values
        switcher= { 
                1:((Sen1 and MWhW and Sen1) or Zr),
                2:((Sen2 and MWhS and Sen2) or Zr),
                5:((Sen5 and SumMWhCC and Sen5) or Zr),
                6:((Sen6 and SumMWhCT and Sen6) or Zr),
                7:((Sen7 and SumMWhRIC and Sen7) or Zr),
                8:((Sen8 and SumMWhA and Sen8) or Zr),
                13:((Sen13 and MWhW and MWhW_crt and Sen13) or Zr),
                15:((Sen15 and MWhW and SumMWhCC and Sen15) or Zr),
                16:((Sen16 and MWhW and SumMWhCT and Sen16) or Zr),
                17:((Sen17 and MWhW and SumMWhRIC and Sen17) or Zr),
                18:((Sen18 and MWhW and SumMWhA and Sen18) or Zr),
                24:((Sen24 and MWhS and MWhS_crt and Sen24) or Zr),
                25:((Sen25 and MWhS and SumMWhCC and Sen25) or Zr),
                26:((Sen26 and MWhS and SumMWhCT and Sen26) or Zr),
                27:((Sen27 and MWhS and SumMWhRIC and Sen27) or Zr),
                28:((Sen28 and MWhS and SumMWhA and Sen28) or Zr)
                }   
        return switcher.get(i,'invalid resource')
    TrueSen (i)                  # passes T Sen to SenList
    SenTrue.append(TrueSen (i))  # fill in SenTrue list

# assigning values to the Sen# [] from TrueSen
Sen1 = SenTrue[0]
Sen2 = SenTrue[1]
Sen5 = SenTrue[2]
Sen6 = SenTrue[3]
Sen7 = SenTrue[4]
Sen8 = SenTrue[5]
Sen13 = SenTrue[6]
Sen15 = SenTrue[7]
Sen16 = SenTrue[8]
Sen17 = SenTrue[9]
Sen18 = SenTrue[10]
Sen24 = SenTrue[11]
Sen25 = SenTrue[12]
Sen26 = SenTrue[13]
Sen27 = SenTrue[14]
Sen28 = SenTrue[15]

#==================Plots all values regardless 'Existing Resource'===================
#---------------matplotlib----------------
import matplotlib.pyplot as plt  

# plot 'Energy by Resources'
plt.figure (1)

names = ['W','Sol','CC','CT','RIC','Aro',
      'W\n + \nBat','W\n + \nCC','W\n + \nCT',
      'W\n + \nRIC','W\n + \nAro',
      'Sol\n + \nBat','Sol\n + \nCC','Sol\n + \nCT',
      'Sol\n + \nRIC','Sol\n + \nAro',
      'W\nCurt','w\nCurt','Sol\nCurt','sol\nCurt']
values= [Sen1[0],Sen2[0],Sen5[0],Sen6[0],Sen7[0],Sen8[0],
         Sen13[0],Sen15[0],Sen16[0],Sen17[0],Sen18[0],
         Sen24[0],Sen25[0],Sen26[0],Sen27[0], Sen28[0],
         Sen1[2],Sen13[2],Sen2[2],Sen24[2]]

plt.title('Energy by Resources')
plt.ylabel('Energy [MWh]')
plt.xlabel('Resources')
plt.bar(names, values, color='b')

plt.show()
#----------------------------------------------------------------------
# plot 'Cost by Resources'
plt.figure (2)

names = ['W','Sol','CC','CT','RIC','Aro',
      'W\n + \nBat','W\n + \nCC','W\n + \nCT',
      'W\n + \nRIC','W\n + \nAro',
      'Sol\n + \nBat','Sol\n + \nCC','Sol\n + \nCT',
      'Sol\n + \nRIC','Sol\n + \nAro',
      'W\nCurt','w\nCurt','Sol\nCurt','sol\nCurt']
values= [Sen1[1],Sen2[1],Sen5[1],Sen6[1],Sen7[1],Sen8[1],
         Sen13[1],Sen15[1],Sen16[1],Sen17[1],Sen18[1],
         Sen24[1],Sen25[1],Sen26[1],Sen27[1], Sen28[1],
         Sen1[3],Sen13[3],Sen2[3],Sen24[3]]

plt.title('Cost by Resources')
plt.ylabel('dollars [$]')
plt.xlabel('Resources')
plt.bar(names, values, color='r')

plt.show()

#----------------------------------------------------------------------
# plot 'Energy Price per Resource'
plt.figure (3)

names = ['W','Sol','CC','CT','RIC','Aro',
      'W\n + \nBat','W\n + \nCC','W\n + \nCT',
      'W\n + \nRIC','W\n + \nAro',
      'Sol\n + \nBat','Sol\n + \nCC','Sol\n + \nCT',
      'Sol\n + \nRIC','Sol\n + \nAro']
values=   [Sen1[4],Sen2[4],Sen5[4],Sen6[4],Sen7[4],Sen8[4],
           Sen13[4],Sen15[4],Sen16[4],Sen17[4],Sen18[4],
           Sen24[4],Sen25[4],Sen26[4],Sen27[4], Sen28[4]]

plt.title('Energy Price per Resource')
plt.ylabel('dollars per Energy [$/MWh]')
plt.xlabel('Resources')
plt.bar(names, values, color='c')

plt.show()



