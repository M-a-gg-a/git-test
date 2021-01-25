# -*- coding: utf-8 -*-
"""
Created on Wed Jan 13 12:50:29 2021

@author: magal
Jessie was here
"""

import numpy as np
import pandas as pd

#################################################################

#_________Snip1 Variables from Module1 & Module 2________________
#---------------from module 1-------------------------------------------
# Existing resource (Wind T or F)
ExistingWind = 1
#------------------------------All resources initialized to 0------------
#-------------------'Existing Wind' Resource----------------------------

MWhS = 0        # MWh in Year for Solar
# curtailment additions [MWh] Solar in a Year
MWhS_crt = 0    # used by battery 'as Energy' when available - Solar
MWhS_Xcrt = 0   # 'Extra curtailment' not stored even when battery in use- Solar
MWhS_Tcrt = 0   # Total Curtailment - When NOT battery in use - Solar

#-------------------'Existing Solar' Resource----------------------------

MWhW = 0        # MWh in Year for Wind
# curtailment additions [MWh] Wind in a Year
MWhW_crt = 0    # used by battery 'as Energy' when available - Wind
MWhW_Xcrt = 0   # 'Extra curtailment' not stored even when battery in use- Wind
MWhW_Tcrt = 0   # Total Curtailment - When NOT battery in use - Wind

#-------------------'Any Existing' Resource----------------------------
# CAPITAL (CC, CT, RIC, A) are used when ONLY additional resource
# CC MWh produced each month
JrCC = 0; FbCC= 0; MrCC = 0; ApCC =0; MyCC = 0;JnCC = 0;
JlCC = 0; AgCC=0; SpCC = 0;OcCC = 0; NvCC = 0; DcCC = 0
# CT MWh produced each month
JrCT = 0; FbCT= 0; MrCT = 0; ApCT =0; MyCT = 0;JnCT = 0;
JlCT = 0; AgCT=0; SpCT = 0;OcCT = 0; NvCT = 0; DcCT = 0
# RIC MWh produced each month
JrRIC = 0; FbRIC= 0; MrRIC = 0; ApRIC =0; MyRIC = 0;JnRIC = 0;
JlRIC = 0; AgRIC=0; SpRIC = 0;OcRIC = 0; NvRIC = 0; DcRIC = 0
# A MWh produced each month
JrA = 0; FbA= 0; MrA = 0; ApA = 0; MyA = 0;JnA = 0;
JlA = 0; AgA= 0; SpA = 0;OcA = 00; NvA = 0; DcA = 0

# CApital_small (Cc, Ct, Ric, _a) are used when COMBINED additional resource
# Cc MWh produced each month
JrCc = 0; FbCc= 0; MrCc = 0; ApCc =0; MyCc = 0;JnCc = 0;
JlCc = 0; AgCc=0; SpCc = 0;OcCc = 0; NvCc = 0; DcCc = 0
# Ct MWh produced each month
JrCt = 0; FbCt= 0; MrCt = 0; ApCt =0; MyCt = 0;JnCt = 0;
JlCt = 0; AgCt=0; SpCt = 0;OcCt = 0; NvCt = 0; DcCt = 0
# Ric MWh produced each month
JrRic = 0; FbRic= 0; MrRic = 0; ApRic =0; MyRic = 0;JnRic = 0;
JlRic = 0; AgRic=0; SpRic = 0;OcRic = 0; NvRic = 0; DcRic = 0
# _a MWh produced each month
Jr_a = 0; Fb_a= 0; Mr_a = 0; Ap_a = 0; My_a = 0;Jn_a = 0;
Jl_a = 0; Ag_a= 0; Sp_a = 0;Oc_a = 0; Nv_a = 0; Dc_a = 0
##################################################################

# Installed capacity using the year 2020 as based
# reads a csv file of an array of 8760 rows with existing capacity in MWh
# variable Ecapacity stores array
csv = '../csvs/Ecapacity.csv'
Ecapacity = pd.read_csv(csv, usecols=[0,1,2,3,4,5,6,7,8,9,10], index_col = ['year'])
Ecapacity.head(4)           # stores information

# reads a csv file of an array of 8760 rows,
# with fake capacity factor CF by renewable resource
# Variable FAKE8760 stores array
csv = '../csvs/FAKE8760.csv'
FAKE8760 = pd.read_csv(csv, usecols=[0,1,2,3,4,5,6,7,8], index_col = ['year'])
FAKE8760.head(4)

# Conditions: temporary variables to run simulation
# these values should come from the user interface part
utilization = .8
factor = 1
years = 2

# Simulation of wind forecast for "one year" or an array of 8760
wind_forecast = pd.DataFrame(Ecapacity.nameplate_MW * FAKE8760.wind_CF * utilization * factor, columns=['wind'])
wind_forecast.head(10)

# Simulation of solar forecast for "one year" or an array of 8760
solar_forecast = pd.DataFrame(Ecapacity.nameplate_MW * FAKE8760.solar_CF * utilization * factor, columns=['solar'])
solar_forecast.head(8)

# Simulation of batteries forecast for "one year" or an array of 8760
batteries_forecast = pd.DataFrame(Ecapacity.nameplate_MW * FAKE8760.batteries_CF * utilization * factor, columns=['batt'])
batteries_forecast.head(8)

# Simulation of natural gas forecast for "one year" or an array of 8760
natgas_forecast = pd.DataFrame(Ecapacity.nameplate_MW * FAKE8760.natgas_CF * utilization * factor, columns=['natgas'])
natgas_forecast.head(8)
