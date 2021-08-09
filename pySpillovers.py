# ==============================
# IMPORT PACKAGE
# ==============================
import pandas as pd, numpy as np
import functions as f

# ==============================
# USER INPUT
# ==============================
df = pd.read_excel('_userInput.xlsx').set_index("SETTINGS")
dateFrom = df.loc['dateFrom','VALUE']
dateTo = df.loc['dateTo','VALUE']
outputMode = df.loc['outputMode','VALUE']
marketDaysMode = df.loc['marketDaysMode','VALUE']
manualMarketDays = df.loc['manualMarketDays','VALUE']
dataYearEnd = df.loc['dataYearEnd','VALUE']
marketDaysYearEnd = df.loc['marketDaysYearEnd','VALUE']

# ==============================
# IMPORT DATA
# ==============================
# Import sectors
sectors = np.genfromtxt('_sectorsList.csv',delimiter=',',dtype="str")

# Import sectorsData, reIndex sectorsData with Date
sectorsData = {}
marketDays = {}
for sector in sectors:
	# sectorsData
	df = pd.read_csv('DailyPrices\\'+sector+'.JK_D.csv') \
		.assign(Date=lambda x: pd.to_datetime(x.Date, format="%d-%m-%Y")) \
		.set_index("Date")
	sectorsData[sector] = df

	# marketDays
	if marketDaysMode == "Manual":
		marketDays = manualMarketDays
	else:
		marketDays[sector] = f.calcMarketDays(sectorsData[sector],marketDaysYearEnd)

	# Filter sectorsData between DateTo and DateFrom
	sectorsData[sector] = sectorsData[sector].loc[dateFrom:dateTo]

# ==============================
# ==DIEBOLD 2012 SPILLOVERS INDEX==
# ==============================
# ==============================
# DATA PREPARATION BASED ON OUTPUTMODE
# ==============================
lnvariance = f.calcLnvariance(sectorsData)

if outputMode == "Volatility Diebold":
	volatility = f.calcVolatilityDiebold(lnvariance,marketDays)
elif outputMode == "Volatility Aslam":
	volatility = f.calcVolatilityAslam(lnvariance,marketDays)

# ==============================
# STATISTIC OF DATA
# ==============================
setStats = f.calcSetStats(volatility)

# ==============================
# Spillovers Table
# ==============================
spilloversTable, lag_order, forecast_horizon = f.calcSpilloversTable(volatility)

# ==============================
# TOTAL, DIRECTIONAL, NET SPILLOVERS
# ==============================

# ==============================
# SENSITIVITY ANALYSIS
# ==============================

# ==============================
# OUTPUT
# ==============================
# Volatility or Return Table and Graph

# Data Spillover Table

# Total Data Spillover Table and Graph

# Directional Volatility Spillovers FROM, TO, and NET Each, NET Pairwaise Table and Graph

print("End of Analysis")