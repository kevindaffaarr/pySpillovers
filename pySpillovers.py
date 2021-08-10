# ==============================
# IMPORT PACKAGE
# ==============================
import pandas as pd, numpy as np
import math
import functions as f
import warnings
warnings.filterwarnings("ignore")

# ===================================================================================================
# ============================================IMPORT DATA============================================
# ===================================================================================================
def getImportData(marketDaysMode=None,marketDaysYearEnd=250,manualMarketDays=250):
	marketDaysMode = None if marketDaysMode is None else marketDaysMode
	marketDaysYearEnd = 250 if marketDaysYearEnd is None else marketDaysYearEnd
	manualMarketDays = 250 if manualMarketDays is None else manualMarketDays

	# Import sectors
	sectors = np.genfromtxt('_sectorsList.csv',delimiter=',',dtype="str")

	# Import sectorsData, reIndex sectorsData with Date
	rawSectorsData = {}
	marketDays = {}
	for sector in sectors:
		# rawSectorsData
		df = pd.read_csv('DailyPrices\\'+sector+'.JK_D.csv') \
			.assign(Date=lambda x: pd.to_datetime(x.Date, format="%d-%m-%Y")) \
			.set_index("Date")
		rawSectorsData[sector] = df

		# marketDays
		if marketDaysMode == "Manual":
			marketDays = manualMarketDays
		else:
			marketDays[sector] = f.calcMarketDays(rawSectorsData[sector],marketDaysYearEnd)
	return rawSectorsData, marketDays, sectors

# ===================================================================================================
# ============Average and Dynamic Spillovers With Constant Lag Order and Forecast Horizon============
# ===================================================================================================
def getAvgSpillovers(lag_order=None,forecast_horizon=None,output=None):
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
	rollingWindow = df.loc['rollingWindow','VALUE']
	
	lag_order = df.loc['lag_order','VALUE'] if lag_order is None else lag_order
	forecast_horizon = df.loc['forecast_horizon','VALUE'] if forecast_horizon is None else forecast_horizon

	lag_order = None if lag_order =='Auto' else lag_order
	forecast_horizon = None if forecast_horizon =='Auto' else forecast_horizon
	rollingWindow = None if rollingWindow =='Auto' else rollingWindow

	# ==============================
	# IMPORT DATA
	# ==============================
	rawSectorsData, marketDays, sectors = getImportData(marketDaysMode,marketDaysYearEnd,manualMarketDays)
	sectorsData = {}
	for sector in sectors:
		sectorsData[sector] = rawSectorsData[sector].loc[dateFrom:dateTo]

	# ==============================
	# DATA PREPARATION BASED ON OUTPUTMODE
	# ==============================
	lnvariance = f.calcLnvariance(sectorsData)

	if outputMode == "Volatility Diebold":
		volatility = f.calcVolatilityDiebold(lnvariance.copy(),marketDays.copy())
	elif outputMode == "Volatility Aslam":
		volatility = f.calcVolatilityAslam(lnvariance.copy(),marketDays.copy())

	# ==============================
	# STATISTIC OF DATA
	# ==============================
	setStats = f.calcSetStats(volatility)

	# ==============================
	# Spillovers Table
	# ==============================
	spilloversTable, lag_order, forecast_horizon = f.calcAvgSpilloversTable(volatility,forecast_horizon,lag_order)

	# ==============================
	# OUTPUT
	# ==============================
	# StatsModel

	# Volatility or Return Table and Graph

	# Data Spillover Table

	# END OF AVERAGE VOLATILITY SPILLOVER===========================================================================
	return spilloversTable, setStats, volatility, lnvariance, lag_order, forecast_horizon

def getRollingSpillovers(lag_order=None,forecast_horizon=None,output=None):
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
	rollingWindow = df.loc['rollingWindow','VALUE']
	
	lag_order = df.loc['lag_order','VALUE'] if lag_order is None else lag_order
	forecast_horizon = df.loc['forecast_horizon','VALUE'] if forecast_horizon is None else forecast_horizon

	lag_order = None if lag_order =='Auto' else lag_order
	forecast_horizon = None if forecast_horizon =='Auto' else forecast_horizon
	rollingWindow = None if rollingWindow =='Auto' else rollingWindow
	# ==============================
	# IMPORT DATA
	# ==============================
	rawSectorsData, marketDays, sectors = getImportData(marketDaysMode,marketDaysYearEnd,manualMarketDays)
	sectorsData = {}
	for sector in sectors:
		# Filter sectorsData between DateTo and DateFrom
		sectorsData[sector] = f.getWithRollingWindow(rawSectorsData[sector],dateFrom,dateTo,rollingWindow)

	# ==============================
	# DATA PREPARATION BASED ON OUTPUTMODE
	# ==============================
	lnvariance = f.calcLnvariance(sectorsData)

	if outputMode == "Volatility Diebold":
		volatility = f.calcVolatilityDiebold(lnvariance.copy(),marketDays.copy())
	elif outputMode == "Volatility Aslam":
		volatility = f.calcVolatilityAslam(lnvariance.copy(),marketDays.copy())

	# ==============================
	# TOTAL, DIRECTIONAL, NET ROLLING SPILLOVERS
	# ==============================
	# rolling Spillovers: 
	# ['total']
	# ['to'][sector]
	# ['from'][sector]
	# ['net'][sector]
	# ['pairwaise'][sector]
	rollingSpillovers = f.calcRollingSpillovers(volatility, forecast_horizon, lag_order,rollingWindow)

	# ==============================
	# OUTPUT
	# ==============================
	# Total, FROM, TO, and NET Each, NET Pairwaise Data Spillover Table and Graph

	return rollingSpillovers, volatility, lnvariance, lag_order, forecast_horizon

# ==============================
# SENSITIVITY ANALYSIS:
# Average and Dynamic Spillovers With Variant Lag Order
# ==============================
def rollingSensitivityAnalysis(variantParam,start,end,lag_order,forecast_horizon,sectors):
	# ==============================
	# ARRAY PREPARATIONS
	# ==============================
	newRollingSpillovers = {}
	newRollingSpillovers['total'] = pd.DataFrame()
	newRollingSpillovers['to'] = {}
	newRollingSpillovers['from'] = {}
	newRollingSpillovers['net'] = {}
	newRollingSpillovers['pairwaise'] = {}

	for sector in sectors:
		newRollingSpillovers['to'][sector] = pd.DataFrame()
		newRollingSpillovers['from'][sector] = pd.DataFrame()
		newRollingSpillovers['net'][sector] = pd.DataFrame()
		newRollingSpillovers['pairwaise'][sector] = {}
		for sectorFrom in sectors:
			newRollingSpillovers['pairwaise'][sector][sectorFrom] = pd.DataFrame()

	sensitivityRange = {}
	sensitivityRange['total'] = pd.DataFrame()
	sensitivityRange['to'] = {}
	sensitivityRange['from'] = {}
	sensitivityRange['net'] = {}
	sensitivityRange['pairwaise'] = {}
	for sector in sectors:
		sensitivityRange['to'][sector] = pd.DataFrame()
		sensitivityRange['from'][sector] = pd.DataFrame()
		sensitivityRange['net'][sector] = pd.DataFrame()
		sensitivityRange['pairwaise'][sector] = {}
		for sectorFrom in sectors:
			sensitivityRange['pairwaise'][sector][sectorFrom] = pd.DataFrame()

	# ==============================
	# ITERATE FOR EACH VARIANTPARAM
	# ==============================
	for i in range(start,end+1,1):
		print('sensitivityAnalysis #'+str(i))
		if variantParam == 'lag_order':
			rollingSpillovers, temp1, temp2, temp3, temp4 = getRollingSpillovers(lag_order=i,forecast_horizon=forecast_horizon,output='sensitivity_lag_order_'+str(i))
			del temp1, temp2, temp3, temp4
		elif variantParam == 'forecast_horizon':
			rollingSpillovers, temp1, temp2, temp3, temp4 = getRollingSpillovers(lag_order=lag_order,forecast_horizon=i,output='sensitivity_forecast_horizon_'+str(i))
			del temp1, temp2, temp3, temp4
		
		newRollingSpillovers['total'][i] = rollingSpillovers['total']
		for sector in sectors:
			newRollingSpillovers['to'][sector][i] = rollingSpillovers['to'][sector]
			newRollingSpillovers['from'][sector][i] = rollingSpillovers['from'][sector]
			newRollingSpillovers['net'][sector][i] = rollingSpillovers['net'][sector]
			for sectorFrom in sectors:
				newRollingSpillovers['pairwaise'][sector][sectorFrom][i] = rollingSpillovers['pairwaise'][sector][sectorFrom]
	
	sensitivityRange['total'] = pd.DataFrame({'min':newRollingSpillovers['total'].min(axis=1),'median':newRollingSpillovers['total'].median(axis=1),'max':newRollingSpillovers['total'].max(axis=1)})
	for sector in sectors:
		sensitivityRange['to'][sector] = pd.DataFrame({'min':newRollingSpillovers['to'][sector].min(axis=1),'median':newRollingSpillovers['to'][sector].median(axis=1),'max':newRollingSpillovers['to'][sector].max(axis=1)})
		sensitivityRange['from'][sector] = pd.DataFrame({'min':newRollingSpillovers['from'][sector].min(axis=1),'median':newRollingSpillovers['from'][sector].median(axis=1),'max':newRollingSpillovers['from'][sector].max(axis=1)})
		sensitivityRange['net'][sector] = pd.DataFrame({'min':newRollingSpillovers['net'][sector].min(axis=1),'median':newRollingSpillovers['net'][sector].median(axis=1),'max':newRollingSpillovers['net'][sector].max(axis=1)})
		for sectorFrom in sectors:
			sensitivityRange['pairwaise'][sector][sectorFrom] = pd.DataFrame({'min':newRollingSpillovers['pairwaise'][sector][sectorFrom].min(axis=1),'median':newRollingSpillovers['pairwaise'][sector][sectorFrom].median(axis=1),'max':newRollingSpillovers['pairwaise'][sector][sectorFrom].max(axis=1)})
	
	# ==============================
	# OUTPUT
	# ==============================
	# Total, FROM, TO, and NET Each, NET Pairwaise Data Spillover Table and Graph
	
	return sensitivityRange
# ==================================================================================================
# ===============================================MAIN===============================================
# ==================================================================================================
# AVERAGE
print('Starting The Machine...')
print('Calc Average Spillovers...')
spilloversTable, setStats, volatility, lnvariance, lag_order, forecast_horizon = getAvgSpillovers()
sectors = volatility.columns
del spilloversTable, setStats, volatility, lnvariance
print('End of Calc Average Spillovers')

# ROLLING
print('Calc Rolling Spillovers')
rollingSpillovers, temp1, temp2, temp3, temp4 = getRollingSpillovers(lag_order,forecast_horizon)
del rollingSpillovers, temp1, temp2, temp3, temp4
print('End of Calc Rolling Spillovers')

# SENSITIVITY
print('Calc Sensitivity Analysis Spillovers: lag_order')
sensitivityRange = rollingSensitivityAnalysis('lag_order',1,lag_order*2,lag_order,forecast_horizon,sectors)

print('Calc Sensitivity Analysis Spillovers: forecast_horizon')
sensitivityRange = rollingSensitivityAnalysis('forecast_horizon',math.floor(0.5*forecast_horizon),math.ceil(1.5*forecast_horizon),lag_order,forecast_horizon,sectors)

print('End of Calc Analysis Spillovers')

print("End of Analysis")