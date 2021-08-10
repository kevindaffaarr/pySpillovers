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
	# setStats
	setStats.to_csv('output\setStats.csv')

	# Volatility Table
	volatility.to_csv('output\\volatility.csv')

	# Volatility Graph
	f.genStackedTimeSeriesChart(\
		df=volatility, \
		filename='Volatilities (Annualized Standard Deviations)', \
		xaxis_title = 'Date', \
		yaxis_title = '%' \
	)

	# Data Spillover Table
	filename = 'output\spilloversTable.csv'
	title = 'Spillover Table\n'
	title = title + 'lag_order,' + str(lag_order) + '\nforecast_horizon,' + str(forecast_horizon) + '\n'
	title = title + 'TO,FROM\n'
	with open(filename,'w') as out:
		out.write(title)
	spilloversTable.to_csv(filename,mode='a')

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
	# ['pairwise'][sectorTo][sectorFrom]
	rollingSpillovers = f.calcRollingSpillovers(volatility, forecast_horizon, lag_order,rollingWindow)

	# ==============================
	# OUTPUT
	# ==============================
	# Total, FROM, TO, and NET Each, NET Pairwise Data Spillover Table and Graph

	return rollingSpillovers, volatility, lnvariance, lag_order, forecast_horizon

# ==============================
# SENSITIVITY ANALYSIS:
# Average and Dynamic Spillovers With Variant Lag Order
# ==============================
def getRollingSensitivityAnalysis(variantParam,start,end,lag_order,forecast_horizon,sectors):
	# sensitivityRange['total']
	# sensitivityRange['to'][sector]
	# sensitivityRange['from'][sector]
	# sensitivityRange['net'][sector]
	# sensitivityRange['pairwiseTo'][sector][sectorFrom]
	# sensitivityRange['pairwiseNet'][sector][sectorFrom]

	# ==============================
	# ARRAY PREPARATIONS
	# ==============================
	newRollingSpillovers = {}
	newRollingSpillovers['total'] = pd.DataFrame()
	newRollingSpillovers['to'] = {}
	newRollingSpillovers['from'] = {}
	newRollingSpillovers['net'] = {}
	newRollingSpillovers['pairwiseTo'] = {}
	newRollingSpillovers['pairwiseNet'] = {}

	for sector in sectors:
		newRollingSpillovers['to'][sector] = pd.DataFrame()
		newRollingSpillovers['from'][sector] = pd.DataFrame()
		newRollingSpillovers['net'][sector] = pd.DataFrame()
		newRollingSpillovers['pairwiseTo'][sector] = {}
		newRollingSpillovers['pairwiseNet'][sector] = {}
		for sectorFrom in sectors:
			newRollingSpillovers['pairwiseTo'][sector][sectorFrom] = pd.DataFrame()
			newRollingSpillovers['pairwiseNet'][sector][sectorFrom] = pd.DataFrame()

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
				newRollingSpillovers['pairwiseTo'][sector][sectorFrom][i] = rollingSpillovers['pairwiseTo'][sector][sectorFrom]
				newRollingSpillovers['pairwiseNet'][sector][sectorFrom][i] = rollingSpillovers['pairwiseNet'][sector][sectorFrom]
	
	# ==============================
	# SENSITIVITY RANGE
	# ==============================
	sensitivityRange = f.calcRollingSensitivityAnalysis(newRollingSpillovers)

	# ==============================
	# OUTPUT
	# ==============================
	# Total, FROM, TO, and NET Each, NET Pairwise Data Spillover Table and Graph
	
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
print('Calc Rolling Spillovers...')
rollingSpillovers, temp1, temp2, temp3, temp4 = getRollingSpillovers(lag_order,forecast_horizon)
del rollingSpillovers, temp1, temp2, temp3, temp4
print('End of Calc Rolling Spillovers')

# SENSITIVITY
print('Calc Sensitivity Analysis Spillovers: lag_order...')
sensitivityRange = getRollingSensitivityAnalysis('lag_order',min(1,math.floor(0.5*lag_order)),math.ceil(1.5*lag_order),lag_order,forecast_horizon,sectors)
del sensitivityRange

print('Calc Sensitivity Analysis Spillovers: forecast_horizon...')
sensitivityRange = getRollingSensitivityAnalysis('forecast_horizon',min(1,math.floor(0.5*forecast_horizon)),math.ceil(1.5*forecast_horizon),lag_order,forecast_horizon,sectors)
del sensitivityRange
print('End of Calc Analysis Spillovers')

print("End of Analysis")