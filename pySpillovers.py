# ==============================
# IMPORT PACKAGE
# ==============================
import pandas as pd, numpy as np
import math
import functions as f

from pathlib import Path
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
	lnreturn = f.calcLnreturn(sectorsData)
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
	# Correlation of Data
	# ==============================
	correlationTable = volatility.corr(method='pearson')

	# ==============================
	# Spillovers Table
	# ==============================
	spilloversTable, lag_order, forecast_horizon = f.calcAvgSpilloversTable(volatility,forecast_horizon,lag_order)

	# ==============================
	# OUTPUT
	# ==============================
	# setStats
	setStats.to_csv('output\setStats.csv')

	# correlationTable
	correlationTable.to_csv('output\correlationTable.csv')
	
	# Volatility Table
	lnreturn.to_csv('output\\lnreturn.csv')
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
	# ['pairwiseTo'][sectorTo][sectorFrom]
	# ['pairwiseNet'][sectorTo][sectorFrom]
	rollingSpillovers = f.calcRollingSpillovers(volatility, forecast_horizon, lag_order,rollingWindow)

	return rollingSpillovers, volatility, lnvariance, lag_order, forecast_horizon

def exportRollingSpillovers(rollingSpillovers,sectors):
	# ==============================
	# OUTPUT
	# ==============================
	# Total, FROM, TO, NET, PairwiseTo, PairwiseNet Data Spillover Table and Graph
	outputDict = {}
	filenameDict = {}
	subplotsOutputDict = {}
	subplotsfilenameDict = {}
	
	outputDict['Total'] = rollingSpillovers['total'][0]
	filenameDict['Total'] = 'Rolling Total Volatility Spillovers'
	
	for column in sectors:
		outputDict['to_'+column] = rollingSpillovers['to'][column]
		filenameDict['to_'+column] = 'Rolling Directional Volatility Spillovers '+column+' - TO OTHERS'
		subplotsOutputDict['to_'+column] = rollingSpillovers['to'][column]
		subplotsfilenameDict['to_'+column] = 'Rolling Directional Volatility Spillovers '+column+'- TO OTHERS'
	f.genSubplotsTimeSeriesChart( \
		subplotsOutputDict, \
		chartNameDict=subplotsfilenameDict, \
		xaxis_title='Date', \
		yaxis_title='%', \
		filename='Rolling Directional Volatility Spillovers All Sectors - TO OTHERS', \
		chartCol=3
	)
	subplotsOutputDict = {}
	subplotsfilenameDict = {}

	for column in sectors:
		outputDict['from_'+column] = rollingSpillovers['from'][column]
		filenameDict['from_'+column] = 'Rolling Directional Volatility Spillovers '+column+' - FROM OTHERS'
		subplotsOutputDict['from_'+column] = rollingSpillovers['from'][column]
		subplotsfilenameDict['from_'+column] = 'Rolling Directional Volatility Spillovers '+column+' - FROM OTHERS'
	f.genSubplotsTimeSeriesChart( \
		subplotsOutputDict, \
		chartNameDict=subplotsfilenameDict, \
		xaxis_title='Date', \
		yaxis_title='%', \
		filename='Rolling Directional Volatility Spillovers All Sectors - FROM OTHERS', \
		chartCol=3
	)
	subplotsOutputDict = {}
	subplotsfilenameDict = {}

	for column in sectors:
		outputDict['net_'+column] = rollingSpillovers['net'][column]
		filenameDict['net_'+column] = 'Rolling Directional Volatility Spillovers '+column+' - NET'
		subplotsOutputDict['net_'+column] = rollingSpillovers['net'][column]
		subplotsfilenameDict['net_'+column] = 'Rolling Directional Volatility Spillovers '+column+' - NET'
	f.genSubplotsTimeSeriesChart( \
		subplotsOutputDict, \
		chartNameDict=subplotsfilenameDict, \
		xaxis_title='Date', \
		yaxis_title='%', \
		filename='Rolling Directional Volatility Spillovers All Sectors - NET', \
		chartCol=3
	)
	subplotsOutputDict = {}
	subplotsfilenameDict = {}

	for sectorTo in sectors:
		for sectorFrom in sectors:
			outputDict['pairwise_'+sectorTo+'_To_'+sectorFrom] = rollingSpillovers['pairwiseTo'][sectorTo][sectorFrom]
			filenameDict['pairwise_'+sectorTo+'_To_'+sectorFrom] = 'Rolling Pairwise '+sectorTo+' To '+sectorFrom
			subplotsOutputDict['pairwise_'+sectorTo+'_To_'+sectorFrom] = rollingSpillovers['pairwiseTo'][sectorTo][sectorFrom]
			subplotsfilenameDict['pairwise_'+sectorTo+'_To_'+sectorFrom] = 'Rolling Pairwise '+sectorTo+' To '+sectorFrom
	# f.genSubplotsTimeSeriesChart( \
	# 	subplotsOutputDict, \
	# 	chartNameDict=subplotsfilenameDict, \
	# 	xaxis_title='Date', \
	# 	yaxis_title='%', \
	# 	filename='Rolling Pairwise Volatility Spillovers All Sectors ', \
	# 	chartCol=3
	# )
	# subplotsOutputDict = {}
	# subplotsfilenameDict = {}

	for sectorTo in sectors:
		for sectorFrom in sectors:		
			outputDict['pairwise_'+sectorTo+'_Net_'+sectorFrom] = rollingSpillovers['pairwiseNet'][sectorTo][sectorFrom]
			filenameDict['pairwise_'+sectorTo+'_Net_'+sectorFrom] = 'Rolling Pairwise Net '+sectorTo+' - '+sectorFrom
			subplotsOutputDict['pairwise_'+sectorTo+'_Net_'+sectorFrom] = rollingSpillovers['pairwiseNet'][sectorTo][sectorFrom]
			subplotsfilenameDict['pairwise_'+sectorTo+'_Net_'+sectorFrom] = 'Rolling Pairwise Net '+sectorTo+' - '+sectorFrom
	# f.genSubplotsTimeSeriesChart( \
	# 	subplotsOutputDict, \
	# 	chartNameDict=subplotsfilenameDict, \
	# 	xaxis_title='Date', \
	# 	yaxis_title='%', \
	# 	filename='Rolling Pairwise NET Volatility Spillovers All Sectors', \
	# 	chartCol=3
	# )
	# subplotsOutputDict = {}
	# subplotsfilenameDict = {}

	# GRAPH
	print('Spitting The Rolling Spillovers Graph...')
	f.genBulkTimeSeriesChart(outputDict,filenameDict,xaxis_title='Date',yaxis_title='%')
	
	# TABLE
	print('Export The Rolling Spillovers Table...')
	filename = 'output\\rollingSpilloversTable.csv'
	header = ''
	for key in filenameDict:
		header = header + filenameDict[key] + ','
	header = header + '\n'
	with open(filename,'w') as out:
		out.write(header)
	outputDict = pd.DataFrame.from_dict(outputDict)
	outputDict.to_csv(filename,mode='a')

	return True

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
	# Total, FROM, TO, NET, PairwiseTo, PairwiseNet Data Spillover Table and Graph
	outputDict = {}
	filenameDict = {}
	subplotsOutputDict = {}
	subplotsfilenameDict = {}
	
	outputDict['Total'] = sensitivityRange['total']
	filenameDict['Total'] = 'Sensitivity Range Rolling Total Volatility Spillovers'
	
	for column in sectors:
		outputDict['to_'+column] = sensitivityRange['to'][column]
		filenameDict['to_'+column] = 'Sensitivity Range Rolling Directional Volatility Spillovers '+column+' - TO OTHERS'
		subplotsOutputDict['to_'+column] = sensitivityRange['to'][column]
		subplotsfilenameDict['to_'+column] = 'Sensitivity Range Rolling Directional Volatility Spillovers '+column+' - TO OTHERS'
	f.genSubplotsRangeChart( \
		subplotsOutputDict, \
		chartNameDict=subplotsfilenameDict, \
		xaxis_title='Date', \
		yaxis_title='%', \
		filename='sensitivity_'+variantParam+'\\'+'Sensitivity Range Rolling Directional Volatility Spillovers All Sectors - TO OTHERS', \
		chartCol=3
	)
	subplotsOutputDict = {}
	subplotsfilenameDict = {}

	for column in sectors:
		outputDict['from_'+column] = sensitivityRange['from'][column]
		filenameDict['from_'+column] = 'Sensitivity Range Rolling Directional Volatility Spillovers '+column+' - FROM OTHERS'
		subplotsOutputDict['from_'+column] = sensitivityRange['from'][column]
		subplotsfilenameDict['from_'+column] = 'Sensitivity Range Rolling Directional Volatility Spillovers '+column+' - FROM OTHERS'
	f.genSubplotsRangeChart( \
		subplotsOutputDict, \
		chartNameDict=subplotsfilenameDict, \
		xaxis_title='Date', \
		yaxis_title='%', \
		filename='sensitivity_'+variantParam+'\\'+'Sensitivity Range Rolling Directional Volatility Spillovers All Sectors - FROM OTHERS', \
		chartCol=3
	)
	subplotsOutputDict = {}
	subplotsfilenameDict = {}

	for column in sectors:
		outputDict['net_'+column] = sensitivityRange['net'][column]
		filenameDict['net_'+column] = 'Sensitivity Range Rolling Directional Volatility Spillovers '+column+' - NET'
		subplotsOutputDict['net_'+column] = sensitivityRange['net'][column]
		subplotsfilenameDict['net_'+column] = 'Sensitivity Range Rolling Directional Volatility Spillovers '+column+' - NET'
	f.genSubplotsRangeChart( \
		subplotsOutputDict, \
		chartNameDict=subplotsfilenameDict, \
		xaxis_title='Date', \
		yaxis_title='%', \
		filename='sensitivity_'+variantParam+'\\'+'Sensitivity Range Rolling Directional Volatility Spillovers All Sectors - NET', \
		chartCol=3
	)
	subplotsOutputDict = {}
	subplotsfilenameDict = {}

	for sectorTo in sectors:
		for sectorFrom in sectors:
			outputDict['pairwise_'+sectorTo+'_To_'+sectorFrom] = sensitivityRange['pairwiseTo'][sectorTo][sectorFrom]
			filenameDict['pairwise_'+sectorTo+'_To_'+sectorFrom] = 'Sensitivity Range Rolling Pairwise '+sectorTo+' To '+sectorFrom
			subplotsOutputDict['pairwise_'+sectorTo+'_To_'+sectorFrom] = sensitivityRange['pairwiseTo'][sectorTo][sectorFrom]
			subplotsfilenameDict['pairwise_'+sectorTo+'_To_'+sectorFrom] = 'Sensitivity Range Rolling Pairwise '+sectorTo+' To '+sectorFrom
	# f.genSubplotsRangeChart( \
	# 	subplotsOutputDict, \
	# 	chartNameDict=subplotsfilenameDict, \
	# 	xaxis_title='Date', \
	# 	yaxis_title='%', \
	# 	filename='sensitivity_'+variantParam+'\\'+'Sensitivity Range Rolling Pairwise Volatility Spillovers All Sectors', \
	# 	chartCol=3
	# )
	# subplotsOutputDict = {}
	# subplotsfilenameDict = {}

	for sectorTo in sectors:
		for sectorFrom in sectors:		
			outputDict['pairwise_'+sectorTo+'_Net_'+sectorFrom] = sensitivityRange['pairwiseNet'][sectorTo][sectorFrom]
			filenameDict['pairwise_'+sectorTo+'_Net_'+sectorFrom] = 'Sensitivity Range Rolling Pairwise Net '+sectorTo+' - '+sectorFrom
			subplotsOutputDict['pairwise_'+sectorTo+'_Net_'+sectorFrom] = sensitivityRange['pairwiseNet'][sectorTo][sectorFrom]
			subplotsfilenameDict['pairwise_'+sectorTo+'_Net_'+sectorFrom] = 'Sensitivity Range Rolling Pairwise Net '+sectorTo+' - '+sectorFrom
	# f.genSubplotsRangeChart( \
	# 	subplotsOutputDict, \
	# 	chartNameDict=subplotsfilenameDict, \
	# 	xaxis_title='Date', \
	# 	yaxis_title='%', \
	# 	filename='sensitivity_'+variantParam+'\\'+'Sensitivity Range Rolling Pairwise NET Volatility Spillovers All Sectors', \
	# 	chartCol=3
	# )
	# subplotsOutputDict = {}
	# subplotsfilenameDict = {}

	# GRAPH
	print('Spitting The Sensitivity Range Rolling Spillovers Graph...')
	f.genBulkRangeChart(outputDict,filenameDict,xaxis_title='Date',yaxis_title='%',folder='sensitivity_'+variantParam+'\\')
	
	# TABLE
	print('Export The Sensitivity Range Rolling Spillovers Table...')
	filename = 'output\\sensitivity_'+variantParam+'\\sensitivityRangeTable.csv'
	df = {(outerKey, innerKey): values for outerKey, innerDict in outputDict.items() for innerKey, values in innerDict.iteritems()}
	df = pd.DataFrame(df)
	df.to_csv(filename)
	return sensitivityRange


# ==================================================================================================
# ===============================================MAIN===============================================
# ==================================================================================================
# ==============================
# CHECK DIRECTORY
# ==============================
Path("output").mkdir(parents=True, exist_ok=True)
Path("output/sensitivity_lag_order").mkdir(parents=True, exist_ok=True)
Path("output/sensitivity_forecast_horizon").mkdir(parents=True, exist_ok=True)

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
export = exportRollingSpillovers(rollingSpillovers,sectors)
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