# ==============================
# IMPORT PACKAGE
# ==============================
import pandas as pd, numpy as np
from statsmodels.tsa.api import VAR
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ==============================
# IMPORT DATA
# ==============================
def calcMarketDays(sectorData,marketDaysYearEnd=None):
	# sectorData is a dataframes of data for a sector/market
	# example: sectorData = pd.Dataframe(columns=['Open','High','Low','Close'])

	# marketDaysYearEnd is a integer of how many market days in the last year of data.
	# usually the data of last/current year is not full until end of year because the year is still running
	
	df = pd.DataFrame(pd.DatetimeIndex(sectorData.index.values).year,columns=["year"]).groupby("year")["year"].count()
	marketDays = df
	if marketDaysYearEnd != None:
		marketDays.loc[marketDays.index[-1]] = marketDaysYearEnd
	return marketDays

def getWithRollingWindow(sectorData,dateFrom,dateTo,rollingWindow=200):
	rollingWindow = 200 if rollingWindow is None else rollingWindow
	rollingWindow = rollingWindow - 1
	sectorData = pd.concat([(sectorData.loc[:dateFrom]).iloc[-rollingWindow:],sectorData.loc[dateFrom:dateTo]])
	return sectorData

# ==============================
# DATA PREPARATION BASED ON OUTPUTMODE
# ==============================
def calcLnvariance (sectorsData):
	# sectorsData is a dict consist of dataframes of data for each sector/market
	# example: sectorsData['AGRI'] = pd.Dataframe(columns=['Open','High','Low','Close'])
	# np.log is natural log
	lnvariance = pd.DataFrame()
	for sector in sectorsData:
		lnvariance[sector] = 0.361*((np.log(sectorsData[sector]['High'])-np.log(sectorsData[sector]['Low']))**2)
	return lnvariance

def calcVolatilityDiebold(lnvariance,marketDays):
	# lnvariance is a dataframe from calcLnvariance function
	# marketDays
	lnvariance = lnvariance.to_dict('series')
	volatility = pd.DataFrame()
	for sector in lnvariance:
		lnvariance[sector] = lnvariance[sector].to_frame('lnvariance')
		if isinstance(marketDays,dict):
			marketDays[sector] = marketDays[sector].to_frame('marketDays')
			lnvariance[sector]['year'] = pd.to_datetime(lnvariance[sector].index.values).year.astype(int)
			lnvariance[sector] = lnvariance[sector].reset_index().merge(marketDays[sector],how='left',on='year').set_index('Date')
		else:
			lnvariance[sector]['marketDays'] = marketDays
		volatility[sector] = 100 * np.sqrt(lnvariance[sector]['marketDays']*lnvariance[sector]['lnvariance'])
	return volatility

def calcVolatilityAslam(lnvariance,marketDays):
	# lnvariance is a dataframe from calcLnvariance function
	# marketDays
	lnvariance = lnvariance.to_dict('series')
	volatility = pd.DataFrame()
	for sector in lnvariance:
		lnvariance[sector] = lnvariance[sector].to_frame('lnvariance')
		if isinstance(marketDays,dict):
			marketDays[sector] = marketDays[sector].to_frame('marketDays')
			lnvariance[sector]['year'] = pd.to_datetime(lnvariance[sector].index.values).year.astype(int)
			lnvariance[sector] = lnvariance[sector].reset_index().merge(marketDays[sector],how='left',on='year').set_index('Date')
		else:
			lnvariance[sector]['marketDays'] = marketDays
		volatility[sector] = np.arcsinh(np.sqrt(lnvariance[sector]['marketDays']*lnvariance[sector]['lnvariance']))
	return volatility

# ==============================
# Calc Sets of Statistic
# ==============================
def calcSetStats(volatility):
	setStats = pd.DataFrame(columns=['mean','median','max','min','stdDev','skew','kurtosis','count'])
	setStats.index.name = 'sector'
	for sector in volatility:
		setStats.loc[sector] = \
			volatility[sector].mean(), \
			volatility[sector].median(), \
			volatility[sector].max(), \
			volatility[sector].min(), \
			volatility[sector].std(), \
			volatility[sector].skew(), \
			volatility[sector].kurtosis(), \
			volatility[sector].count()
	return setStats

# ==============================
# Spillovers Table Based on Diebold Yilmaz 2012
# ==============================
def calcAvgSpilloversTable(volatility, forecast_horizon=10, lag_order=None):
	# ===
	# sources:
	# https://www.statsmodels.org/dev/vector_ar.html
	# https://en.wikipedia.org/wiki/n#Comparison_with_BIC
	# https://groups.google.com/g/pystatsmodels/c/BqMqOIghN78/m/21NkPAEPJgIJ
	# ===
	forecast_horizon = 10 if forecast_horizon is None else forecast_horizon
	
	model = VAR(volatility)
	if lag_order==None:
		results = model.fit(lag_order,ic='aic')
		lag_order = results.k_ar
	else:
		results = model.fit(lag_order)
	
	sigma_u = np.asarray(results.sigma_u)
	sd_u = np.sqrt(np.diag(sigma_u))
	
	fevd = results.fevd(forecast_horizon, sigma_u/sd_u)
	fe = fevd.decomp[:,-1,:]
	fevd = (fe / fe.sum(1)[:,None] * 100)

	cont_incl = fevd.sum(0)
	cont_to = fevd.sum(0) - np.diag(fevd)
	cont_from  = fevd.sum(1) - np.diag(fevd)
	spillover_index = 100*cont_to.sum()/cont_incl.sum()

	names = model.endog_names
	spilloversTable = pd.DataFrame(fevd, columns=names).set_index([names])
	spilloversTable.loc['Cont_To'] = cont_to
	spilloversTable.loc['Cont_Incl'] = cont_incl
	spilloversTable=pd.concat([spilloversTable,pd.DataFrame(cont_from,columns=['Cont_From']).set_index([names])],axis=1)
	spilloversTable=pd.concat([spilloversTable,pd.DataFrame(cont_to-cont_from,columns=['Cont_Net']).set_index([names])],axis=1)
	spilloversTable.loc['Cont_To','Cont_From'] = cont_to.sum()
	spilloversTable.loc['Cont_Incl','Cont_From'] = cont_incl.sum()
	spilloversTable.loc['Cont_Incl','Cont_Net'] = spillover_index

	return spilloversTable, lag_order, forecast_horizon

# ==============================
# Rolling Spillovers Based on Diebold Yilmaz 2012
# ==============================
def calcRollingSpillovers(volatility, forecast_horizon=10, lag_order=None,rollingWindow=200):
	# rollingSpillovers: 
	# [total] : spillover_index
	# [to][sector] : Cont_To[sector]
	# [from][sector] : Cont_From[sector]
	# [net][sector] : Cont_Net[sector]
	# [pairwise][sector_to][sector_from] : spilloversTable.loc['sector_From','sector_To']

	forecast_horizon = 10 if forecast_horizon is None else forecast_horizon
	rollingWindow = 200 if rollingWindow is None else rollingWindow
	
	rollingSpillovers = {}
	rollingSpillovers['total'] = pd.DataFrame()
	rollingSpillovers['to'] = pd.DataFrame()
	rollingSpillovers['from'] = pd.DataFrame()
	rollingSpillovers['net'] = pd.DataFrame()
	rollingSpillovers['pairwiseTo'] = {}
	rollingSpillovers['pairwiseNet'] = {}
	sectors = volatility.columns
	for sector in sectors:
		rollingSpillovers['pairwiseTo'][sector] = pd.DataFrame(columns=sectors)
		rollingSpillovers['pairwiseNet'][sector] = pd.DataFrame(columns=sectors)
	
	for i in range(volatility.shape[0]-(rollingWindow-1)):
		UBound = i+rollingWindow
		df = volatility.iloc[i:UBound]
		spilloversTable, lag_order, forecast_horizon = calcAvgSpilloversTable(df,forecast_horizon,lag_order)
		
		rollingSpillovers['total'] = rollingSpillovers['total'].append(pd.DataFrame([[spilloversTable.loc['Cont_Incl','Cont_Net']]],index=[volatility.iloc[UBound-1].name]))
		rollingSpillovers['to'] = rollingSpillovers['to'].append(pd.DataFrame([spilloversTable.loc['Cont_To']],index=[volatility.iloc[UBound-1].name]))
		rollingSpillovers['from'] = rollingSpillovers['from'].append(pd.DataFrame([spilloversTable['Cont_From']],index=[volatility.iloc[UBound-1].name]))
		rollingSpillovers['net'] = rollingSpillovers['net'].append(pd.DataFrame([spilloversTable['Cont_Net']],index=[volatility.iloc[UBound-1].name]))
		for sector in sectors:
			rollingSpillovers['pairwiseTo'][sector].loc[volatility.iloc[UBound-1].name] = spilloversTable[sector]
			rollingSpillovers['pairwiseNet'][sector].loc[volatility.iloc[UBound-1].name] = spilloversTable[sector]-spilloversTable.loc[sector]

	return rollingSpillovers

# ==============================
# SENSITIVITY ANALYSIS:
# Average and Dynamic Spillovers With Variant Lag Order
# ==============================
def calcRollingSensitivityAnalysis(newRollingSpillovers):
	sectors = list(newRollingSpillovers['to'].keys())
	# ==============================
	# ARRAY PREPARATIONS
	# ==============================
	sensitivityRange = {}
	sensitivityRange['total'] = pd.DataFrame()
	sensitivityRange['to'] = {}
	sensitivityRange['from'] = {}
	sensitivityRange['net'] = {}
	sensitivityRange['pairwiseTo'] = {}
	sensitivityRange['pairwiseNet'] = {}
	for sector in sectors:
		sensitivityRange['to'][sector] = pd.DataFrame()
		sensitivityRange['from'][sector] = pd.DataFrame()
		sensitivityRange['net'][sector] = pd.DataFrame()
		sensitivityRange['pairwiseTo'][sector] = {}
		sensitivityRange['pairwiseNet'][sector] = {}
		for sectorFrom in sectors:
			sensitivityRange['pairwiseTo'][sector][sectorFrom] = pd.DataFrame()
			sensitivityRange['pairwiseNet'][sector][sectorFrom] = pd.DataFrame()

	sensitivityRange['total'] = pd.DataFrame({'min':newRollingSpillovers['total'].min(axis=1),'median':newRollingSpillovers['total'].median(axis=1),'max':newRollingSpillovers['total'].max(axis=1)})
	for sector in sectors:
		sensitivityRange['to'][sector] = pd.DataFrame({'min':newRollingSpillovers['to'][sector].min(axis=1),'median':newRollingSpillovers['to'][sector].median(axis=1),'max':newRollingSpillovers['to'][sector].max(axis=1)})
		sensitivityRange['from'][sector] = pd.DataFrame({'min':newRollingSpillovers['from'][sector].min(axis=1),'median':newRollingSpillovers['from'][sector].median(axis=1),'max':newRollingSpillovers['from'][sector].max(axis=1)})
		sensitivityRange['net'][sector] = pd.DataFrame({'min':newRollingSpillovers['net'][sector].min(axis=1),'median':newRollingSpillovers['net'][sector].median(axis=1),'max':newRollingSpillovers['net'][sector].max(axis=1)})
		for sectorFrom in sectors:
			sensitivityRange['pairwiseTo'][sector][sectorFrom] = pd.DataFrame({'min':newRollingSpillovers['pairwiseTo'][sector][sectorFrom].min(axis=1),'median':newRollingSpillovers['pairwiseTo'][sector][sectorFrom].median(axis=1),'max':newRollingSpillovers['pairwiseTo'][sector][sectorFrom].max(axis=1)})
			sensitivityRange['pairwiseNet'][sector][sectorFrom] = pd.DataFrame({'min':newRollingSpillovers['pairwiseNet'][sector][sectorFrom].min(axis=1),'median':newRollingSpillovers['pairwiseNet'][sector][sectorFrom].median(axis=1),'max':newRollingSpillovers['pairwiseNet'][sector][sectorFrom].max(axis=1)})
	
	return sensitivityRange
# ==============================
# Charting
# ==============================
def genStackedTimeSeriesChart(df,filename,xaxis_title,yaxis_title):
	fig = go.Figure()
	for column in df:
		fig.add_trace(go.Scatter( \
			x=df.index, \
			y=df[column], \
			name=column \
		))
	fig.update_layout(title={'text':filename, 'x':0.5})
	fig.update_layout(legend={"orientation":"h","y":-0.075})
	fig.update_layout(margin=dict(l=50,r=50,b=100,t=50,pad=0))
	fig.update_layout( \
		xaxis_title = xaxis_title, \
		yaxis_title = yaxis_title, \
		template = 'plotly_white' \
	)
	fig.write_image('output\\'+filename+'.png',width=1400,height=1050)
	return True