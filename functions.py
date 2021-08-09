# ==============================
# IMPORT PACKAGE
# ==============================
import pandas as pd, numpy as np
from statsmodels.tsa.api import VAR

# ==============================
# MARKET DAYS
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

# ==============================
# DATA PREPARATION BASED ON OUTPUTMODE
# ==============================
def calcLnvariance (sectorsData):
	# sectorsData is a dict consist of dataframes of data for each sector/market
	# example: sectorsData['AGRI'] = pd.Dataframe(columns=['Open','High','Low','Close'])
	# np.log is natural log
	lnvariance = pd.DataFrame()
	for sector in sectorsData:
		a = np.log(sectorsData[sector]['High'])
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
def calcSpilloversTable(volatility, forecast_horizon=10, lag_order=None):
	# ===
	# sources:
	# https://www.statsmodels.org/dev/vector_ar.html
	# https://en.wikipedia.org/wiki/n#Comparison_with_BIC
	# https://groups.google.com/g/pystatsmodels/c/BqMqOIghN78/m/21NkPAEPJgIJ
	# ===
	model = VAR(volatility)
	results = model.fit(lag_order,ic='aic',verbose=True)
	
	lag_order = results.k_ar
	sigma_u = np.asarray(results.sigma_u)
	sd_u = np.sqrt(np.diag(sigma_u))
	
	fevd = results.fevd(forecast_horizon, sigma_u/sd_u)
	fe = fevd.decomp[:,-1,:]
	fevd = (fe / fe.sum(0)[:,None] * 100)

	cont_incl = fevd.sum(0)
	cont_to = fevd.sum(0) - np.diag(fevd)
	cont_from  = fevd.sum(1) - np.diag(fevd)
	spillover_index = cont_to.sum()/cont_incl.sum()

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