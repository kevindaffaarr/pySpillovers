# ==============================
# IMPORT PACKAGE
# ==============================
import pandas as pd, numpy as np

# ==============================
# DATA PREPARATION BASED ON OUTPUTMODE
# ==============================
def calcLnvariance (sectorData):
	# np.log is natural log
	lnvariance = {}
	for sector in sectorData:
		a = np.log(sectorData[sector]['High'])
		lnvariance[sector] = 0.361*((np.log(sectorData[sector]['High'])-np.log(sectorData[sector]['Low']))**2)
	return lnvariance

def calcVolatilityDiebold(lnvariance,marketDays):
	volatility = {}
	for sector in lnvariance:
		lnvariance[sector] = lnvariance[sector].to_frame('lnvariance')
		marketDays[sector] = marketDays[sector].to_frame('marketDays')
		lnvariance[sector]['year'] = pd.to_datetime(lnvariance[sector].index.values).year.astype(int)
		lnvariance[sector] = lnvariance[sector].reset_index().merge(marketDays[sector],how='left',on='year').set_index('Date')

		volatility[sector] = 100 * np.sqrt(lnvariance[sector]['marketDays']*lnvariance[sector]['lnvariance'])
	return volatility

def calcVolatilityAslam(lnvariance,marketDays):
	volatility = {}
	for sector in lnvariance:
		lnvariance[sector] = lnvariance[sector].to_frame('lnvariance')
		marketDays[sector] = marketDays[sector].to_frame('marketDays')
		lnvariance[sector]['year'] = pd.to_datetime(lnvariance[sector].index.values).year.astype(int)
		lnvariance[sector] = lnvariance[sector].merge(marketDays[sector],how='left',on='year')

		volatility[sector] = np.arcsinh(np.sqrt(lnvariance[sector]['marketDays']*lnvariance[sector]['lnvariance']))
	return volatility

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