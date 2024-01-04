SELECT [visaReprID]
      ,[rcID]
	  ,rc.countryName AS receivingCountryName
	  ,rc.countryCode AS receivingCountryCode
      ,[dYear]
      ,[scCityID]
	  ,sc.countryName AS sendingCountryName
	  ,sc.countryCode AS sendingCountryCode
	  ,c.cityName AS sendingCityName
      ,[ExtSerPro]
      ,[reprByRcID]
	  ,reprby.countryName AS representedByReceivingCountryName
	  ,reprby.countryCode AS representedByReceivingCountryCode
  FROM [visa].[dbo].[visaRepresentations] v
    INNER JOIN countries_cities cc ON v.scCityID = cc.countryCityID
	INNER JOIN countries sc ON cc.countryID = sc.countryID
	INNER JOIN cities c ON cc.cityID = c.cityID
	INNER JOIN countries rc ON v.rcID = rc.countryID
	INNER JOIN countries reprby ON v.reprByRcID = reprby.countryID