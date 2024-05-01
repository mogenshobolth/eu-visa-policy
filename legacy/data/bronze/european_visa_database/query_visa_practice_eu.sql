SELECT [visaPracticeEuID]
      ,[rcID]
	  ,rc.countryName AS receivingCountryName
	  ,rc.countryCode AS receivengCountryCode
      ,[scCityID]
	  ,sc.countryName AS sendingCountryName
	  ,sc.countryCode AS sendingCountryCode
	  ,c.cityName AS sendingCityName
      ,[dYear]
      ,[shortStayAppliedFor]
      ,[shortStayIssued]
      ,[shortStayRefused]
      ,[shortStayRefusalRate]
      ,[issuedA_All]
      ,[issuedA_Mev]
      ,[issuedB]
      ,[issuedC_All]
      ,[issuedC_Mev]
      ,[issuedD]
      ,[issuedDC]
      ,[issuedVTL]
      ,[issuedADS]
      ,[issuedABC]
      ,[issuedABCVTL]
      ,[issuedABCDVTL]
      ,[issuedABCDDCVTL]
      ,[appliedA]
      ,[appliedB]
      ,[appliedC]
      ,[appliedABC]
      ,[notIssuedA]
      ,[notIssuedB]
      ,[notIssuedC]
      ,[notIssuedABC]
  FROM [visa].[dbo].[visaPractice_EU] v
	INNER JOIN countries_cities cc ON v.scCityID = cc.countryCityID
	INNER JOIN countries sc ON cc.countryID = sc.countryID
	INNER JOIN cities c ON cc.cityID = c.cityID
	INNER JOIN countries rc ON v.rcID = rc.countryID
