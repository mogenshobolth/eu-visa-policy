SELECT [visaPracticeEuID]
      ,[rcID]
	  ,rc.countryName AS receivingCountryName
	  ,rc.countryCode AS receivingCountryCode
      ,[scCountryID]
	  ,sc.countryName AS sendingCountryName
	  ,sc.countryCode AS sendingCountryCode
      ,[scCityID]
	  ,c.cityName AS sendingCityName
      ,v.[cityName]
      ,[dYear]
      ,[shortStayAppliedFor]
      ,[shortStayIssued]
      ,[shortStayRefused]
      ,[shortStayRefusalRate]
      ,[AvisasIssued]
      ,[BvisasIssued]
      ,[CvisasIssued]
      ,[VTLvisasIssued]
      ,[TotalABCIssued]
      ,[DvisasIssued]
      ,[DCvisasIssued]
      ,[TotalABCDIssued]
      ,[visasNotIssued]
      ,[visasApplied]
      ,[visaRefusalRate]
      ,[issuedReprOther]
      ,[visasNotIssuedCalc]
  FROM [visa].[dbo].[visaPractice_EU_2004] v
	INNER JOIN countries sc ON v.scCountryID = sc.countryID
	INNER JOIN cities c ON v.scCityID = c.cityID
	INNER JOIN countries rc ON v.rcID = rc.countryID
