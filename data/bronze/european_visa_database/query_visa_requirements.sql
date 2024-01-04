SELECT [visaReqID]
      ,[rcID]
	  ,rc.countryName AS receivingCountryName
	  ,rc.countryCode AS receivingCountryCode
      ,[scID]
	  ,sc.countryName AS sendingCountryName
	  ,sc.countryCode AS sendingCountryCode
      ,[dYear]
      ,[shortStayVisaRequired]
  FROM [visa].[dbo].[visaRequirements] v
    INNER JOIN countries sc ON v.scID = sc.countryID
	INNER JOIN countries rc ON v.rcID = rc.countryID

