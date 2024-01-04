Imports Microsoft.VisualBasic
Imports System.Data.OleDb
Imports System.Data
Imports OfficeOpenXml
Imports System.IO

Public Class impVisaPractice

    '* Import visa practice data for the United Kingdom
    Public Shared Function importVisaPracticeUK(ByVal dYear As Integer) As String

        '* Define database variables
        Dim SQL As String
        Dim db As New OleDbConnection(System.Web.Configuration.WebConfigurationManager.ConnectionStrings("visa").ConnectionString)
        Dim cmd As OleDbCommand

        '* Define excel variables
        Dim xls As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & HttpContext.Current.Server.MapPath("DataImport\conversionTable.xls") & ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1;ImportMixedTypes=Text""")
        Dim importFile As FileInfo

        '* Define other variables
        Dim intRowsInserted As Integer
        Dim intRecordsProcessed As Integer
        Dim strResult As String = ""
        Dim dtCities As DataTable
        Dim dtAltCityNames As DataTable
        Dim intRow As Integer
        Dim intSheet As Integer
        Dim drc() As DataRow
        Dim intRcID As Integer
        Dim strCityName As String
        Dim intCityID As Integer
        Dim strVisitReceived As String
        Dim strVisitIssued As String
        Dim strVisitRefused As String
        Dim strVisitFamilyReceived As String
        Dim strVisitFamilyIssued As String
        Dim strVisitFamilyRefused As String
        Dim strReceived As String
        Dim strIssued As String
        Dim strRefused As String
        Dim strWithdrawn As String
        Dim strLapsed As String
        Dim strDecided As String
        Dim strTableField As String = ""

        Try

            '* Open database
            db.Open()

            '* Open excelsheet with variation in country names
            xls.Open()

            '* Get database table with all countries/cities
            SQL = "SELECT countries_cities.countryCityID, countries.countryID, countries.countryName, countries.countryCode, cities.cityID, cities.cityName, dbo.ufn_getCountryStartYear(countryName) AS countryStartYear " _
                & " FROM  countries_cities INNER JOIN " _
                & "       cities ON countries_cities.cityID = cities.cityID INNER JOIN " _
                & "       countries ON countries_cities.countryID = countries.countryID;"
            cmd = New OleDbCommand(SQL, db)
            dtCities = New DataTable : dtCities.Load(cmd.ExecuteReader)

            '* Open excel sheet with alternative/variations in city names
            cmd = New OleDbCommand("SELECT * FROM [cities$]", xls)
            dtAltCityNames = New DataTable : dtAltCityNames.Load(cmd.ExecuteReader)

            '* Get countryID of the United Kingdom
            SQL = "SELECT countryID FROM countries WHERE countryName = 'United Kingdom'"
            cmd = New OleDbCommand(SQL, db)
            intRcID = cmd.ExecuteScalar()

            '* Open excel file with information on visa practice
            importFile = New FileInfo(HttpContext.Current.Server.MapPath("DataImport\VisaPractice\UK\UK_Visa_ShortStay_Practice_" & dYear & ".xlsx"))
            Using xlPackage As New ExcelPackage(importFile)

                '* Import procedure for 2001, 2002, 2003 and 2004
                If dYear = 2001 Or dYear = 2002 Or dYear = 2003 Or dYear = 2004 Then

                    '* Get worksheet
                    Dim worksheet1 As ExcelWorksheet = xlPackage.Workbook.Worksheets(1)

                    '* Loop through each row in worksheet (skip header)
                    For intRow = 2 To worksheet1.Dimension.End.Row

                        '* Get city name and try to fetch ID from database
                        strCityName = ""
                        intCityID = 0
                        If Not worksheet1.Cells(intRow, 1) Is Nothing Then
                            If worksheet1.Cells(intRow, 1).Value <> "" Then

                                '* Get city name
                                If worksheet1.Cells(intRow, 1).IsRichText Then
                                    strCityName = worksheet1.Cells(intRow, 1).RichText.Text
                                Else
                                    strCityName = worksheet1.Cells(intRow, 1).Value
                                End If

                                '* Replace ' with '' to avoid SQL syntax error
                                strCityName = strCityName.Replace("'", "''")

                                '* Manually set the diplomatic post in "Bahrain" as being in the capital Manama
                                If strCityName = "Bahrain" Then strCityName = "Manama"

                                '* Manually set the diplomatic post in "Ascension Is" as being in Georgetown (Ascension Island) (UK overseas territory)
                                If strCityName = "Ascension Is" Then strCityName = "Georgetown (Ascension Island)"

                                '* Try to fetch cityID from database
                                drc = dtCities.Select("cityName = '" & strCityName & "'")
                                If drc.GetUpperBound(0) >= 0 Then
                                    intCityID = drc(0)("countryCityID")
                                End If

                                '* If city was not found in database table check for name variations in the conversion table
                                If intCityID = 0 Then
                                    drc = dtAltCityNames.Select("AlternativeCityName = '" & strCityName & "'")
                                    If drc.GetUpperBound(0) >= 0 Then
                                        drc = dtCities.Select("cityName = '" & drc(0)("cityName").ToString.Replace("'", "''") & "'")
                                        If drc.GetUpperBound(0) >= 0 Then
                                            intCityID = drc(0)("countryCityID")
                                        End If
                                    End If
                                End If

                            End If

                        End If

                        '* Warn user if city was not found in database
                        If intCityID = 0 Then strResult &= "Error: City not found in database: " & strCityName & "<br>"

                        '* Import visa practice data, if diplomatic post found in database
                        If intCityID <> 0 Then

                            '* Get information on visa practice
                            If IsNumeric(worksheet1.Cells(intRow, 2).Value) Then strVisitReceived = worksheet1.Cells(intRow, 2).Value Else strVisitReceived = "NULL"
                            If IsNumeric(worksheet1.Cells(intRow, 3).Value) Then strVisitIssued = worksheet1.Cells(intRow, 3).Value Else strVisitIssued = "NULL"
                            If IsNumeric(worksheet1.Cells(intRow, 4).Value) Then strVisitRefused = worksheet1.Cells(intRow, 4).Value Else strVisitRefused = "NULL"
                            If dYear = 2004 And IsNumeric(worksheet1.Cells(intRow, 5).Value) Then strVisitFamilyReceived = worksheet1.Cells(intRow, 5).Value Else strVisitFamilyReceived = "NULL"
                            If dYear = 2004 And IsNumeric(worksheet1.Cells(intRow, 6).Value) Then strVisitFamilyIssued = worksheet1.Cells(intRow, 6).Value Else strVisitFamilyIssued = "NULL"
                            If dYear = 2004 And IsNumeric(worksheet1.Cells(intRow, 7).Value) Then strVisitFamilyRefused = worksheet1.Cells(intRow, 7).Value Else strVisitFamilyRefused = "NULL"

                            '* Parse to database
                            SQL = "INSERT INTO visaPractice_UK (rcID,scCityID,dYear,visitReceived,visitIssued,visitRefused,visitFamilyReceived,visitFamilyIssued,visitFamilyRefused) VALUES (" & intRcID & "," & intCityID & "," & dYear & "," & strVisitReceived.Replace(",", ".") & "," & strVisitIssued.Replace(",", ".") & "," & strVisitRefused.Replace(",", ".") & "," & strVisitFamilyReceived.Replace(",", ".") & "," & strVisitFamilyIssued.Replace(",", ".") & "," & strVisitFamilyRefused.Replace(",", ".") & ");"
                            'strResult &= SQL & "<br>"
                            cmd = New OleDbCommand(SQL, db)
                            intRowsInserted += cmd.ExecuteNonQuery()

                        End If

                    Next

                End If

                '* Import procedure for 2005, 2006, 2007 and 2008
                If dYear = 2005 Or dYear = 2006 Or dYear = 2007 Or dYear = 2008 Then

                    '* Loop through data for (1) family visits, (2) other visits and (3) transits
                    For intSheet = 1 To 3

                        '* Get worksheets
                        Dim worksheet1 As ExcelWorksheet = xlPackage.Workbook.Worksheets(intSheet)

                        '* Loop through each row in worksheet (skip header)
                        For intRow = 3 To worksheet1.Dimension.End.Row

                            '* Get city name and try to fetch ID from database
                            strCityName = ""
                            intCityID = 0
                            If Not worksheet1.Cells(intRow, 1) Is Nothing Then
                                If worksheet1.Cells(intRow, 1).Value <> "" Then

                                    '* Get city name
                                    If worksheet1.Cells(intRow, 1).IsRichText Then
                                        strCityName = worksheet1.Cells(intRow, 1).RichText.Text
                                    Else
                                        strCityName = worksheet1.Cells(intRow, 1).Value
                                    End If

                                    '* Replace ' with '' to avoid SQL syntax error
                                    strCityName = strCityName.Replace("'", "''")

                                    '* Manually set London name
                                    If strCityName = "London(UK)" Then strCityName = "London"

                                    '* For 2006, 2007 and 2008 the city name is in (), e.g. "Belarus (Minsk)"
                                    If dYear = 2006 Or dYear = 2007 Or dYear = 2008 Then
                                        strCityName = strCityName.Substring(InStrRev(strCityName, "("), Len(strCityName) - InStrRev(strCityName, "("))
                                        strCityName = strCityName.Replace(")", "")
                                        strCityName = strCityName.Trim
                                    End If

                                    '* Manually set Casablanca name
                                    If strCityName = "Casablanca Ro" Then strCityName = "Casablanca"

                                    '* Manually set the diplomatic post in "Bahrain" as being in the capital Manama
                                    If strCityName = "Bahrain" Then strCityName = "Manama"

                                    '* Manually set the diplomatic post in "Ascension Is" as being in Georgetown (Ascension Island) (UK overseas territory)
                                    If strCityName = "Ascension Is" Then strCityName = "Georgetown (Ascension Island)"

                                    '* Try to fetch cityID from database
                                    drc = dtCities.Select("cityName = '" & strCityName & "' AND " & dYear & " >= countryStartYear")
                                    If drc.GetUpperBound(0) >= 0 Then
                                        intCityID = drc(0)("countryCityID")
                                    End If

                                    '* If city was not found in database table check for name variations in the conversion table
                                    If intCityID = 0 Then
                                        drc = dtAltCityNames.Select("AlternativeCityName = '" & strCityName & "'")
                                        If drc.GetUpperBound(0) >= 0 Then
                                            drc = dtCities.Select("cityName = '" & drc(0)("cityName").ToString.Replace("'", "''") & "' AND " & dYear & " >= countryStartYear")
                                            If drc.GetUpperBound(0) >= 0 Then
                                                intCityID = drc(0)("countryCityID")
                                            End If
                                        End If
                                    End If

                                End If

                            End If

                            '* Warn user if city was not found in database
                            If intCityID = 0 Then strResult &= "Error: City not found in database: " & strCityName & "<br>"

                            '* Import visa practice data, if diplomatic post found in database
                            If intCityID <> 0 Then

                                '* Get information on visa practice
                                '* Apps is coded as received; resolved is coded as decided
                                If IsNumeric(worksheet1.Cells(intRow, 2).Value) Then strReceived = worksheet1.Cells(intRow, 2).Value Else strReceived = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 3).Value) Then strIssued = worksheet1.Cells(intRow, 3).Value Else strIssued = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 4).Value) Then strRefused = worksheet1.Cells(intRow, 4).Value Else strRefused = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 5).Value) Then strWithdrawn = worksheet1.Cells(intRow, 5).Value Else strWithdrawn = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 6).Value) Then strLapsed = worksheet1.Cells(intRow, 6).Value Else strLapsed = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 7).Value) Then strDecided = worksheet1.Cells(intRow, 7).Value Else strDecided = "NULL"

                                '* Set table field depending on sheet being processed
                                If intSheet = 1 Then strTableField = "visitFamily"
                                If intSheet = 2 Then strTableField = "visitOther"
                                If intSheet = 3 Then strTableField = "transit"

                                '* Parse to database; update record if an entry already exist for the city and year; otherwise insert a new record
                                SQL = "IF EXISTS (SELECT visaPracticeUkID FROM visaPractice_UK WHERE rcID = " & intRcID & " AND scCityID = " & intCityID & " AND dYear = " & dYear & ") " _
                                    & " UPDATE visaPractice_UK SET " _
                                    & " " & strTableField & "Received = " & strReceived.Replace(",", ".") _
                                    & "," & strTableField & "Issued = " & strIssued.Replace(",", ".") _
                                    & "," & strTableField & "Refused = " & strRefused.Replace(",", ".") _
                                    & "," & strTableField & "Withdrawn = " & strWithdrawn.Replace(",", ".") _
                                    & "," & strTableField & "Lapsed = " & strLapsed.Replace(",", ".") _
                                    & "," & strTableField & "Decided = " & strDecided.Replace(",", ".") _
                                    & " WHERE rcID = " & intRcID & " AND scCityID = " & intCityID & " AND dYear = " & dYear _
                                    & " ELSE " _
                                    & " INSERT INTO visaPractice_UK (rcID,scCityID,dYear," & strTableField & "Received," & strTableField & "Issued," & strTableField & "Refused," & strTableField & "Withdrawn," & strTableField & "Lapsed," & strTableField & "Decided) VALUES " _
                                    & " (" & intRcID _
                                    & "," & intCityID _
                                    & "," & dYear _
                                    & "," & strReceived.Replace(",", ".") _
                                    & "," & strIssued.Replace(",", ".") _
                                    & "," & strRefused.Replace(",", ".") _
                                    & "," & strWithdrawn.Replace(",", ".") _
                                    & "," & strLapsed.Replace(",", ".") _
                                    & "," & strDecided.Replace(",", ".") & ");"
                                'strResult &= SQL & "<br>"
                                cmd = New OleDbCommand(SQL, db)
                                intRecordsProcessed += cmd.ExecuteNonQuery()

                            End If

                        Next

                    Next

                End If

            End Using

        Catch ex As Exception

            '* Write error
            strResult &= "Error: " & ex.Message

        Finally

            '* Close database
            db.Close()
            db.Dispose()

            '* Close excel import sheet
            xls.Close()
            xls.Dispose()

        End Try

        '* Get status on number of rows inserted
        strResult &= "Rows inserted: " & intRowsInserted & "<br>"
        strResult &= "Rows processed: " & intRecordsProcessed & "<br>"

        '* Return result of processing script
        Return strResult

    End Function

    '* Import visa practice data for the United States
    Public Shared Function importVisaPracticeUS(ByVal dYear As Integer) As String

        '* Define database variables
        Dim SQL As String
        Dim db As New OleDbConnection(System.Web.Configuration.WebConfigurationManager.ConnectionStrings("visa").ConnectionString)
        Dim cmd As OleDbCommand

        '* Define excel variables
        Dim xls As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & HttpContext.Current.Server.MapPath("DataImport\VisaPractice\US\US_Visa_ShortStay_Practice_RefusalRate.xls") & ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1;ImportMixedTypes=Text""")
        Dim xls2 As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & HttpContext.Current.Server.MapPath("DataImport\VisaPractice\US\US_Visa_ShortStay_Practice_Issued.xls") & ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1;ImportMixedTypes=Text""")
        Dim xls3 As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & HttpContext.Current.Server.MapPath("DataImport\conversionTable.xls") & ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1;ImportMixedTypes=Text""")

        '* Define other variables
        Dim intRowsInserted As Integer
        Dim intRecordsProcessed As Integer
        Dim strResult As String = ""
        Dim dtAltCountryNames As DataTable
        Dim dtSc As DataTable
        Dim dtVisaRefRate As DataTable
        Dim dtVisaIssued As DataTable
        Dim intRow As Integer
        Dim drc() As DataRow
        Dim drc2() As DataRow
        Dim intRcID As Integer
        Dim intScID As Integer
        Dim strScName As String
        Dim booScFoundInVisaRefRate As Boolean
        Dim booScFoundInVisaIssued As Boolean
        Dim intScYearEstablished As Integer
        Dim strVisaRefusalRate As String
        Dim strVisaIssued As String
        Dim strVisaIssued2 As String
        Dim strVisaRefusalRate2 As String
        Dim decVisaAppliedFor As Decimal
        Dim decVisaRefused As Decimal
        Dim strTypeB1issued As String
        Dim strTypeB2issued As String
        Dim strTypeB1comb2issued As String
        Dim strTypeB1comb2BCCissued As String
        Dim strTypeB1comb2BCVissued As String

        Try

            '* Open database
            db.Open()

            '* Open excelsheets with visa practice data
            xls.Open()
            xls2.Open()

            '* Open excelsheet with variation in country names
            xls3.Open()

            '* Open excel sheet with visa refusal rate data
            cmd = New OleDbCommand("SELECT * FROM [FY" & dYear & "$]", xls)
            dtVisaRefRate = New DataTable : dtVisaRefRate.Load(cmd.ExecuteReader)

            '* Open excel sheet with visa issued data
            cmd = New OleDbCommand("SELECT * FROM [FY" & dYear.ToString.Replace("20", "") & "$]", xls2)
            dtVisaIssued = New DataTable : dtVisaIssued.Load(cmd.ExecuteReader)

            '* Open excel sheet with alternative/variations in country names
            cmd = New OleDbCommand("SELECT * FROM [countries$]", xls3)
            dtAltCountryNames = New DataTable : dtAltCountryNames.Load(cmd.ExecuteReader)

            '* Get countryID of the United States
            SQL = "SELECT countryID FROM countries WHERE countryName = 'United States of America'"
            cmd = New OleDbCommand(SQL, db)
            intRcID = cmd.ExecuteScalar()

            '* Get database table with all potential sending countries
            SQL = "SELECT countryID, countryName, countryCode, dbo.ufn_getCountryStartYear(countryName) AS countryStartYear FROM countries WHERE countryID <> " & intRcID
            cmd = New OleDbCommand(SQL, db)
            dtSc = New DataTable : dtSc.Load(cmd.ExecuteReader)

            ''* Loop through all visa refusal rate data in the import datafile and ensure that the sending countries are found in database
            'For intRow = 0 To dtVisaRefRate.Rows.Count - 1

            '    '* Get sending country name
            '    strScName = dtVisaRefRate.Rows(intRow).Item("countryName")

            '    '* Replace ' with '' to avoid SQL syntax error
            '    strScName = strScName.Replace("'", "''")
            '    strScName = strScName.Trim

            '    '* Fetch sending country from database
            '    booScFoundInDatabase = False
            '    drc = dtSc.Select("CountryName = '" & strScName & "'")
            '    If drc.GetUpperBound(0) >= 0 Then
            '        booScFoundInDatabase = True
            '    End If

            '    '* If country data could not be found using the main name, try the variations 
            '    If Not booScFoundInDatabase Then
            '        drc = dtAltCountryNames.Select("alternativeName = '" & strScName & "'")
            '        If drc.GetUpperBound(0) >= 0 Then
            '            booScFoundInDatabase = True
            '        End If
            '    End If

            '    '* Warn user if country not find in import dataset
            '    If Not booScFoundInDatabase Then
            '        strResult &= "Error: Country not found: " & strScName & "<br>"
            '    Else
            '        'strResult &= intRow & ":" & strScName & ":" & drc(0)("countryName") & "<br>"
            '    End If

            'Next

            ''* Loop through all visa issued data in the import datafile and ensure that the sending countries are found in database
            'For intRow = 0 To dtVisaIssued.Rows.Count - 1

            '    '* Get sending country name
            '    strScName = dtVisaIssued.Rows(intRow).Item("countryName").ToString

            '    '* Replace ' with '' to avoid SQL syntax error
            '    strScName = strScName.Replace("'", "''")
            '    strScName = strScName.Trim

            '    '* Fetch sending country from database
            '    booScFoundInDatabase = False
            '    drc = dtSc.Select("CountryName = '" & strScName & "'")
            '    If drc.GetUpperBound(0) >= 0 Then
            '        booScFoundInDatabase = True
            '    End If

            '    '* If country data could not be found using the main name, try the variations 
            '    If Not booScFoundInDatabase Then
            '        drc = dtAltCountryNames.Select("alternativeName = '" & strScName & "'")
            '        If drc.GetUpperBound(0) >= 0 Then
            '            booScFoundInDatabase = True
            '        End If
            '    End If

            '    '* Warn user if country not find in import dataset
            '    If Not booScFoundInDatabase And strScName <> "" Then
            '        strResult &= "Error: Country not found: " & strScName & "<br>"
            '    Else
            '        'strResult &= intRow & ":" & strScName & ":" & drc(0)("countryName") & "<br>"
            '    End If

            'Next

            '* Loop through all sending countries and add visa practice
            For intRow = 0 To dtSc.Rows.Count - 1

                '* Reset data on visa refusal rate and visas issued
                strVisaRefusalRate = "NULL"
                strVisaIssued = "NULL"
                strTypeB1issued = "NULL"
                strTypeB2issued = "NULL"
                strTypeB1comb2issued = "NULL"
                strTypeB1comb2BCCissued = "NULL"
                strTypeB1comb2BCVissued = "NULL"

                '* Get sending country ID and name
                intScID = dtSc.Rows(intRow).Item("countryID")
                strScName = dtSc.Rows(intRow).Item("countryName")

                '* Get year country was established
                '** Countries formed prior to 2001 coded as 0
                intScYearEstablished = dtSc.Rows(intRow).Item("countryStartYear")

                '* Replace ' with '' to avoid SQL syntax error
                strScName = strScName.Replace("'", "''")
                strScName = strScName.Trim

                '* Visa refusal rate: Fetch sending country from dataset
                booScFoundInVisaRefRate = False
                drc = dtVisaRefRate.Select("CountryName = '" & strScName & "'")
                If drc.GetUpperBound(0) >= 0 Then
                    booScFoundInVisaRefRate = True
                End If

                '* Visa refusal rate: If country data could not be found using the main name, fetch the variations and search the import data for matches
                If Not booScFoundInVisaRefRate Then
                    drc2 = dtAltCountryNames.Select("countryName = '" & strScName & "'")
                    For i = 0 To drc2.GetUpperBound(0)
                        drc = dtVisaRefRate.Select("CountryName = '" & drc2(i)("alternativeName").ToString.Replace("'", "''") & "'")
                        If drc.GetUpperBound(0) >= 0 Then
                            booScFoundInVisaRefRate = True
                            Exit For
                        End If
                    Next
                End If

                '* Visa refusal rate: If country found in visa practice table then fetch refusal rate
                If booScFoundInVisaRefRate Then

                    '* Get visa refusal rate
                    strVisaRefusalRate = drc(0)("AdjustedRefusalRate")
                    strVisaRefusalRate = strVisaRefusalRate.Replace("%", "")
                    strVisaRefusalRate = strVisaRefusalRate.Trim

                Else

                    '* Warn user if sending country was established in the year but was not found in database
                    If dYear >= intScYearEstablished Then
                        strResult &= "Warning: Sending country not found in visa refusal rate file: " & strScName & "<br>"
                    End If

                End If

                '* Visa issued: Fetch sending country from dataset
                booScFoundInVisaIssued = False
                drc = dtVisaIssued.Select("CountryName = '" & strScName & "'")
                If drc.GetUpperBound(0) >= 0 Then
                    booScFoundInVisaIssued = True
                End If

                '* Visa issued: If country data could not be found using the main name, fetch the variations and search the import data for matches
                If Not booScFoundInVisaIssued Then
                    drc2 = dtAltCountryNames.Select("countryName = '" & strScName & "'")
                    For i = 0 To drc2.GetUpperBound(0)
                        drc = dtVisaIssued.Select("CountryName = '" & drc2(i)("alternativeName").ToString.Replace("'", "''") & "'")
                        If drc.GetUpperBound(0) >= 0 Then
                            booScFoundInVisaIssued = True
                            Exit For
                        End If
                    Next
                End If

                '* Visa issued: If country found in visa practice table then fetch refusal rate
                If booScFoundInVisaIssued Then

                    '* Get visas issued
                    strTypeB1issued = drc(0)("B-1")
                    strTypeB2issued = drc(0)("B-2")
                    strTypeB1comb2issued = drc(0)("B-1,2")
                    strVisaIssued = drc(0)("B-1") + drc(0)("B-1,2") + drc(0)("B-2")
                    strTypeB1comb2BCCissued = drc(0)("B-1,2/BCC")
                    If dYear >= 2007 Then strTypeB1comb2BCVissued = drc(0)("B-1,2/BCV")

                Else

                    '* Warn user if sending country was established in the year but was not found in database
                    If dYear >= intScYearEstablished Then
                        strResult &= "Warning: Sending country not found in visa issued file: " & strScName & "<br>"
                    End If

                End If

                '* Fix: Serbia: Get separate visa issued and refusal rate for "Serbia and Montenegro" and include in calculation
                If dYear >= 2008 And strScName = "Serbia" Then

                    '* Get visas issued
                    drc = dtVisaIssued.Select("CountryName = 'Serbia and Montenegro'")
                    strVisaIssued2 = drc(0)("B-1") + drc(0)("B-1,2") + drc(0)("B-2")
                    strTypeB1issued = CType(strTypeB1issued, Integer) + drc(0)("B-1")
                    strTypeB2issued = CType(strTypeB2issued, Integer) + drc(0)("B-2")
                    strTypeB1comb2issued = CType(strTypeB1comb2issued, Integer) + drc(0)("B-1,2")

                    '* Get visa refusal rate
                    drc = dtVisaRefRate.Select("CountryName = 'Serbia and Montenegro'")
                    strVisaRefusalRate2 = drc(0)("AdjustedRefusalRate").ToString.Replace("%", "").Trim

                    '* Calculate number of visas applied for in total based on visas issued and the refusal rate
                    decVisaAppliedFor = CType(strVisaIssued, Integer) / (1 - CType(strVisaRefusalRate.Replace(".", ","), Decimal) / 100)
                    decVisaAppliedFor += CType(strVisaIssued2, Integer) / (1 - CType(strVisaRefusalRate2.Replace(".", ","), Decimal) / 100)

                    '* Calculate number of visas refused in total
                    decVisaRefused = decVisaAppliedFor - CType(strVisaIssued, Integer) - CType(strVisaIssued2, Integer)

                    '* Re-calculate refusal rate and visas issued
                    strVisaIssued = CType(strVisaIssued, Integer) + CType(strVisaIssued2, Integer)
                    strVisaRefusalRate = Math.Round(decVisaRefused / decVisaAppliedFor * 100, 1)

                End If

                '* Fix: Sudan: Get separate visa issued and refusal rate for "South Sudan" and include in calculation
                If dYear = 2011 And strScName = "Sudan" Then

                    '* Get visas issued
                    drc = dtVisaIssued.Select("CountryName = 'South Sudan'")
                    strVisaIssued2 = drc(0)("B-1") + drc(0)("B-1,2") + drc(0)("B-2")
                    strTypeB1issued = CType(strTypeB1issued, Integer) + drc(0)("B-1")
                    strTypeB2issued = CType(strTypeB2issued, Integer) + drc(0)("B-2")
                    strTypeB1comb2issued = CType(strTypeB1comb2issued, Integer) + drc(0)("B-1,2")

                    '* Get visa refusal rate
                    drc = dtVisaRefRate.Select("CountryName = 'South Sudan'")
                    strVisaRefusalRate2 = drc(0)("AdjustedRefusalRate").ToString.Replace("%", "").Trim

                    '* Calculate number of visas applied for in total based on visas issued and the refusal rate
                    decVisaAppliedFor = CType(strVisaIssued, Integer) / (1 - CType(strVisaRefusalRate.Replace(".", ","), Decimal) / 100)
                    decVisaAppliedFor += CType(strVisaIssued2, Integer) / (1 - CType(strVisaRefusalRate2.Replace(".", ","), Decimal) / 100)

                    '* Calculate number of visas refused in total
                    decVisaRefused = decVisaAppliedFor - CType(strVisaIssued, Integer) - CType(strVisaIssued2, Integer)

                    '* Re-calculate refusal rate and visas issued
                    strVisaIssued = CType(strVisaIssued, Integer) + CType(strVisaIssued2, Integer)
                    strVisaRefusalRate = Math.Round(decVisaRefused / decVisaAppliedFor * 100, 1)

                End If

                '* If country was established in the year, then add data
                If dYear >= intScYearEstablished Then

                    '* Add visa information to database
                    SQL = "IF EXISTS (SELECT visaPracticeUsID FROM visaPractice_US WHERE rcID = " & intRcID & " AND scID = " & intScID & " AND dYear = " & dYear & ") " _
                        & " UPDATE visaPractice_US SET " _
                        & " shortStayRefusalRate = " & strVisaRefusalRate.Replace(",", ".") _
                        & ", shortStayIssued     = " & strVisaIssued.Replace(",", ".") _
                        & ", typeB1issued        = " & strTypeB1issued.Replace(",", ".") _
                        & ", typeB2issued        = " & strTypeB2issued.Replace(",", ".") _
                        & ", typeB1comb2issued   = " & strTypeB1comb2issued.Replace(",", ".") _
                        & ", typeB1comb2BCCissued= " & strTypeB1comb2BCCissued.Replace(",", ".") _
                        & ", typeB1comb2BCVissued= " & strTypeB1comb2BCVissued.Replace(",", ".") _
                        & " WHERE rcID = " & intRcID & " AND scID = " & intScID & " AND dYear = " & dYear _
                        & " ELSE " _
                        & " INSERT INTO visaPractice_US (rcID,scID,dYear,shortStayRefusalRate,shortStayIssued,typeB1issued,typeB2issued,typeB1comb2issued,typeB1comb2BCCissued,typeB1comb2BCVissued) VALUES " _
                        & " (" & intRcID _
                        & "," & intScID _
                        & "," & dYear _
                        & "," & strVisaRefusalRate.Replace(",", ".") _
                        & "," & strVisaIssued.Replace(",", ".") _
                        & "," & strTypeB1issued.Replace(",", ".") _
                        & "," & strTypeB2issued.Replace(",", ".") _
                        & "," & strTypeB1comb2issued.Replace(",", ".") _
                        & "," & strTypeB1comb2BCCissued.Replace(",", ".") _
                        & "," & strTypeB1comb2BCVissued.Replace(",", ".") _
                        & ");"
                    'strResult &= SQL & "<br>"
                    cmd = New OleDbCommand(SQL, db)
                    cmd.ExecuteNonQuery()
                    intRecordsProcessed += 1

                End If

            Next

            '* Set negative refusal rates to 0
            SQL = "UPDATE visaPractice_US SET " _
                & " shortStayRefusalRate = 0 " _
                & " WHERE shortStayRefusalRate < 0 " _
                & "   AND dYear = " & dYear _
                & ";"
            cmd = New OleDbCommand(SQL, db)
            cmd.ExecuteNonQuery()

            '* Calculate number of visas applied for in the year
            SQL = "UPDATE visaPractice_US SET " _
                & " shortStayAppliedFor = shortStayIssued / (1 - shortStayRefusalRate/100) " _
                & " WHERE shortStayRefusalRate <> 100 " _
                & "   AND dYear = " & dYear _
                & ";"
            cmd = New OleDbCommand(SQL, db)
            cmd.ExecuteNonQuery()

            '* Calculate number of visas refused in the year
            SQL = "UPDATE visaPractice_US SET " _
                & " shortStayRefused = shortStayAppliedFor - shortStayIssued " _
                & " WHERE dYear = " & dYear _
                & ";"
            cmd = New OleDbCommand(SQL, db)
            cmd.ExecuteNonQuery()

        Catch ex As Exception

            '* Write error
            strResult &= "Error: " & ex.Message

        Finally

            '* Close database
            db.Close()
            db.Dispose()

            '* Close excel visa practice sheets
            xls.Close()
            xls.Dispose()
            xls2.Close()
            xls2.Dispose()

            '* Close excel conversion table
            xls3.Close()
            xls3.Dispose()

        End Try

        '* Get status on number of rows inserted
        strResult &= "Rows inserted: " & intRowsInserted & "<br>"
        strResult &= "Rows processed: " & intRecordsProcessed & "<br>"

        '* Return result of processing script
        Return strResult

    End Function

    '* Import visa practice data for the European Union (Schengen)
    Public Shared Function importVisaPracticeEU(ByVal dYear As Integer) As String

        '* Define database variables
        Dim SQL As String
        Dim db As New OleDbConnection(System.Web.Configuration.WebConfigurationManager.ConnectionStrings("visa").ConnectionString)
        Dim cmd As OleDbCommand

        '* Define excel variables
        Dim xls As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & HttpContext.Current.Server.MapPath("DataImport\conversionTable.xls") & ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1;ImportMixedTypes=Text""")
        Dim importFile As FileInfo

        '* Define other variables
        Dim intRowsInserted As Integer
        Dim intRecordsProcessed As Integer
        Dim strResult As String = ""
        Dim dtCities As DataTable
        Dim dtAltCityNames As DataTable
        Dim dtAltCountryNames As DataTable
        Dim intRow As Integer
        Dim intCol As Integer
        Dim intSheet As Integer
        Dim drc() As DataRow
        Dim intRcID As Integer
        Dim strRcName As String
        Dim strCityName As String
        Dim intCityID As Integer
        Dim strCountryName As String
        Dim strCountryCode As String
        Dim intScID As Integer
        Dim strScNameAndCity As String
        Dim strScName As String
        Dim strScCity As String
        Dim strStatistic As String
        Dim strIssuedA As String
        Dim strIssuedA_All As String
        Dim strIssuedA_Mev As String
        Dim strIssuedB As String
        Dim strIssuedC As String
        Dim strIssuedC_All As String
        Dim strIssuedC_Mev As String
        Dim strIssuedD As String
        Dim strIssuedDC As String
        Dim strIssuedVTL As String
        Dim strIssuedADS As String
        Dim strIssuedABC As String
        Dim strIssuedABCD As String
        Dim strIssuedABCDVTL As String
        Dim strIssuedABCDDCVTL As String
        Dim strIssuedReprOther As String
        Dim strApplied As String
        Dim strAppliedA As String
        Dim strAppliedB As String
        Dim strAppliedC As String
        Dim strAppliedABC As String
        Dim strNotIssued As String
        Dim strNotIssuedA As String
        Dim strNotIssuedB As String
        Dim strNotIssuedC As String
        Dim strNotIssuedABC As String
        Dim strRefusalRate As String
        Dim booImportRow As Boolean
        Dim strCityID As String

        Try

            '* Open database
            db.Open()

            '* Open excelsheet with variation in country names
            xls.Open()

            '* Get database table with all countries/cities
            SQL = "SELECT countries_cities.countryCityID, countries.countryID, countries.countryName, countries.countryCode, cities.cityID, cities.cityName, dbo.ufn_getCountryStartYear(countryName) AS countryStartYear " _
                & " FROM  countries_cities INNER JOIN " _
                & "       cities ON countries_cities.cityID = cities.cityID INNER JOIN " _
                & "       countries ON countries_cities.countryID = countries.countryID;"
            cmd = New OleDbCommand(SQL, db)
            dtCities = New DataTable : dtCities.Load(cmd.ExecuteReader)

            '* Open excel sheet with alternative/variations in city names
            cmd = New OleDbCommand("SELECT * FROM [cities$]", xls)
            dtAltCityNames = New DataTable : dtAltCityNames.Load(cmd.ExecuteReader)

            '* Open excel sheet with alternative/variations in country names
            cmd = New OleDbCommand("SELECT * FROM [countries$]", xls)
            dtAltCountryNames = New DataTable : dtAltCountryNames.Load(cmd.ExecuteReader)

            '* Open excel file with information on visa practice
            importFile = New FileInfo(HttpContext.Current.Server.MapPath("DataImport\VisaPractice\EU\EU_Visa_ShortStay_Practice_" & dYear & ".xlsx"))
            Using xlPackage As New ExcelPackage(importFile)

                '* 2002/2003/2004 import rutine
                If dYear = 2002 Or dYear = 2003 Or dYear = 2004 Then

                    '* Get worksheet
                    Dim worksheet1 As ExcelWorksheet = xlPackage.Workbook.Worksheets(1)

                    '* Loop through each row in worksheet; skip header
                    For intRow = 2 To worksheet1.Dimension.End.Row

                        '* Get receiving country

                        '* Reset receiving country id
                        intRcID = 0

                        '* Fetch receiving country name
                        If worksheet1.Cells(intRow, 1).IsRichText Then
                            strRcName = worksheet1.Cells(intRow, 1).RichText.Text
                        Else
                            strRcName = worksheet1.Cells(intRow, 1).Value
                        End If

                        '* Try to fetch country from database
                        drc = dtCities.Select("countryName = '" & strRcName & "'")
                        If drc.GetUpperBound(0) >= 0 Then
                            intRcID = drc(0)("countryID")
                        End If

                        '* Warn user if receiving country was not found in database
                        If intRcID = 0 Then
                            strResult &= "Error: Receiving country not found in database: " & strRcName & " (" & intRow & ")<br>"
                        End If

                        '* Get sending country

                        '* Reset variables
                        intScID = 0 : intCityID = 0 : strScCity = "" : strScName = "" : strScNameAndCity = ""

                        If Not worksheet1.Cells(intRow, 2).Value Is Nothing Then

                            '* Get sending country name
                            If worksheet1.Cells(intRow, 2).IsRichText Then
                                strScName = worksheet1.Cells(intRow, 2).RichText.Text
                            Else
                                strScName = worksheet1.Cells(intRow, 2).Value
                            End If

                            '* Get sending country city name
                            If worksheet1.Cells(intRow, 3).IsRichText Then
                                strScCity = worksheet1.Cells(intRow, 3).RichText.Text
                            Else
                                strScCity = worksheet1.Cells(intRow, 3).Value
                            End If

                        Else

                            '* Get sending country name and city as contained in column c 
                            If worksheet1.Cells(intRow, 3).IsRichText Then
                                strScNameAndCity = worksheet1.Cells(intRow, 3).RichText.Text
                            Else
                                strScNameAndCity = worksheet1.Cells(intRow, 3).Value
                            End If

                            '* Split data into country and city
                            If strScNameAndCity.Contains("-") Then
                                If strScNameAndCity.Contains("GUINEA-BUSSAU") Or strScNameAndCity.Contains("Guinea-Bissau") Then
                                    strScName = "Guinea-Bissau"
                                ElseIf strScNameAndCity.Contains("SERBIA-MONTENEGRO") Then
                                    strScName = "Serbia"
                                ElseIf strScNameAndCity.Contains("BOSNIA-HERZEGOVINA") Then
                                    strScName = "Bosnia-Herzegovina"
                                Else
                                    strScName = Left(strScNameAndCity, InStr(strScNameAndCity, "-") - 1).Trim
                                End If
                                strScCity = Right(strScNameAndCity, Len(strScNameAndCity) - InStr(strScNameAndCity, "-")).Trim
                            Else
                                strScName = strScNameAndCity
                                strScCity = strScNameAndCity
                            End If

                        End If

                        '* Replace ' with '' to avoid SQL syntax error
                        strScName = strScName.Replace("'", "''")
                        strScCity = strScCity.Replace("'", "''")

                        '* Fix for Hamilton, Bermuda (part of the UK)
                        If strScName = "BERMUDA" And strScCity = "HAMILTON" Then
                            strScName = "United Kingdom"
                            strScCity = "Hamilton (Bermuda)"
                        End If

                        '* Fix for ARUBA, Netherlands
                        If strScName = "ARUBA" Then
                            strScName = "Netherlands"
                            strScCity = "Aruba"
                        End If

                        '* Fix for NETHERLANDS ANTILLES
                        If strScName = "NETHERLANDS ANTILLES" Or strScName = "NEDERLANDS ANTILLEN" Then
                            strScName = "Netherlands"
                            strScCity = "Willemstad (Curacao)"
                        End If

                        '* Fix for various sending country typing/OCR recognition errors
                        If strScName = "SERBIA-MONTEN" Then strScName = "Serbia and montenegro"
                        If strScName = "UNITED ARAB EMI" Then strScName = "United Arab Emirates"
                        If strScName = "CAP VERDE" Then strScName = "CAPE VERDE"
                        If strScName = "AFGANISTAN" Then strScName = "AFGHANISTAN"
                        If strScName = "FRAN CIA" Then strScName = "France"
                        If strScName = "ITALA" Then strScName = "Italy"
                        If strScName = "JORDAN IA" Then strScName = "Jordan"
                        If strScName = "I CRANIA" Then strScName = "Ukraine"
                        If strScName = "M AI ASI A" Then strScName = "Malaysia"
                        If strScName = "JERUSALEN" Then strScName = "Palestinian Authority" '* ES entry for Jerusalem
                        If strScName = "BOSNIA-HERZOGOWINA" Then strScName = "Bosnia and Herzegovina"
                        If strScName = "COTE DTVOIFcE" Then strScName = "Côte d''Ivoire"
                        If strScName = "RUS SIA" Then strScName = "Russia"
                        If strScName = "SCOTTLAND" Then strScName = "United Kingdom" '* NO entry for Edinburgh listed as in Scotland
                        If strScName = "SL OVENIA" Then strScName = "SLOVENIA"
                        If strScName = "BULGARIS" Then strScName = "Bulgaria"
                        If strScName = "CHINA (TAIWAN)" Then strScName = "Taiwan"
                        If strScName = "KYRG YZSTAN" Then strScName = "KYRGYZSTAN"
                        If strScName = "LUXEMBURG O" Then strScName = "LUXEMBURGO"
                        If strScName = "SLID AFRICA" Then strScName = "South Africa"
                        If strScName = "SOMALI A" Then strScName = "SOMALIA"
                        If strScName = "LET LAND" Then strScName = "LETLAND"
                        If strScName = "SLOVEN IE" Then strScName = "SLOVENIE"
                        If strScName = "SYR IE" Then strScName = "SYRIE"
                        If strScName = "PHILD7EVES" Then strScName = "PHILIPINES"
                        If strScName = "PHILD7ENES" Then strScName = "PHILIPINES"
                        If strScName = "PHILD7PENES" Then strScName = "PHILIPINES"
                        If strScName = "SAO TOME ET PREVCD7E" Then strScName = "SAO TOME ET PRINCIPE"
                        If strScName = "SAO TOME ET PRENCD7E" Then strScName = "SAO TOME ET PRINCIPE"
                        If strScName = "SAO TOME  E PRINCIPE" Then strScName = "SAO TOME ET PRINCIPE"
                        If strScName = "JUGOSLAVIA  (S ERVIA/MO NT EN EGRO)" Then strScName = "Serbia"
                        If strScName = "SEPUBLICA DEMOCRATIC^ DO  CONGO" Then strScName = "Congo (Democratic Republic of)"
                        If strScName = "ALEMAHHA" Then strScName = "Germany"
                        If strScName = "B SNIA" Then strScName = "Bosnia"
                        If strScName = "COL MBIA" Then strScName = "COLOMBIA"
                        If strScName = "tlARROCOS" Then strScName = "MOROCCO"
                        If strScName = "ilO''AMBIQUE" Then strScName = "Mozambique"
                        If strScName = "2UENIA" Then strScName = "Kenya"
                        If strScName = "TIMOR  LESTE" Then strScName = "Timor-Leste"

                        '* Try to fetch sending country from database
                        drc = dtCities.Select("countryName = '" & strScName & "' AND " & dYear & " >= countryStartYear")
                        If drc.GetUpperBound(0) >= 0 Then
                            intScID = drc(0)("countryID")
                        End If

                        '* If sending country was not found in database table check for name variations in the conversion table
                        If intScID = 0 Then
                            drc = dtAltCountryNames.Select("AlternativeName = '" & strScName & "'")
                            If drc.GetUpperBound(0) >= 0 Then
                                drc = dtCities.Select("countryName = '" & drc(0)("countryName").ToString.Replace("'", "''") & "' AND " & dYear & " >= countryStartYear")
                                If drc.GetUpperBound(0) >= 0 Then
                                    intScID = drc(0)("countryID")
                                End If
                            End If
                        End If

                        '* Try to fetch cityID from database
                        drc = dtCities.Select("countryID = " & intScID & " AND cityName = '" & strScCity & "' AND " & dYear & " >= countryStartYear")
                        If drc.GetUpperBound(0) >= 0 Then
                            intCityID = drc(0)("countryCityID")
                        End If

                        '* If city was not found in database table check for name variations in the conversion table
                        If intCityID = 0 Then
                            drc = dtAltCityNames.Select("AlternativeCityName = '" & strScCity & "'")
                            If drc.GetUpperBound(0) >= 0 Then
                                drc = dtCities.Select("countryID = " & intScID & " AND cityName = '" & drc(0)("cityName").ToString.Replace("'", "''") & "' AND " & dYear & " >= countryStartYear")
                                If drc.GetUpperBound(0) >= 0 Then
                                    intCityID = drc(0)("countryCityID")
                                End If
                            End If
                        End If

                        '* Warn user if country was not found in database
                        If intScID = 0 Then strResult &= "Error: Sending country not found in database: " & strScName & " (" & intRow & ")<br>"
                        'If intScID = 0 Then strResult &= strScName & "<br>"

                        '* Warn user if city was not found in database / FOR 2002+2003+2004: Do not code cities missing in database; only code entry at country level
                        'If intCityID = 0 Then strResult &= "Error: City not found in database: " & strScCity & ", " & strScName & " (" & intRow & ")<br>"

                        If intScID <> 0 Then

                            '* Reset data
                            strIssuedA = "NULL" : strIssuedB = "NULL" : strIssuedC = "NULL" : strIssuedVTL = "NULL" : strIssuedABC = "NULL" : strIssuedD = "NULL" : strIssuedDC = "NULL"
                            strIssuedABCD = "NULL" : strNotIssued = "NULL" : strApplied = "NULL" : strRefusalRate = "NULL" : strIssuedReprOther = "NULL" : booImportRow = False

                            '* Get information on visa practice
                            For intCol = 4 To 15
                                If Not worksheet1.Cells(intRow, intCol).Value Is Nothing Then

                                    '* Get data from field; remove blank spaces; remove ".";taken into account rich text or not
                                    If worksheet1.Cells(intRow, intCol).IsRichText Then
                                        strStatistic = worksheet1.Cells(intRow, intCol).RichText.Text.ToString.Replace(" ", "")
                                    Else
                                        strStatistic = worksheet1.Cells(intRow, intCol).Value.ToString.Replace(" ", "")
                                    End If

                                    '* DK: - interpreted as 0
                                    If strRcName = "Denmark" And strStatistic = "-" Then strStatistic = "0"

                                    '* Lithuania: Adjust data in refusal rate column
                                    If strRcName = "Lithuania" And intCol = 14 Then strStatistic = strStatistic.Replace(".", ",").Replace("%", "")

                                    '* Recode percentage values
                                    If intCol = 14 And IsNumeric(strStatistic) Then
                                        If worksheet1.Cells(intRow, intCol).Style.Numberformat.NumFmtID = 9 _
                                            Or worksheet1.Cells(intRow, intCol).Style.Numberformat.NumFmtID = 10 Then '* Column type is percentage
                                            strStatistic = strStatistic * 100
                                        End If
                                        If strRcName = "Czech Republic" Then
                                            strStatistic = strStatistic * 100
                                        End If
                                        If strRcName = "Netherlands" Then
                                            strStatistic = strStatistic / 100
                                        End If
                                    End If

                                    '* Set relevant statistic to fetched value
                                    If IsNumeric(strStatistic) Then
                                        If intCol = 4 Then strIssuedA = strStatistic
                                        If intCol = 5 Then strIssuedB = strStatistic
                                        If intCol = 6 Then strIssuedC = strStatistic
                                        If intCol = 7 Then strIssuedVTL = strStatistic
                                        If intCol = 8 Then strIssuedABC = strStatistic
                                        If intCol = 9 Then strIssuedD = strStatistic
                                        If intCol = 10 Then strIssuedDC = strStatistic
                                        If intCol = 11 Then strIssuedABCD = strStatistic
                                        If intCol = 12 Then strNotIssued = strStatistic
                                        If intCol = 13 Then strApplied = strStatistic
                                        If intCol = 14 Then strRefusalRate = strStatistic.Replace(",", ".")
                                        If intCol = 15 Then strIssuedReprOther = strStatistic
                                        booImportRow = True
                                    Else
                                        If dYear = 2004 Then
                                            If (strRcName = "Italy" And strStatistic = "n.d") _
                                                Or (strRcName = "Austria" And (strStatistic = "-" Or strStatistic = "N.A.")) _
                                                Or (strRcName = "Lithuania" And strStatistic = "-") Then
                                                '* Do not print error in these cases
                                            Else
                                                strResult &= "Warning: Visa statistic not numeric (row " & intRow & ", col " & intCol & ")<br>"
                                            End If
                                        End If
                                        If dYear = 2003 Then
                                            If (strRcName = "Italy" And (strStatistic = "n.d" Or strStatistic = "n.d.")) _
                                                Or (strRcName = "Austria" And (strStatistic = "..." Or strStatistic = "n.a.")) _
                                                Then
                                                '* Do not print error in these cases
                                            Else
                                                strResult &= "Warning: Visa statistic not numeric (row " & intRow & ", col " & intCol & ")<br>"
                                            End If
                                        End If
                                        If dYear = 2002 Then
                                            If (strRcName = "Austria" And (strStatistic = "n.a.")) Then
                                                '* Do not print error in these cases
                                            Else
                                                strResult &= "Warning: Visa statistic not numeric (row " & intRow & ", col " & intCol & ")<br>"
                                            End If
                                        End If
                                    End If

                                End If

                            Next

                            '* Import row, if data found
                            If booImportRow Then

                                '* Parse city id
                                If intCityID = 0 Then
                                    strCityID = "NULL"
                                    strScCity = "'" & strScCity & "'"
                                Else
                                    strCityID = intCityID
                                    strScCity = "NULL"
                                End If

                                '* Compile SQL syntax
                                SQL = "INSERT INTO visaPractice_EU_" & dYear & " (rcID,scCountryID,scCityID,cityName,dYear,AvisasIssued,BvisasIssued,CvisasIssued,VTLvisasIssued,TotalABCIssued,DvisasIssued,DCvisasIssued,TotalABCDIssued,visasNotIssued,visasApplied,visaRefusalRate,issuedReprOther) VALUES " _
                                    & " (" & intRcID _
                                    & "," & intScID _
                                    & "," & strCityID _
                                    & "," & strScCity _
                                    & "," & dYear _
                                    & "," & strIssuedA _
                                    & "," & strIssuedB _
                                    & "," & strIssuedC _
                                    & "," & strIssuedVTL _
                                    & "," & strIssuedABC _
                                    & "," & strIssuedD _
                                    & "," & strIssuedDC _
                                    & "," & strIssuedABCD _
                                    & "," & strNotIssued _
                                    & "," & strApplied _
                                    & "," & strRefusalRate _
                                    & "," & strIssuedReprOther _
                                    & ");"
                                'strResult &= SQL & "<br>"
                                cmd = New OleDbCommand(SQL, db)
                                intRowsInserted += cmd.ExecuteNonQuery()

                            End If

                        End If

                    Next

                End If

                '* 2005 import rutine
                If dYear = 2005 Then

                    '* Get worksheet
                    Dim worksheet1 As ExcelWorksheet = xlPackage.Workbook.Worksheets(1)

                    '* Loop through each row in worksheet
                    For intRow = 1 To worksheet1.Dimension.End.Row

                        '* Only process rows with content
                        If Not worksheet1.Cells(intRow, 1) Is Nothing Then
                            If worksheet1.Cells(intRow, 1).Value <> "" Then

                                If InStr(worksheet1.Cells(intRow, 1).Value.ToString.Replace(" ", ""), "01/01/" & dYear) Then

                                    '* Get receiving country

                                    '* Reset receiving country id
                                    intRcID = 0

                                    '* Fetch receiving country name
                                    If worksheet1.Cells(intRow, 1).IsRichText Then
                                        strRcName = worksheet1.Cells(intRow, 1).RichText.Text
                                    Else
                                        strRcName = worksheet1.Cells(intRow, 1).Value
                                    End If

                                    '* Remove non-country name characters from string
                                    strCountryCode = Left(strRcName.Replace(" ", "").Replace("31/12/" & dYear, "").Replace("01/01/" & dYear, "").Replace("V 1.1", "").Trim, 2)

                                    '* Fix for situations where country code is contained in next column
                                    If strCountryCode = "" Then
                                        If worksheet1.Cells(intRow, 2).IsRichText Then
                                            strRcName = worksheet1.Cells(intRow, 2).RichText.Text
                                        Else
                                            strRcName = worksheet1.Cells(intRow, 2).Value
                                        End If
                                        strCountryCode = Left(strRcName, 2)
                                    End If

                                    '* Fix for Finland
                                    If strCountryCode = "Fl" Then strCountryCode = "FI"

                                    '* Try to fetch country from database
                                    drc = dtCities.Select("countryCode = '" & strCountryCode & "'")
                                    If drc.GetUpperBound(0) >= 0 Then
                                        intRcID = drc(0)("countryID")
                                    End If

                                    '* Warn user if receiving country was not found in database
                                    If intRcID = 0 Then
                                        strResult &= "Error: Receiving country not found in database: " & strRcName & " (" & intRow & ")<br>"
                                    End If

                                Else

                                    '* Import row with data on practice in sending country
                                    If Not worksheet1.Cells(intRow, 1).Value.ToString.Contains("TOTAL") _
                                        And Not worksheet1.Cells(intRow, 1).Value.ToString.Contains("Minimum") _
                                        And Not worksheet1.Cells(intRow, 1).Value.ToString.Contains("Maximum") _
                                        And Not worksheet1.Cells(intRow, 1).Value.ToString.Contains("Average") _
                                        And Not worksheet1.Cells(intRow, 1).Value.ToString.Contains("cities") _
                                        And Not worksheet1.Cells(intRow, 1).Value.ToString.Contains("2005") _
                                        And Not worksheet1.Cells(intRow, 1).Value.ToString.Contains("1.1") _
                                        And Not worksheet1.Cells(intRow, 1).Value = "EE" _
                                        And Not worksheet1.Cells(intRow, 1).Value = "HU" _
                                        Then

                                        '* Get sending country
                                        '* Reset variables
                                        intScID = 0 : intCityID = 0 : strScCity = "" : strScName = "" : strScNameAndCity = ""

                                        '* Get sending country name
                                        If worksheet1.Cells(intRow, 1).IsRichText Then
                                            strScName = worksheet1.Cells(intRow, 1).RichText.Text
                                        Else
                                            strScName = worksheet1.Cells(intRow, 1).Value
                                        End If

                                        '* Get sending country city name
                                        If worksheet1.Cells(intRow, 2).IsRichText Then
                                            strScCity = worksheet1.Cells(intRow, 2).RichText.Text
                                        Else
                                            strScCity = worksheet1.Cells(intRow, 2).Value
                                        End If

                                        '* Replace ' with '' to avoid SQL syntax error
                                        strScName = strScName.Replace("'", "''")
                                        strScCity = strScCity.Replace("'", "''")

                                        '* Fix for Hamilton, Bermuda (part of the UK)
                                        If strScName = "BERMUDA" And strScCity = "HAMILTON" Then
                                            strScName = "United Kingdom"
                                            strScCity = "Hamilton (Bermuda)"
                                        End If

                                        '* Try to fetch sending country from database
                                        drc = dtCities.Select("countryName = '" & strScName & "' AND " & dYear & " >= countryStartYear")
                                        If drc.GetUpperBound(0) >= 0 Then
                                            intScID = drc(0)("countryID")
                                        End If

                                        '* If sending country was not found in database table check for name variations in the conversion table
                                        If intScID = 0 Then
                                            drc = dtAltCountryNames.Select("AlternativeName = '" & strScName & "'")
                                            If drc.GetUpperBound(0) >= 0 Then
                                                drc = dtCities.Select("countryName = '" & drc(0)("countryName").ToString.Replace("'", "''") & "' AND " & dYear & " >= countryStartYear")
                                                If drc.GetUpperBound(0) >= 0 Then
                                                    intScID = drc(0)("countryID")
                                                End If
                                            End If
                                        End If

                                        '* Try to fetch cityID from database
                                        drc = dtCities.Select("countryID = " & intScID & " AND cityName = '" & strScCity & "' AND " & dYear & " >= countryStartYear")
                                        If drc.GetUpperBound(0) >= 0 Then
                                            intCityID = drc(0)("countryCityID")
                                        End If

                                        '* If city was not found in database table check for name variations in the conversion table
                                        If intCityID = 0 Then
                                            drc = dtAltCityNames.Select("AlternativeCityName = '" & strScCity & "'")
                                            If drc.GetUpperBound(0) >= 0 Then
                                                drc = dtCities.Select("countryID = " & intScID & " AND cityName = '" & drc(0)("cityName").ToString.Replace("'", "''") & "' AND " & dYear & " >= countryStartYear")
                                                If drc.GetUpperBound(0) >= 0 Then
                                                    intCityID = drc(0)("countryCityID")
                                                End If
                                            End If
                                        End If

                                        '* Reset information on visa practice
                                        strIssuedA_All = "NULL" : strIssuedB = "NULL" : strIssuedC_All = "NULL" : strAppliedC = "NULL"
                                        strIssuedABC = "NULL" : strAppliedABC = "NULL" : strNotIssuedABC = "NULL" : strIssuedVTL = "NULL"
                                        strIssuedD = "NULL" : strIssuedDC = "NULL" : strIssuedABCDDCVTL = "NULL" : strIssuedADS = "NULL"

                                        '* Get information on visa practice
                                        For intCol = 3 To 15
                                            If Not worksheet1.Cells(intRow, intCol).Value Is Nothing Then

                                                '* Get data from field; remove blank spaces; remove ".";taken into account rich text or not
                                                If worksheet1.Cells(intRow, intCol).IsRichText Then
                                                    strStatistic = worksheet1.Cells(intRow, intCol).RichText.Text.ToString.Replace(".", "").Replace(" ", "")
                                                Else
                                                    strStatistic = worksheet1.Cells(intRow, intCol).Value.ToString.Replace(".", "").Replace(" ", "")
                                                End If

                                                '* Set relevant statistic to fetched value
                                                If IsNumeric(strStatistic) Then
                                                    If intCol = 3 Then strIssuedA_All = strStatistic
                                                    If intCol = 4 Then strIssuedB = strStatistic
                                                    If intCol = 5 Then strIssuedC_All = strStatistic
                                                    If intCol = 6 Then strAppliedC = strStatistic
                                                    If intCol = 7 Then strIssuedADS = strStatistic
                                                    If intCol = 8 Then strIssuedABC = strStatistic
                                                    If intCol = 9 Then strAppliedABC = strStatistic
                                                    If intCol = 10 Then strNotIssuedABC = strStatistic
                                                    If intCol = 12 Then strIssuedVTL = strStatistic
                                                    If intCol = 13 Then strIssuedD = strStatistic
                                                    If intCol = 14 Then strIssuedDC = strStatistic
                                                    If intCol = 15 Then strIssuedABCDDCVTL = strStatistic
                                                Else
                                                    '* Do not print error for column as it contains the percentage refused
                                                    If intCol <> 11 Then strResult &= "Warning: Visa statistic not numeric (row " & intRow & ", col " & intCol & ")<br>"
                                                End If
                                            End If
                                        Next

                                        '* Parse to database, if row contains visa data
                                        If strIssuedA_All <> "NULL" Or strIssuedB <> "NULL" Or strIssuedC_All <> "NULL" Or strAppliedC <> "NULL" Or strIssuedADS <> "NULL" _
                                            Or strAppliedABC <> "NULL" Or strNotIssuedABC <> "NULL" Or strIssuedVTL <> "NULL" Or strIssuedD <> "NULL" Or strIssuedDC <> "NULL" _
                                            Or strIssuedABCDDCVTL <> "NULL" Then

                                            '* Warn user if country was not found in database
                                            If intScID = 0 Then strResult &= "Error: Sending country not found in database: " & strScName & " (" & intRow & ")<br>"

                                            '* Warn user if city was not found in database
                                            If intCityID = 0 Then strResult &= "Error: City not found in database: " & strScCity & ", " & strScName & " (" & intRow & ")<br>"

                                            '* If diplomatic post found in database, insert data
                                            If intCityID <> 0 Then

                                                SQL = "INSERT INTO visaPractice_EU (rcID,scCityID,dYear,issuedA_All,issuedB,issuedC_All,issuedVTL,issuedADS,issuedABC,issuedD,issuedDC,issuedABCDDCVTL,appliedC,appliedABC,notIssuedABC) VALUES " _
                                                    & " (" & intRcID _
                                                    & "," & intCityID _
                                                    & "," & dYear _
                                                    & "," & strIssuedA_All _
                                                    & "," & strIssuedB _
                                                    & "," & strIssuedC_All _
                                                    & "," & strIssuedVTL _
                                                    & "," & strIssuedADS _
                                                    & "," & strIssuedABC _
                                                    & "," & strIssuedD _
                                                    & "," & strIssuedDC _
                                                    & "," & strIssuedABCDDCVTL _
                                                    & "," & strAppliedC _
                                                    & "," & strAppliedABC _
                                                    & "," & strNotIssuedABC _
                                                    & ");"
                                                'strResult &= SQL & "<br>"
                                                cmd = New OleDbCommand(SQL, db)
                                                intRowsInserted += cmd.ExecuteNonQuery()

                                            End If

                                        End If

                                    End If

                                End If

                            End If

                        End If

                    Next

                End If

                '* 2006, 2007 2008, 2009 import rutine
                If dYear = 2006 Or dYear = 2007 Or dYear = 2008 Or dYear = 2009 Then

                    '* Get worksheet
                    Dim worksheet1 As ExcelWorksheet = xlPackage.Workbook.Worksheets(1)

                    '* Loop through each row in worksheet
                    For intRow = 1 To worksheet1.Dimension.End.Row

                        '* Only process rows with content
                        If Not worksheet1.Cells(intRow, 1) Is Nothing Then
                            If worksheet1.Cells(intRow, 1).Value <> "" Then

                                '* Get receiving country name and try to fetch ID from database
                                If InStr(worksheet1.Cells(intRow, 1).Value.ToString.Replace(" ", ""), "31/12/" & dYear) Then

                                    '* Reset receiving country id
                                    intRcID = 0

                                    '* Fetch receiving country name
                                    If worksheet1.Cells(intRow, 1).IsRichText Then
                                        strRcName = worksheet1.Cells(intRow, 1).RichText.Text
                                    Else
                                        strRcName = worksheet1.Cells(intRow, 1).Value
                                    End If

                                    '* Remove non-country name characters from string
                                    strRcName = strRcName.Replace(" ", "").Replace("-31/12/" & dYear, "").Replace("1/01/" & dYear, "").Trim

                                    '* Fix for Czech Republic
                                    If strRcName = "CZECHREPUBLIC" Then strRcName = "CZECH REPUBLIC"

                                    '* Try to fetch receiving country from database
                                    drc = dtCities.Select("countryName = '" & strRcName & "'")
                                    If drc.GetUpperBound(0) >= 0 Then
                                        intRcID = drc(0)("countryID")
                                    End If

                                    '* If receiving country was not found in database table check for name variations in the conversion table
                                    If intRcID = 0 Then
                                        drc = dtAltCountryNames.Select("AlternativeName = '" & strRcName & "'")
                                        If drc.GetUpperBound(0) >= 0 Then
                                            drc = dtCities.Select("countryName = '" & drc(0)("countryName").ToString.Replace("'", "''") & "' AND " & dYear & " >= countryStartYear")
                                            If drc.GetUpperBound(0) >= 0 Then
                                                intRcID = drc(0)("countryID")
                                            End If
                                        End If
                                    End If

                                    '* Warn user if receiving country was not found in database
                                    If intRcID = 0 Then strResult &= "Error: Receiving country not found in database: " & strRcName & " (" & intRow & ")<br>"

                                Else

                                    '* Import row with data on practice in sending country

                                    '* Get sending country name and city

                                    '* Reset variables
                                    intScID = 0 : intCityID = 0 : strScCity = "" : strScName = "" : strScNameAndCity = ""

                                    '* Get country - city name
                                    If worksheet1.Cells(intRow, 1).IsRichText Then
                                        strScNameAndCity = worksheet1.Cells(intRow, 1).RichText.Text
                                    Else
                                        strScNameAndCity = worksheet1.Cells(intRow, 1).Value
                                    End If

                                    '* Replace ' with '' to avoid SQL syntax error
                                    strScNameAndCity = strScNameAndCity.Replace("'", "''")

                                    '* Split country - city into two variables
                                    strScName = Left(strScNameAndCity, InStr(strScNameAndCity, "-") - 2).Trim
                                    strScCity = Right(strScNameAndCity, Len(strScNameAndCity) - InStrRev(strScNameAndCity, "-")).Trim

                                    '* Manual fix for country names which contains an "-"
                                    If InStr(strScNameAndCity, "GUINEA-BISSAU") Then
                                        strScName = "GUINEA-BISSAU"
                                    ElseIf InStr(strScNameAndCity, "TIMOR-LESTE") Then
                                        strScName = "TIMOR-LESTE"
                                    End If

                                    '* Manual fixes of OCR-recognition and typing errors in raw data
                                    If strScName = "T1MO" Then
                                        strScName = "TIMOR-LESTE"
                                    ElseIf strScName = "BELARU" Then
                                        strScName = "BELARUS"
                                    ElseIf strScName = "TURKE" Then
                                        strScName = "TURKEY"
                                    ElseIf strScName = "COTE D´lVOIRE" Or strScName = "COTE LVIVOIRE" Or strScName = "COTED''IVOIRE" Or strScName = "COTEDWOIRE" Then
                                        strScName = "Cote d''Ivoire"
                                    ElseIf strScName = "GUINE" Then
                                        strScName = "GUINEA"
                                    ElseIf strScName = "VE N EZ U E LA" Then
                                        strScName = "VENEZUELA"
                                    ElseIf strScName = "COSTARICA" Then
                                        strScName = "COSTA RICA"
                                    ElseIf strScName = "UGAND" Then
                                        strScName = "UGANDA"
                                    ElseIf strScName = "GERMAN" Then
                                        strScName = "GERMANY"
                                    ElseIf strScName = "NEWZEALAND" Then
                                        strScName = "NEW ZEALAND"
                                    ElseIf strScName = "ANGOL" Then
                                        strScName = "ANGOLA"
                                    ElseIf strScName = "MYANMA" Then
                                        strScName = "MYANMAR"
                                    ElseIf strScName = "CANAD" Then
                                        strScName = "CANADA"
                                    End If

                                    'VALENCIA, SPAIN

                                    '* Fix for Valencia (Spain)
                                    If strScName = "SPAIN" And strScCity = "VALENCIA" Then
                                        strScCity = "Valencia (Spain)"
                                    End If

                                    '* Fix for Hamilton, Bermuda (part of the UK)
                                    If strScName = "BERMUDA" And strScCity = "HAMILTON" Then
                                        strScName = "United Kingdom"
                                        strScCity = "Hamilton (Bermuda)"
                                    End If

                                    '* Fix for Pristina 2008: In 2008 visa data it is still listed under Serbia; in the database Kosovo is coded as independent this year
                                    If dYear = 2008 And strScName = "SERBIA" And strScCity = "PRISTINA" Then
                                        strScName = "Kosovo"
                                        strScCity = "Pristina"
                                    End If

                                    '* Try to fetch sending country from database
                                    drc = dtCities.Select("countryName = '" & strScName & "' AND " & dYear & " >= countryStartYear")
                                    If drc.GetUpperBound(0) >= 0 Then
                                        intScID = drc(0)("countryID")
                                    End If

                                    '* If sending country was not found in database table check for name variations in the conversion table
                                    If intScID = 0 Then
                                        drc = dtAltCountryNames.Select("AlternativeName = '" & strScName & "'")
                                        If drc.GetUpperBound(0) >= 0 Then
                                            drc = dtCities.Select("countryName = '" & drc(0)("countryName").ToString.Replace("'", "''") & "' AND " & dYear & " >= countryStartYear")
                                            If drc.GetUpperBound(0) >= 0 Then
                                                intScID = drc(0)("countryID")
                                            End If
                                        End If
                                    End If

                                    '* Try to fetch cityID from database
                                    drc = dtCities.Select("countryID = " & intScID & " AND cityName = '" & strScCity & "' AND " & dYear & " >= countryStartYear")
                                    If drc.GetUpperBound(0) >= 0 Then
                                        intCityID = drc(0)("countryCityID")
                                    End If

                                    '* If city was not found in database table check for name variations in the conversion table
                                    If intCityID = 0 Then
                                        drc = dtAltCityNames.Select("AlternativeCityName = '" & strScCity & "'")
                                        If drc.GetUpperBound(0) >= 0 Then
                                            drc = dtCities.Select("countryID = " & intScID & " AND cityName = '" & drc(0)("cityName").ToString.Replace("'", "''") & "' AND " & dYear & " >= countryStartYear")
                                            If drc.GetUpperBound(0) >= 0 Then
                                                intCityID = drc(0)("countryCityID")
                                            End If
                                        End If
                                    End If

                                    '* Warn user if country was not found in database
                                    If intScID = 0 Then strResult &= "Error: Sending country not found in database: " & strScName & " (" & intRow & ")<br>"

                                    '* Warn user if city was not found in database
                                    If intCityID = 0 Then
                                        strResult &= "Error: City not found in database: " & strScCity & ", " & strScName & " (" & intRow & ")<br>"
                                    End If

                                    '* Debug: Show receiving country / sending country / sending country city identified
                                    'strResult &= intRcID & ":" & strRcName & ":" & intScID & ":" & intCityID & ":" & strScName & ":" & strScCity & "<br>"

                                    '* If diplomatic post found in database, insert data
                                    If intCityID <> 0 Then

                                        '* Reset information on visa practice
                                        strIssuedA_All = "NULL" : strIssuedB = "NULL" : strIssuedC_All = "NULL" : strAppliedC = "NULL"
                                        strIssuedABC = "NULL" : strAppliedABC = "NULL" : strNotIssuedABC = "NULL" : strIssuedVTL = "NULL"
                                        strIssuedD = "NULL" : strIssuedDC = "NULL" : strIssuedABCDDCVTL = "NULL" : strIssuedADS = "NULL"

                                        '* Get information on visa practice
                                        For intCol = 2 To 13
                                            If Not worksheet1.Cells(intRow, intCol).Value Is Nothing Then

                                                '* Get data from field; remove blank spaces; remove ".";taken into account rich text or not
                                                If worksheet1.Cells(intRow, intCol).IsRichText Then
                                                    strStatistic = worksheet1.Cells(intRow, intCol).RichText.Text.ToString.Replace(".", "").Replace(" ", "")
                                                Else
                                                    strStatistic = worksheet1.Cells(intRow, intCol).Value.ToString.Replace(".", "").Replace(" ", "")
                                                End If

                                                '* Set relevant statistic to fetched value
                                                If IsNumeric(strStatistic) Then
                                                    If intCol = 2 Then strIssuedA_All = strStatistic
                                                    If intCol = 3 Then strIssuedB = strStatistic
                                                    If intCol = 4 Then strIssuedC_All = strStatistic
                                                    If intCol = 5 Then strAppliedC = strStatistic
                                                    If intCol = 6 Then strIssuedABC = strStatistic
                                                    If intCol = 7 Then strAppliedABC = strStatistic
                                                    If intCol = 8 Then strNotIssuedABC = strStatistic
                                                    If intCol = 9 Then strIssuedVTL = strStatistic
                                                    If intCol = 10 Then strIssuedD = strStatistic
                                                    If intCol = 11 Then strIssuedDC = strStatistic
                                                    If intCol = 12 Then strIssuedABCDDCVTL = strStatistic
                                                    If intCol = 13 Then strIssuedADS = strStatistic
                                                Else
                                                    strResult &= "Warning: Visa statistic not numeric (row " & intRow & ", col " & intCol & ")<br>"
                                                End If
                                            End If
                                        Next

                                        '* Parse to database
                                        SQL = "INSERT INTO visaPractice_EU (rcID,scCityID,dYear,issuedA_All,issuedB,issuedC_All,issuedVTL,issuedADS,issuedABC,issuedD,issuedDC,issuedABCDDCVTL,appliedC,appliedABC,notIssuedABC) VALUES " _
                                            & " (" & intRcID _
                                            & "," & intCityID _
                                            & "," & dYear _
                                            & "," & strIssuedA_All _
                                            & "," & strIssuedB _
                                            & "," & strIssuedC_All _
                                            & "," & strIssuedVTL _
                                            & "," & strIssuedADS _
                                            & "," & strIssuedABC _
                                            & "," & strIssuedD _
                                            & "," & strIssuedDC _
                                            & "," & strIssuedABCDDCVTL _
                                            & "," & strAppliedC _
                                            & "," & strAppliedABC _
                                            & "," & strNotIssuedABC _
                                            & ");"
                                        'strResult &= SQL & "<br>"
                                        cmd = New OleDbCommand(SQL, db)
                                        intRowsInserted += cmd.ExecuteNonQuery()

                                    End If

                                End If

                            End If
                        End If

                    Next

                End If

                '* Import procedure for 2010
                If dYear = 2010 Then

                    '* Get data from Schengen and non-Schengen sheets
                    For intSheet = 1 To 2

                        '* Get worksheet
                        Dim worksheet1 As ExcelWorksheet = xlPackage.Workbook.Worksheets(intSheet)

                        '* Loop through each row in worksheet (skip header)
                        For intRow = 2 To worksheet1.Dimension.End.Row

                            '* Get receiving country name and try to fetch ID from database
                            strCountryCode = ""
                            intRcID = 0
                            If Not worksheet1.Cells(intRow, 3) Is Nothing Then
                                If worksheet1.Cells(intRow, 3).Value <> "" Then

                                    '* Get city name
                                    If worksheet1.Cells(intRow, 3).IsRichText Then
                                        strCountryCode = worksheet1.Cells(intRow, 3).RichText.Text
                                    Else
                                        strCountryCode = worksheet1.Cells(intRow, 3).Value
                                    End If

                                    '* Fix for Greece
                                    If strCountryCode = "EL" Then strCountryCode = "GR"

                                    '* Replace ' with '' to avoid SQL syntax error
                                    strCountryCode = strCountryCode.Replace("'", "''")

                                    '* Try to fetch country from database
                                    drc = dtCities.Select("countryCode = '" & strCountryCode & "'")
                                    If drc.GetUpperBound(0) >= 0 Then
                                        intRcID = drc(0)("countryID")
                                    End If

                                End If
                            End If

                            '* Warn user if receiving country was not found in database
                            If intRcID = 0 Then strResult &= "Error: Receiving country not found in database: " & strCountryCode & " (" & intRow & ")<br>"

                            '* Get sending country name and try to fetch ID from database
                            strCountryName = ""
                            intScID = 0
                            If Not worksheet1.Cells(intRow, 1) Is Nothing Then
                                If worksheet1.Cells(intRow, 1).Value <> "" Then

                                    '* Get city name
                                    If worksheet1.Cells(intRow, 1).IsRichText Then
                                        strCountryName = worksheet1.Cells(intRow, 1).RichText.Text
                                    Else
                                        strCountryName = worksheet1.Cells(intRow, 1).Value
                                    End If

                                    '* Replace ' with '' to avoid SQL syntax error
                                    strCountryName = strCountryName.Replace("'", "''")

                                    '* Try to fetch country from database
                                    drc = dtCities.Select("countryName = '" & strCountryName & "' AND " & dYear & " >= countryStartYear")
                                    If drc.GetUpperBound(0) >= 0 Then
                                        intScID = drc(0)("countryID")
                                    End If

                                    '* If country was not found in database table check for name variations in the conversion table
                                    If intScID = 0 Then
                                        drc = dtAltCountryNames.Select("AlternativeName = '" & strCountryName & "'")
                                        If drc.GetUpperBound(0) >= 0 Then
                                            drc = dtCities.Select("countryName = '" & drc(0)("countryName").ToString.Replace("'", "''") & "' AND " & dYear & " >= countryStartYear")
                                            If drc.GetUpperBound(0) >= 0 Then
                                                intScID = drc(0)("countryID")
                                            End If
                                        End If
                                    End If

                                End If
                            End If

                            '* Warn user if country was not found in database
                            If intScID = 0 Then strResult &= "Error: Sending country not found in database: " & strCountryName & " (" & intRow & ")<br>"

                            '* Get city name and try to fetch ID from database
                            strCityName = ""
                            intCityID = 0
                            If Not worksheet1.Cells(intRow, 2) Is Nothing Then
                                If worksheet1.Cells(intRow, 2).Value <> "" Then

                                    '* Get city name
                                    If worksheet1.Cells(intRow, 2).IsRichText Then
                                        strCityName = worksheet1.Cells(intRow, 2).RichText.Text
                                    Else
                                        strCityName = worksheet1.Cells(intRow, 2).Value
                                    End If

                                    '* Replace ' with '' to avoid SQL syntax error
                                    strCityName = strCityName.Replace("'", "''")

                                    '* Fix for Valencia (Spain)
                                    If strCountryName = "SPAIN" And strCityName = "VALENCIA" Then
                                        strCityName = "Valencia (Spain)"
                                    End If

                                    '* Try to fetch cityID from database
                                    drc = dtCities.Select("countryID = " & intScID & " AND cityName = '" & strCityName & "' AND " & dYear & " >= countryStartYear")
                                    If drc.GetUpperBound(0) >= 0 Then
                                        intCityID = drc(0)("countryCityID")
                                    End If

                                    '* If city was not found in database table check for name variations in the conversion table
                                    If intCityID = 0 Then
                                        drc = dtAltCityNames.Select("AlternativeCityName = '" & strCityName & "'")
                                        If drc.GetUpperBound(0) >= 0 Then
                                            drc = dtCities.Select("countryID = " & intScID & " AND cityName = '" & drc(0)("cityName").ToString.Replace("'", "''") & "' AND " & dYear & " >= countryStartYear")
                                            If drc.GetUpperBound(0) >= 0 Then
                                                intCityID = drc(0)("countryCityID")
                                            End If
                                        End If
                                    End If

                                End If
                            End If

                            '* Warn user if city was not found in database
                            If intCityID = 0 Then strResult &= "Error: City not found in database: " & strCityName & ", " & strCountryName & " (" & intRow & ")<br>"

                            '* Import visa practice data, if diplomatic post found in database
                            If intCityID <> 0 Then

                                '* Get information on visa practice
                                If IsNumeric(worksheet1.Cells(intRow, 4).Value) Then strIssuedA_All = worksheet1.Cells(intRow, 4).Value Else strIssuedA_All = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 5).Value) Then strIssuedA_Mev = worksheet1.Cells(intRow, 5).Value Else strIssuedA_Mev = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 6).Value) Then strIssuedB = worksheet1.Cells(intRow, 6).Value Else strIssuedB = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 7).Value) Then strIssuedC_All = worksheet1.Cells(intRow, 7).Value Else strIssuedC_All = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 8).Value) Then strIssuedC_Mev = worksheet1.Cells(intRow, 8).Value Else strIssuedC_Mev = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 10).Value) Then strAppliedC = worksheet1.Cells(intRow, 10).Value Else strAppliedC = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 11).Value) Then strIssuedABC = worksheet1.Cells(intRow, 11).Value Else strIssuedABC = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 12).Value) Then strAppliedABC = worksheet1.Cells(intRow, 12).Value Else strAppliedABC = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 13).Value) Then strNotIssuedABC = worksheet1.Cells(intRow, 13).Value Else strNotIssuedABC = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 15).Value) Then strIssuedVTL = worksheet1.Cells(intRow, 15).Value Else strIssuedVTL = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 16).Value) Then strIssuedD = worksheet1.Cells(intRow, 16).Value Else strIssuedD = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 17).Value) Then strIssuedDC = worksheet1.Cells(intRow, 17).Value Else strIssuedDC = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 18).Value) Then strIssuedABCDDCVTL = worksheet1.Cells(intRow, 18).Value Else strIssuedABCDDCVTL = "NULL"

                                '* Parse to database
                                SQL = "INSERT INTO visaPractice_EU (rcID,scCityID,dYear,issuedA_All,issuedA_Mev,issuedB,issuedC_All,issuedC_Mev,issuedVTL,issuedABC,issuedD,issuedDC,issuedABCDDCVTL,appliedC,appliedABC,notIssuedABC) VALUES " _
                                    & " (" & intRcID _
                                    & "," & intCityID _
                                    & "," & dYear _
                                    & "," & strIssuedA_All _
                                    & "," & strIssuedA_Mev _
                                    & "," & strIssuedB _
                                    & "," & strIssuedC_All _
                                    & "," & strIssuedC_Mev _
                                    & "," & strIssuedVTL _
                                    & "," & strIssuedABC _
                                    & "," & strIssuedD _
                                    & "," & strIssuedDC _
                                    & "," & strIssuedABCDDCVTL _
                                    & "," & strAppliedC _
                                    & "," & strAppliedABC _
                                    & "," & strNotIssuedABC _
                                    & ");"
                                'strResult &= SQL & "<br>"
                                cmd = New OleDbCommand(SQL, db)
                                intRowsInserted += cmd.ExecuteNonQuery()

                            End If

                        Next

                    Next

                End If

                '* Import procedure for 2011
                If dYear = 2011 Then

                    '* Get data from Schengen and non-Schengen sheets
                    For intSheet = 1 To 2

                        '* Get worksheet
                        Dim worksheet1 As ExcelWorksheet = xlPackage.Workbook.Worksheets(intSheet)

                        '* Loop through each row in worksheet (skip header)
                        For intRow = 2 To worksheet1.Dimension.End.Row

                            '* Get receiving country name and try to fetch ID from database
                            strCountryName = ""
                            intRcID = 0
                            If Not worksheet1.Cells(intRow, 1) Is Nothing Then
                                If worksheet1.Cells(intRow, 1).Value <> "" Then

                                    '* Get city name
                                    If worksheet1.Cells(intRow, 1).IsRichText Then
                                        strCountryName = worksheet1.Cells(intRow, 1).RichText.Text
                                    Else
                                        strCountryName = worksheet1.Cells(intRow, 1).Value
                                    End If

                                    '* Replace ' with '' to avoid SQL syntax error
                                    strCountryName = strCountryName.Replace("'", "''")

                                    '* Try to fetch country from database
                                    drc = dtCities.Select("countryName = '" & strCountryName & "'")
                                    If drc.GetUpperBound(0) >= 0 Then
                                        intRcID = drc(0)("countryID")
                                    End If

                                    '* If country was not found in database table check for name variations in the conversion table
                                    If intRcID = 0 Then
                                        drc = dtAltCountryNames.Select("AlternativeName = '" & strCountryName & "'")
                                        If drc.GetUpperBound(0) >= 0 Then
                                            drc = dtCities.Select("countryName = '" & drc(0)("countryName").ToString.Replace("'", "''") & "' AND " & dYear & " >= countryStartYear")
                                            If drc.GetUpperBound(0) >= 0 Then
                                                intRcID = drc(0)("countryID")
                                            End If
                                        End If
                                    End If

                                End If
                            End If

                            '* Warn user if country was not found in database
                            If intRcID = 0 Then strResult &= "Error: Receiving country not found in database: " & strCountryName & " (" & intRow & ")<br>"

                            '* Get sending country name and try to fetch ID from database
                            strCountryName = ""
                            intScID = 0
                            If Not worksheet1.Cells(intRow, 3) Is Nothing Then
                                If worksheet1.Cells(intRow, 3).Value <> "" Then

                                    '* Get city name
                                    If worksheet1.Cells(intRow, 3).IsRichText Then
                                        strCountryName = worksheet1.Cells(intRow, 3).RichText.Text
                                    Else
                                        strCountryName = worksheet1.Cells(intRow, 3).Value
                                    End If

                                    '* Replace ' with '' to avoid SQL syntax error
                                    strCountryName = strCountryName.Replace("'", "''")

                                    '* Try to fetch country from database
                                    drc = dtCities.Select("countryName = '" & strCountryName & "' AND " & dYear & " >= countryStartYear")
                                    If drc.GetUpperBound(0) >= 0 Then
                                        intScID = drc(0)("countryID")
                                    End If

                                    '* If country was not found in database table check for name variations in the conversion table
                                    If intScID = 0 Then
                                        drc = dtAltCountryNames.Select("AlternativeName = '" & strCountryName & "'")
                                        If drc.GetUpperBound(0) >= 0 Then
                                            drc = dtCities.Select("countryName = '" & drc(0)("countryName").ToString.Replace("'", "''") & "' AND " & dYear & " >= countryStartYear")
                                            If drc.GetUpperBound(0) >= 0 Then
                                                intScID = drc(0)("countryID")
                                            End If
                                        End If
                                    End If

                                End If
                            End If

                            '* Warn user if country was not found in database
                            If intScID = 0 Then strResult &= "Error: Sending country not found in database: " & strCountryName & " (" & intRow & ")<br>"

                            '* Get city name and try to fetch ID from database
                            strCityName = ""
                            intCityID = 0
                            If Not worksheet1.Cells(intRow, 4) Is Nothing Then
                                If worksheet1.Cells(intRow, 4).Value <> "" Then

                                    '* Get city name
                                    If worksheet1.Cells(intRow, 4).IsRichText Then
                                        strCityName = worksheet1.Cells(intRow, 4).RichText.Text
                                    Else
                                        strCityName = worksheet1.Cells(intRow, 4).Value
                                    End If

                                    '* Replace ' with '' to avoid SQL syntax error
                                    strCityName = strCityName.Replace("'", "''")

                                    '* Fix for Valencia (Spain)
                                    If strCountryName = "SPAIN" And strCityName = "VALENCIA" Then
                                        strCityName = "Valencia (Spain)"
                                    End If

                                    '* Try to fetch cityID from database
                                    drc = dtCities.Select("countryID = " & intScID & " AND cityName = '" & strCityName & "' AND " & dYear & " >= countryStartYear")
                                    If drc.GetUpperBound(0) >= 0 Then
                                        intCityID = drc(0)("countryCityID")
                                    End If

                                    '* If city was not found in database table check for name variations in the conversion table
                                    If intCityID = 0 Then
                                        drc = dtAltCityNames.Select("AlternativeCityName = '" & strCityName & "'")
                                        If drc.GetUpperBound(0) >= 0 Then
                                            drc = dtCities.Select("countryID = " & intScID & " AND cityName = '" & drc(0)("cityName").ToString.Replace("'", "''") & "' AND " & dYear & " >= countryStartYear")
                                            If drc.GetUpperBound(0) >= 0 Then
                                                intCityID = drc(0)("countryCityID")
                                            End If
                                        End If
                                    End If

                                End If
                            End If

                            '* Warn user if city was not found in database
                            If intCityID = 0 Then strResult &= "Error: City not found in database: " & strCityName & " (" & intRow & ")<br>"

                            '* Import visa practice data, if diplomatic post found in database
                            If intCityID <> 0 Then

                                '* Get information on visa practice
                                If IsNumeric(worksheet1.Cells(intRow, 5).Value) Then strIssuedA_All = worksheet1.Cells(intRow, 5).Value Else strIssuedA_All = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 6).Value) Then strIssuedA_Mev = worksheet1.Cells(intRow, 6).Value Else strIssuedA_Mev = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 7).Value) Then strAppliedA = worksheet1.Cells(intRow, 7).Value Else strAppliedA = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 8).Value) Then strNotIssuedA = worksheet1.Cells(intRow, 8).Value Else strNotIssuedA = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 9).Value) Then strIssuedC_All = worksheet1.Cells(intRow, 9).Value Else strIssuedC_All = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 10).Value) Then strIssuedC_Mev = worksheet1.Cells(intRow, 10).Value Else strIssuedC_Mev = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 12).Value) Then strAppliedC = worksheet1.Cells(intRow, 12).Value Else strAppliedC = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 13).Value) Then strNotIssuedC = worksheet1.Cells(intRow, 13).Value Else strNotIssuedC = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 14).Value) Then strIssuedABC = worksheet1.Cells(intRow, 14).Value Else strIssuedABC = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 15).Value) Then strAppliedABC = worksheet1.Cells(intRow, 15).Value Else strAppliedABC = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 16).Value) Then strNotIssuedABC = worksheet1.Cells(intRow, 16).Value Else strNotIssuedABC = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 18).Value) Then strIssuedVTL = worksheet1.Cells(intRow, 18).Value Else strIssuedVTL = "NULL"
                                If IsNumeric(worksheet1.Cells(intRow, 19).Value) Then strIssuedABCDVTL = worksheet1.Cells(intRow, 19).Value Else strIssuedABCDVTL = "NULL"

                                '* Parse to database
                                SQL = "INSERT INTO visaPractice_EU (rcID,scCityID,dYear,issuedA_All,issuedA_Mev,issuedC_All,issuedC_Mev,issuedVTL,issuedABC,issuedABCVTL,appliedA,appliedC,appliedABC,notIssuedA,notIssuedC,notIssuedABC) VALUES " _
                                    & " (" & intRcID _
                                    & "," & intCityID _
                                    & "," & dYear _
                                    & "," & strIssuedA_All _
                                    & "," & strIssuedA_Mev _
                                    & "," & strIssuedC_All _
                                    & "," & strIssuedC_Mev _
                                    & "," & strIssuedVTL _
                                    & "," & strIssuedABC _
                                    & "," & strIssuedABCDVTL _
                                    & "," & strAppliedA _
                                    & "," & strAppliedC _
                                    & "," & strAppliedABC _
                                    & "," & strNotIssuedA _
                                    & "," & strNotIssuedC _
                                    & "," & strNotIssuedABC _
                                    & ");"
                                'strResult &= SQL & "<br>"
                                cmd = New OleDbCommand(SQL, db)
                                intRowsInserted += cmd.ExecuteNonQuery()

                            End If

                        Next

                    Next

                End If

                '* Import procedure for 2012
                If dYear = 2012 Then

                    '* Get data from Schengen and non-Schengen sheets
                    Dim worksheet1 As ExcelWorksheet = xlPackage.Workbook.Worksheets(1)

                    '* Loop through each row in worksheet (skip header)
                    For intRow = 2 To worksheet1.Dimension.End.Row

                        '* Get receiving country name and try to fetch ID from database
                        strCountryName = ""
                        intRcID = 0
                        If Not worksheet1.Cells(intRow, 1) Is Nothing Then
                            If worksheet1.Cells(intRow, 1).Value <> "" Then

                                '* Get city name
                                If worksheet1.Cells(intRow, 1).IsRichText Then
                                    strCountryName = worksheet1.Cells(intRow, 1).RichText.Text
                                Else
                                    strCountryName = worksheet1.Cells(intRow, 1).Value
                                End If

                                '* Replace ' with '' to avoid SQL syntax error
                                strCountryName = strCountryName.Replace("'", "''")

                                '* Try to fetch country from database
                                drc = dtCities.Select("countryName = '" & strCountryName & "'")
                                If drc.GetUpperBound(0) >= 0 Then
                                    intRcID = drc(0)("countryID")
                                End If

                                '* If country was not found in database table check for name variations in the conversion table
                                If intRcID = 0 Then
                                    drc = dtAltCountryNames.Select("AlternativeName = '" & strCountryName & "'")
                                    If drc.GetUpperBound(0) >= 0 Then
                                        drc = dtCities.Select("countryName = '" & drc(0)("countryName").ToString.Replace("'", "''") & "' AND " & dYear & " >= countryStartYear")
                                        If drc.GetUpperBound(0) >= 0 Then
                                            intRcID = drc(0)("countryID")
                                        End If
                                    End If
                                End If

                            End If
                        End If

                        '* Warn user if country was not found in database
                        If intRcID = 0 Then strResult &= "Error: Receiving country not found in database: " & strCountryName & " (" & intRow & ")<br>"

                        '* Get sending country name and try to fetch ID from database
                        strCountryName = ""
                        intScID = 0
                        If Not worksheet1.Cells(intRow, 3) Is Nothing Then
                            If worksheet1.Cells(intRow, 3).Value <> "" Then

                                '* Get city name
                                If worksheet1.Cells(intRow, 3).IsRichText Then
                                    strCountryName = worksheet1.Cells(intRow, 3).RichText.Text
                                Else
                                    strCountryName = worksheet1.Cells(intRow, 3).Value
                                End If

                                '* Replace ' with '' to avoid SQL syntax error
                                strCountryName = strCountryName.Replace("'", "''")

                                '* Try to fetch country from database
                                drc = dtCities.Select("countryName = '" & strCountryName & "' AND " & dYear & " >= countryStartYear")
                                If drc.GetUpperBound(0) >= 0 Then
                                    intScID = drc(0)("countryID")
                                End If

                                '* If country was not found in database table check for name variations in the conversion table
                                If intScID = 0 Then
                                    drc = dtAltCountryNames.Select("AlternativeName = '" & strCountryName & "'")
                                    If drc.GetUpperBound(0) >= 0 Then
                                        drc = dtCities.Select("countryName = '" & drc(0)("countryName").ToString.Replace("'", "''") & "' AND " & dYear & " >= countryStartYear")
                                        If drc.GetUpperBound(0) >= 0 Then
                                            intScID = drc(0)("countryID")
                                        End If
                                    End If
                                End If

                            End If
                        End If

                        '* Warn user if country was not found in database
                        If intScID = 0 Then strResult &= "Error: Sending country not found in database: " & strCountryName & " (" & intRow & ")<br>"

                        '* Get city name and try to fetch ID from database
                        strCityName = ""
                        intCityID = 0
                        If Not worksheet1.Cells(intRow, 4) Is Nothing Then
                            If worksheet1.Cells(intRow, 4).Value <> "" Then

                                '* Get city name
                                If worksheet1.Cells(intRow, 4).IsRichText Then
                                    strCityName = worksheet1.Cells(intRow, 4).RichText.Text
                                Else
                                    strCityName = worksheet1.Cells(intRow, 4).Value
                                End If

                                '* Replace ' with '' to avoid SQL syntax error
                                strCityName = strCityName.Replace("'", "''")

                                '* Fix for Valencia (Spain)
                                If strCountryName = "SPAIN" And strCityName = "VALENCIA" Then
                                    strCityName = "Valencia (Spain)"
                                End If

                                '* Try to fetch cityID from database
                                drc = dtCities.Select("countryID = " & intScID & " AND cityName = '" & strCityName & "' AND " & dYear & " >= countryStartYear")
                                If drc.GetUpperBound(0) >= 0 Then
                                    intCityID = drc(0)("countryCityID")
                                End If

                                '* If city was not found in database table check for name variations in the conversion table
                                If intCityID = 0 Then
                                    drc = dtAltCityNames.Select("AlternativeCityName = '" & strCityName & "'")
                                    If drc.GetUpperBound(0) >= 0 Then
                                        drc = dtCities.Select("countryID = " & intScID & " AND cityName = '" & drc(0)("cityName").ToString.Replace("'", "''") & "' AND " & dYear & " >= countryStartYear")
                                        If drc.GetUpperBound(0) >= 0 Then
                                            intCityID = drc(0)("countryCityID")
                                        End If
                                    End If
                                End If

                            End If
                        End If

                        '* Warn user if city was not found in database
                        If intCityID = 0 Then strResult &= "Error: City not found in database: " & strCityName & " (" & intRow & ")<br>"

                        '* Import visa practice data, if diplomatic post found in database
                        If intCityID <> 0 Then

                            '* Get information on visa practice
                            If IsNumeric(worksheet1.Cells(intRow, 5).Value) Then strIssuedA_All = worksheet1.Cells(intRow, 5).Value Else strIssuedA_All = "NULL"
                            If IsNumeric(worksheet1.Cells(intRow, 6).Value) Then strIssuedA_Mev = worksheet1.Cells(intRow, 6).Value Else strIssuedA_Mev = "NULL"
                            If IsNumeric(worksheet1.Cells(intRow, 7).Value) Then strAppliedA = worksheet1.Cells(intRow, 7).Value Else strAppliedA = "NULL"
                            If IsNumeric(worksheet1.Cells(intRow, 8).Value) Then strNotIssuedA = worksheet1.Cells(intRow, 8).Value Else strNotIssuedA = "NULL"
                            If IsNumeric(worksheet1.Cells(intRow, 9).Value) Then strIssuedC_All = worksheet1.Cells(intRow, 9).Value Else strIssuedC_All = "NULL"
                            If IsNumeric(worksheet1.Cells(intRow, 10).Value) Then strIssuedC_Mev = worksheet1.Cells(intRow, 10).Value Else strIssuedC_Mev = "NULL"
                            If IsNumeric(worksheet1.Cells(intRow, 12).Value) Then strAppliedC = worksheet1.Cells(intRow, 12).Value Else strAppliedC = "NULL"
                            If IsNumeric(worksheet1.Cells(intRow, 13).Value) Then strNotIssuedC = worksheet1.Cells(intRow, 13).Value Else strNotIssuedC = "NULL"
                            If IsNumeric(worksheet1.Cells(intRow, 14).Value) Then strIssuedABC = worksheet1.Cells(intRow, 14).Value Else strIssuedABC = "NULL"
                            If IsNumeric(worksheet1.Cells(intRow, 15).Value) Then strAppliedABC = worksheet1.Cells(intRow, 15).Value Else strAppliedABC = "NULL"
                            If IsNumeric(worksheet1.Cells(intRow, 16).Value) Then strNotIssuedABC = worksheet1.Cells(intRow, 16).Value Else strNotIssuedABC = "NULL"
                            If IsNumeric(worksheet1.Cells(intRow, 18).Value) Then strIssuedVTL = worksheet1.Cells(intRow, 18).Value Else strIssuedVTL = "NULL"
                            If IsNumeric(worksheet1.Cells(intRow, 19).Value) Then strIssuedABCDVTL = worksheet1.Cells(intRow, 19).Value Else strIssuedABCDVTL = "NULL"

                            '* Parse to database
                            SQL = "INSERT INTO visaPractice_EU (rcID,scCityID,dYear,issuedA_All,issuedA_Mev,issuedC_All,issuedC_Mev,issuedVTL,issuedABC,issuedABCVTL,appliedA,appliedC,appliedABC,notIssuedA,notIssuedC,notIssuedABC) VALUES " _
                                & " (" & intRcID _
                                & "," & intCityID _
                                & "," & dYear _
                                & "," & strIssuedA_All _
                                & "," & strIssuedA_Mev _
                                & "," & strIssuedC_All _
                                & "," & strIssuedC_Mev _
                                & "," & strIssuedVTL _
                                & "," & strIssuedABC _
                                & "," & strIssuedABCDVTL _
                                & "," & strAppliedA _
                                & "," & strAppliedC _
                                & "," & strAppliedABC _
                                & "," & strNotIssuedA _
                                & "," & strNotIssuedC _
                                & "," & strNotIssuedABC _
                                & ");"
                            'strResult &= SQL & "<br>"
                            cmd = New OleDbCommand(SQL, db)
                            intRowsInserted += cmd.ExecuteNonQuery()

                        End If

                    Next

                End If

            End Using

        Catch ex As Exception

            '* Write error
            strResult &= "Error: " & ex.Message & "(" & intRow & ")"

        Finally

            '* Close database
            db.Close()
            db.Dispose()

            '* Close excel import sheet
            xls.Close()
            xls.Dispose()

        End Try

        '* Get status on number of rows inserted
        strResult &= "Rows inserted: " & intRowsInserted & "<br>"
        strResult &= "Rows processed: " & intRecordsProcessed & "<br>"

        '* Return result of processing script
        Return strResult

    End Function

End Class
