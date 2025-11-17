Attribute VB_Name = "Module1"


#If VBA7 Then
    Public Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
        (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, _
         ByVal dwReserved As LongPtr, ByVal lpfnCB As LongPtr) As Long
#Else
    Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
        (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
         ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If

Public Type rowType
    pkColumn As Variant
    pkColumn2 As Variant
    cell As range
    county As String
    salesDate As Date
    addDate As Date
    Status As String
    foreclosureStartAt As Date
    taxDeedStartAt As Date
    CaseNumber As String
    plaintiffName As String
    defendantName As String
    defendantShorten As String
    isCancelled As Boolean
    isRelisted As Boolean
    action As String
    sheetName As String
    finalJudgment As Double
    openingBid As Double
    AssessedValue As Double
    PrimaryPlaintiff As String
    CertificateHolderName As String
    PlaintiffMaxBid As String
    address As String
    city As String
    zip As String
    ParcelID As String
    bedrooms As Integer
    bathrooms As Integer
    LivingSF As Integer
    PropertyType As String
    PropApprAddress As String
    CourtDocAddress As String
    MyBid As Double
    saleType As String
    PA_link As String
    PA_link_Address As String
    saleURL As String
    CourtDoc_link As String
    home As String
    homePhone As String
    workPhone As String
    email As String
    filePath As String
    notes As String
    xlWS As Object
    xlApp As Object
    xlWB As Object
    fontSize As Integer
    rowHeight As Integer
    listSource As String
End Type

Dim conn As Object
Public mySemaphore As Boolean
Public OkHOAs As Variant
Public theDataArray As Variant
'Option Explicit

Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long

Private Const Timeout As Long = 5000 ' 5 seconds


Function ProcessCounty(ByVal county As String, ByVal filePath As String, ByRef rt_orginal As rowType, _
                        Optional ByVal sheetName As String = "Sales", _
                        Optional ByVal paTemplate As String, Optional ByVal saleType As String = "Foreclosure", _
                        Optional paTemplateAddr As String, _
                        Optional ByVal foreclosureStartAt As Date, Optional ByVal taxDeedStartAt As Date _
                        ) As Boolean

    ' Define variables
    Dim SaleDate As Date
    Dim addDate As Date
    Dim CaseNumber As String
    Dim Status As String
    Dim finalJudgment As Double
    Dim openingBid As Double
    Dim AssessedValue As Double
    Dim PrimaryPlaintiff As String
    Dim CertificateHolderName As String
    Dim PlaintiffMaxBid As String
    Dim address As String
    Dim city As String
    Dim zip As String
    Dim ParcelID As String
    Dim PA_link As String
    Dim MyBid As Double
    Dim skipIt As Boolean
    Dim salesList As String
    Dim PA_link_Address As String
    Dim rt As rowType
    
    Dim sheetName2 As String: sheetName2 = "Vars"
    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets(sheetName2)
    ws2.range("B3").Value = ""
    ws2.range("B4").Value = ""
    ws2.range("B5").Value = ""
    ws2.range("B15").Value = ""
    ws2.range("B16").Value = ""
    ws2.range("B17").Value = ""
    ws2.range("B18").Value = ""

   
    UpdateParameterValue "VBA_status", "running"
    rt.county = rt_orginal.county
    rt.saleType = rt_orginal.saleType
    rt.sheetName = rt_orginal.sheetName
    rt.filePath = rt_orginal.filePath
    rt.home = rt_orginal.home
    rt.PropApprAddress = getPAlink(countyName:=rt.county)
    rt.CourtDocAddress = getCourtlink(countyName:=rt.county)
    rt.fontSize = 6
    rt.rowHeight = 12
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    ws.range("a1").Value = "Sale Date"
    Call makeNamePartArray
      
    If LCase(rt.saleType) = "foreclosure" Then
        Select Case rt.county
          Case "Brevard"
              '  rt.CourtDocAddress = "http://vweb2.brevardclerk.us/Foreclosures/foreclosure_sales.html"
                ProcessBrevardCounty rt:=rt, url:="http://vweb2.brevardclerk.us/Foreclosures/foreclosure_sales.html}" ', rt.sheetName
                GoTo exit_proc:
          Case "Lake"
              rt.saleType = "foreclosure"
             ' rt.CourtDocAddress = "https://www.lakecountyclerk.org/record_searches/public_sales_calendar/"
              ProcessLakeCounty rt ', "Lake", "https://www.lakecountyclerk.org/record_searches/public_sales_calendar/", rt.sheetName
              GoTo exit_proc:
          Case "Osceola"
              rt.saleType = "foreclosure"
              rt.filePath = "C:\Users\philk\Downloads\OsceolaSalesList.txt"
           '   rt.CourtDocAddress = "https://courts.osceolaclerk.com/BenchmarkWeb/Home.aspx/Search"
              ProcessOsceolaCounty rt ', "Osceola", "C:\Users\philk\Downloads\OsceolaSalesList.txt", rt.sheetName
              GoTo exit_proc:
          Case "Highlands"
              rt.saleType = "foreclosure"
              'rt.FilePath = "C:\Users\philk\Desktop\HighlandsSalesList.txt"""
              ProcessHighlandsCounty rt
              GoTo exit_proc
          Case "Sumter"
              rt.saleType = "foreclosure"
              rt.filePath = "C:\Users\philk\Desktop\SumterSalesList.txt"
              'rt.CourtDocAddress = "https://www.civitekflorida.com/ocrs/county/60/"
              ProcessSumterCounty rt
              GoTo exit_proc:
        End Select
    End If
    If Dir(rt.filePath) = "" Then Exit Function
    If fileDateHasAlreadyBeenProcessed(localFileName:=rt.filePath, countyName:=rt.county, salesType:=rt.saleType) Then
        Exit Function
    End If
    On Error Resume Next
    ' Open the CSV file
    Open rt.filePath For Input As #1
    If Err.Number <> 0 Then
        ' File is already open - close it and try again with a new file number
        Close #1
        Dim fileNum As Integer
        fileNum = FreeFile()
        Open filePath For Input As #1
    End If
    On Error GoTo 0
    skipIt = True
    ' Loop through each row of the CSV file
    Dim lineNumber As Long
    Dim lineData As String
    Dim i As Integer
    On Error GoTo exit_proc
    Do While Not EOF(1)
        If EOF(1) Then Exit Do
'        If Len(Trim(lineData)) > 0 Then
            lineNumber = lineNumber + 1
            Line Input #1, lineData
            If Len(Trim(lineData)) > 0 Then ' check if the line is not empty
            
            'On Error Resume Next
            ' Split the line into an array
            Dim arrData() As String
                ' Split the line into an array, while escaping commas inside quotes
            Dim regex As Object
            Set regex = CreateObject("VBScript.RegExp")
            regex.Pattern = ",(?=(?:[^""]*""[^""]*"")*(?![^""]*""))"
            regex.Global = True
            arrData = Split(regex.Replace(lineData, "|"), "|")
            
            If Err.Number <> 0 Then
                Debug.Print "Error on line " & lineNumber & ": " & lineData
                Err.Clear
                GoTo EndLoop
            End If
            On Error GoTo ErrorHandler
            If skipIt Then GoTo EndLoop
            'For each element in the array, check if it is equal to "TIMESHARE"
            For i = 1 To UBound(arrData)
                If Replace(arrData(i), """", "") = "TIMESHARE" Then
                    GoTo EndLoop
                End If
            Next
            
            ' Load each field into a variable
            rt.salesDate = CDate(Replace(arrData(0), """", "")) ': rt.salesDate = saleDate
            'if the sale date is before the user specificied cutoff then skip this parcel
            If (rt.saleType = "Foreclosure" And rt.salesDate < foreclosureStartAt) Or (rt.saleType = "Tax Deed" And rt.salesDate < taxDeedStartAt) Then
                GoTo EndLoop:
            End If
            'rt.addDate = CDate(arrData(1))
            rt.addDate = CDate(Replace(arrData(1), """", ""))
            rt.CaseNumber = Replace(arrData(2), """", "")
            rt.Status = Replace(arrData(3), """", "")
            If Not Replace(arrData(4), """", "") = "" Then rt.finalJudgment = CDbl(Replace(arrData(4), """", ""))
            If Not Replace(arrData(5), """", "") = "" Then rt.openingBid = CDbl(Replace(arrData(5), """", ""))
            If Not Replace(arrData(6), """", "") = "" Then rt.AssessedValue = CDbl(Replace(arrData(6), """", ""))
            rt.PrimaryPlaintiff = Replace(arrData(7), """", "")
            rt.CertificateHolderName = Replace(arrData(8), """", "")
            If Not Replace(arrData(9), """", "") = "" Then rt.PlaintiffMaxBid = Replace(arrData(9), """", "")
            rt.address = Replace(arrData(10), """", "")
            rt.city = Replace(arrData(11), """", "")
          '  Debug.Print Address
            rt.zip = Replace(arrData(12), """", "")
            rt.ParcelID = Replace(arrData(13), """", "")
            rt.PA_link = IIf(rt.ParcelID <> "", Replace(paTemplate, "<<PID>>", rt.ParcelID), "")
            rt.PA_link_Address = IIf(rt.address <> "", Replace(paTemplateAddr, "<<address>>", Replace(Trim(rt.address), " ", "%20")), "") ': rt.PA_link_Address = PA_link_Address
            rt.isCancelled = UCase(rt.Status) = "CANCELED" ': rt.isCancelled = isCancelled
         '   Debug.Print county & " " & CaseNumber & " " & saletype & " " & saleDate
            Dim myRow As rowType
            rt.action = rt_orginal.action
            rt.PA_link = rt_orginal.PA_link
            rt.PA_link_Address = rt_orginal.PA_link_Address
            rt.CourtDoc_link = rt_orginal.CourtDoc_link
            UpdateOrAddRow2 rt
            GoTo EndLoop
     
ErrorHandler:
                Debug.Print "Error for " & county & " on line " & lineNumber & ": " & lineData & ", " & Err.description
                Err.Clear
                GoTo exit_proc:
                
EndLoop:
            skipIt = False
            'Next lineNumber
 '           End If
        End If
    Loop
exit_proc:
    ' Close the file
    UpdateParameterValue "VBA_status", "completed"
    Call urlDateHasAlreadyBeenProcessed(countyName:=rt.county, salesType:=rt.saleType)
    Close #1
    Kill filePath
    Call DeleteQuickSearchFiles

End Function

Function ProcessCounty2(ByVal county As String, ByVal filePath As String, ByRef rt_orginal As rowType, _
                        Optional ByVal sheetName As String = "Sales", _
                        Optional ByVal paTemplate As String, Optional ByVal saleType As String = "Foreclosure", _
                        Optional paTemplateAddr As String, _
                        Optional ByVal foreclosureStartAt As Date, Optional ByVal taxDeedStartAt As Date _
                        ) As Boolean

    ' Define variables
    Dim SaleDate As Date
    Dim addDate As Date
    Dim CaseNumber As String
    Dim Status As String
    Dim finalJudgment As Double
    Dim openingBid As Double
    Dim AssessedValue As Double
    Dim PrimaryPlaintiff As String
    Dim CertificateHolderName As String
    Dim PlaintiffMaxBid As String
    Dim address As String
    Dim city As String
    Dim zip As String
    Dim ParcelID As String
    Dim PA_link As String
    Dim MyBid As Double
    Dim skipIt As Boolean
    Dim salesList As String
    Dim PA_link_Address As String
    Dim rt As rowType
    Dim totalLines As Long
    rt.county = rt_orginal.county
    rt.saleType = rt_orginal.saleType
    rt.sheetName = rt_orginal.sheetName
    rt.filePath = rt_orginal.filePath
    rt.home = rt_orginal.home
    rt.PropApprAddress = getPAlink(countyName:=rt.county)
    rt.CourtDocAddress = getCourtlink(countyName:=rt.county)
      
    'Sub InsertRowIntoMySQL()
    'Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Define the connection string
    conn.ConnectionString = "Driver={MySQL ODBC 8.0 ANSI Driver};Server=localhost;Port=3306;Database='RealDeal';User=root;Password='password';"
    
    ' Open the database connection
    conn.Open
    
    ' Define your SQL query for inserting data
    Dim strSQL As String
    strSQL = "INSERT INTO your_table (column1, column2, column3) VALUES ('value1', 'value2', 'value3');"
    
    ' Execute the SQL query
'    conn.Execute strSQL
    
    ' Close the database connection
    conn.Close
    
    ' Release the connection object
    Set conn = Nothing

      
    If LCase(rt.saleType) = "foreclosure" Then
        Select Case rt.county
          Case "Brevard"
              '  rt.CourtDocAddress = "http://vweb2.brevardclerk.us/Foreclosures/foreclosure_sales.html"
                ProcessBrevardCounty rt:=rt, url:="http://vweb2.brevardclerk.us/Foreclosures/foreclosure_sales.html}" ', rt.sheetName
                GoTo exit_proc:
          Case "Lake"
              rt.saleType = "foreclosure"
             ' rt.CourtDocAddress = "https://www.lakecountyclerk.org/record_searches/public_sales_calendar/"
              ProcessLakeCounty rt ', "Lake", "https://www.lakecountyclerk.org/record_searches/public_sales_calendar/", rt.sheetName
              GoTo exit_proc:
          Case "Osceola"
              rt.saleType = "foreclosure"
              rt.filePath = "C:\Users\philk\Downloads\OsceolaSalesList.txt"
           '   rt.CourtDocAddress = "https://courts.osceolaclerk.com/BenchmarkWeb/Home.aspx/Search"
              ProcessOsceolaCounty rt ', "Osceola", "C:\Users\philk\Downloads\OsceolaSalesList.txt", rt.sheetName
              GoTo exit_proc:
          Case "Highlands"
              rt.saleType = "foreclosure"
              'rt.FilePath = "C:\Users\philk\Desktop\HighlandsSalesList.txt"""
              ProcessHighlandsCounty rt
              GoTo exit_proc
          Case "Sumter"
              rt.saleType = "foreclosure"
              rt.filePath = "C:\Users\philk\Desktop\SumterSalesList.txt"
              'rt.CourtDocAddress = "https://www.civitekflorida.com/ocrs/county/60/"
              ProcessSumterCounty rt
              GoTo exit_proc:
        End Select
    End If
    If Dir(rt.filePath) = "" Then Exit Function
    If fileDateHasAlreadyBeenProcessed(localFileName:=rt.filePath, countyName:=rt.county, salesType:=rt.saleType) Then
        Exit Function
    End If
    On Error Resume Next
    ' Open the CSV file
    Open rt.filePath For Input As #1
    If Err.Number <> 0 Then
        ' File is already open - close it and try again with a new file number
        Close #1
        Dim fileNum As Integer
        fileNum = FreeFile()
        Open filePath For Input As #1

    End If
'    Do While Not EOF(fileNum)
'        Line Input #fileNum, Line
'        totalLines = totalLines + 1
'    Loop
'    Dim sheetName2 As String: sheetName = "Vars"
'    Dim ws As Worksheet
'
'    Set ws = ThisWorkbook.Sheets(sheetName2)
'    ws.range("c3").Value = "of " & totalLines
    fileNum = FreeFile()
    Open filePath For Input As #1
        If Err.Number <> 0 Then
        ' File is already open - close it and try again with a new file number
        Close #1
        'Dim fileNum As Integer
        fileNum = FreeFile()
        Open filePath For Input As #1
    End If

    On Error GoTo 0
    skipIt = True
    ' Loop through each row of the CSV file
    Dim lineNumber As Long
    Dim lineData As String
    Dim i As Integer
    Do While Not EOF(1)
        lineNumber = lineNumber + 1
        Line Input #1, lineData
        If Len(Trim(lineData)) > 0 Then ' check if the line is not empty
        
        On Error Resume Next
        ' Split the line into an array
        Dim arrData() As String
            ' Split the line into an array, while escaping commas inside quotes
        Dim regex As Object
        Set regex = CreateObject("VBScript.RegExp")
        regex.Pattern = ",(?=(?:[^""]*""[^""]*"")*(?![^""]*""))"
        regex.Global = True
        arrData = Split(regex.Replace(lineData, "|"), "|")
        
        If Err.Number <> 0 Then
            Debug.Print "Error on line " & lineNumber & ": " & lineData
            Err.Clear
            GoTo EndLoop
        End If
        On Error GoTo ErrorHandler
        If skipIt Then GoTo EndLoop
        'For each element in the array, check if it is equal to "TIMESHARE"
        For i = 1 To UBound(arrData)
            If Replace(arrData(i), """", "") = "TIMESHARE" Then
                GoTo EndLoop
            End If
        Next
        
        ' Load each field into a variable
        rt.salesDate = CDate(Replace(arrData(0), """", "")) ': rt.salesDate = saleDate
        'if the sale date is before the user specificied cutoff then skip this parcel
        If (rt.saleType = "Foreclosure" And rt.salesDate < foreclosureStartAt) Or (rt.saleType = "Tax Deed" And rt.salesDate < taxDeedStartAt) Then
            GoTo EndLoop:
        End If
        rt.addDate = CDate(arrData(1))
        'AddDate = CDate(Replace(arrData(1), """", ""))
        rt.CaseNumber = Replace(arrData(2), """", "")
        rt.Status = Replace(arrData(3), """", "")
        If Not Replace(arrData(4), """", "") = "" Then rt.finalJudgment = CDbl(Replace(arrData(4), """", ""))
        If Not Replace(arrData(5), """", "") = "" Then rt.openingBid = CDbl(Replace(arrData(5), """", ""))
        If Not Replace(arrData(6), """", "") = "" Then rt.AssessedValue = CDbl(Replace(arrData(6), """", ""))
        rt.PrimaryPlaintiff = Replace(arrData(7), """", "")
        rt.CertificateHolderName = Replace(arrData(8), """", "")
        If Not Replace(arrData(9), """", "") = "" Then rt.PlaintiffMaxBid = Replace(arrData(9), """", "")
        rt.address = Replace(arrData(10), """", "")
        rt.city = Replace(arrData(11), """", "")
      '  Debug.Print Address
        rt.zip = Replace(arrData(12), """", "")
        rt.ParcelID = Replace(arrData(13), """", "")
        rt.PA_link = IIf(rt.ParcelID <> "", Replace(paTemplate, "<<PID>>", rt.ParcelID), "")
        rt.PA_link_Address = IIf(rt.address <> "", Replace(paTemplateAddr, "<<address>>", Replace(Trim(rt.address), " ", "%20")), "") ': rt.PA_link_Address = PA_link_Address
        rt.isCancelled = rt.Status = "Canceled" ': rt.isCancelled = isCancelled
     '   Debug.Print county & " " & CaseNumber & " " & saletype & " " & saleDate
        Dim myRow As rowType
        rt.action = rt_orginal.action
        rt.PA_link = rt_orginal.PA_link
        rt.PA_link_Address = rt_orginal.PA_link_Address
        rt.CourtDoc_link = rt_orginal.CourtDoc_link
        UpdateOrAddRow2 rt
        GoTo EndLoop
 
ErrorHandler:
            Debug.Print "Error for " & county & " on line " & lineNumber & ": " & lineData & ", " & Err.description
            Err.Clear
            GoTo EndLoop
            
EndLoop:
        skipIt = False
        'Next lineNumber
        End If
    Loop
exit_proc:
    ' Close the file
    Close #1
    Kill filePath
    Call urlDateHasAlreadyBeenProcessed(countyName:=rt.county, salesType:=rt.saleType)

End Function

Function ProcessSumterCounty(rt As rowType) As Boolean
    ' Create an instance of the FileSystemObject
    Dim FSO As Object
    Dim regex As New RegExp
    Dim regexSalesDate As New RegExp
    Dim fileNum As Integer
    Dim lineText As String
    Dim caseNum As String
    Dim parties As String
    Dim judgment As String
    Dim salesDate As String
    Dim address As String
    Dim isInRecord As Boolean
    Dim csz As String ' List of Sumter County cities
    Dim sumterCountyCities As Variant
    sumterCountyCities = Array("Sumterville", "The Villages", "Lake Panasoffkee", "Webster") ' Add more city names as needed
    Call PdfGetter(rt:=rt)
    rt.filePath = "C:\Users\philk\Dropbox\Real Estate\Deal Finder\temp.txt"
    regex.Pattern = "\d{4}-\w{2}-\d{6}"
    regexSalesDate.Pattern = "(\b\w+,\s\w+\s\d{1,2},\s\d{4}\b)"
    fileNum = FreeFile
    Open rt.filePath For Input As fileNum
    isInRecord = False
    parties = ""
    judgment = ""
    address = ""
    Do Until EOF(fileNum)
        Line Input #fileNum, lineText
        If InStr(lineText, "sales file = ") > 0 Then
            rt.saleURL = Trim(Split(lineText, " = ")(1))
        ElseIf regex.Test(lineText) Then
            If isInRecord Then
                UpdateOrAddRow2 rt:=rt
                judgment = ""
                parties = ""
                address = ""
            End If
            ' If we were previously processing a record, update the row and reset variables
            ' If we're currently processing a record, extract relevant information
            caseNum = regex.Execute(lineText)(0)
            rt.CaseNumber = caseNum
            isInRecord = True
        End If
        ' If we're currently processing a record, extract relevant information
        If isInRecord Then
            If InStr(1, lineText, "$") > 0 Then
                    ' This line contains the final judgment amount
                '    Dim parts() As String
                parts = Split(lineText, "$")
                rt.defendantName = Trim(parts(0))
                parts = Split(parts(1), " ")
                rt.finalJudgment = CDbl(Trim(parts(0)))
                rt.address = Replace(lineText, "CANCELLED", "")
                If InStr(UCase(lineText), "CANCELLED") > 0 Then
                    rt.defendantName = Trim(lastLineText)
                    rt.isCancelled = True
                    rt.address = Replace(lineText, "CANCELLED", "")
                End If
                rt.address = Replace(rt.address, rt.defendantName, "")
                rt.address = Replace(rt.address, Trim(parts(0)), "")
                rt.address = Trim(Replace(rt.address, "$", ""))
            ElseIf InStr(1, lineText, caseNum) = 0 Then
                ' This line does not contain the case number or judgment amount.
                ' If judgment has not been found yet, assume it's part of the parties, else assume it's part of the address
                If judgment = "" Then
                    parties = parties & " " & lineText
                    rt.isCancelled = InStr(UCase(parties), "CANCELLED") > 0
                    parties = Replace(parties, "CANCELLED", "")
                    parties = Replace(parties, "- VS-", "-VS-")
                    parties = Replace(parties, "-VS -", "-VS-")
                    rt.PrimaryPlaintiff = Trim(Split(parties, " -VS- ")(0))
                    rt.defendantName = Trim(Split(parties, " -VS- ")(1))
                Else
                    If regexSalesDate.Test(lineText) Then
                        ' This line contains the sales date
                        salesDate = regexSalesDate.Execute(lineText)(0)
                    Else
                        ' This line contains part of the address
                        address = address & " " & lineText
                    End If
                End If
            End If
        End If
        
        ' Check if the line contains the city, state, and zip
        If InStr(1, lineText, ", FL ") > 0 Then
            csz = lineText
            rt.city = Trim(Mid(csz, 1, InStr(csz, ", FL ") - 1))
            Dim cszTemp As String
            cszTemp = Mid(csz, InStr(csz, ", FL ") + 5) ' Extract zip code and sales date together
 
            rt.zip = Left(cszTemp, 5) ' Extract first 5 characters as zip code
            Dim salesDateStr As String
            salesDateStr = Trim(Mid(cszTemp, 6)) ' Extract the remaining characters as sales date string
            rt.salesDate = ConvertToDate(salesDateStr) ' Convert the sales date string to a date type

            ' Check if the city matches any Sumter County city names
            Dim i As Long
            For i = LBound(sumterCountyCities) To UBound(sumterCountyCities)
                If UCase(rt.city) = UCase(sumterCountyCities(i)) Then
                    rt.city = sumterCountyCities(i)
                    Exit For
                End If
            Next i
        End If
        lastLineText = lineText
    Loop
    ' Close the text file
    Close fileNum
    
    ProcessSumterCounty = True
End Function



' Custom function to convert sales date string to a date type
Function ConvertToDate(dateStr As String) As Date
    Dim parts() As String
    parts = Split(dateStr, ", ")
    Dim month As String
    Dim day As Integer
    Dim year As Integer
    month = Split(parts(1), " ")(0)
    day = Split(parts(1), " ")(1)
    year = Val(parts(2))
    ConvertToDate = DateValue(month & " " & day & ", " & year)
End Function

Function ProcessSumterCountyOLD(rt As rowType) ', ByVal county As String, ByVal FilePath As String, PropApprAddress As String, CourtDocAddress As String, Optional ByVal sheetName As String = "Sales", Optional saletype As String = "Foreclosure") As Boolean
    ' Create an instance of the FileSystemObject
    Dim content As String
    Dim FSO As Object
'    Dim PropApprAddress As String:  PropApprAddress = getPAlink(countyName:=county)
'    Dim CourtDocAddress As String:  CourtDocAddress = getCourtlink(countyName:=county)
    Call PdfGetter(rt:=rt)
    If Dir(rt.filePath) = "" Then Exit Function
    If fileDateHasAlreadyBeenProcessed(localFileName:=rt.filePath, countyName:=rt.county, salesType:=rt.saleType) Then
        Exit Function
    End If

    ' Define the regular expression pattern for matching the case number and sales date
    Dim regex As New RegExp
    regex.Pattern = "\d{4}-\w{2}-\d{6}"
    Dim regexSalesDate As New RegExp
    regexSalesDate.Pattern = "\b\w+,\s\w+\s\d{1,2},\s\d{4}\b"

    ' Open the text file for reading
    Dim fileNum As Integer
    fileNum = FreeFile
    Open rt.filePath For Input As fileNum

    ' Loop through each line of the text file
    Dim lineText As String
    Dim caseNum As String
    Dim parties As String
    Dim judgment As String
    Dim salesDate As String
    Dim address As String
    Dim isInRecord As Boolean
    Dim plaintiffName As String
    Dim defendantName As String
    isInRecord = False
    Do Until EOF(fileNum)
        Line Input #fileNum, lineText

        ' Check if the line contains a case number
        If regex.Test(lineText) Then

            ' If we were previously processing a record, print it out
            If isInRecord Then
                rt.isCancelled = InStr(UCase(parties), "CANCELLED") > 0
                parties = Replace(parties, "CANCELLED", "")
                rt.PrimaryPlaintiff = Trim(Split(parties, " -VS- ")(0))
                rt.defendantName = Trim(Split(parties, " -VS- ")(1))
                Dim city As String
                Dim zip As String

                                
                Dim parts() As String
                parts = Split(csz, ", FL ")
                
                rt.city = parts(0)
                If UBound(parts) > 0 Then rt.zip = parts(1)
                
             '   Debug.Print caseNum & vbTab & parties & vbTab & judgment & vbTab & address; ""
'                Debug.Print "caseNum = " & caseNum & vbCrLf & _
'                "plaintiffName = " & plaintiffName & vbCrLf & _
'                "defendantName = " & defendantName & vbCrLf & _
'                "judgment = " & judgment & vbCrLf & _
'                "address = " & Address & vbCrLf & _
'                "City = " & city & vbCrLf & _
'                "Zip = " & zip & vbCrLf & _
'                "isCancelled = " & isCancelled & vbCrLf & _
'                "sales date = " & salesDate & vbCrLf
                UpdateOrAddRow2 rt:=rt ', county:=county, saletype:=saletype, salesDate:=salesDate, CaseNumber:=caseNum, plaintiffName:=plaintiffName, _
                           defendantName:=defendantName, isCancelled:=isCancelled, sheetName:=sheetName, _
                           finalJudgment:=CDbl(judgment), _
                           Address:=Address, city:=city, zip:=zip, _
                           PA_link_Address:=PropApprAddress, CourtDoc_link:=CourtDocAddress
             '   End If
                
            End If

            ' Extract the case number from the line
            rt.CaseNumber = regex.Execute(lineText)(0)
            isInRecord = True
            parties = ""
            finalJudgment = ""
            rt.address = ""
        End If

        ' If we're currently processing a record, add the line to the appropriate field
        If isInRecord Then
            If InStr(lineText, "$") > 0 Then
                ' This line contains the final judgment amount
                finalJudgment = Replace(lineText, "$", "")
            ElseIf InStr(lineText, rt.CaseNumber) = 0 Then
                ' This line does not contain the case number or judgment amount.
                ' if judgement has not been found yet, assume it's part of the parties, else assume it's part of the address
                If finalJudgment = "" Then
                    parties = parties & " " & lineText
                Else
                    If regexSalesDate.Test(lineText) Then
                    ' This line contains the sales date
                        salesDate = regexSalesDate.Execute(lineText)(0)
                        ' Convert the sales date to a date type
                        If Len(salesDate) > 0 Then
                            rt.salesDate = DateValue(Mid(salesDate, InStr(salesDate, ",") + 2))
                        End If

                        'salesDate = DateValue(Mid(dateStr, InStr(regexSalesDate.Execute(lineText)(0), ",") + 2))
                   
                    Else
                       ' address = address & " " & lineText
                        If InStr(lineText, "FL") > 0 Then
                            csz = Trim(lineText)
                        Else
                            rt.address = rt.address & " " & Trim(lineText)
                        End If
                    End If
                End If
            Else
                ' This line contains the case number, so assume the previous record is complete
          '      Debug.Print caseNum & vbTab & parties & vbTab & judgment & vbTab & address
                'caseNum = ""
                parties = ""
                judgment = ""
                address = ""
                'salesDate
            End If

            ' If the line doesn't contain the case number, judgment amount, or address, assume it's part of the parties field
'            If InStr(lineText, caseNum) = 0 And InStr(lineText, "$") = 0 Then
'                parties = parties & " " & lineText
'            End If
        End If
    Loop

    ' If we were processing a record at the end of the file, print it out
    If isInRecord Then
        Debug.Print rt.CaseNumber & vbTab & parties & vbTab & judgment & vbTab & address
    End If

    ' Close the text file
    Close fileNum
End Function

Function ProcessOsceolaCounty(rt As rowType) ', ByVal county As String, ByVal FilePath As String, ByVal sheetName As String, Optional saletype As String = "Foreclosure") As Boolean
    ' Create an instance of the FileSystemObject
    Dim content As String
    Dim i As Long
    Dim salesLine As String
    'Dim salesDate As Date
   ' Dim CaseNumber As String
    Dim parties As String
    'Dim isCancelled As Boolean
   ' Dim plaintiffName As String
   ' Dim defendantName As String
   ' Dim PropApprAddress As String:  PropApprAddress = getPAlink(countyName:=county): rt.PropApprAddress = PropApprAddress
   ' Dim CourtDocAddress As String:  CourtDocAddress = getCourtlink(countyName:=county): rt.CourtDocAddress = CourtDocAddress
   rt.PropApprAddress = getPAlink(countyName:=rt.county)
   rt.CourtDocAddress = getCourtlink(countyName:=rt.county)
    
    If True Then
        Dim FSO As Object
        Call PdfGetter(rt:=rt)
        If Dir(rt.filePath) = "" Then Exit Function
        If fileDateHasAlreadyBeenProcessed(localFileName:=rt.filePath, countyName:=rt.county, salesType:=rt.saleType) Then
            Exit Function
        End If
        Set FSO = CreateObject("Scripting.FileSystemObject")
        
        ' Check if the file exists
        If Not FSO.FileExists(rt.filePath) Then
            MsgBox "File not found: " & rt.filePath, vbExclamation
            ProcessOsceolaCounty = False
            Exit Function
        End If
        
        ' Open the text file and read the content
        Dim textStream As Object
        Set textStream = FSO.OpenTextFile(rt.filePath, 1, False)
        content = textStream.ReadAll
        textStream.Close
    Else
        content = GetPDFContent(rt.filePath)
    End If
    
    ' Split the content into an array of lines
    Dim lines() As String
    s = "Sale Date Case Number Party Name"""
    content = Mid(content, InStrRev(content, s) + Len(s))
    lines = Split(content, vbCrLf)
    
    For i = 0 To UBound(lines)
        If IsDate(Left(lines(i), 10)) Then
            ' Start a new sales line
            If salesLine <> "" Then
                'Debug.Print salesDate
               ' Debug.Print caseNumber
               ' Debug.Print parties
                rt.isCancelled = InStr(1, parties, "cancelled", vbTextCompare) > 0 ': rt.isCancelled = isCancelled
                If isCancelled Then
                    parties = Replace(parties, "cancelled", "", , , vbTextCompare)
                End If
                Dim partyNames As Variant
                partyNames = Split(parties, "vs.", 2)
                If UBound(partyNames) > 0 Then
                    rt.PrimaryPlaintiff = Trim(partyNames(0)) ': rt.plaintiffName = plaintiffName
                    rt.defendantName = Trim(partyNames(1)) ': rt.defendantName = defendantName
                Else
                    rt.PrimaryPlaintiff = "" ': rt.plaintiffName = plaintiffName
                    rt.defendantName = "" ': rt.defendantName = defendantName
                End If
                'Debug.Print "Cancelled: " & isCancelled
               ' Debug.Print "Plaintiff Name: " & plaintiffName
               ' Debug.Print "Defendant Name: " & defendantName
                UpdateOrAddRow2 rt:=rt ', county:=county, salesDate:=salesDate, CaseNumber:=CaseNumber, _
                plaintiffName:=plaintiffName, defendantName:=defendantName, _
                PA_link:=PropApprAddress, isCancelled:=isCancelled, sheetName:=sheetName, _
                CourtDoc_link:=CourtDocAddress
            End If
            rt.salesDate = Left(lines(i), 10) ': rt.salesDate = salesDate
            salesLine = Mid(lines(i), 12)
            rt.CaseNumber = Left(salesLine, 17) ': rt.CaseNumber = CaseNumber
            parties = Trim(Mid(salesLine, 18))
            
        Else
        
            ' Concatenate with previous line
            parties = parties & " " & Trim(lines(i))
        End If
    Next i
    
    ' Print the last sales line
    If salesLine <> "" Then
       UpdateOrAddRow2 rt:=rt ', county:=county, salesDate:=salesDate, CaseNumber:=CaseNumber, plaintiffName:=plaintiffName, _
       defendantName:=defendantName, isCancelled:=isCancelled, sheetName:=sheetName, _
        PA_link_Address:=PropApprAddress, CourtDoc_link:=CourtDocAddress
    End If
    Dim col As String: col = GetColumnByHeaderName2(headerName:="Cancelled", getLetter:=True)
    Dim col2 As String: col2 = GetColumnByHeaderName2(headerName:="Case Number", getLetter:=True)
    Dim col3 As String: col3 = GetColumnByHeaderName2(headerName:="Sale Date", getLetter:=True)
    Dim col4 As String: col4 = GetColumnByHeaderName2(headerName:="County", getLetter:=True)
    
 'now loop through all Osceola County sales and if it is not found in the current sales list mark it as Cancelled
    Dim lastRow As Long
    ' Change the sheet name and column letters as needed
    With ThisWorkbook.Sheets("Sales")
        lastRow = .Cells(.Rows.Count, col2).End(xlUp).row ' assuming case numbers are in column B
        For i = 2 To lastRow ' assuming data starts in row 2
            If IsDate(.Cells(i, col3).Value) Then
           ' Debug.Print .Cells(I, "D").Value
           ' Debug.Print .Cells(I, "A").Value
                If .Cells(i, col4).Value = "Osceola" And .Cells(i, col3).Value >= Date Then ' assuming County is in column C and Sale date is in column D
                    CaseNumber = .Cells(i, col2).Value
                    rt.isCancelled = Not InStr(content, CaseNumber) > 0 ' check if case number is found in the content string
                    .Cells(i, col).Value = IIf(rt.isCancelled, "X", "") ' save result to column B (assuming column B is unused)
                End If
            End If
        Next i
    End With
End Function
    
Function ProcessOsceolaCountyPDF(ByVal county As String, ByVal url As String, ByVal sheetName As String)
    ' Declare variables
    Dim acroApp As Object
    Dim acroAVDoc As Object
    Dim acroPDDoc As Object
    Dim acroPage As Object
    Dim acroWord As Object
    Dim acroText As Object
    Dim pageNum As Integer
    Dim pageText As String
    Dim salesDate As Date
    Dim CaseNumber As String
    Dim plaintiffName As String
    Dim defendantName As String
    Dim isCancelled As Boolean
    Dim i As Integer
    Dim PropApprAddress As String
    
    ' Open the PDF file in Adobe Acrobat
    Set acroApp = CreateObject("AcroExch.App")
    Set acroAVDoc = CreateObject("AcroExch.AVDoc")
    If acroAVDoc.Open(url, "") Then
        Set acroPDDoc = acroAVDoc.GetPDDoc()
        If acroPDDoc.GetNumPages() > 0 Then
            ' Loop through each page of the PDF file
            For pageNum = 0 To acroPDDoc.GetNumPages() - 1
                Set acroPage = acroPDDoc.AcquirePage(pageNum)
                acroPage.RemoveAllAnnots ' Remove annotations from page to avoid errors
                Set acroWord = CreateObject("AcroExch.Word")
                acroWord.SetActiveHighlightColor RGB(255, 255, 255) ' Set highlight color to white to avoid visibility
                Set acroText = acroWord.GetChars(0, 0, 10000, 10000, acroPage)
                pageText = acroText.GetText()
                ' Update the sales data sheet with the extracted data
                If Not IsEmpty(salesDate) And Not IsEmpty(CaseNumber) And Not IsEmpty(plaintiffName) And Not IsEmpty(defendantName) Then
               '     UpdateOrAddRow county, salesDate, CaseNumber, plaintiffName, defendantName, isCancelled, sheetName, PA_link_Address:=PropApprAddress, CourtDoc_link:=CourtDocAddress
                End If
                
                Set acroPage = Nothing
            Next pageNum
        End If
        acroPDDoc.Close
    End If
    acroAVDoc.Close True
    acroApp.Exit
    Set acroApp = Nothing
End Function

Private Function DownloadFile(ByVal url As String, ByVal localPath As String) As Boolean

    ' Download the file from the URL to the local path
    Dim winHttpReq As Object
    Set winHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    winHttpReq.Open "GET", url, False
    winHttpReq.send
    
    If winHttpReq.Status <> 200 Then
        DownloadFile = False
    Else
        Dim fileStream As Object
        Set fileStream = CreateObject("ADODB.Stream")
        
        fileStream.Open
        fileStream.Type = 1 'binary
        fileStream.write winHttpReq.responseBody
        fileStream.SaveToFile localPath, 2 'overwrite
        fileStream.Close
        
        DownloadFile = True
    End If
    
End Function

Private Function ConvertPdfToTextOLD(ByVal pdfFile As String) As String

    ' Convert the PDF file to text using the pdftotext command-line tool
    Dim textFile As String
    textFile = Environ$("temp") & "\" & Format(Now, "yyyymmddhhmmss") & ".txt"
    
    Dim cmd As String
    cmd = "pdftotext -layout """ & pdfFile & """ """ & textFile & """"
    
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    
    shell.Run cmd, 0, True
    
    ' Read the text file into a string and remove it
    Dim text As String
    
    On Error Resume Next
    Dim fileNum As Integer
    fileNum = FreeFile
    Open textFile For Input As fileNum
    text = Input(LOF(fileNum), fileNum)
    Close
End Function

Function ConvertPdfToText(ByVal pdfPath As String) As String
    On Error GoTo ErrorHandler
    
    Dim text As String
    
    ' Use a Shell command to run the "pdftotext" command-line tool to convert the PDF to text
    Dim cmd As String
    cmd = "cmd /c ""pdftotext.exe -enc UTF-8 """ & pdfPath & """ -"""
    
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    
    Dim exec As Object
    Set exec = shell.exec(cmd)
    
    ' Read the output of the command and append it to the text string
    While Not exec.StdOut.AtEndOfStream
        text = text & exec.StdOut.ReadAll()
    Wend
    
    ConvertPdfToText = text
    
    Exit Function

ErrorHandler:
    MsgBox "Error: " & Err.description
    ConvertPdfToText = ""
End Function



Function ProcessBrevardCounty(ByRef rt As rowType, url As String) ', sheetName As String, Optional saletype As String = "Foreclosure")
'    Dim PropApprAddress As String:  PropApprAddress = getPAlink(countyName:=rt.county)
'    Dim CourtDocAddress As String:  CourtDocAddress = getCourtlink(countyName:=rt.county)
    
    rt.county = "Brevard"
    rt.saleType = "Foreclosure"
    
    If urlDateHasAlreadyBeenProcessed(countyName:=rt.county, salesType:=rt.saleType) Then
        Exit Function
    End If
    
    ' Send an HTTP request to the specified URL
    Dim http As New MSXML2.XMLHTTP60
    http.Open "GET", url, False
    http.send

    ' Create a new HTML document and load the response from the HTTP request
    Dim htmlDoc As New MSHTML.HTMLDocument
    htmlDoc.body.innerHTML = http.responseText

    ' Find the table that contains the foreclosure sales data
    Dim tbl As MSHTML.HTMLTable
    Dim tables As MSHTML.IHTMLElementCollection
    Set tables = htmlDoc.getElementsByTagName("table")
    For Each tbl In tables
        If tbl.getAttribute("border") = "2" And tbl.getAttribute("cellpadding") = "2" And tbl.getAttribute("cellspacing") = "1" Then
            Exit For
        End If
    Next tbl

    ' If the table was found, loop through its rows and extract the data
    If Not tbl Is Nothing Then
        Dim row As MSHTML.HTMLTableRow
        Dim isFirstRow As Boolean
        isFirstRow = True
        For Each row In tbl.Rows
            ' Skip the first row as it is a header
            If isFirstRow Then
                isFirstRow = False
                GoTo SkipHere
            End If
            ' Extract the data from the current row
'            Dim CaseNumber As String
            Dim caseTitle As String
            Dim comment As String
'            Dim saleDate As Date
'            Dim isCancelled As Boolean
            
            rt.CaseNumber = Trim(row.Cells(0).innerText)
            caseTitle = Trim(row.Cells(1).innerText)
            comment = Trim(row.Cells(2).innerText)
            rt.salesDate = CDate(Trim(row.Cells(3).innerText))
            rt.isCancelled = InStr(UCase(comment), "CANCELLED") > 0
            
            ' Split the case title into plaintiff and defendant
            Dim arrParts() As String
            arrParts = Split(caseTitle, "VS")
            Dim plaintiff As String
            Dim defendant As String
            If UBound(arrParts) >= 1 Then
                rt.plaintiffName = Trim(arrParts(0))
                rt.PrimaryPlaintiff = rt.plaintiffName
                rt.defendantName = Trim(arrParts(1))
            Else
                rt.plaintiffName = caseTitle
                rt.PrimaryPlaintiff = rt.plaintiffName
                rt.defendantName = ""
            End If
            
            
            ' Update the sales data in the sheet
            UpdateOrAddRow2 rt:=rt ', county:=county, salesDate:=saleDate, CaseNumber:=CaseNumber, _
            plaintiffName:=plaintiff, defendantName:=defendant, _
            PA_link:=rt.PropApprAddress, isCancelled:=isCancelled, sheetName:=sheetName, PA_link_Address:=rt.PropApprAddress, CourtDoc_link:=rt.CourtDocAddress
SkipHere:
        Next row
    End If
End Function




Function ProcessLakeCounty(rt As rowType) ', county As String, URL As String, sheetName As String, Optional saletype As String = "Foreclosure")
    rt.county = "Lake"
    rt.listSource = "special"
    Dim namePartArray() As String
    'namePartArray = makeNamePartArray()
    UpdateParameterValue "VBA_status", "running"
    Call makeNamePartArray
    If urlDateHasAlreadyBeenProcessed(countyName:=rt.county, salesType:=rt.saleType) Then
        Exit Function
    End If
    Dim httpRequest As Object
    Dim IE As Object
    Dim doc As Object
    Dim startIndex As Long
    Dim endIndex As Long
    'rt.county = "Lake"
    Dim url As String: url = "https://www.lakecountyclerk.org/record_searches/public_sales_calendar/"
    rt.PropApprAddress = getPAlink(countyName:=rt.county)
    rt.CourtDocAddress = getCourtlink(countyName:=rt.county)
    
    ' Delete the sheet if it already exists
    On Error Resume Next
    Application.DisplayAlerts = False
    'Sheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Create a new instance of Internet Explorer
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = False
    
    ' Navigate to the URL and wait for the page to load
    IE.Navigate url
    Do While IE.Busy Or IE.ReadyState <> 4
        DoEvents
    Loop
    
    ' Get the rendered HTML content of the page
    Set doc = IE.Document
    
    ' Extract the desired text
    Dim downloadedText As String
    downloadedText = doc.body.outerText
    
    ' Find the index of the text after the phrase "for more information and payment terms"
    startIndex = InStr(1, downloadedText, "for more information and payment terms", vbTextCompare)
    If startIndex > 0 Then
        startIndex = startIndex + Len("for more information and payment terms")
        endIndex = Len(downloadedText)
    Else
        ' If the phrase is not found, extract the entire outertext
        startIndex = 1
        endIndex = Len(downloadedText)
    End If
    
    ' Extract the desired text
    downloadedText = Mid(downloadedText, startIndex, endIndex - startIndex + 1)
    
    ' Split the text into separate lines and remove empty lines
    Dim textLines() As String
    textLines = Split(downloadedText, vbNewLine)
    Dim i As Long
    
    ' Create a new worksheet
'    Dim newSheet As Worksheet
'    Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
'   ' newSheet.Name = sheetName
'
'    ' Add headers to the worksheet
'    newSheet.Range("A1").Value = columnName1
    
    ' Add each line to the worksheet if it starts with "Mon", "Tue", "Wed", "Thu", "Fri", or "Sat"
    For i = 0 To UBound(textLines) - 1
        Dim line1 As String
        Dim strResult As String
        Dim line2 As String
        Dim s As String
        If UBound(textLines) < i + 5 Then Exit For
        line1 = Trim(textLines(i))
        line2 = Trim(textLines(i + 5))
        'If Left(line1, 3) Like "[M,T,W,T,F,S]*" And Left(line2, 2) = "20" Then
        If Left(line1, 3) Like "Mon*" Or _
               Left(line1, 3) Like "Tue*" Or _
               Left(line1, 3) Like "Wed*" Or _
               Left(line1, 3) Like "Thu*" Or _
               Left(line1, 3) Like "Fri*" Then
               rt.salesDate = DateValue(Mid(line1, 6) & " " & year(Date))
            
            If rt.salesDate < DateTime.Now Then
                 ' Date is in the past, so add a year
                 rt.salesDate = DateAdd("yyyy", 1, rt.salesDate)
            End If
               
            Dim strText As String: strText = line2
'            ' Split the string using ":" as the delimiter
            Dim arrParts() As String: arrParts = Split(strText, ":")
               rt.CaseNumber = Trim(arrParts(0))
             ' Split the second part of the array using " vs " as the delimiter
            arrParts(1) = Replace(arrParts(1), " vs ", "|")
            arrParts = Split(arrParts(1), "|")
            rt.plaintiffName = Trim(arrParts(0))
                Dim intPos As Integer
            s = Trim(arrParts(1))
            intPos = InStr(1, s, "Canceled", vbTextCompare)
            ' Extract the text before "Canceled"
            If intPos > 0 Then
                strResult = Left(s, intPos - 1)
            Else
                strResult = s
            End If
            rt.defendantName = strResult
            rt.defendantShorten = GetNameInSerarchFormat(rt.defendantName)
            rt.isCancelled = InStr(arrParts(1), "Canceled") > 0
            UpdateOrAddRow2 rt:=rt ', county:=county, salesDate:=salesDate, CaseNumber:=CaseNumber, plaintiffName:=plaintiff, defendantName:=defendent, isCancelled:=isCancelled, sheetName:=sheetName, PA_link_Address:=PropApprAddress, CourtDoc_link:=CourtDocAddress
        End If
    Next
    IE.Quit
    UpdateParameterValue "LakeCountyStatus", "dirty" 'signifies to UiPath the Lake County is read for special processing
    UpdateParameterValue "VBA_status", "completed"
End Function

Function GetParameterValue(parameterName As String) As Variant
    Dim ws As Worksheet
    Dim rng As range
    Dim foundCell As range

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Parameters")

    ' Find the parameter name in the "name" column
    Set rng = ws.Columns("A") ' Assuming "name" column is column A
    Set foundCell = rng.Find(parameterName, LookIn:=xlValues)

    ' Check if the parameter name is found
    If Not foundCell Is Nothing Then
        ' Return the corresponding value in the "value" column
        GetParameterValue = foundCell.Offset(0, 1).Value
    Else
        ' Return an error value or handle the case as needed
        GetParameterValue = CVErr(xlErrValue)
        ' Or display a message if you prefer MsgBox:
        ' MsgBox "Parameter not found: " & parameterName
    End If
End Function


Sub UpdateParameterValue(parameterName As String, newValue As Variant)
    Dim ws As Worksheet
    Dim rng As range
    Dim foundCell As range

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Parameters")

    ' Find the parameter name in the "name" column
    Set rng = ws.Columns("A") ' Assuming "name" column is column A
    Set foundCell = rng.Find(parameterName, LookIn:=xlValues)

    ' Check if the parameter name is found
    If Not foundCell Is Nothing Then
        ' Update the corresponding value in the "value" column (assuming "value" column is next to "name" column)
        foundCell.Offset(0, 1).Value = newValue
        ' MsgBox "Value updated successfully for parameter: " & parameterName
    Else
        ' Add the parameter if it doesn't exist
        ws.Cells(ws.Rows.Count, "A").End(xlUp).Offset(1, 0).Value = parameterName
        ws.Cells(ws.Rows.Count, "A").End(xlUp).Offset(0, 1).Value = newValue
        ' MsgBox "Parameter added: " & parameterName
    End If
End Sub




Sub formatCaseNumberColumn()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Find column number of "Sale Date" header
    Dim colNum As Long
    colNum = 0
    For Each cell In ws.range("1:1")
        If cell.Value = "Sale Date" Then
            colNum = cell.Column
            Exit For
        End If
    Next cell
    
    ' If "Sale Date" header not found, exit sub
    If colNum = 0 Then
        MsgBox "Column header 'Sale Date' not found.", vbExclamation
        Exit Sub
    End If
    
    ' Sort by "Sale Date"
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Cells(1, colNum), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.UsedRange
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Find column number of "Case Number" header
    colNum = 0
    For Each cell In ws.range("1:1")
        If cell.Value = "Case Number" Then
            colNum = cell.Column
            Exit For
        End If
    Next cell
    
    ' If "Case Number" header not found, exit sub
    If colNum = 0 Then
        MsgBox "Column header 'Case Number' not found.", vbExclamation
        Exit Sub
    End If
    
    ' Set up initial values for formatting loop
    Dim lastSaleDate As Date
    Dim weekNumber As Integer
    Dim isFirstRow As Boolean
    isFirstRow = True
    Dim s As Integer: s = GetColumnNumberByHeaderName(headerName:="Sale Date")
    s = 1
    ' Loop through cells in "Case Number" column
    For Each cell In ws.range(ws.Cells(2, colNum), ws.Cells(ws.Cells(ws.Rows.Count, colNum).End(xlUp).row, colNum))
        ' Get week number of current "Sale Date"
        If isFirstRow Then
            weekNumber = WorksheetFunction.IsoWeekNum(ws.Cells(2, s))
            lastSaleDate = ws.Cells(2, 1).Value
            isFirstRow = False
        ElseIf cell.Offset(-1, 0).Value <> "" And cell.Offset(-1, 0).Value <> cell.Value Then
            weekNumber = weekNumber + 1
        End If
        

Dim colorDict As Object
Set colorDict = CreateObject("Scripting.Dictionary")

If IsDate(cell.Value) Then
    If (WorksheetFunction.IsoWeekNum(cell.Value) <> weekNumber Or year(cell.Value) <> year(lastSaleDate)) Then
        If Not colorDict.Exists(weekNumber) Then
            ' If the color for this week hasn't been set yet, set it based on the week number
            If weekNumber Mod 2 = 0 Then ' Even week number
                colorDict.Add weekNumber, RGB(135, 206, 250) ' Light blue color
            Else ' Odd week number
                colorDict.Add weekNumber, RGB(255, 192, 203) ' Light red color
            End If
        End If
        ' Set the background color for this cell based on the week number
        cell.Interior.color = colorDict(WorksheetFunction.IsoWeekNum(cell.Value))
        weekNumber = WorksheetFunction.IsoWeekNum(cell.Value)
    Else
        ' Set the same background color as the previous cell in the same week
        cell.Interior.color = colorDict(weekNumber)
    End If
    lastSaleDate = cell.Value
End If




        
        ' Set font color
       ' SetFontColor cell
    Next cell
    
    ' Toggle filter if last filter state was on
    If lastFilterState = True Then
        ws.range("1:1").AutoFilter
    End If
End Sub

Function GetDirections(homeAddress As String, otherAddress As String) As String
    Dim directionsUrl As String
    
    ' Construct the URL for the Google Maps directions map
    directionsUrl = "https://www.google.com/maps/dir/?api=1" & _
                    "&origin=" & Replace(Replace(homeAddress, " ", "+"), ",", "") & _
                    "&destination=" & Replace(Replace(otherAddress, " ", "+"), ",", "")
    
    GetDirections = directionsUrl
End Function


Sub formatCaseNumber(ByRef cell As range, county As String)
   ' s = Me.GetColumnByCounty("Court Docs", county)
  '  If makeHyperlinkFormula <> "" Then
       ' cell.Value = makeHyperlinkFormula(s, cell.Value)
  '  End If
End Sub

Sub SetFontColor(ByRef cell As range)
    If cell.Interior.color <> RGB(255, 255, 255) And cell.Font.color <> RGB(255, 255, 255) Then
        If (0.299 * cell.Interior.color.Red + 0.587 * cell.Interior.color.Green + 0.114 * cell.Interior.color.Blue) > 128 Then
            cell.Font.color = RGB(255, 255, 255)
        End If
    End If
End Sub

Sub SortBySaleDate()
    On Error GoTo ErrorHandler:
    Dim ee As String
'    Set ws = ActiveSheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("sales")
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.range(ws.Cells(2, 1), ws.Cells(ws.Rows.Count, 1).End(xlUp)), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.UsedRange
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
exit_proc:
    Exit Sub
ErrorHandler:
    Debug.Print "failed sortBySaleDate , see loc 55 " & ee; Err.Number & " " & Err.description
    Err.Clear
    Resume exit_proc
End Sub


Sub SortBy(Optional fieldName As String, Optional Sheet As String = "sales", Optional Order As String = "ASC")
    On Error GoTo ErrorHandler:
    Dim ee As String
    Dim col As Integer: col = 1 'default column is the sales column
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("sales")
    If fieldName <> "" Then col = GetColumnNumberByHeaderNamePLUS1(fieldName)
    SORT_ORDER_CONSTANT = IIf(Order = "ASC", xlAscending, xlDescending)
    
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.range(ws.Cells(2, col), ws.Cells(ws.Rows.Count, col).End(xlUp)), _
            SortOn:=xlSortOnValues, Order:=SORT_ORDER_CONSTANT, DataOption:=xlSortNormal
        .SetRange ws.UsedRange
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
exit_proc:
    Exit Sub
ErrorHandler:
    Debug.Print "failed SortBy , see loc 38 " & ee; Err.Number & " " & Err.description
    Err.Clear
    Resume exit_proc
End Sub

Function GetDistance(homeAddress As String, otherAddress As String) As Double
    Dim xmlHttpRequest As Object
    Dim url As String
    Dim distance As Double
    
    ' Construct the URL for the Distance Matrix API request
    url = "https://maps.googleapis.com/maps/api/distancematrix/xml?units=imperial" & _
          "&origins=" & Replace(Replace(homeAddress, " ", "+"), ",", "") & _
          "&destinations=" & Replace(Replace(otherAddress, " ", "+"), ",", "") & _
          "&key=YOUR_API_KEY"
    
    ' Send the HTTP request to the Distance Matrix API
    Set xmlHttpRequest = CreateObject("MSXML2.XMLHTTP")
    xmlHttpRequest.Open "GET", url, False
    xmlHttpRequest.send
    
    ' Parse the response from the Distance Matrix API and extract the distance value
    distance = CDbl(Split(Split(xmlHttpRequest.responseText, "<text>")(2), " ")(0))
    
    GetDistance = distance
End Function

Function GetTravelTime(homeAddress As String, otherAddress As String) As Double
    Dim xmlHttpRequest As Object
    Dim url As String
    Dim travelTime As Double
    
    ' Construct the URL for the Distance Matrix API request
    url = "https://maps.googleapis.com/maps/api/distancematrix/xml?units=imperial" & _
          "&origins=" & Replace(Replace(homeAddress, " ", "+"), ",", "") & _
          "&destinations=" & Replace(Replace(otherAddress, " ", "+"), ",", "") & _
          "&key=YOUR_API_KEY"
    
    ' Send the HTTP request to the Distance Matrix API
    Set xmlHttpRequest = CreateObject("MSXML2.XMLHTTP")
    xmlHttpRequest.Open "GET", url, False
    xmlHttpRequest.send
    
    ' Parse the response from the Distance Matrix API and extract the travel time value
    travelTime = CDbl(Split(Split(xmlHttpRequest.responseText, "<text>")(3), " ")(0))
    
    GetTravelTime = travelTime
End Function

'Function GetDirections(homeAddress As String, otherAddress As String) As String
'    Dim directionsUrl As String
'
'    ' Construct the URL for the Google Maps directions map
'    directionsUrl = "https://www.google.com/maps/dir/?api=1" & _
'                    "&origin=" & Replace(Replace(homeAddress, " ", "+"), ",", "") & _
'                    "&destination=" & Replace(Replace(otherAddress, " ", "+"), ",", "")
'
'    GetDirections = directionsUrl
'End Function


Sub FormatSaleDateColumnAll()
    Dim ws As Worksheet
    Dim colNum As Long
    Dim cell As Variant
    
    'Set ws = ActiveSheet
    Set ws = ThisWorkbook.Sheets("sales")
    
    ' Find column number of "Sale Date" header
    colNum = 0
    For Each cell In ws.range("1:1")
        If cell.Value = "Sale Date" Then
            colNum = cell.Column
            Exit For
        End If
    Next cell
    
    ' If "Sale Date" header not found, exit sub
'    If colNum = 0 Then
'        MsgBox "Column header 'Sale Date' not found.", vbExclamation
'        Exit Sub
'    End If
    'ppaakk this is not working
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Cells(2, colNum), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.UsedRange
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Loop through "Sale Date" column and toggle color
    Dim lastDate As Date
    Dim i As String
    Dim h As String
    Dim c As String
    i = GetColumnNumberByHeaderName(headerName:="Interested")
    h = GetColumnNumberByHeaderName(headerName:="HOA")
    c = GetColumnNumberByHeaderName(headerName:="Cancelled")
    Dim colorIndex As Integer
    ' Loop through each cell in the "Sale Date" column
    For Each cell In ws.range(ws.Cells(2, colNum), ws.Cells(ws.Rows.Count, colNum).End(xlUp))
     '   Call setFormulaCells(cell)
        If IsDate(cell.Value) Then
            cell.NumberFormat = "m/d ddd"
            If (cell.Offset(0, i).Value = "" Or cell.Offset(0, i).Value = "yes") And _
                    cell.Offset(0, h).Value <> "x" And cell.Offset(0, c).Value <> "x" Then
                ' Check if the date has changed since the last cell
                If cell.Value <> lastDate Then
                    ' Update the color index
                    colorIndex = (colorIndex + 1) Mod 2
                    
                    ' Update the last date seen
                    lastDate = cell.Value
                End If
                
                ' Set the background color of the cell based on the color index
                If colorIndex = 0 Then
                    cell.Interior.color = RGB(173, 216, 230) ' light blue
                Else
                    cell.Interior.color = RGB(255, 204, 204) ' light red
                End If
            End If
        End If
    Next cell
End Sub

Sub FormatSaleDateColumnAll2()
    Dim ws As Worksheet
    Dim colNum As Long
    Dim lastDate As Date
    Dim colorIndex As Integer
    Dim weekIndex As Integer
    Dim cell As range
    Dim i As Long, h As Long, c As Long
    Dim color1 As Long, color2 As Long
    Dim weekDict As Object
    Set weekDict = CreateObject("Scripting.Dictionary")

    Set ws = ActiveSheet
    
    ' Find column number of "Sale Date" header
    colNum = 0
    For Each cell In ws.range("1:1")
        If cell.Value = "Sale Date" Then
            colNum = cell.Column
            Exit For
        End If
    Next cell
    
    ' If "Sale Date" header not found, exit sub
    If colNum = 0 Then
        MsgBox "Column header 'Sale Date' not found.", vbExclamation
        Exit Sub
    End If
    
    ' Define colors to use
    color1 = RGB(173, 216, 230) ' light blue
    color2 = RGB(255, 204, 204) ' light red
    
    ' Loop through "Sale Date" column and toggle color
    i = GetColumnNumberByHeaderName(headerName:="Interested")
    h = GetColumnNumberByHeaderName(headerName:="HOA")
    c = GetColumnNumberByHeaderName(headerName:="Cancelled")
    For Each cell In ws.range(ws.Cells(2, colNum), ws.Cells(ws.Rows.Count, colNum).End(xlUp))
        cell.NumberFormat = "m/d ddd"
        If (cell.Offset(0, i).Value = "" Or cell.Offset(0, i).Value = "yes") And _
                cell.Offset(0, h).Value <> "x" And cell.Offset(0, c).Value <> "x" Then
            ' Check if the date has changed since the last cell
            If cell.Value <> lastDate Then
                ' Update the week index
                If Not weekDict.Exists(year(cell.Value) & "_" & WorksheetFunction.RoundUp(WorksheetFunction.day(cell.Value) / 7, 0)) Then
                    weekIndex = weekIndex + 1
                    weekDict(year(cell.Value) & "_" & WorksheetFunction.RoundUp(WorksheetFunction.day(cell.Value) / 7, 0)) = weekIndex Mod 2
                End If
                
                ' Update the color index
                colorIndex = weekDict(year(cell.Value) & "_" & WorksheetFunction.RoundUp(WorksheetFunction.day(cell.Value) / 7, 0))
                
                ' Update the last date seen
                lastDate = cell.Value
            End If
            
            ' Set the background color of the cell to the right of "Sale Date"
            If colorIndex = 0 Then
                cell.Offset(0, 1).Interior.color = color1
            Else
                cell.Offset(0, 1).Interior.color = color2
            End If
        End If
    Next cell
End Sub




Sub formatSalesDateColumn()
    Dim lastRow As Long
    Dim currentColor As Long
    Dim currentDate As Date
    
    ' Change "Sales Date" to the name of your column
    With ActiveSheet.range("Sale Date")
        lastRow = .Cells(.Cells.Count).End(xlUp).row
        
        ' Set initial color to light blue
        currentColor = RGB(173, 216, 230)
        
        For i = 2 To lastRow
            If .Cells(i, 1).Value <> currentDate Then
                ' Switch color every time the date changes
                If currentColor = RGB(173, 216, 230) Then
                    currentColor = RGB(255, 204, 204)
                Else
                    currentColor = RGB(173, 216, 230)
                End If
                currentDate = .Cells(i, 1).Value
            End If
            
            ' Set format and color for the cell
            .Cells(i, 1).NumberFormat = "m/d ddd"
            .Cells(i, 1).Interior.color = currentColor
        Next i
    End With
End Sub


Sub formatSaleDateCell(ByRef match As range)
    On Error GoTo ErrorHandler:
    match.NumberFormat = "m/d ddd"
    Dim ee As String
  '  ColorByContrast match
   ' SetFontColor match
    If Weekday(match.Value, vbMonday) Mod 2 = 0 Then
        match.Interior.color = RGB(173, 216, 230)
    Else
        match.Interior.color = RGB(255, 204, 204)
    End If
exit_proc:
    Exit Sub
ErrorHandler:
    Debug.Print "failed formatSaleDateCell , see loc 52 " & ee; Err.Number & " " & Err.description
    Err.Clear
    Resume exit_proc
End Sub

Sub ColorByContrast(ByRef cell As range)
    On Error GoTo ErrorHandler:
    ' get the color of the cell above
    Dim above As range
    Set above = cell.Offset(-1, 0)
    Dim aboveColor As Long
    aboveColor = above.Interior.color

    ' check if the date is the same as above
    Dim cellDate As Date
    cellDate = cell.Value
    Dim aboveDate As Date
    aboveDate = above.Value
    Dim useSameColor As Boolean
    useSameColor = (cellDate = aboveDate)

    ' get the next available contrasting color
    Dim contrastColors(1 To 8) As Long
    contrastColors(1) = RGB(255, 192, 0)
    contrastColors(2) = RGB(0, 176, 80)
    contrastColors(3) = RGB(255, 0, 0)
    contrastColors(4) = RGB(0, 112, 192)
    contrastColors(5) = RGB(255, 0, 255)
    contrastColors(6) = RGB(0, 176, 240)
    contrastColors(7) = RGB(255, 128, 128)
    contrastColors(8) = RGB(128, 0, 128)
    
    Dim nextColorIndex As Integer
    Dim nextColor As Long
    If useSameColor Then
        nextColor = aboveColor
    Else
        i = "ee"
        nextColorIndex = Application.WorksheetFunction.match(aboveColor, contrastColors, 0) + 1
        If nextColorIndex > UBound(contrastColors) Then
            nextColorIndex = 1
        End If
        nextColor = contrastColors(nextColorIndex)
    End If

    ' apply the color to the cell
    cell.Interior.color = nextColor

    ' check if the color is dark and set the font color to white
    If ColorIsDark(nextColor) Then
        cell.Font.color = RGB(255, 255, 255)
    End If
exit_proc:
    Exit Sub
ErrorHandler:
    Debug.Print "failed ColorByContrast , see loc 51 " & ee; Err.Number & " " & Err.description
    Resume exit_proc
End Sub

Function ColorIsDark(color As Long) As Boolean
    Dim r As Integer, g As Integer, b As Integer
    r = color Mod 256
    g = (color \ 256) Mod 256
    b = (color \ 65536) Mod 256
    Dim brightness As Double
    brightness = Sqr(0.241 * r ^ 2 + 0.691 * g ^ 2 + 0.068 * b ^ 2)
    ColorIsDark = (brightness < 128)
End Function




Function ColorContrast(ByVal color1 As Long, ByVal color2 As Long) As Double
    Dim r1 As Double, g1 As Double, b1 As Double, r2 As Double, g2 As Double, b2 As Double
    
    r1 = (color1 Mod 256) / 255
    g1 = ((color1 \ 256) Mod 256) / 255
    b1 = ((color1 \ 65536) Mod 256) / 255
    
    r2 = (color2 Mod 256) / 255
    g2 = ((color2 \ 256) Mod 256) / 255
    b2 = ((color2 \ 65536) Mod 256) / 255
    
    Dim luminance1 As Double, luminance2 As Double
    luminance1 = 0.2126 * r1 + 0.7152 * g1 + 0.0722 * b1
    luminance2 = 0.2126 * r2 + 0.7152 * g2 + 0.0722 * b2
    
    Dim contrastRatio As Double
    contrastRatio = (Larger(luminance1, luminance2) + 0.05) / (Smaller(luminance1, luminance2) + 0.05)
    
    ColorContrast = contrastRatio
End Function

Function Larger(ByVal num1 As Double, ByVal num2 As Double) As Double
    If num1 > num2 Then
        Larger = num1
    Else
        Larger = num2
    End If
End Function

Function Smaller(ByVal num1 As Double, ByVal num2 As Double) As Double
    If num1 < num2 Then
        Smaller = num1
    Else
        Smaller = num2
    End If
End Function

Sub RemoveHyperlinkFormat(rng As range)
    Dim cell As range
    
    For Each cell In rng
        If cell.Hyperlinks.Count = 0 Then
            cell.Font.Underline = xlUnderlineStyleNone
            cell.Font.colorIndex = xlAutomatic
            cell.Font.FontStyle = "Regular"
        End If
    Next cell
End Sub


Public Sub populateEditRow(pkColumn As Variant, xlWS As Object, ByRef cell As range, _
                        ByVal county As String, ByVal salesDate As Date, ByVal CaseNumber As String, _
                           ByVal plaintiffName As String, ByVal defendantName As String, Optional ByVal isCancelled As Boolean, _
                           Optional ByVal action As String = "update", _
                           Optional ByVal sheetName As String = "Sales", Optional ByVal finalJudgment As Double, Optional ByVal openingBid As Double, _
                           Optional ByVal AssessedValue As Double, Optional ByVal PrimaryPlaintiff As String, _
                           Optional ByVal CertificateHolderName As String, Optional ByVal PlaintiffMaxBid As String, _
                           Optional ByVal address As String, Optional ByVal city As String, Optional ByVal zip As String, _
                           Optional ByVal ParcelID As String, Optional ByVal MyBid As Double, Optional ByVal saleType As String = "Foreclosure", _
                           Optional ByVal PA_link As String, Optional ByVal PA_link_Address As String, Optional ByVal CourtDoc_link As String, _
                           Optional ByVal home As String = "2665 Hempel Ave, Windermere FL 34786")
    Dim paLink As String
    i = "a"
    On Error GoTo ErrorHandler
    If action <> "update" Then
        cell.Value = salesDate
        cell.Offset(0, GetColumnNumberByHeaderName("County", sheetName)).Value = county
        Call populateLink2(range:=cell, headerName:="Zillo", path:="https://www.zillow.com/homes/" & Replace(address & "_" & city & ", FL " & zip, " ", "-") & "_rb")
    End If
    cell.Offset(0, GetColumnNumberByHeaderName("Days", sheetName)).Formula = "=-TODAY() + DATE(" & year(salesDate) & "," & month(salesDate) & "," & day(salesDate) & ")"
    cell.Offset(0, GetColumnNumberByHeaderName("Cancelled", sheetName)).Value = IIf(isCancelled, "x", "")
    Call populateLink2(range:=cell, headerName:="Sale", path:=getSaleHyperlink(county:=county, SaleDate:=salesDate, saleType:=saleType))
    cell.Offset(0, GetColumnNumberByHeaderName("Plaintiff Name", sheetName)).Value = plaintiffName
    cell.Offset(0, GetColumnNumberByHeaderName("Defendant Name", sheetName)).Value = defendantName
    Call formatCaseNumber(cell.Offset(0, GetColumnNumberByHeaderName("Case Number", sheetName)), county)
    i = "d"
    address = address & ", " & city & " FL " & " " & zip
    cell.Offset(0, GetColumnNumberByHeaderName("Map", sheetName)).Value = makeHyperlinkFormula(GetDirections(home, address), "c")
    cell.Offset(0, GetColumnNumberByHeaderName("Court Documents", sheetName)).Value = makeHyperlinkFormula(CourtDoc_link, "c")
    cell.Offset(0, GetColumnNumberByHeaderName("Prop Appr", sheetName)).Value = makeHyperlinkFormula(PA_link, "c")
    cell.Offset(0, GetColumnNumberByHeaderName("Prop Appr 2", sheetName)).Value = makeHyperlinkFormula(PA_link_Address, "c")
    i = "e"
    doPID cell, county, ParcelID
    i = "f"
    cell.Offset(0, GetColumnNumberByHeaderName("Case Number", sheetName)).Value = CaseNumber
 '   cell.Offset(0, 5).Value = IIf(isCancelled, "x", "")
    i = "g"
    cell.Offset(0, GetColumnNumberByHeaderName("Sales Type", sheetName)).Value = saleType
    cell.Offset(0, GetColumnNumberByHeaderName("Address", sheetName)).Value = address
    cell.Offset(0, GetColumnNumberByHeaderName("CSZ", sheetName)).Value = city & ", " & zip
    cell.Offset(0, GetColumnNumberByHeaderName("Judgement", sheetName)).Value = finalJudgment
    cell.Offset(0, GetColumnNumberByHeaderName("AssessedValue", sheetName)).Value = AssessedValue
    cell.Offset(0, GetColumnNumberByHeaderName("PlaintiffMaxBid", sheetName)).Value = PlaintiffMaxBid
    cell.Offset(0, GetColumnNumberByHeaderName("Openning Bid", sheetName)).Value = openingBid
    cell.Offset(0, GetColumnNumberByHeaderName("HOA", sheetName)).Value = isHOA(plaintiffName)
    cell.Offset(0, GetColumnNumberByHeaderName("Sale2", sheetName)).Value = getSaleHyperlink(asFormula:=True, county:=county, SaleDate:=salesDate, saleType:=saleType)
    i = "h"
    Call formatSaleDateCell(cell)
    formatCells cell
    Debug.Print IIf(action = "update", "Updated ", "Added ") & cell.Value & " " & county & " " & CaseNumber & " " & saleType & " " & address
    SleepUntilReady
exit_proc:
     Exit Sub
ErrorHandler:
    Debug.Print "failed populateEditRow , see loc 521" & i & " case#:" & CaseNumber; " Err#: " & Err.Number & " " & Err.description
    Err.Clear
    Resume exit_proc
End Sub

Public Sub formatCells(cell As range, Optional sheetName As String = "sales", Optional fontSize As Integer = 6, Optional currencyFormat As String = "$#,##0.00")
    cell.Offset(0, GetColumnNumberByHeaderName("Days", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("County", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("Address", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("Case Number", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("Court Documents", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("Defendant Name", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("Map", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("Plaintiff Name", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("Prop Appr 2", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("Prop Appr", sheetName)).Font.Size = fontSize
    'cell.Offset(0, GetColumnNumberByHeaderName("SalePath", sheetName)).Font.Size = 2
    cell.Offset(0, GetColumnNumberByHeaderName("Sales Type", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("Address", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("CSZ", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("Address", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("Judgement", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("Judgement", sheetName)).NumberFormat = currencyFormat
    cell.Offset(0, GetColumnNumberByHeaderName("Notes", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("Interested", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("PlaintiffMaxBid", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("AssessedValue", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("Openning Bid", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("Sale Date", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("Zestimate", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("Zestimate", sheetName)).NumberFormat = currencyFormat
    cell.Offset(0, GetColumnNumberByHeaderName("Equity", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("Equity", sheetName)).NumberFormat = currencyFormat
    cell.Offset(0, GetColumnNumberByHeaderName("Equity %", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("Equity %", sheetName)).NumberFormat = "0%"
    cell.Offset(0, GetColumnNumberByHeaderName("TimeStamp", sheetName)).Font.Size = fontSize
    cell.Offset(0, GetColumnNumberByHeaderName("TimeStamp", sheetName)).Value = Now() 'CDate(Date)
    cell.Offset(0, GetColumnNumberByHeaderName("PlaintiffMaxBid", sheetName)).NumberFormat = currencyFormat
    cell.Offset(0, GetColumnNumberByHeaderName("Openning Bid", sheetName)).NumberFormat = currencyFormat
    cell.Offset(0, GetColumnNumberByHeaderName("AssessedValue", sheetName)).NumberFormat = currencyFormat
    cell.Offset(0, GetColumnNumberByHeaderName("TimeStamp", sheetName)).NumberFormat = "mm/dd/yy hh:mm:ss"
    
End Sub

Public Sub UpdateOrAddRow(ByVal county As String, ByVal salesDate As Date, ByVal CaseNumber As String, rt As rowType, _
                           ByVal plaintiffName As String, ByVal defendantName As String, Optional ByVal isCancelled As Boolean, _
                           Optional ByVal sheetName As String = "Sales", Optional finalJudgment As Double, Optional openingBid As Double, _
                           Optional AssessedValue As Double, Optional PrimaryPlaintiff As String, _
                           Optional CertificateHolderName As String, Optional PlaintiffMaxBid As String, _
                           Optional address As String, Optional city As String, Optional zip As String, _
                           Optional ParcelID As String, Optional MyBid As Double, Optional saleType As String = "Foreclosure", _
                           Optional PA_link As String, Optional PA_link_Address As String, Optional CourtDoc_link As String, _
                           Optional home As String = "2665 Hempel Ave, Windermere FL 34786")
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlWS As Object
    Dim h As String
    Dim action As String
    On Error GoTo ErrorHandler:
    
    i = "a"
    Set xlApp = GetObject(, "Excel.Application")
    Set xlWB = ThisWorkbook
    Set xlWS = xlWB.Worksheets(sheetName)
    ' Find primary key columns
    Dim headers As Variant
    headers = xlWS.Rows(1).Value
    Dim pkColumn As Variant
    Dim pkColumn2 As Variant
    'this feilds make up the primary key which is always distinct
    pkColumn = Application.match("Sale Date", headers, 0)
    pkColumn2 = Application.match("Case Number", headers, 0)
    If IsError(pkColumn) Or IsError(pkColumn2) Then
    '    MsgBox "Primary key columns not found in sheet " & sheetName
        Exit Sub
    End If
    i = "b"
    ' Find matching row
    Dim cell As range
    Dim pkRange As range
    Set pkRange = xlWS.range(xlWS.Cells(2, pkColumn), xlWS.Cells(xlWS.Cells(xlWS.Rows.Count, pkColumn).End(xlUp).row, pkColumn2))
    Dim match As range
    'now see if you can find a row with this primary key
    For Each cell In pkRange.Cells
        'If cell.Value = salesDate And cell.Offset(0, 1).Value = CaseNumber Then
        If cell.Value = salesDate _
            And cell.Offset(0, GetColumnNumberByHeaderName(headerName:="County")).Value = county _
            And cell.Offset(0, GetColumnNumberByHeaderName(headerName:="Case Number")).Value = CaseNumber Then
            Set match = cell
            Call formatSaleDateCell(match)
            i = "bb"
            Exit For
        End If
    Next cell
    i = "c"
    If Not match Is Nothing Then '
        Set cell = match
        action = "update"
    Else ' Add new row
       ' Application.Run "addNewRow", pkColumn, xlWS
        Dim newRow As range
        Set newRow = xlWS.Cells(xlWS.Cells(xlWS.Rows.Count, pkColumn).End(xlUp).row + 1, pkColumn)
        Set cell = newRow
        action = "add"
    End If
    h = 0
    Do While True
        h = h + 1
        ' Create a rowtype object
        'Call populateEditRow2(rt:=myRow)
        DoEvents
        If rowExists(salesDate:=salesDate, county:=county, CaseNumber:=CaseNumber) Then
            h = 0
            Exit Do
        ElseIf h > 10 Then
            Debug.Print "Unable to add row after " & h & " attempts."
                h = 0
            Exit Do
        Else
            ' Erase the last row of the immediate window
           ' If h > 1 Then Application.SendKeys "{DEL}", True
            Debug.Print "Attempt " & h
            SleepUntilReady (h)
        End If
    Loop
exit_proc:
    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Debug.Print "failed UpdateOrAddRow , see loc 515" & i & " case#:" & CaseNumber; " Err#: " & Err.Number & " " & Err.description
    Err.Clear
    Resume exit_proc
End Sub

Sub populateZillo(rt As rowType)
    If InStr(rt.cell.Offset(0, GetColumnNumberByHeaderName("Zillo", rt.sheetName)).Formula, "/homes/_,-FL-_rb") Then
        Dim zpath As String: zpath = "https://www.zillow.com/homes/" & Replace(rt.address & "_" & rt.city & ", FL " & rt.zip, " ", "-") & "_rb"
        rt.cell.Offset(0, GetColumnNumberByHeaderName("ZilloPath", rt.sheetName)).Value = zpath
        Call populateLink2(range:=rt.cell, headerName:="Zillo", path:=zpath)
    End If
End Sub

Sub populateEditRow2(ByRef rt As rowType)
    On Error GoTo ErrorHandler
    Dim i As String: i = "a'"
    '"SPECIAL" is for Counties such as Lake that many of thier fields are processed via UiPath.  This tag causes data fields subseqently populated in UiPath to not be overwritten upon updates of LAKE county
    isNotSpecial = UCase(rt.listSource) <> "SPECIAL"
    
    If rt.action <> "update" Then
        rt.cell.Value = rt.salesDate
        rt.cell.Offset(0, GetColumnNumberByHeaderName("County", rt.sheetName)).Value = rt.county
        If isNotSpecial Then
            Call populateLink2(range:=rt.cell, headerName:="Zillo", path:="https://www.zillow.com/homes/" & Replace(rt.address & "_" & rt.city & ", FL " & rt.zip, " ", "-") & "_rb")
            populateZillo rt
        End If
    End If
        
    If isNotSpecial Then 'do not write over values already solved in UiPath
        If InStr(rt.cell.Offset(0, GetColumnNumberByHeaderName("Zillo", rt.sheetName)).Formula, "/homes/_,-FL-_rb") Then
            populateZillo rt
        End If
        Dim address As String: address = rt.address & ", " & rt.city & " FL " & " " & rt.zip
        rt.cell.Offset(0, GetColumnNumberByHeaderName("Map", rt.sheetName)).Value = makeHyperlinkFormula(GetDirections(rt.home, address), "c")
        rt.cell.Offset(0, GetColumnNumberByHeaderName("Prop Appr", rt.sheetName)).Value = makeHyperlinkFormula(rt.PA_link, "c")
        rt.cell.Offset(0, GetColumnNumberByHeaderName("Prop Appr 2", rt.sheetName)).Value = makeHyperlinkFormula(rt.PA_link_Address, "c")
        rt.cell.Offset(0, GetColumnNumberByHeaderName("Address", rt.sheetName)).Value = rt.address
        rt.cell.Offset(0, GetColumnNumberByHeaderName("CSZ", rt.sheetName)).Value = rt.city & ", " & rt.zip
        rt.cell.Offset(0, GetColumnNumberByHeaderName("Judgement", rt.sheetName)).Value = rt.finalJudgment
    
    End If
    
    'rt.cell.Offset(0, GetColumnNumberByHeaderName("Rownumber", rt.sheetName)).Formula =        "=ROW(" & rt.cell.address & ")"
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Rownumber", rt.sheetName)).Formula = Replace("=ROW(" & rt.cell.address & ")", "$", "")
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Days", rt.sheetName)).Formula = "=-TODAY() + DATE(" & year(rt.salesDate) & "," & month(rt.salesDate) & "," & day(rt.salesDate) & ")"
    If rt.isRelisted Then
        rt.cell.Offset(0, GetColumnNumberByHeaderName("Cancelled", rt.sheetName)).Value = "rl"
    ElseIf rt.isCancelled Then
        rt.cell.Offset(0, GetColumnNumberByHeaderName("Cancelled", rt.sheetName)).Value = "x"
    Else
        rt.cell.Offset(0, GetColumnNumberByHeaderName("Cancelled", rt.sheetName)).Value = ""
    End If
    Dim path2URL As String
    If rt.saleURL <> "" Then
        path2URL = rt.saleURL
    Else
        path2URL = getSaleHyperlink(county:=rt.county, SaleDate:=rt.salesDate, saleType:=rt.saleType)
    End If
    Call populateLink2(range:=rt.cell, headerName:="Sale", path:=path2URL)
   ' Call populateLink2(range:=rt.cell, headerName:="Sale", path:=getSaleHyperlink(county:=rt.county, SaleDate:=rt.salesDate, saletype:=rt.saletype))
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Plaintiff Name", rt.sheetName)).Value = rt.PrimaryPlaintiff
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Defendant Name", rt.sheetName)).Value = rt.defendantName
    'i hid "Defendant Name Short" column so it cant be found.  so i am just adding a number here
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Defendant Name", rt.sheetName) + 1).Value = rt.defendantShorten
    Call formatCaseNumber(rt.cell.Offset(0, GetColumnNumberByHeaderName("Case Number", rt.sheetName)), rt.county)
    i = "d"
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Court Documents", rt.sheetName)).Value = makeHyperlinkFormula(rt.CourtDoc_link, "c")
    i = "e"
    doPID rt.cell, rt.county, rt.ParcelID
    i = "f"
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Case Number", rt.sheetName)).Value = rt.CaseNumber
    i = "g"
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Sales Type", rt.sheetName)).Value = rt.saleType
   ' rt.cell.Offset(0, GetColumnNumberByHeaderName("ZilloPath", rt.sheetName)).Value = getSocialMediaLink(rt)
    rt.cell.Offset(0, GetColumnNumberByHeaderName("AssessedValue", rt.sheetName)).Value = rt.AssessedValue
    rt.cell.Offset(0, GetColumnNumberByHeaderName("PlaintiffMaxBid", rt.sheetName)).Value = rt.PlaintiffMaxBid
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Openning Bid", rt.sheetName)).Value = rt.openingBid
    rt.cell.Offset(0, GetColumnNumberByHeaderName("HOA", rt.sheetName)).Value = isHOA(name:=rt.PrimaryPlaintiff, county:=rt.county)
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Add Date", rt.sheetName)).Value = rt.addDate
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Days Noticed", rt.sheetName)).Value = DateDiff("d", rt.addDate, rt.salesDate)
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Sale2", rt.sheetName)).Value = getSaleHyperlink(asFormula:=True, county:=rt.county, SaleDate:=rt.salesDate, saleType:=rt.saleType)
    i = "h"
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Address Invalid", rt.sheetName)).Value = xIfFalse(IsValidAddress(address:=rt.address))
    'enter formula for equity & equity percentage
    'i use a formula as Zestimate is found after this vba and i dont want ot have to run vba again from UiPath
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Address Invalid", rt.sheetName)).Value = xIfFalse(IsValidAddress(address:=rt.address))
    rt.cell.Offset(0, GetColumnNumberByHeaderName("SaleDateYYYYMMDD", rt.sheetName)).Value = Format(rt.salesDate, "YYYYMMDD")
    Dim ZestimateCR As String
    Dim JudgementCR As String
    Dim MaxBidCR As String
    Dim theFormula As String
    'get the cell addresses for needed parts
    ZestimateCR = GetColumnByHeaderName2("Zestimate", getLetter:=True) & rt.cell.row
    RedFinEstimateCR = GetColumnByHeaderName2("RedFin Estimate", getLetter:=True) & rt.cell.row
    'RedFinEstimateCR = GetColumnByHeaderName2("RedFin Estimate", getLetter:=True) & rt.cell.row
    RealtorEstimateCR = GetColumnByHeaderName2("Realtor.com Estimate", getLetter:=True) & rt.cell.row
    JMV_CR = GetColumnByHeaderName2("Just Market Value", getLetter:=True) & rt.cell.row
    JudgementCR = GetColumnByHeaderName2("Judgement", getLetter:=True) & rt.cell.row
    MaxBidCR = GetColumnByHeaderName2("PlaintiffMaxBid", getLetter:=True) & rt.cell.row
    equity_CR = GetColumnByHeaderName2("Equity", getLetter:=True) & rt.cell.row
    FMV_CR = GetColumnByHeaderName2("FMV", getLetter:=True) & rt.cell.row
    'construct formula for FMV
  '  FMV_Formula = "=IFERROR(IF(AVERAGE(" & ZestimateCR & ":" & RedFinEstimateCR & ")> 0, AVERAGE(" & ZestimateCR & ":" & RedFinEstimateCR & ")," & JMV_CR & "/avg_JMV2FMV),  "" )"
    FMV_Formula = "=IFERROR(AVERAGE(" & ZestimateCR & ":" & RedFinEstimateCR & "), IF(AND(ISNUMBER(" & JMV_CR & "), " & JMV_CR & "<>0, ISNUMBER(avg_JMV2FMV), avg_JMV2FMV<>0), " & JMV_CR & "/avg_JMV2FMV, """"))"
  '  FMV_Formula = "=IF(AVERAGE(" & ZestimateCR & ":" & RedFinEstimateCR & ")>0, AVERAGE(" & ZestimateCR & ":" & RedFinEstimateCR & "), IF(" & JMV_CR & "/avg_JMV2FMV>0, " & JMV_CR & "/avg_JMV2FMV, " & Chr(34) & Chr(34) & "))"
   ' FMV_Formula = "=IF(OR(AVERAGE(" & ZestimateCR & ":" & RedFinEstimateCR & ")=0, ISERROR(AVERAGE(" & ZestimateCR & ":" & RedFinEstimateCR & "))), "", AVERAGE(" & ZestimateCR & ":" & RedFinEstimateCR & "))"
  ' FMV_Formula = "=IFERROR(AVERAGEIFS(" & ZestimateCR & ":" & RedFinEstimateCR & ", " & ZestimateCR & ":" & RedFinEstimateCR & ", " & Chr(34) & ">0" & Chr(34) & "),  " & Chr(34) & Chr(34) & ")"
    rt.cell.Offset(0, GetColumnNumberByHeaderName("FMV", rt.sheetName)).Value = FMV_Formula
    'construct formula for equity
    theFormula = "=IF(" & FMV_CR & "<> " & Chr(34) & Chr(34) & ", " & FMV_CR & " - MAX(" & JudgementCR & ", " & MaxBidCR & "), " & Chr(34) & Chr(34) & ")"
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Equity", rt.sheetName)).Value = theFormula
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Equity", rt.sheetName)).Font.Size = rt.fontSize
    'construct formula for equity %
    theFormula = "=IF(" & equity_CR & "<> " & Chr(34) & Chr(34) & ", " & equity_CR & "/" & FMV_CR & "," & Chr(34) & Chr(34) & ")"
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Equity %", rt.sheetName)).Value = theFormula
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Equity %", rt.sheetName)).Font.Size = rt.fontSize
    'construct formaula to test if ther is a estimated value that has not been at least attempted.  This is necessary for UiPath as only two filters are possible
    theFormula = "=OR(" & ZestimateCR & "=" & Chr(34) & Chr(34) & " ," & RedFinEstimateCR & "=" & Chr(34) & Chr(34) & "," & RealtorEstimateCR & "=" & Chr(34) & Chr(34) & "," & JMV_CR & "=" & Chr(34) & Chr(34) & ")"
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Value Needed", rt.sheetName)).Value = theFormula
    
    Call setFormulaCells(rt)
    
    Call formatSaleDateCell(rt.cell)
    formatCells rt.cell
    Debug.Print IIf(rt.action = "update", "Updated ", "Added ") & rt.salesDate & " " & rt.county & " " & rt.CaseNumber & " " & rt.saleType & " " & rt.address & ", " & rt.city
    SleepUntilReady
exit_proc:
     Exit Sub
ErrorHandler:
    Debug.Print "failed populateEditRow2 , see loc 523" & i & " case#:" & rt.CaseNumber; " Err#: " & Err.Number & " " & Err.description
    Err.Clear
    Resume exit_proc
End Sub

Sub setFormulaCells(ByRef rt As rowType)
'enter formula for equity & equity percentage
    'i use a formula as Zestimate is found after this vba and i dont want ot have to run vba again from UiPath
    Dim ZestimateCR As String
    Dim JudgementCR As String
    Dim MaxBidCR As String
    Dim theFormula As String
    'get the cell addresses for needed parts
    ZestimateCR = GetColumnByHeaderName2("Zestimate", getLetter:=True) & rt.cell.row
    RedFinEstimateCR = GetColumnByHeaderName2("RedFin Estimate", getLetter:=True) & rt.cell.row
    'RedFinEstimateCR = GetColumnByHeaderName2("RedFin Estimate", getLetter:=True) & rt.cell.row
    RealtorEstimateCR = GetColumnByHeaderName2("Realtor.com Estimate", getLetter:=True) & rt.cell.row
    JMV_CR = GetColumnByHeaderName2("Just Market Value", getLetter:=True) & rt.cell.row
    JudgementCR = GetColumnByHeaderName2("Judgement", getLetter:=True) & rt.cell.row
    MaxBidCR = GetColumnByHeaderName2("PlaintiffMaxBid", getLetter:=True) & rt.cell.row
    equity_CR = GetColumnByHeaderName2("Equity", getLetter:=True) & rt.cell.row
    FMV_CR = GetColumnByHeaderName2("FMV", getLetter:=True) & rt.cell.row
    SalePrice_CR = GetColumnByHeaderName2("Sale Price", getLetter:=True) & rt.cell.row
    
    'construct formula for FMV
  '  FMV_Formula = "=IFERROR(IF(AVERAGE(" & ZestimateCR & ":" & RedFinEstimateCR & ")> 0, AVERAGE(" & ZestimateCR & ":" & RedFinEstimateCR & ")," & JMV_CR & "/avg_JMV2FMV),  "" )"
    FMV_Formula = "=IFERROR(AVERAGE(" & ZestimateCR & ":" & RedFinEstimateCR & "), IF(AND(ISNUMBER(" & JMV_CR & "), " & JMV_CR & "<>0, ISNUMBER(avg_JMV2FMV), avg_JMV2FMV<>0), " & JMV_CR & "/avg_JMV2FMV, """"))"
  '  FMV_Formula = "=IF(AVERAGE(" & ZestimateCR & ":" & RedFinEstimateCR & ")>0, AVERAGE(" & ZestimateCR & ":" & RedFinEstimateCR & "), IF(" & JMV_CR & "/avg_JMV2FMV>0, " & JMV_CR & "/avg_JMV2FMV, " & Chr(34) & Chr(34) & "))"
   ' FMV_Formula = "=IF(OR(AVERAGE(" & ZestimateCR & ":" & RedFinEstimateCR & ")=0, ISERROR(AVERAGE(" & ZestimateCR & ":" & RedFinEstimateCR & "))), "", AVERAGE(" & ZestimateCR & ":" & RedFinEstimateCR & "))"
  ' FMV_Formula = "=IFERROR(AVERAGEIFS(" & ZestimateCR & ":" & RedFinEstimateCR & ", " & ZestimateCR & ":" & RedFinEstimateCR & ", " & Chr(34) & ">0" & Chr(34) & "),  " & Chr(34) & Chr(34) & ")"
    rt.cell.Offset(0, GetColumnNumberByHeaderName("FMV", rt.sheetName)).Value = FMV_Formula
    'construct formula for equity
    theFormula = "=IF(" & FMV_CR & "<> " & Chr(34) & Chr(34) & ", " & FMV_CR & " - MAX(" & JudgementCR & ", " & MaxBidCR & "), " & Chr(34) & Chr(34) & ")"
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Equity", rt.sheetName)).Value = theFormula
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Equity", rt.sheetName)).Font.Size = rt.fontSize
    'construct formula for equity %
    theFormula = "=IF(" & equity_CR & "<> " & Chr(34) & Chr(34) & ", " & equity_CR & "/" & FMV_CR & "," & Chr(34) & Chr(34) & ")"
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Equity %", rt.sheetName)).Value = theFormula
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Equity %", rt.sheetName)).Font.Size = rt.fontSize
    'construct formaula to test if ther is a estimated value that has not been at least attempted.  This is necessary for UiPath as only two filters are possible
    theFormula = "=OR(" & ZestimateCR & "=" & Chr(34) & Chr(34) & " ," & RedFinEstimateCR & "=" & Chr(34) & Chr(34) & "," & RealtorEstimateCR & "=" & Chr(34) & Chr(34) & "," & JMV_CR & "=" & Chr(34) & Chr(34) & ")"
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Value Needed", rt.sheetName)).Value = theFormula
    theFormula = "=IFERROR(" & SalePrice_CR & "/" & FMV_CR & "," & Chr(34) & Chr(34) & ")"
    rt.cell.Offset(0, GetColumnNumberByHeaderName("Sale Price to FMV", rt.sheetName)).Value = theFormula
    Call formatSaleDateCell(rt.cell)
    formatCells rt.cell
    Debug.Print IIf(rt.action = "update", "Updated ", "Added ") & rt.salesDate & " " & rt.county & " " & rt.CaseNumber & " " & rt.saleType & " " & rt.address & ", " & rt.city
    SleepUntilReady
exit_proc:
     Exit Sub
ErrorHandler:
    Debug.Print "failed populateEditRow2 , see loc 523" & i & " case#:" & rt.CaseNumber; " Err#: " & Err.Number & " " & Err.description
    Err.Clear
    Resume exit_proc
End Sub



Function IsValidAddress(address As String) As Boolean
    ' Check if the address is not empty, has at least 10 characters,
    ' does not contain the string "unknown" (case-insensitive),
    ' and the first character is not zero
    If Len(address) >= 10 And InStr(1, UCase(address), "UNKNOWN", vbTextCompare) = 0 And Not Left(address, 1) Like "0" Then
        ' Check if the first character is a digit
        If IsNumeric(Left(address, 1)) Then
            IsValidAddress = True
        End If
    End If
End Function

Function xIf(x As Boolean, Optional objective As Boolean = True) As String
    If (x And objective) Or (Not x And Not objective) Then
        xIf = "x"
    Else
        xIf = ""
    End If
End Function

Function xIfFalse(x As Boolean) As String
    xIfFalse = xIf(x, False)
End Function

Function xIfTrue(x As Boolean) As Variant
    xIfFalse = xIf(x, True)
End Function


Public Sub UpdateOrAddRow2(rt As rowType, Optional currencyFormat As String = "$#,##0.00")
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlWS As Object
    Dim h As String
    Dim i As String
    Dim action As String
    On Error GoTo ErrorHandler:
    
    i = "a"
    rt.isRelisted = False
    Set rt.xlApp = GetObject(, "Excel.Application")
    Set rt.xlWB = ThisWorkbook
    Set rt.xlWS = rt.xlWB.Worksheets(rt.sheetName)
    ' Find primary key columns
    Dim headers As Variant
    headers = rt.xlWS.Rows(1).Value
    'this feilds make up the primary key which is always distinct
    rt.pkColumn = Application.match("Sale Date", headers, 0)
    rt.pkColumn2 = Application.match("Case Number", headers, 0)
    If IsError(rt.pkColumn) Or IsError(rt.pkColumn2) Then
        Exit Sub
    End If
    i = "b"
    ' Find matching row
    Dim cell As range
    Dim pkRange As range
    Set pkRange = rt.xlWS.range(rt.xlWS.Cells(2, rt.pkColumn), rt.xlWS.Cells(rt.xlWS.Cells(rt.xlWS.Rows.Count, rt.pkColumn).End(xlUp).row, rt.pkColumn2))
    Dim match As range
    'now see if you can find a row with this primary key
    
    For Each cell In pkRange.Cells
        If Not IsDate(cell.Value) Then
            Set match = Nothing
    '       Rows.Delete
          ' On Error Resume Next
'        ElseIf cell.Value = rt.salesDate _
'            And cell.Offset(0, GetColumnNumberByHeaderName(headerName:="County")).Value = rt.county _
'            And cell.Offset(0, GetColumnNumberByHeaderName(headerName:="Case Number")).Value = rt.CaseNumber Then
'                Set match = cell
'                Call formatSaleDateCell(match)
'                I = "bb"
'                Exit For
        ElseIf cell.Offset(0, GetColumnNumberByHeaderName(headerName:="County")).Value = rt.county _
            And cell.Offset(0, GetColumnNumberByHeaderName(headerName:="Case Number")).Value = rt.CaseNumber Then
            If Not rt.isRelisted Then rt.isRelisted = cell.Value <> rt.salesDate
            If cell.Value = rt.salesDate Then
                Set match = cell
                Call formatSaleDateCell(match)
                i = "bb"
                Exit For
            End If
            cell.rowHeight = rt.rowHeight
        End If
    Next cell
    i = "c"
    rt.PA_link = getPAlink(countyName:=rt.county, PID:=rt.ParcelID)
    rt.PA_link_Address = getPAlink(countyName:=rt.county, address:=rt.address, What2Get:="ByAddress")
    rt.CourtDoc_link = getCourtlink(countyName:=rt.county)
    If Not match Is Nothing Then '
        Set rt.cell = match
        rt.action = "update"
        populateEditRow2 rt
    Else ' Add new row
       ' Application.Run "addNewRow", rt
       rt.action = "add"
       AddNewRow rt
    End If
    Dim sheetName As String: sheetName = "Vars"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
'    ws.range("a1").Value = "Sale Date"
'now update the monitor on the var page to show progress
    If IsNumeric(ws.range("b3").Value) Then
        ws.range("b3").Value = ws.range("b3").Value + 1
        ws.range("b15").Value = ws.range("b15").Value + 1
        ws.range("b18").Value = rt.salesDate
        ws.range("c18").Value = "in " & DateDiff("d", Now(), rt.salesDate) & " days"
        ws.range("a19").Value = IIf(rt.saleType = "Foreclosure", "Case Number", "File Number")
        ws.range("b19").Value = rt.CaseNumber
        ws.range("b20").Value = rt.address
        ws.range("b21").Value = rt.city & ", FL " & rt.zip
        ws.range("b22").Value = rt.county
        ws.range("b23").Value = rt.finalJudgment
        ws.range("b23").NumberFormat = currencyFormat
        ws.range("b24").Value = rt.PrimaryPlaintiff
        ws.range("b25").Value = rt.defendantName
        ws.range("b26").Value = IIf(rt.isCancelled, "CANCELLED", "")
       ' ws.range("b27").Value = IIf(cell.Offset(0, Trim(GetColumnNumberByHeaderName("HOA", sheetName)).Value) = "", "", "Plaintiff is an HOA")
        If rt.action = "update" Then
            ws.range("b16").Value = ws.range("b16").Value + 1
        Else
            ws.range("b17").Value = ws.range("b17").Value + 1
        End If
    Else
        ws.range("b3").Value = 1
    End If
exit_proc:
    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Debug.Print "failed UpdateOrAddRow2 , see loc 519" & i & " case#:" & rt.CaseNumber; " Err#: " & Err.Number & " " & Err.description
    Err.Clear
    Resume exit_proc
End Sub

Sub AddNewRow(rt As rowType)
    ' Acquire the global variable
    While mySemaphore = True
        Sleep 1
    Wend
    
    Dim xlWB As Object
    Dim xlWS As Object
    Dim action As String
    On Error GoTo ErrorHandler:
    
    Set xlApp = GetObject(, "Excel.Application")
    Set xlWB = ThisWorkbook
    Set xlWS = xlWB.Worksheets(rt.sheetName)
    h = 0
    
    ' Unhide the worksheet
    xlWS.Visible = True
    Do While True
        h = h + 1
        Dim newRow As range
        Set newRow = xlWS.Cells(xlWS.Cells(xlWS.Rows.Count, rt.pkColumn).End(xlUp).row + 1, rt.pkColumn)
        Set rt.cell = newRow
        populateEditRow2 rt
        DoEvents
        If rowExists(salesDate:=rt.salesDate, county:=rt.county, CaseNumber:=rt.CaseNumber) Then
            h = 0
            DoEvents
            Exit Do
        ElseIf h > 10 Then
            Debug.Print "Unable to add row after " & h & " attempts."
                h = 0
            Exit Do
        Else
            ' Erase the last row of the immediate window
           ' If h > 1 Then Application.SendKeys "{DEL}", True
            Debug.Print "Attempt " & h
            populateEditRow2 rt
            SleepUntilReady (h)
        End If
    Loop
    mySemaphore = False
exit_proc:
    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Debug.Print "failed addNewRow , see loc 518" & i & " case#:" & CaseNumber; " Err#: " & Err.Number & " " & Err.description
    Err.Clear
    Resume exit_proc
End Sub

Sub AddNewRow2()
    Dim ws As Worksheet
    Dim rng As range
    
    Set ws = ThisWorkbook.Sheets("sales")
    Set rng = ws.range("A1")
    
    rng.Insert Shift:=xlDown
    rng.Value = today
End Sub


Sub PopulateWithThreads(ByRef cell As range, _
                        ByVal county As String, ByVal salesDate As Date, ByVal CaseNumber As String, _
                           ByVal plaintiffName As String, ByVal defendantName As String, Optional ByVal isCancelled As Boolean, _
                           Optional ByVal action As String = "update", _
                           Optional ByVal sheetName As String = "Sales", Optional ByVal finalJudgment As Double, Optional ByVal openingBid As Double, _
                           Optional ByVal AssessedValue As Double, Optional ByVal PrimaryPlaintiff As String, _
                           Optional ByVal CertificateHolderName As String, Optional ByVal PlaintiffMaxBid As String, _
                           Optional ByVal address As String, Optional ByVal city As String, Optional ByVal zip As String, _
                           Optional ByVal ParcelID As String, Optional ByVal MyBid As Double, Optional ByVal saleType As String = "Foreclosure", _
                           Optional ByVal PA_link As String, Optional ByVal PA_link_Address As String, Optional ByVal CourtDoc_link As String, _
                           Optional ByVal home As String = "2665 Hempel Ave, Windermere FL 34786")
    
    If action = "Add" Then
        h = 0
        Do While True
            h = h + 1
'            populateEditRow action:=action, cell:=cell, county:=county, saletype:=saletype, salesDate:=salesDate, CaseNumber:=CaseNumber, _
'                            plaintiffName:=PrimaryPlaintiff, defendantName:=defendantName, isCancelled:=isCancelled, sheetName:=sheetName, _
'                            finalJudgment:=finalJudgment, AssessedValue:=AssessedValue, PrimaryPlaintiff:=PrimaryPlaintiff, _
'                            PlaintiffMaxBid:=PlaintiffMd, openingBid:=openingBid, Address:=Address, city:=city, zip:=zip, _
'                            ParcelID:=ParcelID, PA_link:=PA_link, PA_link_Address:=PA_link_Address, CourtDoc_link:=CourtDocAddress, _
'                            CertificateHolderName:=CertificateHolderNameaxBi
            DoEvents
            If rowExists(salesDate:=salesDate, county:=county, CaseNumber:=CaseNumber) Then
                h = 0
                Exit Do
            ElseIf h > 10 Then
                Debug.Print "Unable to add row after " & h & " attempts."
                    h = 0
                Exit Do
            Else
                ' Erase the last row of the immediate window
               ' If h > 1 Then Application.SendKeys "{DEL}", True
                Debug.Print "Attempt " & h
                SleepUntilReady (h)
            End If
        Loop
    Else
        Application.Run "AddWithThread", cell, "county_value", #4/29/2023#, "CaseNumber_value", _
            "plaintiffName_value", "defendantName_value", False, "update", "Sales", 1000, 500, 750, _
            "PrimaryPlaintiff_value", "CertificateHolderName_value", "PlaintiffMaxBid_value", "Address_value", "city_value", "zip_value", _
            "ParcelID_value", 0, "Foreclosure", "PA_link_value", "PA_link_Address_value", "CourtDoc_link_value"
    End If

End Sub

Sub AddWithThread(ByRef cell As range, _
                  ByVal county As String, ByVal salesDate As Date, ByVal CaseNumber As String, _
                  ByVal plaintiffName As String, ByVal defendantName As String, Optional ByVal isCancelled As Boolean, _
                  Optional ByVal action As String = "update", _
                  Optional ByVal sheetName As String = "Sales", Optional ByVal finalJudgment As Double, Optional ByVal openingBid As Double, _
                  Optional ByVal AssessedValue As Double, Optional ByVal PrimaryPlaintiff As String, _
                  Optional ByVal CertificateHolderName As String, Optional ByVal PlaintiffMaxBid As String, _
                  Optional ByVal address As String, Optional ByVal city As String, Optional ByVal zip As String, _
                  Optional ByVal ParcelID As String, Optional ByVal MyBid As Double, Optional ByVal saleType As String = "Foreclosure", _
                  Optional ByVal PA_link As String, Optional ByVal PA_link_Address As String, Optional ByVal CourtDoc_link As String, _
                  Optional ByVal home As String = "2665 Hempel Ave, Windermere FL 34786")
    ' code here



'Sub AddWithThread(ByRef cell As range, _
'                        ByVal county As String, ByVal salesDate As Date, ByVal CaseNumber As String, _
'                           ByVal plaintiffName As String, ByVal defendantName As String, Optional ByVal isCancelled As Boolean, _
'                           Optional ByVal action As String = "update", _
'                           Optional ByVal sheetName As String = "Sales", Optional ByVal finalJudgment As Double, Optional ByVal openingBid As Double, _
'                           Optional ByVal AssessedValue As Double, Optional ByVal PrimaryPlaintiff As String, _
'                           Optional ByVal CertificateHolderName As String, Optional ByVal PlaintiffMaxBid As String, _
'                           Optional ByVal Address As String, Optional ByVal city As String, Optional ByVal zip As String, _
'                           Optional ByVal ParcelID As String, Optional ByVal MyBid As Double, Optional ByVal saletype As String = "Foreclosure", _
'                           Optional ByVal PA_link As String, Optional ByVal PA_link_Address As String, Optional ByVal CourtDoc_link As String, _
'                           Optional ByVal home As String = "2665 Hempel Ave, Windermere FL 34786")

      h = 0
        Do While True
            h = h + 1
'            populateEditRow action:=action, cell:=cell, county:=county, saletype:=saletype, salesDate:=salesDate, CaseNumber:=CaseNumber, _
'                            plaintiffName:=PrimaryPlaintiff, defendantName:=defendantName, isCancelled:=isCancelled, sheetName:=sheetName, _
'                            finalJudgment:=finalJudgment, AssessedValue:=AssessedValue, PrimaryPlaintiff:=PrimaryPlaintiff, _
'                            PlaintiffMaxBid:=PlaintiffMaxBid, openingBid:=openingBid, Address:=Address, city:=city, zip:=zip, _
'                            ParcelID:=ParcelID, PA_link:=PA_link, PA_link_Address:=PA_link_Address, CourtDoc_link:=CourtDocAddress
            DoEvents
            If rowExists(salesDate:=salesDate, county:=county, CaseNumber:=CaseNumber) Then
                h = 0
                Exit Do
            ElseIf h > 10 Then
                Debug.Print "Unable to add row after " & h & " attempts."
                    h = 0
                Exit Do
            Else
                ' Erase the last row of the immediate window
               ' If h > 1 Then Application.SendKeys "{DEL}", True
                Debug.Print "Attempt " & h
                SleepUntilReady (h)
            End If
        Loop
End Sub
Private Function rowExists(ByVal county As String, ByVal salesDate As Date, ByVal CaseNumber As String) As Boolean
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlWS As Object
    Dim h As String
    Dim action As String
    On Error GoTo ErrorHandler:
    sheetName = "Sales"
    i = "a"
    Set xlApp = GetObject(, "Excel.Application")
    Set xlWB = ThisWorkbook
    Set xlWS = xlWB.Worksheets(sheetName)
    ' Find primary key columns
    Dim headers As Variant
    headers = xlWS.Rows(1).Value
    Dim pkSaleDate As Variant
    Dim pkCaseNumber As Variant
    Dim pkCounty As Variant
    'this feilds make up the primary key which is always distinct
    pkSaleDate = 1 'Application.match("Sale Date", headers, 0)
    pkCaseNumber = Application.match("Case Number", headers, 0) - 1
    pkCounty = Application.match("County", headers, 0) - 1
    If IsError(pkSaleDate) Or IsError(pkCaseNumber) Then
    '    MsgBox "Primary key columns not found in sheet " & sheetName
        Exit Function
    End If
    i = "b"
    ' Find matching row
    Dim cell As range
    Dim pkRange As range
    Set pkRange = xlWS.range(xlWS.Cells(2, 1), xlWS.Cells(xlWS.Cells(xlWS.Rows.Count, 1).End(xlUp).row, 1))
    'Set pkRange = xlWS.range(xlWS.Cells(2, pkSaleDate), xlWS.Cells(xlWS.Cells(xlWS.Rows.Count, pkSaleDate).End(xlUp).row, pkCaseNumber))
    Dim match As range
    'now see if you can find a row with this primary key
    'a = pkCounty
  '  b = GetColumnNumberByHeaderName(headerName:="Case Number")
    For Each cell In pkRange.Cells
        If IsDate(cell.Value) Then
            If cell.Value = salesDate _
                And cell.Offset(0, pkCounty).Value = county _
                And cell.Offset(0, pkCaseNumber).Value = CaseNumber Then
                Set match = cell
                rowExists = True
                Exit For
            End If
        End If
    Next cell
exit_proc:
    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    Exit Function
ErrorHandler:
    Debug.Print "failed rowExists , see loc 511" & i & " case#:" & CaseNumber; " Err#: " & Err.Number & " " & Err.description
    Err.Clear
    Resume exit_proc
End Function

Sub doPID(cell As range, county As String, ByVal ParcelID As String)
    sheetName = "Sales"
    Dim paLink As String
    On Error GoTo ErrorHandler
    If ParcelID = "" Then Exit Sub
    cell.Offset(0, GetColumnNumberByHeaderName("PID", sheetName)).Value = ParcelID
    cell.Offset(0, GetColumnNumberByHeaderName("PID", sheetName)).Font.Size = 6
    cell.Offset(0, GetColumnNumberByHeaderName("Prop Appr", sheetName)).Font.Size = 6
    paLink = Trim(getPAlink(countyName:=county, PID:=ParcelID))
    If paLink = "" Then
        RemoveHyperlinkFormat cell.Offset(0, GetColumnNumberByHeaderName("PID", sheetName))
        Exit Sub
    End If
    If ParcelID <> "MULTIPLE PARCELS" And county <> "Pinellas" Then
        link = IIf(Trim(ParcelID) = "", "", "=hyperlink(" & Chr(34) & paLink & Chr(34) & ", " & Chr(34) & "<<Click>>" & Chr(34) & ")")
        cell.Offset(0, GetColumnNumberByHeaderName("PAPath", sheetName)).Value = paLink
        cell.Offset(0, GetColumnNumberByHeaderName("Prop Appr", sheetName)).Value = link
        If Trim(paLink) <> "" And UCase(Trim(paLink)) <> "NA" Then
            ParcelID = IIf(Trim(ParcelID) = "", "", "=hyperlink(" & Chr(34) & paLink & Chr(34) & ", " & Chr(34) & ParcelID & Chr(34) & ")")
        End If
    End If
exit_proc:
    Exit Sub
ErrorHandler:
    Debug.Print "failed doPID , see loc 43 "; Err.Number & " " & Err.description
    Resume exit_proc
End Sub



Sub CopyCellsToGoogleSheet(newRow As range)
    'To copy the copyRange to a Google Sheet, you need to first enable the Google Sheets API and set up the OAuth2 credentials. Once that's done, you can use the Google.Apis.Sheets.v4 .NET library to interact with the Google Sheets API from VBA.
    
    Dim ws As Worksheet
    Set ws = newRow.Worksheet
    
    Dim saleDate2Col As Long
    saleDate2Col = WorksheetFunction.match("Sale Date2", ws.Rows(1), 0)
    
    If saleDate2Col > 0 Then
        Dim rowAbove As range
        Set rowAbove = newRow.Offset(-1, 0)
        Dim copyRange As range
        Set copyRange = ws.range(ws.Cells(rowAbove.row, saleDate2Col + 1), ws.Cells(rowAbove.row, ws.Columns.Count))
        
        'Set up the Google Sheets API service
        Dim service As Object
        Set service = CreateObject("Google.Apis.Sheets.v4.SheetsService")
        service.ApplicationName = "MyApp"
        service.UseApplicationDefaultCredentials = True
        
        'Set the ID and name of the target sheet
        Dim spreadsheetId As String
        spreadsheetId = "1g6t4LXHJnSN1T7PXgItiJpxQqF4reD4IOhJu-F0Nbpo"
        Dim sheetName As String
        sheetName = "Potential Bids"
        
        'Define the target range in A1 notation
        Dim targetRange As String
        targetRange = "A1"
        
        'Calculate the target range based on the size of the copy range
        Dim numRows As Long
        numRows = copyRange.Rows.Count
        Dim numCols As Long
        numCols = copyRange.Columns.Count
        targetRange = targetRange & ":" & Cells(numRows, numCols).address(False, False)
        
        'Copy the values to the target range in the Google Sheet
        Dim valueRange As Object
        Set valueRange = CreateObject("Google.Apis.Sheets.v4.Data.ValueRange")
        valueRange.Values = copyRange.Value
        service.spreadsheets.Values.Update spreadsheetId, sheetName & "!" & targetRange, valueRange, "USER_ENTERED"
    Else
        MsgBox "Sale Date2 column not found!"
    End If
    
End Sub


Sub populateLink(cell As range, path As String, Optional politeText As String = "<<Click>>", Optional fontSize As Integer = 6)
   ' Dim cell As Range
  '  Set cell = match.Offset(0, GetColumnNumberByHeaderName("Sale", sheetName))
    With cell
        .Value = makeHyperlinkFormula(path:=path, politeText:=politeText)
        .Font.Size = fontSize
        .HorizontalAlignment = xlCenter
    End With
End Sub

Sub populateLink2(range As range, headerName As String, path As String, Optional sheetName As String = "Sales", Optional politeText As String = "<<Click>>", Optional fontSize As Integer = 6)
    range.Offset(0, GetColumnNumberByHeaderName(headerName & "Path", sheetName)).Value = path
    range.Offset(0, GetColumnNumberByHeaderName(headerName & "Path", sheetName)).Font.Size = 2
    Dim cell As range
    Set cell = range.Offset(0, GetColumnNumberByHeaderName(headerName, sheetName))
    With cell
        .Value = makeHyperlinkFormula(path:=path, politeText:=politeText)
        .Font.Size = fontSize
        .HorizontalAlignment = xlCenter
    End With
End Sub



Function urlDateHasAlreadyBeenProcessed(countyName As String, salesType As String, Optional hoursToWait As Integer = 1) As Boolean
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Activity")
    
    ' Get the row number of the county in the appropriate column
    Dim countyColumn As range
    Select Case salesType
        Case "Tax Deed"
            Set countyColumn = ws.range("a:a")
        Case "Foreclosure"
            Set countyColumn = ws.range("a:a")
        Case Else
            ' Invalid sales type
            Exit Function
    End Select
    Dim countyRow As range
    Set countyRow = countyColumn.Find(countyName, LookIn:=xlValues, LookAt:=xlWhole)
    If countyRow Is Nothing Then
        ' County not found
        Exit Function
    End If
    
    ' Get the last run date in the appropriate column
    Dim lastRunDateColumn As range
    Select Case salesType
        Case "Tax Deed"
            Set lastRunDateColumn = ws.range("C:C")
        Case "Foreclosure"
            Set lastRunDateColumn = ws.range("B:B")
    End Select
    Dim lastRunDateCell As range
    Set lastRunDateCell = ws.Cells(countyRow.row, lastRunDateColumn.Column)
    Dim lastRunDate As Date
    lastRunDate = lastRunDateCell.Value
    urlDateHasAlreadyBeenProcessed = True
    If Now() > lastRunDate + hoursToWait And Not (IsEmpty(lastRunDateCell.Value) Or lastRunDateCell.Value = "") Then
        urlDateHasAlreadyBeenProcessed = False
        lastRunDateCell.Value = Now()
    End If
    ' Compare the dates and update the last run date if necessary
    If (IsEmpty(lastRunDateCell.Value) Or lastRunDateCell.Value = "") Then
        urlDateHasAlreadyBeenProcessed = False
        lastRunDateCell.Value = Now()
    End If


End Function





Function fileDateHasAlreadyBeenProcessed(localFileName As String, countyName As String, salesType As String) As Boolean
    ' Get the last modified date of the file
    Dim lastTimeDownloaded As Date
    lastTimeDownloaded = FileDateTime(localFileName)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Activity")
    
    ' Get the row number of the county in the appropriate column
    Dim countyColumn As range
    Select Case salesType
        Case "Tax Deed"
            Set countyColumn = ws.range("a:a")
        Case "Foreclosure"
            Set countyColumn = ws.range("a:a")
        Case Else
            ' Invalid sales type
            Exit Function
    End Select
    Dim countyRow As range
    Set countyRow = countyColumn.Find(countyName, LookIn:=xlValues, LookAt:=xlWhole)
    If countyRow Is Nothing Then
        ' County not found
        Exit Function
    End If
    
    ' Get the last run date in the appropriate column
    Dim lastRunDateColumn As range
    Select Case salesType
        Case "Tax Deed"
            Set lastRunDateColumn = ws.range("C:C")
        Case "Foreclosure"
            Set lastRunDateColumn = ws.range("B:B")
    End Select
    Dim lastRunDateCell As range
    Set lastRunDateCell = ws.Cells(countyRow.row, lastRunDateColumn.Column)
    Dim lastRunDate As Date
    lastRunDate = lastRunDateCell.Value
    fileDateHasAlreadyBeenProcessed = True
    If lastTimeDownloaded > lastRunDate And Not (IsEmpty(lastRunDateCell.Value) Or lastRunDateCell.Value = "") Then
        fileDateHasAlreadyBeenProcessed = False
        lastRunDateCell.Value = lastTimeDownloaded
    End If
    ' Compare the dates and update the last run date if necessary
    If (IsEmpty(lastRunDateCell.Value) Or lastRunDateCell.Value = "") Then
        fileDateHasAlreadyBeenProcessed = False
        lastRunDateCell.Value = lastTimeDownloaded
    End If
    
    
End Function

Function getPAlink(countyName As String, Optional PID As String, Optional address As String, Optional What2Get As String = "ByPID") As String
    On Error GoTo ErrorHandler
    Dim ee As String
    If What2Get = "ByPID" And PID <> "" Then
        getPAlink = Replace(getLinks(countyName, "PAtemplate"), "<<PID>>", PID)
        getPAlink = IIf(getPAlink = "", getLinks(countyName, "PA"), getPAlink)
    ElseIf What2Get = "ByAddress" And address <> "" Then
        getPAlink = Replace(getLinks(countyName, "paTemplateAddr"), "<<PID>>", PID)
        getPAlink = IIf(address <> "", Replace(getPAlink, "<<address>>", Replace(Trim(address), " ", "%20")), "")
    Else
        getPAlink = getLinks(countyName, "PA")
    End If
exit_proc:
    Exit Function
ErrorHandler:
    Debug.Print "failed formatSaleDateCell , see loc 42 " & ee; Err.Number & " " & Err.description
    Err.Clear
    Resume exit_proc
End Function

Function getCourtlink(countyName As String, Optional CaseNumber As String) As String
    If CaseNumber <> "" Then
  '      getCourtlink = Replace(getLinks(countyName, "Court Docs"), "<<Click>>")
    Else
        getCourtlink = getLinks(countyName, "Court Docs")
    End If
End Function

Function getLinks(countyName As String, LinkType As String) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Activity")
    
    ' Get the row number of the county in the appropriate column
    Dim countyColumn As range
    Set countyColumn = ws.range("a:a")
    Dim countyRow As range
    Dim linkAddressColumn As range
    Set countyRow = countyColumn.Find(countyName, LookIn:=xlValues, LookAt:=xlWhole)
    If countyRow Is Nothing Then
        ' County not found
        Exit Function
    End If
    
    Dim columnNumber As Integer
    columnNumber = GetColumnNumberByHeaderName(headerName:=LinkType, worksheetName:="Activity") + 1
    Set linkAddressColumn = ws.range(columnNumber & ":" & columnNumber)
    getLinks = ws.Cells(countyRow.row, columnNumber)
    Dim lastRunDateColumn As range
    Select Case LinkType
        Case "List"
            Set linkAddressColumn = ws.range("B:B")
        Case "PA"
            Set linkAddressColumn = ws.range("h:h")
        Case "PAtemplate"
            Set linkAddressColumn = ws.range("K:K")
        Case "Court Docs"
            Set linkAddressColumn = ws.range("i:i")
        Case "public records"
            Set linkAddressColumn = ws.range("E:E")
    End Select
    Dim linkAddressCell As range
    Set linkAddressCell = ws.Cells(countyRow.row, linkAddressColumn.Column)
    getLinks = linkAddressCell
End Function

Public Function isHOA(name As String, Optional county As String = "") As String
    Dim keywords As Variant
    isHOA = ""
    If name = "" Then Exit Function
    keywords = Array("RESIDENTS ASSOCIATION", "NEIGHBORHOOD ASSOCIATION", "OWNERS' ASSOCIATION", _
    "HOMEOWNERS", "CONDO", "HOA", "COA", "TOWNHOMES ASSOCIATION", "OWNERS ASSOCIATION", "MASTER ASSOCIATION", _
    "owner's Association", "Community Association", "HOMEOWNERS ASSOCIATION", "OWNERSHIP RESORTS", "HIILTON RESORTS", _
     "Resort Association", "POINCIANA VILLAGES", "VACATION RESORTS", "MASTER CORPORATION")
    
    
    If county <> "" Then
        isHOA = getOkHOAtype(name:=name, county:=county)
    End If
    
    If isHOA = "" Then
        For i = 0 To UBound(keywords)
            If InStr(UCase(name), UCase(keywords(i))) > 0 Then
                isHOA = "x"
                Exit For
            End If
        Next i
    End If
End Function

Function getOkHOAtype(name As String, county As String) As String
    Dim data As Variant
    Dim description As String
    Dim row As range
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    'Read the data from the worksheet
    data = Sheets("vars").range("okHOA").Value
    
    'Loop through the data and find the matching row
    For i = LBound(data, 1) + 1 To UBound(data, 1)
       ' Set row = data(i, 1)
        
        If data(i, 1) = name And data(i, 2) = county Then
            description = data(i, 3)
            Exit For
        End If
    Next i
    
    'If the row was not found, return an empty string
    If description = "" Then
        getOkHOAtype = ""
    Else
        getOkHOAtype = description
    End If
exit_proc:
    Exit Function
ErrorHandler:
    Debug.Print "Failed getOkHOAtype, see loc 19: " & Err.Number & " " & Err.description
    Err.Clear
    Resume exit_proc
End Function

Public Function isOkHOAold(name As String, county As String) As String
    
    Dim OkHOA As range
    
    On Error GoTo ErrorHandler
    'Declare the OkHOA object
    Set OkHOA = range("OkHOA")
    ee = "a"
    'Find the row number of the name string
    Dim nameRow As Long
    nameRow = OkHOA.Find(name).row
    
    'Find the row number of the county string
    Dim countyRow As Long
    countyRow = OkHOA.Find(county).row + 1
    
    'Return the value of the OkHOA object at the row number specified by the county string
    If countyRow > 0 Then
 '     getOkHOAtype = OkHOA(countyRow)
    Else
    '  getOkHOAtype = ""
    End If

exit_proc:
    Exit Function
ErrorHandler:
    Debug.Print "failed GetColumnNumberByHeaderName , see loc 19 " & ee; Err.Number & " " & Err.description
    Err.Clear
    Resume exit_proc
End Function


Public Function GetColumnNumberByHeaderNamePLUS1(ByVal headerName As String, Optional ByVal worksheetName As String = "Sales") As Long
'this is a fix as the below function is returning the wrong value and i have jsut been adding one. try to correct over time
    GetColumnNumberByHeaderNamePLUS1 = GetColumnNumberByHeaderName(headerName, worksheetName) + 1
End Function


Public Function GetColumnNumberByHeaderName(ByVal headerName As String, Optional ByVal worksheetName As String = "Sales") As Long
   'DEPRECATED
    ' Declare variables
    Dim ee As String
    Dim headerRange As range
    Dim headerCell As range
    Dim columnNumber As Long
    On Error GoTo ErrorHandler
    ' Set the header range to the first row of the specified worksheet
    Set headerRange = Worksheets(worksheetName).range("1:1")
    
    ' Search for the header name in the header range, including hidden cells
    Set headerCell = headerRange.Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False, SearchFormat:=True)
    
    ' If the header is found, return the column number
    If Not headerCell Is Nothing Then
        columnNumber = headerCell.Column - 1
    Else
        ' If the header is not found, return 0
        columnNumber = 0 ' Indicates header not found
    End If
    
    ' Return the column number
    GetColumnNumberByHeaderName = columnNumber
exit_proc:
    Exit Function
ErrorHandler:
    Debug.Print "failed GetColumnNumberByHeaderName , see loc 53 " & ee; Err.Number & " " & Err.description
    Err.Clear
    Resume exit_proc
End Function

Public Function GetColumnByHeaderName2(ByVal headerName As String, Optional ByVal worksheetName As String = "Sales", Optional ByVal getLetter As Boolean = False) As Variant
    ' Declare variables
    Dim headerRange As range
    Dim headerCell As range
    Dim columnNumber As Long
    Dim columnLetter As String
    
    On Error GoTo ErrorHandler
    
    ' Set the header range to the first row of the specified worksheet
    Set headerRange = Worksheets(worksheetName).Rows(1)
    
    ' Search for the header name in the header range, including hidden cells
    Set headerCell = headerRange.Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False, SearchFormat:=True)
    
    ' If the header is found, retrieve the column number or letter based on the 'getLetter' parameter
    If Not headerCell Is Nothing Then
        columnNumber = headerCell.Column
        columnLetter = Split(Cells(1, columnNumber).address, "$")(1)
        
        If getLetter Then
            GetColumnByHeaderName2 = columnLetter
        Else
            GetColumnByHeaderName2 = columnNumber
        End If
    Else
        ' If the header is not found, return 0 or an empty string based on the 'getLetter' parameter
        If getLetter Then
            GetColumnByHeaderName2 = ""
        Else
            GetColumnByHeaderName2 = 0 ' Indicates header not found
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Failed GetColumnByHeaderName2: " & Err.Number & " - " & Err.description
    Err.Clear
    Resume Next
End Function

Function Slice(arr As Variant, start As Long, Optional length As Long) As Variant
    If IsMissing(length) Then
        length = UBound(arr) - start + 1
    End If
    Dim result() As Variant
    ReDim result(0 To length - 1)
    Dim i As Long
    For i = 0 To length - 1
        result(i) = arr(start + i)
    Next i
    Slice = result
End Function

Function GetPDFContent(ByVal pdfUrl As String) As String
'used to return the text from a url rendering for processing
    'Declare variables
    Dim IE As Object
    Dim objShell As Object
    Dim objWsh As Object
    Dim strClipboard As String
    Dim content As String

    'Create a new instance of Internet Explorer
    Set IE = CreateObject("InternetExplorer.Application")
    
    'Navigate to the web URL of the PDF file
    IE.Navigate pdfUrl

    'Wait for the page to load completely
    Do While IE.ReadyState <> 4
        DoEvents
    Loop

    'Send the "Ctrl+C" keystrokes to copy the contents of the PDF file to the clipboard
    Set objShell = CreateObject("WScript.Shell")
    objShell.SendKeys "^a^c"
    
    'Retrieve the contents of the clipboard
    Set objWsh = CreateObject("WScript.Shell")
    strClipboard = objWsh.exec("cmd /c echo off | clip").StdOut.ReadAll
    
    'Assign the contents of the clipboard to the "content" variable
    content = strClipboard
    
    'Close Internet Explorer
    IE.Quit
    
    'Return the contents of the PDF file as a string variable
    GetPDFContent = content
End Function

Function IsStreetType(ByVal word As String) As Boolean

'Declare variables
Dim isValid As Boolean

'Set default value
isValid = False

'Check if the word is a valid street type
If InStr(UCase(word), "AVE") > 0 Or _
    InStr(UCase(word), "ST") > 0 Or _
    InStr(UCase(word), "DR") > 0 Or _
    InStr(UCase(word), "RD") > 0 Or _
    InStr(UCase(word), "CT") > 0 Or _
    InStr(UCase(word), "LN") > 0 Or _
    InStr(UCase(word), "PL") > 0 Or _
    InStr(UCase(word), "WY") > 0 Or _
    InStr(UCase(word), "HWY") > 0 Or _
    InStr(UCase(word), "PKWY") > 0 Or _
    InStr(UCase(word), "CIR") > 0 Or _
    InStr(UCase(word), "SQ") > 0 Or _
    InStr(UCase(word), "TRL") > 0 Or _
    InStr(UCase(word), "HTS") > 0 Or _
    InStr(UCase(word), "APT") > 0 Or _
    InStr(UCase(word), "STE") > 0 Or _
    InStr(UCase(word), "RM") > 0 Or _
    InStr(UCase(word), "FL") > 0 Then

        'The word is a valid street type
        isValid = True
    End If

'Return the result
IsStreetType = isValid

End Function

Sub Highlands()
'ProcessHighlandsCounty "C:\Users\philk\Desktop\HighlandsSalesList.txt"
End Sub
Sub ProcessHighlandsCounty(rt As rowType) ', ByVal rt.FilePath As String, Optional ByVal rt.sheetName As String = "Sales", Optional rt.saletype As String = "Foreclosure")

'Declare variables
'Dim rt.county As String
'Dim rt.saletype As Date
'Dim rt.plaintiffName As String
'Dim rt.defendantName As String
'Dim rt.isCancelled As Boolean
Dim amount As String
Dim legalDescription As String
'Dim rt.city As String

'Declare other variables
Dim dataLine As String
Dim lineNumber As Long
'Dim rt.CaseNumber As String
Dim salesDataStart As Boolean
Dim salesDataEnd As Boolean
Dim dontAdvance As Boolean: dontAdvance = False
Dim fileNum As Integer
Dim tblRow As Long
Dim regex As Object
Dim regex2 As Object
'Dim rt.finalJudgment As String
Dim matches As Object
Dim match As Object
Dim modifiedString As String
'Dim PropApprAddress As String:  PropApprrt.Address = getPAlink(countyName:=rt.county)
'Dim CourtDocAddress As String:  CourtDocrt.Address = getCourtlink(countyName:=rt.county)
'rt.county = "Highlands"
Call PdfGetter(rt:=rt)
rt.PropApprAddress = getPAlink(countyName:=rt.county)
rt.CourtDocAddress = getCourtlink(countyName:=rt.county)
rt.filePath = GetSaleTypeOrCounty2("filepath")
rt.filePath = "C:\Users\philk\Desktop\HighlandsSalesList.txt"


' Create a regular expression object to match currency amounts
Set regex = CreateObject("VBScript.RegExp")
regex.Pattern = "\$[\d,]+\.\d{2}" ' Matches amounts in the format $35,175.84
Set regex2 = CreateObject("VBScript.RegExp")
regex2.Pattern = "\b2\w*?X\b"
'Dim rt.Address As String

' Define the list of cities from Highlands rt.county, FL
Dim highlandsCities As Variant
highlandsCities = Array("Avon Park", "Avon Par", "Lake Placid", "Sebring")
Dim potentialPlaintiffs As Variant
potentialPlaintiffs = Array("US BANK TRUST NA", "CARRINGTON MORTGAGE SERVICES LLC", "SUN N LAKE OF SEBRING", "US BANK NATIONAL", "LAKEVIEW LOAN SERVICING", "THE BANK OF NEW YORK", "THE BANK OF NEW YORK MELLON")

'Set initial values
rt.county = "Highlands"
salesDataStart = False
salesDataEnd = False
tblRow = 2

'Open file
fileNum = FreeFile()
Open rt.filePath For Input As #fileNum

'Loop through file
While Not EOF(fileNum)
    'Read line from file
    If Not dontAdvance Then
        Line Input #fileNum, dataLine
    End If
    
    'see if this is the line with the address
    For Each pltf In potentialPlaintiffs
      If InStr(UCase(dataLine), UCase(pltf)) > 1 Then
          ' If the line ends with a Highlands rt.county city, then it is the rt.Address line
          rt.plaintiffName = pltf
          dataLine = Replace(dataLine, pltf, "")
          Exit For
      End If
    Next pltf
    
    'Check if line indicates start of sales data, this is where the sales data starts and the header ends
    If InStr(dataLine, "Amount") > 0 Then
        salesDataStart = True
        lineNumber = 0
    End If
    'Process sales data
If salesDataStart And Not salesDataEnd Then
    lineNumber = lineNumber + 1
    If lineNumber > 1 Then 'skip first row after table header
        'Split line into variables
        Dim data() As String
        If Not dontAdvance Then
            data = Split(dataLine, vbTab)
        End If
        'rt.saletypeStr = Trim(data(0))
        If LineAfterDate And IsNumeric(Trim(data(0))) Then
            If Val(Trim(data(0))) >= year(Now()) And Val(Trim(data(0))) < 2050 Then
                saletypeStr = DateSerial(Trim(data(0)), month(Replace(saletypeStr, ",", "")), day(Replace(saletypeStr, ",", "")))
            End If
            LineAfterDate = False
        ElseIf IsDate(Replace(data(0), ",", "")) Then
        'problem is when date is long format and on same line as other stuff.  write logic to get is out of unused date or to check is first word is a month name and then test if all left of the third space is a date
            saletypeStr = Trim(data(0))
            LineAfterDate = True
        ElseIf regex2.Test(data(0)) Then
            If regex2.Execute(data(0))(0) <> "" Then
            ' dontAdvance = True
             'Dim regex As Object
             'Dim matches As Object
             'Dim match As Object
             'Dim modifiedString As String
             
             ' Create a regular expression object to match words that begin with a "2" and end with an "X"
             'Set regex = CreateObject("VBScript.RegExp")
             'regex.Pattern = "\b2\w*?X\b"
             
             ' Find all matches of the pattern in data(0)
             Set matches = regex2.Execute(data(0))
             
             ' Replace each match with a pipe symbol
             modifiedString = data(0)
             For Each match In matches
               '  Debug.Print "Case Number:  " & match
                 CaseNumberNew = match
                 modifiedString = Replace(modifiedString, match.Value, "|")
             Next match
             dontAdvance = Len(modifiedString) > 1
             If dontAdvance Then data(0) = modifiedString
             ' Display the modified string for testing purposes
             'Debug.Print "Modified data(0): " & modifiedString
             If rt.CaseNumber <> "" And rt.CaseNumber <> CaseNumberNew Then
'                 Debug.Print saletypeStr
'                 Debug.Print "Case Number:  " & rt.CaseNumber
'                 Debug.Print rt.Address & ", " & rt.city
'                 Debug.Print "rt.finalJudgment: " & rt.finalJudgment
'                 Debug.Print "Paintifff: " & rt.plaintiffName
'                 Debug.Print "Defendent: " & rt.defendantName
'                Debug.Print "Unused Data: " & data(0)
                rt.isCancelled = False
                rt.salesDate = CDate(saletypeStr)
                UpdateOrAddRow2 rt:=rt ', rt.county:=rt.county, salesDate:=rt.salesDate, rt.CaseNumber:=rt.CaseNumber, rt.plaintiffName:=rt.plaintiffName, _
                           rt.defendantName:=rt.defendantName, rt.sheetName:=rt.sheetName, _
                           finalrt.finalJudgment:=Val(rt.finalJudgment), _
                           rt.Address:=rt.Address, city:=rt.city, PA_link_rt.Address:=rt.PropApprrt.Address, CourtDoc_link:=rt.CourtDocrt.Address
                 'Debug.Print "**************************"
                 rt.CaseNumber = ""
                 rt.address = ""
                 rt.city = ""
                 rt.finalJudgment = 0#
                 rt.plaintiffName = ""
                 rt.defendantName = ""
                 unusedData = ""
             End If
             If CaseNumberNew <> "" Then
                 rt.CaseNumber = CaseNumberNew
             End If
            End If
        ElseIf regex.Test(data(0)) Then
            dontAdvance = TrueCaseNumberNew
            rt.finalJudgment = regex.Execute(data(0))(0)
            data(0) = regex.Replace(data(0), "|")
            rt.finalJudgment = Replace(rt.finalJudgment, "$", "") ' Remove the $ symbol from the rt.finalJudgment amount
            rt.finalJudgment = Replace(rt.finalJudgment, ",", "") ' Remove commas from the rt.finalJudgment amount
            ' Display the extracted rt.finalJudgment amount and the modified data(0) string for testing purposes
            rt.defendantName = Trim(Split(data(0), "|")(0))
            If InStr(rt.defendantName, " ") > 0 Then
                parts = Split(rt.defendantName, " ")
                ct = UBound(parts)
                rt.defendantName = Trim(parts(ct) & " " & parts(ct - 1))
                'rt.defendantName = Trim(Split(rt.defendantName, " ")(0) & " " & Split(rt.defendantName, " ")(1))
            End If
            data(0) = Replace(data(0), rt.defendantName, "")

        Else
            dontAdvance = False
            For Each city In highlandsCities
            'if this like ", Sebring" or  the word before " Sebring" is a valid street type
                If (UCase(Right(Trim(data(0)), Len(city) + 2)) = ", " & UCase(city)) _
                Or ((UCase(Right(Trim(data(0)), Len(city) + 1)) = " " & UCase(city)) And IsStreetType(Mid(Trim(data(0)), InStr(1, data(0), " ") + 1, Len(city) - 2))) Then
                    ' If the line ends with a Highlands rt.county city, then it is the rt.Address line
                    rt.city = city
                    rt.address = Replace(Trim(data(0)), ", " & Trim(UCase(city)), "")
                 '   Debug.Print rt.saletypeStr & ", " & rt.Address, " " & rt.city
                    Exit For
                End If
            Next city
        End If
        
'        Debug.Print rt.saletypeStr
'        Debug.Print "Case Number:" & rt.CaseNumber
'        Debug.Print rt.Address, " " & rt.city
'        Debug.Print "rt.finalJudgment: " & rt.finalJudgment
'     '   Debug.Print "Modified data(0): " & data(0)
'         Debug.Print "**************************"
        ' year = isyear((Replace(data(0), ",", ""))

        tblRow = tblRow + 1 'increment table row counter
    End If
End If

    Wend
    
    'Close file
    Close #fileNum
    
End Sub

Sub ExtractZestimate()
    ' Open the webpage
    Dim IE As Object
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = True
    IE.Navigate "https://www.zillow.com/homedetails/523-Canterbury-Ln-Kissimmee-FL-34741/46322231_zpid/"

    ' Wait for the page to load
    Do While IE.ReadyState <> 4 Or IE.Busy
        Application.Wait DateAdd("s", 1, Now)
    Loop

    ' Wait for the Zestimate element to appear
    Dim zestimateElement As Object
    Do Until Not zestimateElement Is Nothing
        Set zestimateElement = IE.Document.querySelector("a[data-za-label='Zestimates']")
        If Not zestimateElement Is Nothing Then
            Dim zestimateValue As String
            zestimateValue = zestimateElement.NextSibling.innerText
            Debug.Print "Zestimate: " & zestimateValue
            Exit Do
        Else
            Application.Wait DateAdd("s", 1, Now)
        End If
    Loop

    If zestimateElement Is Nothing Then
        Debug.Print "Could not find Zestimate."
    End If

    ' Close the browser
    IE.Quit
End Sub

Sub CopyGoogleTemplateCellsToNewRow(newRow As range)
    
    Dim ws As Worksheet
    Set ws = newRow.Worksheet
    
    Dim saleDate2Col As Long
    saleDate2Col = WorksheetFunction.match("Sale Date2", ws.Rows(1), 0)
    
    If saleDate2Col > 0 Then
        Dim rowAbove As range
        Set rowAbove = newRow.Offset(-1, 0)
        Dim copyRange As range
        Set copyRange = ws.range(ws.Cells(rowAbove.row, saleDate2Col + 1), ws.Cells(rowAbove.row, ws.Columns.Count))
        copyRange.Copy newRow.Offset(0, saleDate2Col)
    Else
        MsgBox "Sale Date2 column not found!"
    End If
    
End Sub


Function getSaleHyperlink(county As String, SaleDate As Date, Optional saleType As String = "foreclosure", _
                          Optional asFormula As Boolean = False) As String
    Dim hyperlink As String
    Dim ee As String
    Dim sDate As String: sDate = Format(SaleDate, "mm/dd/yyyy")
    On Error GoTo ErrorHandler
    ''https://pinellas.realforeclose.com/index.cfm?zaction=AUCTION&Zmethod=DAYLIST&AUCTIONDATE=03/17/2023
    hyperlink = "https://<<COUNTY>>.realforeclose.com/index.cfm?zaction=AUCTION&zmethod=PREVIEW&AuctionDate=" & sDate
    If LCase(saleType) = "foreclosure" Then
'        hyperlink = _
'            IIf(UCase(county) = "BREVARD", "http://vweb2.brevardclerk.us/Foreclosures/foreclosure_sales.html", _
'            IIf(UCase(county) = "LAKE", "https://www.lakecountyclerk.org/record_searches/public_sales_calendar/", _
'            IIf(UCase(county) = "OSCEOLA", "https://courts.osceolaclerk.com/reports/CivilMortgageForeclosuresWeb.pdf", _
'            IIf(county <> "", "https://" & _
'                IIf(UCase(county) = "ORANGE", "www.myorangeclerk", _
'                IIf(UCase(county) = "POLK", "polk", _
'                IIf(UCase(county) = "PINELLAS", "pinellas", _
'                IIf(UCase(county) = "SEMINOLE", "seminole", _
'                IIf(UCase(county) = "FLAGLER", "flagler", _
'                IIf(UCase(county) = "VOLUSIA", "volusia", "")))))) & _
'                ".realforeclose.com/index.cfm?zaction=AUCTION&zmethod=PREVIEW&AuctionDate=" & sDate, _
'            "<<Click Me>>"))))
        Select Case county
        Case "BREVARD"
            hyperlink = "http://vweb2.brevardclerk.us/Foreclosures/foreclosure_sales.html"
        Case "LAKE"
            hyperlink = "https://www.lakecountyclerk.org/record_searches/public_sales_calendar/"
        Case "OSCEOLA"
            hyperlink = "https://courts.osceolaclerk.com/reports/CivilMortgageForeclosuresWeb.pdf"
        Case "LAKE"
            hyperlink = "https://www.lakecountyclerk.org/record_searches/public_sales_calendar/"
        Case "ORANGE"
            county = "www.myorangeclerk"
        Case Else
            ' Invalid sales type
       '     Exit Function
        End Select
        hyperlink = Replace(hyperlink, "<<COUNTY>>", county)
    Else
'        getSaleHyperlink = ""
'        hyperlink = _
'            IIf(UCase(county) = "PINELLAS", "https://pinellas.realtaxdeed.com/index.cfm?zaction=AUCTION&Zmethod=DAYLIST&AUCTIONDATE=" & sDate, "")
           ' https://manatee.realforeclose.com/index.cfm?zaction=AUCTION&Zmethod=PREVIEW&AUCTIONDATE=04/03/2023
        hyperlink = "https://<<COUNTY>>.realtaxdeed.com/index.cfm?zaction=AUCTION&Zmethod=DAYLIST&AUCTIONDATE="
        hyperlink = Replace(hyperlink, "<<COUNTY>>", county) & sDate
    End If
    If asFormula Then
        hyperlink = makeHyperlinkFormula(path:=hyperlink, politeText:="<<Click>>")
    End If
    getSaleHyperlink = hyperlink
exit_proc:
    Exit Function
ErrorHandler:
    Debug.Print "failed getSaleHyperlink , see loc 54 " & ee; Err.Number & " " & Err.description
    Err.Clear
    Resume exit_proc
End Function
    
Function makeHyperlinkFormula(path As String, Optional politeText As String) As String
    Dim asFormula As Boolean: asFormula = True
    politeText = IIf(LCase(politeText) = "c", "<<Click>>", politeText)
    prompt = IIf(politeText <> "", politeText, path)
    makeHyperlinkFormula = IIf(asFormula, "=HYPERLINK(""" & path & """,""" & prompt & """)", path)
End Function

Sub clearFilter(Optional ws As Worksheet)
    Dim col As Variant
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets("Sales")
  
    For Each col In ws.UsedRange.Columns
        col.AutoFilter Field:=1, Criteria1:="*"
    Next col
    
    With ws
        If .AutoFilterMode Then .AutoFilterMode = False ' turn off existing filter if it exists
        .UsedRange.AutoFilter Field:=1, Criteria1:="<>""" ' apply filter to first column
    End With
    
End Sub



Sub toggleFilter(Optional ByVal filterOn As Boolean = False)
    Dim ws As Worksheet
    On Error GoTo ErrorHandler
    Set ws = ActiveSheet
    
    If filterOn Then
        ws.AutoFilterMode = True
    Else
        If ws.AutoFilterMode Then
            ws.AutoFilterMode = False
        End If
    End If
exit_proc:
    Exit Sub
ErrorHandler:
    Debug.Print "failed toggleFilter , see loc 50 " & ee; Err.Number & " " & Err.description
    Err.Clear
    Resume exit_proc
End Sub

Sub SortByForeclosureColumn()
    Dim ws As Worksheet
    Dim colNum As Long
    Dim cell As range
    Set ws = ThisWorkbook.Worksheets("Activity")
    
    ' Find column number of "Foreclosure" header
    colNum = 0
    For Each cell In ws.range("1:1")
        If cell.Value = "Foreclosure" Then
            colNum = cell.Column
            Exit For
        End If
    Next cell
    
    ' If "Foreclosure" header not found, exit sub
    If colNum = 0 Then
        MsgBox "Column header 'Foreclosure' not found.", vbExclamation
        Exit Sub
    End If
    
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Cells(1, colNum), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.UsedRange
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub FilterSalesWorksheet()
    Dim ws As Worksheet
    Dim filterRange As range
    Dim today As Date
    
    Set ws = ThisWorkbook.Worksheets("Sales")
    today = Date
    
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    SortBySaleDate
   ' button4.Caption = "Unfilter"
    
    ' Find the range to filter
    Set filterRange = ws.range("A1").CurrentRegion
    
    ' Apply the filters
    filterRange.AutoFilter Field:=GetColumnNumberByHeaderName("Cancelled") + 1, Criteria1:="<>x"
    filterRange.AutoFilter Field:=GetColumnNumberByHeaderName("HOA") + 1, Criteria1:="<>x"
    filterRange.AutoFilter Field:=GetColumnNumberByHeaderName("Sales Type") + 1, Criteria1:="<>Tax Deed"
    filterRange.AutoFilter Field:=GetColumnNumberByHeaderName("Interested") + 1, Criteria1:="=yes", Operator:=xlOr, Criteria2:="="
    filterRange.AutoFilter Field:=1, Criteria1:=">=" & today
    
    ' Turn off autofilter mode if there are no visible rows
    If filterRange.SpecialCells(xlCellTypeVisible).Rows.Count <= 1 Then
        filterRange.AutoFilter Field:=GetColumnNumberByHeaderName("Cancelled")
    End If
End Sub

Sub FilterTaxDeed()
    Dim ws As Worksheet
    Dim filterRange As range
    Dim today As Date
    
    Set ws = ThisWorkbook.Worksheets("Sales")
    today = Date
    
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    SortBySaleDate
   ' button4.Caption = "Unfilter"
    
    ' Find the range to filter
    Set filterRange = ws.range("A1").CurrentRegion
    
    ' Apply the filters
    filterRange.AutoFilter Field:=GetColumnNumberByHeaderName("Cancelled") + 1, Criteria1:="<>x"
    filterRange.AutoFilter Field:=GetColumnNumberByHeaderName("Sales Type") + 1, Criteria1:="Tax Deed"
    filterRange.AutoFilter Field:=GetColumnNumberByHeaderName("Interested") + 1, Criteria1:="=yes", Operator:=xlOr, Criteria2:="="
    filterRange.AutoFilter Field:=1, Criteria1:=">=" & today
    
    ' Turn off autofilter mode if there are no visible rows
    If filterRange.SpecialCells(xlCellTypeVisible).Rows.Count <= 1 Then
        filterRange.AutoFilter Field:=GetColumnNumberByHeaderName("Cancelled")
    End If
End Sub

Sub FilterSalesWorksheet2()
    Dim ws As Worksheet
    Dim filterRange As range
    Dim today As Date
    
    Set ws = ThisWorkbook.Worksheets("Sales")
    today = Date
    
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    'Call SortBy(fieldName:="Equity %", Order:="decending")
    Call SortBy(fieldName:="Sale Date", Order:="ascending")
   ' button4.Caption = "Unfilter"
    
    ' Find the range to filter
    Set filterRange = ws.range("A1").CurrentRegion
    
    ' Apply the filters
    filterRange.AutoFilter Field:=GetColumnNumberByHeaderNamePLUS1("Cancelled"), Criteria1:="<>x"
    filterRange.AutoFilter Field:=GetColumnNumberByHeaderNamePLUS1("HOA"), Criteria1:="<>x"
    filterRange.AutoFilter Field:=GetColumnNumberByHeaderName("Sales Type") + 1, Criteria1:="<>Tax Deed"
    filterRange.AutoFilter Field:=GetColumnNumberByHeaderNamePLUS1("Interested"), Criteria1:="=yes", Operator:=xlOr, Criteria2:="="
    filterRange.AutoFilter Field:=1, Criteria1:=">=" & today
    filterRange.AutoFilter Field:=GetColumnNumberByHeaderNamePLUS1("Zestimate"), Criteria1:="<> 'none'", Operator:=xlOr, Criteria2:="<>"
    filterRange.AutoFilter Field:=GetColumnNumberByHeaderNamePLUS1("Judgement"), Criteria1:="<>0"
    filterRange.AutoFilter Field:=GetColumnNumberByHeaderNamePLUS1("Equity"), Criteria1:=">0"
    filterRange.AutoFilter Field:=GetColumnNumberByHeaderNamePLUS1("Equity %"), Criteria1:=">0.3"
    
    ' Turn off autofilter mode if there are no visible rows
    If filterRange.SpecialCells(xlCellTypeVisible).Rows.Count <= 1 Then
        filterRange.AutoFilter Field:=GetColumnNumberByHeaderNamePLUS1("Cancelled")
    End If
    
    
    Call SortBy(fieldName:="Sale Date", Order:="ascending")
End Sub

Sub GatherStatsAndFormat()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentDate As Date
    
    ' Set the worksheet variable
    Set ws = ThisWorkbook.Sheets("Sales") ' Change "Sales" to your actual sheet name
    
    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Get the current date and time
    currentDate = Now
    
    ' Loop through each row (starting from row 2)
    For i = 2 To lastRow
        ' Check if the date in column A is greater than the current date

        
        If IsDate(ws.Cells(i, 1).Value) And ws.Cells(i, 1).Value > currentDate Then
            ' Set the font size for each cell in the row (adjust the size as needed)
            ws.Rows(i).Font.Size = 6 ' Change 12 to your desired font size
            'next calc running average of ratio of JMVs to FMVs to be used when only a JMV can be found
            'get sum of the estimates present
'            x = cell.Offset(0, GetColumnNumberByHeaderName("Zestimate", sheetName)).Value + _
'                cell.Offset(0, GetColumnNumberByHeaderName("Realtor.com Estimate", sheetName)).Value + _
'                cell.Offset(0, GetColumnNumberByHeaderName("RedFin Estimate", sheetName)).Value
                
           ' Get values from the specified columns
            zestimate = ws.Cells(i, GetColumnNumberByHeaderName("Zestimate")).Value
            realtorEstimate = ws.Cells(i, GetColumnNumberByHeaderName("Realtor.com Estimate")).Value
            redfinEstimate = ws.Cells(i, GetColumnNumberByHeaderName("RedFin Estimate")).Value

            ' Initialize count and sum for averaging
            countValidValues = 0
            averageValue = 0
        
            ' Check if each value is numeric, then accumulate and count valid values
            If IsNumeric(zestimate) Then
                averageValue = averageValue + zestimate
                countValidValues = countValidValues + 1
            End If
        
            If IsNumeric(realtorEstimate) Then
                averageValue = averageValue + realtorEstimate
                countValidValues = countValidValues + 1
            End If
        
            If IsNumeric(redfinEstimate) Then
                averageValue = averageValue + redfinEstimate
                countValidValues = countValidValues + 1
            End If
        
            ' Calculate the average if there are valid values
            If countValidValues > 0 Then
                averageValue = averageValue / countValidValues
               ' MsgBox "Average Value: " & averageValue
                'proceed only if there is a JMV present.  Note, this may be recounted many times over the next few months
                If IsNumeric(ws.Cells(i, GetColumnNumberByHeaderName("Just Market Value")).Value) And averageValue > 0 Then
                   x = ws.Cells(i, GetColumnNumberByHeaderName("Just Market Value")).Value / averageValue
                  ' Debug.Print ("FMV:  " & averageValue & ", JMV:  " & ws.Cells(i, GetColumnNumberByHeaderName("Just Market Value")).Value)
                   If x > 0.5 And x < 1 Then
                        a = GetParameterValue("avg_JMV2FMV")
                        d = GetParameterValue("dataPoints_JMV2FMV")
                        a = (a * d + x) / (d + 1)
                        UpdateParameterValue "avg_JMV2FMV", a
                        UpdateParameterValue "dataPoints_JMV2FMV", d + 1
                   End If
                End If
            Else
               ' MsgBox "No valid numeric values found."
            End If
                
        End If
    Next i
End Sub


Sub FilterSalesWorksheet3()
    Dim ws As Worksheet
    Dim filterRange As range
    Dim today As Date
    
    Set ws = ThisWorkbook.Worksheets("Sales")
    today = Date
    
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    Call SortBy(fieldName:="Sale Price to FMV", Order:="Ascending")
   ' button4.Caption = "Unfilter"
    
    ' Find the range to filter
    Set filterRange = ws.range("A1").CurrentRegion
    
    ' Apply the filters
    filterRange.AutoFilter Field:=GetColumnNumberByHeaderNamePLUS1("Cancelled"), Criteria1:="<>x"
    filterRange.AutoFilter Field:=GetColumnNumberByHeaderNamePLUS1("HOA"), Criteria1:="<>x"
    filterRange.AutoFilter Field:=GetColumnNumberByHeaderNamePLUS1("Zestimate"), Criteria1:="<> 'none'", Operator:=xlOr, Criteria2:="<>"
    filterRange.AutoFilter Field:=GetColumnNumberByHeaderNamePLUS1("Sold To"), Criteria1:="3PB"
    filterRange.AutoFilter Field:=GetColumnNumberByHeaderNamePLUS1("Sale Price to FMV"), Criteria1:="> 20%", Operator:=xlAnd, Criteria2:="< 75%"
   
    ' Turn off autofilter mode if there are no visible rows
    If filterRange.SpecialCells(xlCellTypeVisible).Rows.Count <= 1 Then
        filterRange.AutoFilter Field:=GetColumnNumberByHeaderNamePLUS1("Cancelled")
    End If
End Sub


Sub FilterBySelectedCaseNum()
    'Declare variables
    Dim currentRow As Long
    Dim currentValue As String
    Dim filterRange As range
    Dim county As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("Sales")
    
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
        
    'Get the current row and value
    currentRow = Selection.row
    currentValue = Cells(currentRow, 2).Value
    county = Cells(currentRow, 4).Value
    
    'Create a filter range
    Set filterRange = range("A1:D" & Cells(Rows.Count, 1).End(xlUp).row)
    
    'Clear the existing filter
    'filterRange.AutoFilter Field:=2, Criteria1:=currentValue
    
    'Clear the existing filter
    'filterRange.AutoFilter Field:=4, Criteria1:=currentValue
    
    'Filter the worksheet on the county
    filterRange.AutoFilter Field:=4, Criteria1:=county
    
    'Filter the worksheet on the current value
    filterRange.AutoFilter Field:=2, Criteria1:=currentValue

End Sub



Function GetColumnByCounty(columnName As String, county As String) As Long
    Dim ws As Worksheet
    Dim headerRange As range
    Dim columnNum As Long
    
    Set ws = ThisWorkbook.Worksheets("Activity")
    Set headerRange = ws.Rows(1)
    
    ' Find the column with the specified header
    columnNum = Application.match(columnName, headerRange, 0)
    
    ' Find the column with the header matching the county argument
    countyColumn = Application.match(county, headerRange, 0)
    
    ' Return the column number of the specified column in the same row as the county column
    GetColumnByCounty = ws.Cells(headerRange.row, columnNum).Offset(0, countyColumn - 1).Column
End Function

Sub ProcessFromUiPath(Optional county As String, _
                      Optional saleType As String = "foreclosure", _
                      Optional filePath As String = "C:\Users\philk\Downloads\")
    On Error GoTo ErrorHandler:
    Dim sheetName As String: sheetName = "Vars"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    Dim rt As rowType
    'Dim salesType As Variant
    'Dim county As String
    'Dim saleType As String
    'Dim filePath As String: filePath = "C:\Users\philk\Downloads\"
    rt.sheetName = "Sales"
    rt.fontSize = 12
    rt.home = "2665 Hempel Ave, Windermere FL 34786"
   ' rt.county = ThisWorkbook.Sheets("vars").range("b3").Value
    If county <> "" Then
        rt.county = county
        rt.saleType = saleType
    Else
        If ws.range("b3").Value = "completed" Then
            MsgBox ("County cannot be named 'completed'")
            Exit Sub
        End If
        If ws.range("b3").Value = "" Then
            MsgBox ("County cannot be null")
            Exit Sub
        End If
        
        rt.county = ws.range("b3").Value
        saleType = ws.range("b4").Value
        rt.saleType = saleType
    End If
    ws.range("b3").Value = ""
    ws.range("b4").Value = ""
    ws.range("b5").Value = ""
    ws.range("b15").Value = ""
    ws.range("b16").Value = ""
    ws.range("b17").Value = ""
    ws.range("b18").Value = ""
    'filePath = ThisWorkbook.Sheets("vars").range("b5").Value
    If filePath <> "" Then
        rt.filePath = filePath
    ElseIf UCase(county) <> "OSCEOLA" Then
        rt.filePath = filePath & GetLatestFile()
    End If
   ' MsgBox ("UiPath is processing the " & rt.county & " County " & rt.saleType & " sales")
    ws.range("b3").Value = "running"
    'UpdateParameterValue "VBA_status", "running"
    Call ProcessCounty(rt_orginal:=rt, saleType:=saleType, county:=rt.county, filePath:=filePath)
    'ProcessCounty(ByVal county As String, ByVal filePath As String, ByRef rt_orginal As rowType
    ws.range("b3").Value = "completed"
   ' UpdateParameterValue "VBA_status", "completed"
    ws.range("b4").Value = ""
    ws.range("b5").Value = ""
    Call SetRowHeight
    Call DeleteQuickSearchFiles
    
exit_proc:
    Set ws = Nothing
    Debug.Print "Finished Processing Foreclosure Sale"
    Exit Sub
ErrorHandler:
    Debug.Print "failed ProcessFromUiPath, see loc 503  Err#: " & Err.Number & " " & Err.description
    Err.Clear
    Resume exit_proc
End Sub

Sub SetRowHeight()

Dim i As Double
Dim lastRow As Double

lastRow = ActiveSheet.UsedRange.Rows.Count

For i = lastRow To 1 Step -1
    ActiveSheet.Rows(i).rowHeight = 12
Next i

End Sub

Sub setupUiPath()
'temp utility to manually set up the last sales file downloaded.  in future this will be done automatically
    UpdateParameterValue "current_sale_type", GetSaleTypeOrCounty2()
    UpdateParameterValue "current_County_2b_loaded", GetSaleTypeOrCounty2("County")
End Sub

Sub ProcessSelected(foreclosureStartAt As Date, taxDeedStartAt As Date, rt As rowType)
   ' Dim saleTypes() As Variant
    Dim result() As Variant
    Dim latestFiles() As String
    Dim latestFilePaths() As String
    Dim latestFileCount As Integer
    Dim filePath As String: filePath = "C:\Users\philk\Downloads\"
    Dim x As String
    Dim county As String
    Dim saleTypes(1) As String
    Dim latestTaxDeed As String
    Dim latestForeclosure As String
    x = GetSaleTypeOrCounty2()
    If UCase(x) = "BOTH" Then
        saleTypes(0) = "Foreclosure"
        saleTypes(1) = "Tax Deed"
    Else
        saleTypes(0) = x
    End If
    rt.county = GetSaleTypeOrCounty2("County")
    'rt.county = countyIf
    'rt.FilePath = GetSaleTypeOrCounty2("filepath")
    ' Determine the latest file(s) for each sale type
    latestFileCount = UBound(saleTypes) - LBound(saleTypes) + 1
    ReDim latestFiles(latestFileCount - 1)
    ReDim latestFilePaths(latestFileCount - 1)
    Call setupUiPath
    'Call getPA_Templates(countyName:=rt.county, paTemplate:=rt.PA_link, paTemplateAddr:=rt.PA_link_Address)
    
    If UCase(x) <> "BOTH" Then
        rt.saleType = saleTypes(0)
        'rt.FilePath = FilePath
        rt.filePath = filePath & GetLatestFile()
        Call ProcessCounty(rt_orginal:=rt, saleType:=saleTypes(0), county:=rt.county, filePath:=filePath & GetLatestFile(), foreclosureStartAt:=foreclosureStartAt, taxDeedStartAt:=taxDeedStartAt)
    ElseIf UBound(saleTypes) > 0 Then
        ' Determine which file is the latest for foreclosure and tax deed
        latestFiles = GetLatestFiles()
        latestTaxDeed = filePath & latestFiles(0)
     '   result = DetermineFileType(latestTaxDeed)
        latestForeclosure = filePath & latestFiles(1)
'        If IsEmpty(result) Then
'            Debug.Print "Unable to determine county and sale type."
'        Else
'            Debug.Print "County: " & result(0)
'            Debug.Print "Sale type: " & result(1)
'        End If
'      '  result = DetermineFileType(latestForeclosure)
'        If IsEmpty(result) Then
'            Debug.Print "Unable to determine county and sale type."
'        Else
'            Debug.Print "County: " & result(0)
'            Debug.Print "Sale type: " & result(1)
'        End If
'
        If latestForeclosure = "" And latestTaxDeed = "" Then
            MsgBox "No files found for selected sale types."
            Exit Sub
        ElseIf latestForeclosure = "" Then
            Call ProcessCounty(rt_orginal:=rt, saleType:="Tax Deed", county:=rt.county, filePath:=latestFilePaths(0), taxDeedStartAt:=taxDeedStartAt)
        ElseIf latestTaxDeed = "" Then
            Call ProcessCounty(rt_orginal:=rt, saleType:="Freclosure", county:=rt.county, filePath:=latestFilePaths(0))
        ElseIf latestFiles(0) = latestForeclosure Then
            Call ProcessCounty(rt_orginal:=rt, saleType:="Foreclosure", county:=rt.county, filePath:=latestFilePaths(0), foreclosureStartAt:=foreclosureStartAt)
            Call ProcessCounty(rt_orginal:=rt, saleType:="Tax Deed", county:=rt.county, filePath:=latestFilePaths(1), taxDeedStartAt:=taxDeedStartAt)
        Else
            rt.saleType = "Foreclosure"
            rt.filePath = latestForeclosure
            Call ProcessCounty(rt_orginal:=rt, saleType:="Foreclosure", county:=county, filePath:=latestForeclosure, foreclosureStartAt:=rt.foreclosureStartAt)
            rt.saleType = "Tax Deed"
            rt.filePath = latestTaxDeed
            Call ProcessCounty(rt_orginal:=rt, saleType:="Tax Deed", county:=county, filePath:=latestTaxDeed, taxDeedStartAt:=taxDeedStartAt)
        End If
    End If
    Call GatherStatsAndFormat
End Sub

Function GetLatestFiles(Optional folderPath As String = "C:\Users\philk\Downloads") As Variant
    Dim latestFiles(1) As String
    Dim latestDates(1) As Date
    Dim FSO As Object
    Dim file As Variant
    Dim FileDate As Date
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    latestFiles(0) = ""
    latestDates(0) = DateValue("1/1/1970")
    latestFiles(1) = ""
    latestDates(1) = DateValue("1/1/1970")
    
    For Each file In FSO.GetFolder(folderPath).Files
        If LCase(Right(file.name, 4)) = ".csv" Then
            FileDate = file.DateLastModified
            If FileDate > latestDates(0) Then
                latestDates(1) = latestDates(0)
                latestFiles(1) = latestFiles(0)
                latestDates(0) = FileDate
                latestFiles(0) = file.name
            ElseIf FileDate > latestDates(1) Then
                latestDates(1) = FileDate
                latestFiles(1) = file.name
            End If
        End If
    Next file
    
    GetLatestFiles = latestFiles
    Set FSO = Nothing
End Function

Function GetLatestFile(Optional folderPath As String = "C:\Users\philk\Downloads") As String
    Dim latestFile As String
    Dim latestDate As Date
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")

    ' Find latest file with matching sale type
    latestFile = ""  'C:\Users\philk\Downloads
    latestDate = DateValue("1/1/1970")
    For Each file In FSO.GetFolder(folderPath).Files
        If LCase(Right(file.name, 4)) = ".csv" Then
            FileDate = file.DateLastModified
            If FileDate > latestDate Then
                latestDate = FileDate
                latestFile = file.name
            End If
        End If
    Next file

    GetLatestFile = latestFile
    Set FSO = Nothing
End Function



Function GetSaleTypeOrCounty(Optional info As String) As Variant
    Dim selectedRange As range
    Dim result As Variant
    Dim cell As range
    Dim header As String
    
    header = IIf(LCase(info) = "county", "row", "column")
    
    ' Get the selected range
    Set selectedRange = Selection
    
    ' Get the row header(s)
    If LCase(header) = "row" Then
        ReDim result(1 To selectedRange.Rows.Count)
        For Each cell In selectedRange.Cells
            result(cell.row - selectedRange.row + 1) = Cells(cell.row, 1).Value
        Next cell
    ' Get the column header(s)
    ElseIf LCase(header) = "column" Then
        ReDim result(1 To selectedRange.Columns.Count)
        For Each cell In selectedRange.Cells
            result(cell.Column - selectedRange.Column + 1) = Cells(1, cell.Column).Value
        Next cell
    Else
        MsgBox "Invalid argument. Please enter 'row' or 'column' as the header argument."
        Exit Function
    End If
    
    ' Return the result as a string or an array
    If UBound(result) = 1 Then
        GetSaleTypeOrCounty = result(1)
    Else
        GetSaleTypeOrCounty = result
    End If
End Function

Function GetSaleTypeOrCounty2(Optional info As String) As String
    Dim selectedRange As range
    Set selectedRange = Selection
    'if looking for County return the County, but if looking for sale type, return which one is selected.
    'if more than one cell is selected regardless of which cells assume that both sales types are have been selected.
    If LCase(info) = "county" Then
        GetSaleTypeOrCounty2 = Cells(selectedRange.row, 1).Value
    ElseIf LCase(info) = "filepath" Then
        GetSaleTypeOrCounty2 = Cells(selectedRange.row, 13).Value
    Else
       GetSaleTypeOrCounty2 = IIf(selectedRange.Cells.Count = 1, Cells(1, selectedRange.Column).Value, "BOTH")
    End If
End Function

Function DetermineFileType(ByVal filePath As String) As Variant

    Dim polkCities() As Variant
    polkCities = Array("AUBURNDALE", "LAKELAND", "FROSTPROOF", "DAVENPORT", "BARTOW", "HAINES CITY", "LAKE WALES", "POINCIANA", "POLK CITY", "WINTER HAVEN", "HAINES", "BRADLEY", "KISSIMMEE", "MULBERRY", "POLK")
    
    Dim orangeCities() As Variant
    orangeCities = Array("ORLANDO", "APOPKA", "OCOEE", "WINTER GARDEN", "WINTER PARK", "WINDERMERE", "ZELLWOOD", "GOTHA", "OAKLAND", "EATONVILLE")
    
    Dim lakeCities() As Variant
    lakeCities = Array("CLERMONT", "MOUNT DORA", "TAVARES", "EUSTIS", "LEESBURG", "GROVELAND", "FRUITLAND PARK", "MINNEOLA", "MASCOTTE", "UMATILLA", "ASTATULA", "HOWEY-IN-THE-HILLS", "LADY LAKE", "MONTVERDE", "CLERMONT (3442)", "THE VILLAGES", "GRAND ISLAND", "SORRENTO", "YALAHA", "PAISLEY")
    
    Dim volusiaCities() As Variant
    volusiaCities = Array("DAYTONA BEACH", "PORT ORANGE", "NEW SMYRNA BEACH", "DELAND", "ORMOND BEACH", "EDGEWATER", "ORANGE CITY", "DELTONA", "HOLLY HILL", "SOUTH DAYTONA", "LAKE HELEN", "PIERSON")
    
    Dim indianRiverCities() As Variant
    indianRiverCities = Array("VERO BEACH", "SEBASTIAN", "FELLSMERE", "INDIAN RIVER SHORES", "ORCHID")
    
    Dim brevardCities() As Variant
    brevardCities = Array("PALM BAY", "MELBOURNE", "TITUSVILLE", "COCOA", "ROCKLEDGE", "MERRITT ISLAND", "COCOA BEACH", "CABANA COLONY", "WEST MELBOURNE", "MIMS", "PORT ST JOHN", "MELBOURNE BEACH", "VIERA", "SATELLITE BEACH", "GRANT-VALKARIA", "INDIALANTIC", "MALABAR", "COCOA WEST", "PALM SHORES")
    
    Dim seminoleCities() As Variant
    seminoleCities = Array("CASSELBERRY", "OVIEDO", "SANFORD", "WINTER SPRINGS", "LAKE MARY", "LONGWOOD", "ALTAMONTE SPRINGS", "GENEVA")
    
    Dim osceolaCities() As Variant
    osceolaCities = Array("KISSIMMEE", "ST CLOUD", "CELEBRATION", "HARMONY")
    
    Dim hillsboroughCities() As Variant
    hillsboroughCities = Array("APOLLO BEACH", "BAKERVIEW", "BLOOMINGDALE", "BOYETTE", "BRANDON", "CARROLLWOOD", "CHEVAL", "CITRUS PARK", "DOVER", "EAST LAKE", "GIBSONTON", "KNIGHTS", "LUTZ", "MANGO", "NORTHDALE", "ODESSA", "PALM RIVER", "PARRISH", "PLANT CITY", "PROGRESS VILLAGE", "RIVERVIEW", "RUSKIN", "SEFFNER", "SUN CITY", "TAMPA", "TEMPLE TERRACE", "THONOTOSASSA", "TOWN 'N' COUNTRY", "TRINITY", "VALRICO", "WESTCHASE", "WIMAUMA")

    Dim hernandoCities() As Variant
    hernandoCities = Array("BROOKSVILLE", "HERNANDO BEACH", "MASARYKTOWN", "NORTH WEEKI WACHEE", "PINE ISLAND", "SPRING HILL", "WEEKI WACHEE", "BAYPORT", "NOBLETON", "ISTACHATTA")

    Dim manateeCities() As Variant
    manateeCities = Array("ANNA MARIA", "BRADENTON", "BRADENTON BEACH", "ELLENTON", "LONGBOAT KEY", "MYAKKA CITY", "PALMETTO")

    Dim sumterCities() As Variant
    sumterCities = Array("THE VILLAGES")
    
    Dim marionCities() As Variant
    marionCities = Array("OCALA")
    
    Dim flaglerCities() As Variant
    flaglerCities = Array("PALM COAST")
    
    Dim pascoCities() As Variant
    pascoCities = Array("DADE CITY", "HOLIDAY", "HUDSON", "LAND O LAKES", "LUTZ", "NEW PORT RICHEY", "PORT RICHEY", "SAN ANTONIO", "SPRING HILL", "WESLEY CHAPEL", "ZEPHYRHILLS")
    
    ' Load CSV data into array
    Dim csvData() As String
    csvData = Split(CreateObject("Scripting.FileSystemObject").OpenTextFile(csvPath, 1).ReadAll, vbCrLf)
    
    ' Determine county and sale type
    Dim county As String
    Dim saleType As String
    Dim isForeclosure As Boolean
    For row = LBound(csvData, 1) + 1 To UBound(csvData, 1) ' skip header row
        Dim fields() As String
        fields = Split(csvData(row), ",")
        Dim city As String
        city = Trim(fields(1))
        Dim finalJudgment As Double
        finalJudgment = CDbl(Trim(fields(7)))
        Dim openingBid As Double
        openingBid = CDbl(Trim(fields(8)))
        If county = "" Then
            If InArray(city, polkCities) Then
                county = "Polk"
            ElseIf InArray(city, orangeCities) Then
                county = "Orange"
            ElseIf InArray(city, lakeCities) Then
                county = "Lake"
            ElseIf InArray(city, volusiaCities) Then
                county = "Volusia"
            ElseIf InArray(city, indianRiverCities) Then
                county = "Indian River"
            ElseIf InArray(city, brevardCities) Then
                county = "Brevard"
            ElseIf InArray(city, seminoleCities) Then
                county = "Seminole"
            ElseIf InArray(city, osceolaCities) Then
                county = "Osceola"
            ElseIf InArray(city, hillsboroughCities) Then
                county = "Hillsborough"
            ElseIf InArray(city, sumterCities) Then
                county = "Sumter"
            ElseIf InArray(city, marionCities) Then
                county = "Marion"
            ElseIf InArray(city, flaglerCities) Then
                county = "Flagler"
            ElseIf InArray(city, pascoCities) Then
                county = "Pasco"
            ElseIf InArray(city, sarasotaCities) Then
                county = "Sarasota"
            ElseIf InArray(city, okeechobeeCities) Then
                county = "Okeechobee"
            ElseIf InArray(city, hernandoCities) Then
                county = "Hernando"
            ElseIf InArray(city, manateeCities) Then
                county = "Manatee"
            End If
        End If
        If finalJudgment > 0 And openingBid = 0 Then
            saleType = "Foreclosure"
            isForeclosure = True
        ElseIf finalJudgment = 0 And openingBid > 0 Then
            saleType = "Tax Deed"
            isForeclosure = False
        End If
        If county <> "" And saleType <> "" Then
            Exit For
        End If
    Next row
    ' Return county and sale type
    GetCountyAndSaleType = Array(county, saleType)
End Function

Function InArray(valToFind As Variant, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = valToFind Then
            InArray = True
            Exit Function
        End If
    Next i
    InArray = False
End Function

Sub SleepUntilReady(Optional ByVal seconds As Long = 1)
    Dim endTime As Date
    endTime = Now + TimeSerial(0, 0, seconds)
    Do While Not Application.Ready
    'Do While Application.CalculationState <> xlDone
        DoEvents
        If Now >= endTime Then Exit Do
        Sleep seconds
    Loop
    Sleep seconds
End Sub


Sub Sleep(Optional ByVal seconds As Long = 1)
    Dim end_time As Double
    end_time = Timer + seconds
    
    Do While Timer < end_time
        DoEvents
    Loop
End Sub

Sub getPA_Templates(ByRef countyName As String, ByRef paTemplate As String, ByRef paTemplateAddr As String)
    Dim sheetName As String: sheetName = "Activity"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' Get the row number of the county in the appropriate column
    Dim countyColumn As range
    Set countyColumn = ws.range("a:a")
    
    For Each countyCell In countyColumn.Cells
        If Not IsEmpty(countyCell.Value) Then
            countyName = countyCell.Value
            paTemplate = countyCell.Offset(0, GetColumnNumberByHeaderName("PA template", sheetName)).Value
            paTemplateAddr = countyCell.Offset(0, GetColumnNumberByHeaderName("PA template Address", sheetName)).Value
        End If
    Next
End Sub


Sub CopyWorksheetToGoogleSheetsOLD(ByVal sourceWorksheetName As String, ByVal destinationWorksheetName As String, ByVal destinationGoogleWorkbookName As String)

Dim xlApp As Object
Dim xlWB As Object
Dim xlWS As Object
Dim gsApp As Object
Dim gsWb As Object
Dim gsWS As Object

On Error GoTo ErrorHandler

Set xlApp = GetObject(, "Excel.Application")
Set xlWB = xlApp.Workbooks("foreclosures.xlsm")
Set xlWS = xlWB.Sheets(sourceWorksheetName)

Set gsApp = CreateObject("GoogleAppsScriptApp")
Set gsWb = gsApp.Workbooks.Create(destinationGoogleWorkbookName)
Set gsWS = gsWb.Sheets.Add(destinationWorksheetName)

For Each cell In xlWS.UsedRange
gsWS.Cells(cell.row, cell.Column).Value = cell.Value
Next cell

gsWb.Save

xlApp.Quit

ErrorHandler:

If Err.Number = 438 Then
MsgBox "Google Sheets is not installed."
Else
MsgBox Err.description
End If

End Sub
'
'Sub CopyWorksheetToGoogleSheets(ByVal sourceWorksheetName As String, ByVal destinationWorksheetName As String, ByVal destinationGoogleWorkbookName As String)
'
'Dim xlApp As Object
'Dim xlWB As Object
'Dim xlWS As Object
'Dim gsSpreadsheet As GoogleAppsScript.Spreadsheet
'
'On Error GoTo HandleError
'
'Set xlApp = GetObject(, "Excel.Application")
'Set xlWB = xlApp.Workbooks("Foreclosures.xlsm")
'Set xlWS = xlWB.Sheets(sourceWorksheetName)
'
'Set gsSpreadsheet = New GoogleAppsScript.Spreadsheet
'
'gsSpreadsheet.Create (destinationGoogleWorkbookName)
'
'gsSpreadsheet.Sheets.Add (destinationWorksheetName)
'
'For Each cell In xlWS.UsedRange
'gsSpreadsheet.Sheets(destinationWorksheetName).Cells(cell.row, cell.Column).Value = cell.Value
'Next cell
'
'gsSpreadsheet.Save
'
'xlApp.Quit
'
'Exit Sub
'
'HandleError:
'
'If Err.Number = 438 Then
'MsgBox "The workbook " & destinationGoogleWorkbookName & " already exists."
'Else
'MsgBox "An error occurred: " & Err.Description
'End If
'
'Resume
'
'End Sub
Sub ReplaceWorksheets()

    Dim wbSource As Workbook
    Dim wbDest As Workbook
    Dim wsSales As Worksheet
    Dim wsActivity As Worksheet
    
    ' Set the source workbook
    Set wbSource = ThisWorkbook
    
    ' Set the destination workbook
    Set wbDest = Workbooks("Distressed Properties.xlsx")
    
    ' Open the destination workbook
    wbDest.Activate
    
    ' Delete the existing "sales" worksheet in the destination workbook
    On Error Resume Next
    Application.DisplayAlerts = False
    wbDest.Worksheets("sales").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Delete the existing "activity" worksheet in the destination workbook
    On Error Resume Next
    Application.DisplayAlerts = False
    wbDest.Worksheets("activity").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Copy the "sales" worksheet to the destination workbook and rename it
    Set wsSales = wbSource.Worksheets("sales")
    s = "https://docs.google.com/spreadsheets/d/1g6t4LXHJnSN1T7PXgItiJpxQqF4reD4IOhJu-F0Nbpo/edit#gid=720271773"
    'Call CopyWorksheetToGoogleSheets(sourceWorksheetName:="sales", destinationWorksheetName:="sales", destinationGoogleWorkbookName:=s)
    wsSales.Copy Before:=wbDest.Worksheets(1)
    Set wsSales = wbDest.Worksheets(1)
    wsSales.name = "sales"
    
    ' Clear the "interested" and "notes" columns in the "sales" worksheet
    wsSales.range("P2:P" & wsSales.Cells(wsSales.Rows.Count, "P").End(xlUp).row).ClearContents
    wsSales.range("T2:T" & wsSales.Cells(wsSales.Rows.Count, "T").End(xlUp).row).ClearContents
    
    ' Delete all columns to the right of the "Timestamp" column in the "sales" worksheet
    Dim lastColumn As Long
    lastColumn = wsSales.Cells(2, wsSales.Columns.Count).End(xlToLeft).Column
    If lastColumn > 28 Then
        wsSales.range(wsSales.Cells(1, 29), wsSales.Cells(1, lastColumn)).EntireColumn.Delete
    End If
    
    ' Copy the "activity" worksheet to the destination workbook and rename it
    Set wsActivity = wbSource.Worksheets("activity")
    wsActivity.Copy Before:=wbDest.Worksheets(1)
    Set wsActivity = wbDest.Worksheets(1)
    wsActivity.name = "activity"
    
End Sub

Sub DeleteQuickSearchFiles()
    Dim FSO As Object
    Dim folder As Object
    Dim file As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set folder = FSO.GetFolder("C:\Users\philk\Downloads\")
    
    For Each file In folder.Files
        If InStr(file.name, "QuickSearch") > 0 Then
            file.Delete
        End If
    Next file
End Sub

Sub GetPhysicalAddressOLD(ByVal url As String)
    Dim HTML As New HTMLDocument
    Dim Elements As Object
    Dim address As String
    
    'Get the HTML content from the webpage
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", url, False
        .send
        HTML.body.innerHTML = .responseText
    End With
    
    'Find the element with the label "Physical Address"
    Set Elements = HTML.getElementsByTagName("label")
    For Each element In Elements
    Debug.Print element.innerText
        If element.innerText = "Physical Address" Then
            address = element.NextSibling.innerText
            Exit For
        End If
    Next element
    
    'Load the address into an Excel variable
   ' Debug.Print Address
   ' ThisWorkbook.Sheets("test").range("A1").Value = Address
End Sub

Sub GetPhysicalAddress(ByVal url As String)
    Dim HTML As New HTMLDocument
    Dim Elements As Object
    Dim address As String
    
    'Get the HTML content from the webpage
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", url, False
        .send
        HTML.body.innerHTML = .responseText
    End With
    
    'Find the th element with the label "Physical Address"
    Set Elements = HTML.getElementsByTagName("td")
    For Each element In Elements
        Debug.Print element.innerText
        If element.innerText = "Physical Address" Then
            address = element.NextSibling.innerText
            Exit For
        End If
    Next element
    
   ' Debug.Print Address
    'Load the address into an Excel variable
    ThisWorkbook.Sheets("test").range("A1").Value = address

End Sub

Function getSocialMediaLink(rt As rowType)
    '=T4578&" & Char(10) " '&TEXT(A4579,"mmmm dd")&IF(RIGHT(DAY(A4579),1)=1,"st",IF(RIGHT(DAY(A4579),1)=2,"nd",IF(RIGHT(DAY(A4579),1)=3,"rd","th")))&" "&D4579&" County"&" "&H4579&" Sale "&" & Char(10) " &AC4579&" & Char(10) " &U4579&", "&V4579&IF(W4579<>""," & Char(10) " &"Judgement:  "&DOLLAR(W4579,2),"")7
    
'  "
'September 17th Putnam County foreclosure Sale
'https://www.zillow.com/homes/_,-FL-_rb
', ,
'Judgement:  $593,201.48"
'  =T6&CHAR(10)&TEXT(A6,"mmmm dd")&IF(RIGHT(DAY(A6),1)=1,"st",IF(RIGHT(DAY(A6),1)=2,"nd",IF(RIGHT(DAY(A6),1)=3,"rd","th")))&" "&D6&" County"&" "&H6&" Sale "&CHAR(10)&AL6&CHAR(10)&U6&", "&V6&IF(W6<>"",CHAR(10)&"FMV:  "&DOLLAR(W6,2),"")&IF(X6<>"",CHAR(10)&"Judgement:  "&DOLLAR(X6,2),"")&IF(AB6<>"",CHAR(10)&"Equity:  "&DOLLAR(AB6,2),"")
    
    s = Format(rt.salesDate, "mmmm d")
    Dim ZilloPathCR As String
    Dim NotesCR As String
    Dim JudgementCR As String
    Dim MaxBidCR As String
    'Dim equityFormula As String
    JudgementCR = GetColumnByHeaderName2(headerName:="Judgement", getLetter:=True) & rt.cell.row
    MaxBidCR = GetColumnByHeaderName2(headerName:="PlaintiffMaxBid", getLetter:=True) & rt.cell.row
    ZilloPathCR = GetColumnByHeaderName2(headerName:="ZilloPath", getLetter:=True) & rt.cell.row
    NotesCR = GetColumnByHeaderName2(headerName:="Notes", getLetter:=True) & rt.cell.row
   ' equityFormula = "=IF(AND(" & ZestimateCR &  CHAR(34) & " <>'',UPPER(" & ZestimateCR &  CHAR(34) & )<>'NONE')," & ZestimateCR & "& CHAR(34) & -MAX(" & JudgementCR & "," & MaxBidCR & ") & CHAR(34) & ,'')"
    
    x = "=" & NotesCR & " & Char(10) "
   ' Debug.Print x
    x = x + _
    "& """ & s & IIf(Right(s, 1) = "1", "st", IIf(Right(s, 1) = "2", "nd", IIf(Right(s, 1) = "3", "rd", "th"))) & " "
    x = x + _
    rt.county & " County " & rt.saleType & " Sale" & Chr(34) & " & Char(10) " & " & " & _
    ZilloPathCR & " & Char(10)"
    'Debug.Print x
    
    If rt.address <> "" Or False Then
        x = x + _
        " & " & Chr(10) & rt.address & " & Char(34) " & _
        " & " & rt.city & Chr(34) & " , FL " & Chr(34) & " & rt.zip  & Char(10) "
        Debug.Print x
    End If
    
    x = x '+ _
   ' "IF(" & ZilloPathCR & notEqualToBlank & ",CHAR(10) & Char(34) & " FMV:  Char(34) & DOLLAR(ZilloPathCR,2),"")'&IF(X6<>"",CHAR(10)&"Judgement:  "&DOLLAR(X6,2),"")&IF(AB6<>"",CHAR(10)&"Equity:  "&DOLLAR(AB6,2),"")
 'x = x & " & CHAR(10) & CHAR(34) & "" FMV: "" & CHAR(34) & IF(" & ZilloPathCR & " <> " & notEqualToBlank() & "," & CHAR(10) & CHAR(34) & """"" & " FMV: "" & CHAR(34) & DOLLAR(" & ZilloPathCR & ",2), """") & IF(" & MaxBidCR & "<> " & notEqualToBlank() & "," & CHAR(10) & CHAR(34) & "Judgement: " & CHAR(34) & " & DOLLAR(" & MaxBidCR & ",2), """") & IF(" & JudgementCR & "<> " & notEqualToBlank() & "," & CHAR(10) & CHAR(34) & "Equity: " & CHAR(34) & " & DOLLAR(" & JudgementCR & ",2), """")"
  
   ' IIf(finalJudgment <> "", "Judgement:  " & Format(rt.finalJudgment, "$0.00"), "")
   ' Debug.Print x
    
   '=T5 & CHAR(10) & "June 05th Osceola County foreclosure Sale"  & CHAR(10) &AC5 & CHAR(10)  & CHAR(10) &" , FL " & rt.zip  & CHAR(10)
    
    getSocialMediaLink = ""
    
End Function

Function notEqualToBlank()
    notEqualToBlank = "<>" & Chr(34) & Chr(34)
End Function


Sub LoadOkHOAs()

    'Declare variables
    Dim ws As Worksheet
   ' Dim OkHOAs As Variant
    Dim i As Long
    Dim col1 As String: col1 = GetColumnByHeaderName2(headerName:="HOA Name", getLetter:=True, worksheetName:="Vars")
    Dim col2 As String: col2 = GetColumnByHeaderName2(headerName:="County", getLetter:=True, worksheetName:="Vars")
    Dim col3 As String: col3 = GetColumnByHeaderName2(headerName:="Type", getLetter:=True, worksheetName:="Vars")
    
    'Set worksheet variable
    Set ws = ThisWorkbook.Sheets("vars")
    
    
    'Create multidimensional OkHOAsay
    ReDim OkHOAs(ws.Cells(col1 & "1").Cells.Count, 3)
    
    'Populate OkHOAsay with data from columns G, H, and I
'    For i = 1 To OkHOAs.Rows.Count
'        OkHOAs(i, 1) = ws.Cells(i, 8).Value
'        OkHOAs(i, 2) = ws.Cells(i, 9).Value
'        OkHOAs(i, 3) = ws.Cells(i, 10).Value
'    Next i
    
    'Print OkHOAsay contents
    For i = 1 To OkHOAs.Rows.Count
        Debug.Print OkHOAs(i, 1), OkHOAs(i, 2), OkHOAs(i, 3)
    Next i

End Sub

Sub LoadNameParts()

    'Declare variables
    Dim ws As Worksheet
   ' Dim OkHOAs As Variant
    Dim i As Long
    Dim col1 As String: col1 = GetColumnByHeaderName2(headerName:="Name Part", getLetter:=True, worksheetName:="Vars")
    Dim col2 As String: col2 = GetColumnByHeaderName2(headerName:="type", getLetter:=True, worksheetName:="Vars")
    
    'Set worksheet variable
    Set ws = ThisWorkbook.Sheets("vars")
    
    
    'Create multidimensional OkHOAsay
    ReDim nameParts(ws.Cells(col1 & "1").Cells.Count, 3)
    
    'Populate OkHOAsay with data from columns G, H, and I
'    For i = 1 To OkHOAs.Rows.Count
'        OkHOAs(i, 1) = ws.Cells(i, 8).Value
'        OkHOAs(i, 2) = ws.Cells(i, 9).Value
'        OkHOAs(i, 3) = ws.Cells(i, 10).Value
'    Next i
    
    'Print OkHOAsay contents
    For i = 1 To OkHOAs.Rows.Count
        Debug.Print OkHOAs(i, 1), OkHOAs(i, 2), OkHOAs(i, 3)
    Next i

End Sub

Sub makeNamePartArray()
   ' Set worksheet variable
    Set ws = ThisWorkbook.Sheets("vars")
    Dim col1 As String: col1 = GetColumnByHeaderName2(headerName:="NamePart", getLetter:=True, worksheetName:="Vars")
    Dim col2 As String: col2 = GetColumnByHeaderName2(headerName:="PartType", getLetter:=True, worksheetName:="Vars")
    ' Determine the number of rows with values in the specified column
    numRows = Application.WorksheetFunction.CountA(ws.Columns(col1))
    ' Resize the theDataArray based on the number of rows with values
    ReDim theDataArray(numRows - 1, 1)
    Dim rowIndex As Long
    Dim theDataArrayIndex As Long
    theDataArrayIndex = 0

    ' Loop through the rows and populate the theDataArray with values from the specified column
    For rowIndex = 1 To ws.Rows.Count
        If ws.Cells(rowIndex, col1).Value <> "" Then
            theDataArray(theDataArrayIndex, 0) = ws.Cells(rowIndex, col1).Value
            theDataArray(theDataArrayIndex, 1) = ws.Cells(rowIndex, col2).Value
            theDataArrayIndex = theDataArrayIndex + 1
        End If
    Next rowIndex
    If inputStr <> "" Then
       ' Debug.Print (GetNameInSerarchFormat(inputStr:=inputStr, theDataArray:=theDataArray))
    End If
End Sub

Sub testit4()
Dim x As String
'x = "UNKNOWN HEIRS OF THE ESTATE OF KAREN RAYBE, ET AL"
'x = "UNKNOWN HEIRS OF THE ESTATE OF JACK W. PENTURFF, E..."
'x = "THE UNKNOWN HEIRS OF THE ESTATE OF LEWIS B. BODAME..."
'x = "UNKNOWN SPOUSE OF THE ESTATE OF JESSIE M. GREEN, E..."
x = "UNKNOWN HEIRS OF MARY ANN SCOTT, ET AL"

Debug.Print (x)
x = GetNameInSerarchFormat(x)
Debug.Print (x)
End Sub

Function GetNameInSerarchFormat(inputStr As String, Optional dataArray As Variant) As String


    Dim i As Long

 '   Dim dataArray() As String
    Dim typeAction As String
    Dim ws As Worksheet
    Dim numRows As Integer
    Dim col1 As String: col1 = GetColumnByHeaderName2(headerName:="NamePart", getLetter:=True, worksheetName:="Vars")
    Dim col2 As String: col2 = GetColumnByHeaderName2(headerName:="PartType", getLetter:=True, worksheetName:="Vars")
    Dim pos As Long

'GetNameInSerarchFormat = "test"


    If Not IsEmpty(theDataArray) Then
        dataArray = theDataArray
    Else
'         'Set worksheet variable
        Set ws = ThisWorkbook.Sheets("vars")
        ' Determine the number of rows with values in the specified column
        numRows = Application.WorksheetFunction.CountA(ws.Columns(col1))
        ' Resize the dataArray based on the number of rows with values
        ReDim dataArray(numRows - 1, 1)
        Dim rowIndex As Long
        Dim dataArrayIndex As Long
        dataArrayIndex = 0

        ' Loop through the rows and populate the dataArray with values from the specified column
        For rowIndex = 1 To ws.Rows.Count
            If ws.Cells(rowIndex, col1).Value <> "" Then
                dataArray(dataArrayIndex, 0) = ws.Cells(rowIndex, col1).Value
                dataArray(dataArrayIndex, 1) = ws.Cells(rowIndex, col2).Value
                dataArrayIndex = dataArrayIndex + 1
            End If
        Next rowIndex
    End If

    For i = LBound(dataArray, 1) + 1 To UBound(dataArray, 1)
        namePart = dataArray(i, 0)
        typeAction = dataArray(i, 1)

        If typeAction = "after" Then
            pos = InStr(inputStr, namePart)
            If pos > 0 Then
               ' inputStr = Right(inputStr, pos + Len(namePart) - 1)
          '      inputStr = Trim(Right(inputStr, pos - 2))
                inputStr = Trim(Right(inputStr, Len(inputStr) - pos - Len(namePart) + 1))
            End If
        ElseIf typeAction = "before" Then
            pos = InStr(inputStr, namePart)
            If pos > 0 Then
                inputStr = Left(inputStr, pos - 1)
            End If
        ElseIf typeAction = "" Then
            inputStr = Replace(inputStr, namePart, "")
        End If
    Next i

    inputStr = Replace(inputStr, ".", "")
    inputStr = Replace(inputStr, ";", "")
    inputStr = Replace(inputStr, ",", "")

    GetNameInSerarchFormat = RearrangeWords(inputStr)
    
End Function
Function RearrangeWords(ByVal inputText As String) As String
    Dim words() As String
    words = Split(inputText, " ")
    
    If UBound(words) > 0 Then
        Dim lastWord As String
        lastWord = words(UBound(words))
        
        Dim remainingWords As String
        remainingWords = ""
        
        For i = 0 To UBound(words) - 1
            If remainingWords <> "" Then
                remainingWords = remainingWords & " "
            End If
            remainingWords = remainingWords & words(i)
        Next i
        
        RearrangeWords = lastWord & " " & remainingWords
    Else
        RearrangeWords = inputText
    End If
End Function


Function FileRead(filePath As String) As String

Dim file As Object
Dim data As String

Set file = CreateObject("Scripting.FileSystemObject")
data = file.OpenTextFile(filePath, ForReading).ReadAll()

FileRead = data

End Function


Sub ProcessDesotoCounty(rt As rowType)

Dim data As String
Dim foreclosures As Variant
Dim group As Object
Dim sale_date As String
Dim foreclosure As Object

Dim foreclosures_objects As Object

Call PdfGetter(rt:=rt)
'data = FileRead("C:\Users\philk\Downloads\DeSotoSalesList.txt")
data = "C:\Users\philk\Dropbox\Real Estate\Deal Finder\temp.txt"

foreclosures = Split(data, vbCrLf & vbCrLf)

' Create a new object using the CreateObject function.
Set foreclosures_objects = CreateObject("Object")

For Each group In foreclosures_objects
  sale_date = Mid(group, 1, InStr(group, vbCrLf) - 1)
  foreclosure = []

  ' Define the sale_date property for the foreclosure object.
  foreclosure.sale_date = sale_date

  For Each Line In group.Split(vbCrLf)
    If Line <> sale_date Then
      Dim match As RegExp
      match = RegExp.match(Line, "(?P<case_no>\d+) (?P<plaintiff>\w+) (?P<defendant>\w+) (?P<fj>\d+) (?P<amount>\d+) (?P<legal_desc>\w+) (?P<street_address>\w+)")
      If match Then
'        foreclosure ["case_no"] = match.Groups("case_no").Item.Value
'        foreclosure ["plaintiff"] = match.Groups("plaintiff").Item.Value
'        foreclosure ["defendant"] = match.Groups("defendant").Item.Value
'        foreclosure ["fj"] = match.Groups("fj").Item.Value
'        foreclosure ["amount"] = match.Groups("amount").Item.Value
'        foreclosure ["legal_desc"] = match.Groups("legal_desc").Item.Value
'        foreclosure ["street_address"] = match.Groups("street_address").Item.Value
      End If
    End If
  Next
  foreclosures_objects.Add (foreclosure)
Next

End Sub

Sub download_foreclosure_pdf()
  Dim pdf_data As String
  Dim file_name As String

  'Download the PDF file
  pdf_data = DownloadFile2("https://www.desotoclerk.com/wp-content/uploads/2023/06/UPDATED-Foreclosure-6-5-2023.pdf")

  'Save the PDF data to a variable
  file_name = "foreclosure_data.pdf"
 ' SaveFile pdf_data, file_name
End Sub

Function DownloadFile2(file_url As String) As String
  Dim file_data As String
  Dim FSO As Object
  Dim ts As Object

  'Create a File System Object
  Set FSO = CreateObject("Scripting.FileSystemObject")

  'Create a temporary file
  Set ts = FSO.CreateTextFile("temp.pdf", True)

  'Download the file
  Dim req As Object
  Set req = CreateObject("Microsoft.XMLHTTP")
  req.Open "GET", file_url, False
  req.send
  file_data = req.responseBody

  'Close the temporary file
  ts.Close

  'Return the file data
  DownloadFile2 = file_data
End Function


Sub CallPython()
    Dim pdf_content As String
    Dim pdf_file_path As String
    pdf_url = "https://www.desotoclerk.com/wp-content/uploads/2023/06/UPDATED-Foreclosure-6-5-2023.pdf"
    pdf_content = python.GetPDFContent(pdf_file_path)
    MsgBox pdf_content
End Sub






Sub CallPdfGetterOLD()
    Dim pdf_url As String
    Dim processID As LongPtr
    Dim cmdCall
'python pdf_reader.py "https://www.desotoclerk.com/wp-content/uploads/2023/06/UPDATED-Foreclosure-6-5-2023.pdf" "C://Users//philk//Dropbox//Real Estate//Deal Finder//temp.txt"
    file_path = "C://Users//philk//Dropbox//Real Estate//Deal Finder//temp.txt"
    pdf_url = "https://www.desotoclerk.com/wp-content/uploads/2023/06/UPDATED-Foreclosure-6-5-2023.pdf"
  '  processID = ShellExecute(0, "open", "python", "C:\Users\philk\Dropbox\Real Estate\Deal Finder\pdfGetter.py """ & pdf_url & """", vbNullString, vbNormalFocus)
  processID = ShellExecute(0, "open", "python", "C:\Users\philk\Dropbox\Real Estate\Deal Finder\pdfGetter.py """ & pdf_url & """", vbNullString, vbNormalFocus)
  
  'cmdCall = "C:\Users\philk\Dropbox\Real Estate\Deal Finder\pdfGetter.py " & pdf_url & " " & file_path
  'Debug.Print cmdCall
'    processID = ShellExecute(0, "open", "python3", "C:\Users\philk\Dropbox\Real Estate\Deal Finder\pdfGetter.py """ & pdf_url & """ " & file_path, vbNullString, vbNormalFocus)
   ' processID = ShellExecute(0, "open", "python3", cmdCall, vbNullString, vbNormalFocus)
    WaitForSingleObject processID, Timeout
End Sub


'Private Const Timeout As Long = 5000 ' 5 seconds

Sub PdfGetter(rt As rowType)
    'Dim pdf_url As String
   ' Dim file_path As String
    Dim processID As LongPtr
    
   ' 'pdf_url = "https://www.desotoclerk.com/wp-content/uploads/2023/06/UPDATED-Foreclosure-6-5-2023.pdf"
  '  rt.filePath = "C:\Users\philk\Dropbox\Real Estate\Deal Finder\temp.txt"
    ''county = "DeSoto"
    ''url = "https://www.desotoclerk.com/public-sales/foreclosures/"""
    
  '  processID = ShellExecute(0, "open", "python", """C:\Users\philk\Dropbox\Real Estate\Deal Finder\pdfGetter.py"" """ & rt.county & """ """ & rt.filepath & """", vbNullString, vbNormalFocus)
  ''  processID = ShellExecute(0, "open", "python", """C:\Users\philk\Dropbox\Real Estate\Deal Finder\pdfGetter.py"" """ & pdf_url & """", vbNullString, vbNormalFocus)
   ' 'processID = ShellExecute(0, "open", "python", """C:\Users\philk\Dropbox\Real Estate\Deal Finder\pdfGetter.py""", vbNullString, vbNormalFocus)
    WaitForSingleObject processID, Timeout
End Sub

Private Function updateAllParcels()
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlWS As Object
    Dim h As String
    Dim action As String
    
    On Error GoTo ErrorHandler:
    
    sheetName = "Sales"
    i = "a"
    
    Set xlApp = GetObject(, "Excel.Application")
    Set xlWB = ThisWorkbook
    Set xlWS = xlWB.Worksheets(sheetName)
    
    ' Find primary key columns
    Dim headers As Variant
    headers = xlWS.Rows(1).Value
    
    Dim pkColumn As Variant
    
    ' These fields make up the primary key, which is always distinct
    pkColumn = Application.match("Sale Date", headers, 0)
    
    If IsError(pkColumn) Then
        ' MsgBox "Primary key columns not found in sheet " & sheetName
        Exit Function
    End If
    
    i = "b"
    
    ' Find matching rows
    Dim cell As range
    Dim pkRange As range
    Set pkRange = xlWS.range(xlWS.Cells(2, pkColumn), xlWS.Cells(xlWS.Rows.Count, pkColumn).End(xlUp))
         
    ' Clear existing filters
    xlWS.AutoFilterMode = False
    
    ' Apply new filters to pkRange
'    pkRange.AutoFilter Field:=GetColumnNumberByHeaderName("Cancelled") + 1, Criteria1:="<>x"
'    pkRange.AutoFilter Field:=GetColumnNumberByHeaderName("HOA") + 1, Criteria1:="<>x"
'    pkRange.AutoFilter Field:=GetColumnNumberByHeaderName("Interested") + 1, Criteria1:="=yes", Operator:=xlOr, Criteria2:="="
'    pkRange.AutoFilter Field:=1, Criteria1:=">=" & Date
'Dim columnMapping As Object
'Set columnMapping = CreateObject("Scripting.Dictionary")

'map("bed") = GetColumnNumberByHeaderName(headerName:="Beds")
'map("bath") = GetColumnNumberByHeaderName(headerName:="Bath")
'map("lr") = GetColumnNumberByHeaderName(headerName:="LivingSqFt")
'map("lot") = GetColumnNumberByHeaderName(headerName:="Lot Size")
'map("yb") = GetColumnNumberByHeaderName(headerName:="Year Built")

    
    ' Now see if you can find a row with this primary key
    For Each cell In pkRange.Cells
        If IsDate(cell.Value) Then
            If cell.Value >= Date Then
                If cell.Offset(0, GetColumnNumberByHeaderName("Cancelled", sheetName)).Value <> "x" And _
                    cell.Offset(0, GetColumnNumberByHeaderName("HOA", sheetName)).Value <> "x" And _
                    cell.Offset(0, GetColumnNumberByHeaderName("Interested", sheetName)).Value <> "no" And _
                    cell.Offset(0, GetColumnNumberByHeaderName("Bedrooms", sheetName)).Value = "" Then
                    parcelInfoGetter cell
                End If
            End If
        End If
    Next cell
    
exit_proc:
    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    Exit Function

ErrorHandler:
    Debug.Print "failed rowExists, see loc 511" & i & " case#:" & CaseNumber; " Err#: " & Err.Number & " " & Err.description
    Err.Clear
    Resume exit_proc
End Function

Sub getParcelInfo()
    Dim currentCell As range
    Dim currentRow As Long
    currentRow = Selection.row
    Set currentCell = Cells(currentRow, 1)
    Call parcelInfoGetter(cell:=currentCell)
End Sub

Function parcelInfoGetter(ByRef cell As range)
    On Error GoTo ErrorHandler:
    Dim url As String
    Dim bedColumn As Long
    Dim filePath As String
    Dim fileNumber As Integer
    Dim bed As String
    Dim bath As String
    Dim livingArea As String
    Dim sheetName: sheetName = "Sales"
    Dim rt As rowType
    Set rt.cell = cell
    If rt.cell.Offset(0, GetColumnNumberByHeaderName("address", sheetName)).Value = "" Then Exit Function
    filePath = "C:\Users\philk\Dropbox\Real Estate\Deal Finder\zemp.txt"
    'url = "https://www.bing.com/search?pglt=43&q=2665+hempel+ave%2C+windermere%2C+fl+34786"
    'bedColumn = map("bed")
    Dim startTime As Date
    startTime = Now()
    url = "https:/www.bing.com/search?pglt=43&q=" & _
            Replace(rt.cell.Offset(0, GetColumnNumberByHeaderName("address", sheetName)).Value, " ", "+") & _
            "%2C+" & rt.cell.Offset(0, GetColumnNumberByHeaderName("CSZ", sheetName)).Value & _
            "%2C+fl" '& _
           ' rt.cell.Offset(0, GetColumnNumberByHeaderName("Zip", sheetName)).Value
    processID = ShellExecute(0, "open", "python", """C:\Users\philk\Dropbox\Real Estate\Deal Finder\getParcelInfo.py"" """ & url, vbNullString, vbNormalFocus)
    
    WaitForSingleObject processID, Timeout
    SleepUntilFileChanges startTime, filePath

    ' Open the file
    fileNumber = FreeFile
    Open filePath For Input As fileNumber
    
    ' Read the file content
    rt.notes = Input$(LOF(fileNumber), fileNumber)
    
    ' Close the file
    Close fileNumber
    i = "a"
    ' Extract bed, bath, and living area values
    startIndex = InStr(rt.notes, "Beds") - 2
    If startIndex >= 0 Then
        endIndex = InStr(startIndex, rt.notes, " ")
        If startIndex > 0 And IsNumeric(Mid(rt.notes, startIndex, endIndex - startIndex)) Then
            rt.bedrooms = Val(Mid(rt.notes, startIndex, endIndex - startIndex))
        End If
    End If
        ' Extract bath count
    'startIndex = InStr(rt.notes, "Baths") - 2
    'endIndex = InStr(startIndex, rt.notes, " ")
    'bathCount = Mid(rt.notes, startIndex, endIndex - startIndex)
    
    i = "b"
    startIndex = InStr(rt.notes, " sqft") - 5
    If startIndex >= 0 Then
        endIndex = InStr(startIndex, rt.notes, " ")
        If startIndex > 0 And IsNumeric(Mid(rt.notes, startIndex, endIndex - startIndex)) Then
            rt.LivingSF = Val(Mid(rt.notes, startIndex, endIndex - startIndex))
        End If
    End If
    i = "c"
    If IsNumeric(GetValueBetweenStrings(rt.notes, "Beds", "Baths")) Then
        rt.bathrooms = Val(GetValueBetweenStrings(rt.notes, "Beds", "Baths"))
    End If
   ' rt.bathrooms = GetValueBetweenStrings(rt.notes, "Beds", "Baths")
    'livingArea = GetValueBetweenStrings(rt.notes, "sqft", "houses")
    'Call ParseBingText(rt.cell:=rt.cell, content:=rt.notes)
    ParseBingText rt
    ' Display the extracted values
   ' MsgBox "Bed: " & bed & vbCrLf & "Bath: " & bath & vbCrLf & "Living Area: " & livingArea
   rt.cell.Offset(0, GetColumnNumberByHeaderName("Bedrooms", sheetName)).Value = rt.bedrooms
   rt.cell.Offset(0, GetColumnNumberByHeaderName("Bathrooms", sheetName)).Value = rt.bathrooms
   rt.cell.Offset(0, GetColumnNumberByHeaderName("Living SF", sheetName)).Value = rt.LivingSF
   rt.cell.Offset(0, GetColumnNumberByHeaderName("Property Type", sheetName)).Value = rt.PropertyType
   Debug.Print "Beds: " & rt.bedrooms & ", Baths: " & rt.bathrooms & ", SqFt: " & rt.LivingSF & ", Type: " & rt.PropertyType
exit_proc:
    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    Exit Function

ErrorHandler:
    Debug.Print "failed parcelInfoGetter, see loc 514" & i; " Err#: " & Err.Number & " " & Err.description
    Err.Clear
    Resume exit_proc
End Function

Function GetValueBetweenStrings(fullString As String, startString As String, endString As String) As String
    Dim startPos As Long
    Dim endPos As Long
    
    startPos = InStr(fullString, startString) + Len(startString)
    endPos = InStr(startPos, fullString, endString)
    If endPos - startPos >= 0 Then
        x = Trim(Mid(fullString, startPos, endPos - startPos))
    End If
    GetValueBetweenStrings = x
End Function


Sub SleepUntilFileChanges(startTime As Date, ByVal fileName As String)
  Dim FSO As Object
  Dim FileDate As Date
  Dim OldFileDate As Date

  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  'OldFileDate = fso.GetFile(FileName).DateLastModified

  Do While True
    FileDate = FSO.GetFile(fileName).DateLastModified
    If FileDate > startTime Then
      Exit Do
    End If
    Sleep
  Loop
End Sub

Function ParseBingText(ByRef rt As rowType) '(cell As range, content As String)
    On Error GoTo ErrorHandler:
    Dim lines() As String
    Dim mfh As Variant
    mfh = Array("manufactured", "NEIGHBORHOOD ASSOCIATION")
    rt.address = rt.cell.Offset(0, GetColumnNumberByHeaderName("Address")).Value
    lines = Split(rt.notes, vbCrLf)
    
    For i = 0 To UBound(lines)
        If rt.bedrooms <= 0 And InStr(lines(i), "Beds") > 0 Then
            rt.bedrooms = IIf(IsNumeric(lines(i + 1)), CInt(lines(i + 1)), rt.bedrooms)
        End If
        If rt.bathrooms <= 0 And rt.bathrooms > 0 And InStr(lines(i), "Baths") > 0 Then
            rt.bathrooms = IIf(IsNumeric(lines(i + 1)), CInt(lines(i + 1)), rt.bathrooms)
        End If
        'manufactured homes
        If rt.PropertyType = "" And InStr(lines(i), "Beds") > 0 Then
            For j = 0 To UBound(mfh)
                If InStr(UCase(lines(i)), UCase(mfh(j))) > 0 Then
                    rt.PropertyType = "MFH"
                    Exit For
                End If
            Next j
        End If
    Next i
  '  Debug.Print "Beds: " & rt.bedrooms & ", Baths: " & rt.bathrooms
exit_proc:
    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    Exit Function

ErrorHandler:
    Debug.Print "failed ParseBingText, see loc 513" & i & " case#:" & CaseNumber; " Err#: " & Err.Number & " " & Err.description
    Err.Clear
    Resume exit_proc
End Function

   
    
Sub GetNextAuction(ByRef oldestCounty As String, ByRef oldestSaleType As String, ByRef listURL As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim countyCol As Integer, foreclosureCol As Integer, taxDeedCol As Integer
    Dim daysFCCol As Integer, daysTDCol As Integer, ForeclosureListURLCol As Integer, taxDeedListURLCol As Integer
    Dim county As String
    Dim oldestForeclosureDate As Date, oldestTaxDeedDate As Date
    Dim tempOldestSaleType As String
    Dim tempListURL As String
    
    ' Define the worksheet
    Set ws = ThisWorkbook.Worksheets("Activity")
    
    ' Define the columns
    countyCol = 1 ' Assuming County column is in column A
    foreclosureCol = 2 ' Assuming Foreclosure column is in column B
    taxDeedCol = 3 ' Assuming Tax Deed column is in column C
    daysFCCol = 4 ' Assuming Days FC column is in column D
    daysTDCol = 5 ' Assuming Days TD column is in column E
    ForeclosureListURLCol = 6 ' Assuming List column is in column F
    taxDeedListURLCol = 7 ' Assuming Tax DeedList column is in column G
    ForeclousreYN_COl = 16
    taxDeedYN_COl = 17
    
    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, countyCol).End(xlUp).row
    
    ' Initialize variables
    oldestForeclosureDate = Date
    oldestTaxDeedDate = Date
    
    ' Define the range excluding the header row
    Dim columnBRange As range
    Set columnBRange = ws.range(ws.Cells(3, foreclosureCol), ws.Cells(lastRow, foreclosureCol))
    
    ' Loop through the data
    For Each cell In columnBRange
        ' Check if the row corresponds to a foreclosure with "x" in column P (assuming column 16)
        If ws.Cells(cell.row, 16).Value = "x" Then
           ' county = ws.Cells(cell.row, countyCol).value
            If IsDate(cell.Value) Then
                If cell.Value < oldestForeclosureDate Then
                    oldestForeclosureDate = cell.Value
                End If
            End If
        End If
        
        ' Check if the row corresponds to a tax deed with "x" in column Q (assuming column 17)
        If ws.Cells(cell.row, 17).Value = "x" Then
         '   county = ws.Cells(cell.row, countyCol).value
            If IsDate(ws.Cells(cell.row, taxDeedCol).Value) Then
                If ws.Cells(cell.row, taxDeedCol).Value < oldestTaxDeedDate Then
                    oldestTaxDeedDate = ws.Cells(cell.row, taxDeedCol).Value
                End If
            End If
        End If
    Next cell
    ' Determine the oldest sale type
If oldestForeclosureDate <= oldestTaxDeedDate Then
    tempOldestSaleType = "Foreclosure"
    
    ' Loop through the cells in the foreclosure column to find the row with the oldest date
    For i = 2 To lastRow
        If ws.Cells(i, foreclosureCol).Value = oldestForeclosureDate Then
            ' Check if the corresponding row's "List" column has "Yes" value
            If ws.Cells(i, ForeclousreYN_COl).Value = "x" Then
                tempListURL = "Foreclosure List"
                county = ws.Cells(i, 1).Value
                tempListURL = ws.Cells(i, ForeclosureListURLCol).Value
                Exit For ' Exit the loop once the list type is found
            End If
        End If
    Next i
Else
    tempOldestSaleType = "Tax Deed"
    
    ' Loop through the cells in the tax deed column to find the row with the oldest date
    For i = 2 To lastRow
        If ws.Cells(i, taxDeedCol).Value = oldestTaxDeedDate Then
            ' Check if the corresponding row's "TaxDeedList" column has "Yes" value
            If ws.Cells(i, taxDeedYN_COl).Value = "x" Then
                tempListURL = "Tax Deed List"
                county = ws.Cells(i, 1).Value
                tempListURL = ws.Cells(i, taxDeedListURLCol).Value
                Exit For ' Exit the loop once the list type is found
            End If
        End If
    Next i
End If

    ' Set the return values
    oldestCounty = county
    oldestSaleType = tempOldestSaleType
    listURL = tempListURL
End Sub
Sub resetAuctionDate(ByVal county As String, ByVal saleType As String, ByVal timestamp)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim countyCol As Integer, foreclosureCol As Integer, taxDeedCol As Integer
    Dim daysFCCol As Integer, daysTDCol As Integer
    Dim ForeclousreYN_COl As Integer, taxDeedYN_COl As Integer
    
    ' Define the worksheet
    Set ws = ThisWorkbook.Worksheets("Activity")
    
    ' Define the columns
    countyCol = 1 ' Assuming County column is in column A
    foreclosureCol = 2 ' Assuming Foreclosure column is in column B
    taxDeedCol = 3 ' Assuming Tax Deed column is in column C
    daysFCCol = 4 ' Assuming Days FC column is in column D
    daysTDCol = 5 ' Assuming Days TD column is in column E
    ForeclousreYN_COl = 16 ' Assuming Foreclosure Yes/No column is in column P
    taxDeedYN_COl = 17 ' Assuming Tax Deed Yes/No column is in column Q
    
    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, countyCol).End(xlUp).row
    
    ' Loop through the data
    For i = 2 To lastRow
        ' Check if the row corresponds to the given county and sale type
        If ws.Cells(i, countyCol).Value = county Then
            If saleType = "Foreclosure" And ws.Cells(i, ForeclousreYN_COl).Value = "x" Then
                ws.Cells(i, foreclosureCol).Value = timestamp ' Set foreclosure date to now
            ElseIf saleType = "Tax Deed" And ws.Cells(i, taxDeedYN_COl).Value = "x" Then
                ws.Cells(i, taxDeedCol).Value = timestamp ' Set tax deed date to now
            End If
        End If
    Next i
End Sub

Sub MainSub()
    Dim oldestCounty As String
    Dim oldestSaleType As String
    Dim listURL As String
    
    ' Call the subroutine and capture the returned values
    GetNextAuction oldestCounty, oldestSaleType, listURL
    
    ' Now you can use the returned values as needed
    MsgBox "The oldest county is: " & oldestCounty & vbCrLf & _
           "The oldest sale type is: " & oldestSaleType & vbCrLf & _
           "The list type is: " & listURL
End Sub


Sub DownloadAuctions()
    Dim oldestCounty As String
    Dim oldestSaleType As String
    Dim listURL As String
    Dim timestamp As Date
    Dim filePath As String
    Dim fileName As String
    Dim FSO As Object
    
    ' Create FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Call the subroutine and capture the returned values
    GetNextAuction oldestCounty, oldestSaleType, listURL
    
    ' Save current time to timestamp
    timestamp = Now
    
    ' Open listURL in Edge
    ShellExecute 0, "open", "msedge.exe", listURL, "", 1
    
    ' Wait for Edge to open
    Application.Wait (Now + TimeValue("0:00:05")) ' Wait for 5 seconds for Edge to open
    
    ' Message box asking user to download and click OK
    Dim response As VbMsgBoxResult
    response = MsgBox("Please download " & oldestCounty & " County " & oldestSaleType & " Sale. " & vbCrLf & _
                  "Do you want to continue downloading or process downloaded?" & vbCrLf & _
                  "Click YES to continue downloading, NO to process downloaded, or Cancel to do nothing.", vbYesNoCancel)
    'Dim timestamp As Date: Date = Now
    
    If response = vbYes Then
        ' Get last file from c:\downloads
        filePath = "C:\Users\philk\Downloads"
        fileName = Dir(filePath & "\*.*", vbNormal)
        Dim lastFile As String
        Dim lastFileCreated As Date
        
        Do While fileName <> ""
            If FileDateTime(filePath & "\" & fileName) > lastFileCreated Then
                lastFileCreated = FileDateTime(filePath & "\" & fileName)
                lastFile = fileName
            End If
            fileName = Dir
        Loop
        
        ' Move the last file if created after timestamp
        If lastFileCreated > timestamp Then
            ' Create "auctions" folder if it doesn't exist
            If Not FSO.FolderExists(filePath & "\Auctions") Then
                FSO.CreateFolder filePath & "\Auctions"
            End If
            
            ' Move the file and rename it
            On Error Resume Next
            FSO.MoveFile Source:=filePath & "\" & lastFile, Destination:=filePath & "\Auctions\" & oldestCounty & "_" & oldestSaleType & "_" & Format(timestamp, "YYYYMMDD_HHMMSS") & ".csv"
            If Err.Number <> 0 Then
                MsgBox "Error moving the file: " & Err.description, vbExclamation
            End If
            On Error GoTo 0
            
            ' Update the sale date of the county to the current date in the Activities worksheet
            Dim ws As Worksheet
            Set ws = ThisWorkbook.Worksheets("Activity")
            Dim countyRow As Long
            countyRow = Application.match(oldestCounty, ws.Columns(1), 0)
            If Not IsError(countyRow) Then
                colNum = IIf(saleType = "Foreclosure", 2, 3)
                ws.Cells(countyRow, colNum).Value = DateValue(Now)
            End If
        End If
        
        ' Delete files with "QUICKSEARCH" in their names
        fileName = Dir(filePath & "\*QUICKSEARCH*")
        Do While fileName <> ""
            Kill filePath & "\" & fileName
            fileName = Dir
        Loop
        resetAuctionDate oldestCounty, oldestSaleType, timestamp
         Dim url As String
    Dim jsonData As String
    Dim success As Boolean

    ' Set the URL of the Wix API endpoint

    ' Define the JSON data to be sent in the request body
    jsonData = "{""request"":{""path"":[""myPath1"",""myPath2""],""headers"":{""Accept"":""*/*"",""Content-Type"":""application/json""},""body"":{""json"":{""county"":""" & county & """,""saleType"":""" & saleType & """}}}}"
    ' Call the CallWixAPI function and get the result
    success = UpdateCountyInWIX(jsonData:=jsonData)

    ' Do something based on the success status
    If success Then
        Debug.Print "API call was successful."
    Else
        Debug.Print "API call failed."
    End If
    ElseIf response = vbNo Then
        ProcessAuctionQueue
    ElseIf response = vbCancel Then
    ' Code to handle Cancel response
    ' Add your logic here
    End If
End Sub



Sub ProcessAuctionQueue()
    Dim FSO As Object
    Dim auctionFolder As Object
    Dim file As Object
    Dim filePath As String
    Dim fileName As String
    Dim county As String
    Dim saleType As String
    Dim listURL As String
    Dim lastFileCreated As String
    
    ' Create FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Set the path to the auction folder
    filePath = "C:\Users\philk\Downloads\Auctions" ' Change this to your actual folder path
    
    ' Check if the auction folder exists
    If FSO.FolderExists(filePath) Then
        ' Get the auction folder
        Set auctionFolder = FSO.GetFolder(filePath)
        
        ' Loop through each file in the auction folder
        For Each file In auctionFolder.Files
            ' Get the file name
            fileName = filePath & "\" & file.name
            
            ' Extract county, sale type, and list URL from the file name
            ' Assuming the file name format is "County_SaleType_YYYYMMDD_HHMMSS.csv"
            county = Split(Split(fileName, "_")(0), filePath & "\")(1)
            saleType = Split(Split(fileName, "_")(1), ".")(0)
            lastFileCreated = Split(Split(fileName, "_")(2), ".")(0)
            
            ' Call function to process the auction
            ProcessFromUiPath county, saleType, fileName
            
            ' Delete the file from the auction folder
            On Error Resume Next
            FSO.DeleteFile fileName
            
            ' Save or update parameter
            UpdateParameterValue parameterName:="lastUpdate_" & county & "_" & saleType, newValue:=lastFileCreated
        Next file
    Else
        MsgBox "Auctions folder does not exist!"
    End If
End Sub

Function UpdateCountyInWIX(jsonData As String, Optional url As String = "https://werkhardor.wixsite.com/freeforeclosurelist/_functions-dev/countyLastUpdated") As Boolean
    Dim httpRequest As Object

    ' Create a new HTTP request object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")

    ' Open a connection to the specified URL
    httpRequest.Open "PUT", url, False

    ' Set the request headers
    httpRequest.setRequestHeader "Content-Type", "application/json"
    ' You can add other headers if necessary

    ' Send the JSON data in the request body
    httpRequest.send jsonData

    ' Check if the request was successful (status code 200)
    If httpRequest.Status = 200 Then
        ' Print the response from the server
        Debug.Print "Request successful. Response: " & httpRequest.responseText
        CallWixAPI = True
    Else
        ' Print an error message
        Debug.Print "Error: " & httpRequest.StatusText
        CallWixAPI = False
    End If
End Function







'    Dim content As String
'    Dim i As Long
'    Dim fso As Object
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Dim textStream As Object
'    Set textStream = fso.OpenTextFile(rt.filePath, 1, False)
'    content = textStream.ReadAll
'    textStream.Close
Sub Main()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Activity")
    Dim sheetName As String
    Dim countyColumn As range
    Dim paTemplate As String
    Dim paTemplateAddr As String
    Dim countyName As String
    Dim foreclosureStartAt As Date
    Dim taxDeedStartAt As Date
    Dim countyCell As range
    Dim rt As rowType
    Dim salesType As Variant
    ' Initialize the global variable
    mySemaphore = False
    rt.sheetName = "Sales"
    rt.home = "2665 Hempel Ave, Windermere FL 34786"
    rt.fontSize = 6
    
    Set countyColumn = ws.range("A:A")
       
    For Each countyCell In countyColumn.Cells
        If countyCell.Value = "Start at" Then
            foreclosureStartAt = countyCell.Offset(0, GetColumnNumberByHeaderName("Foreclosure", "Vars")).Value
            taxDeedStartAt = countyCell.Offset(0, GetColumnNumberByHeaderName("Tax Deed", "Vars")).Value
            rt.foreclosureStartAt = foreclosureStartAt
            rt.taxDeedStartAt = taxDeedStartAt
        End If
    Next
    'Call toggleFilter
    clearFilter
    Call ProcessSelected(foreclosureStartAt:=foreclosureStartAt, taxDeedStartAt:=taxDeedStartAt, rt:=rt)
    
'    If False Then
'    For Each countyCell In countyColumn.Cells
'        If Not IsEmpty(countyCell.Value) Then
'            sheetName = "Activity"
'            countyName = countyCell.Value
'       'paTemplate = countyCell.Offset(0, GetColumnNumberByHeaderName("PA template", sheetName)).Value
'            paTemplateAddr = countyCell.Offset(0, GetColumnNumberByHeaderName("PA template Address", sheetName)).Value
'       '     Call getPA_Templates(countyCell:=countyCell, sheetName:=sheetName, countyName:=countyName, paTemplate:=paTemplate, paTemplateAddr:=paTemplateAddr)
'            sheetName = "Sales"
'
'            For Each salesType In Array("Foreclosure", "Tax Deed")
'                Dim filePath As String
'                If UCase(countyName) <> "COUNTY" Then
'                    If salesType = "Tax Deed" Then
'                        filePath = "C:\Users\philk\Downloads\QuickSearch" & Replace(countyName, " ", "") & "TD.csv"
'                    Else
'                        filePath = "C:\Users\philk\Downloads\QuickSearch" & Replace(countyName, " ", "") & ".csv"
'                    End If
'
'                    Call ProcessCounty(rt_orginal:=rt, county:=countyName, filePath:=filePath, sheetName:=sheetName, saletype:=salesType, paTemplate:=paTemplate, paTemplateAddr:=paTemplateAddr)
'                End If
'
'            Next salesType
'        End If
'    Next countyCell
'    End If
    'ProcessBrevardCounty rt:=rt, URL:="http://vweb2.brevardclerk.us/Foreclosures/foreclosure_sales.html" ', sheetName
    ProcessLakeCounty rt ', "Lake", "https://www.lakecountyclerk.org/record_searches/public_sales_calendar/", sheetName
    'ProcessOsceolaCounty "Osceola", "https://courts.osceolaclerk.com/reports/CivilMortgageForeclosuresWeb.pdf", sheetName
 '   ProcessOsceolaCounty rt ', "Osceola", "C:\Users\philk\Downloads\OsceolaSalesList.txt", sheetName
    'ProcessSumterCounty "Sumter", "C:\Users\philk\Desktop\SumterSalesList.txt", sheetName
  '  ProcessCounty "Highlands", "C:\Users\philk\Desktop\HighlandsSalesList.txt", sheetName, ""
    Call SortByForeclosureColumn
    Call FormatSaleDateColumnAll
    'Call toggleFilter(True)
    Call FilterSalesWorksheet
    UpdateParameterValue "VBA_status", "completed"
    Dim i
    For i = 1 To 3 ' Loop 3 times.
     Beep
    Next i
    Debug.Print "Successful Completion"
End Sub
