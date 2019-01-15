Imports System.Net
Imports System.Text
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO.Directory
Imports System.IO
Imports System.Xml
Imports ClosedXML.Excel
Imports Microsoft.Web.Administration

Module Module1

    Sub Main()
        'create variables
        Dim notificationNo As String()
        Dim notificationDate As String()
        Dim notificationDesc As String()
        Dim sgID As String()
        Dim township As String()
        Dim functionalLoc As String()
        Dim account As String()
        Dim name As String()
        Dim customerTown As String()
        Dim streetNo As String()
        Dim streetAddress As String()
        Dim postcode As String()
        Dim code As String()
        Dim shortTxt As String()
        Dim task As String()
        Dim taskTxt As String()
        Dim createdOn As String()
        Dim filePath As String
        Dim wb As XLWorkbook
        'Dim lastcol As Integer
        Dim newCase, updateCase, queryCase As String
        Dim service As BizagiWebReference02.EntityManagerSOA = New BizagiWebReference02.EntityManagerSOA()
        Dim insertService As InsertService.WorkflowEngineSOA = New InsertService.WorkflowEngineSOA
        Dim xmldoc As XmlDocument = New XmlDocument()
        'data file
        Dim dataFile As String

        'set web services
        Dim serviceSR = New StreamReader("C:\inetpub\wwwroot\WebTaskVB\WebTaskVB\service.txt")
        service.Url = serviceSR.ReadLine
        Dim insertServiceSR = New StreamReader("C:\inetpub\wwwroot\WebTaskVB\WebTaskVB\insertService.txt")
        insertService.Url = insertServiceSR.ReadLine

        'set filepath to folder with excel files
        filePath = "C:\inetpub\wwwroot\WebTaskVB\WebTaskVB\Settings.txt"
        ' Open the file to read from.
        Dim readText() As String = File.ReadAllLines(filePath)
        Dim s As String
        For Each s In readText
            dataFile = s
        Next

        ''get virtual folder path
        'Dim serverManager As New ServerManager()
        '' get the site (e.g. default)
        'Dim site As Site = serverManager.Sites.FirstOrDefault(Function(r) r.Name = "Default Web Site")
        '' get the application that you are interested in
        'Dim myApp As Application = site.Applications("/" & dataFile)
        '' get the physical path of the virtual directory
        'Console.WriteLine(myApp.VirtualDirectories(0).PhysicalPath)

        'loop through files in folder
        Dim Directory As New IO.DirectoryInfo(dataFile)
        Dim allFiles As IO.FileInfo() = Directory.GetFiles("*.xls")
        Dim singleFile As IO.FileInfo
        ReDim notificationNo(50000)
        ReDim notificationDate(50000)
        ReDim notificationDesc(50000)
        ReDim sgID(50000)
        ReDim township(50000)
        ReDim functionalLoc(50000)
        ReDim account(50000)
        ReDim name(50000)
        ReDim customerTown(50000)
        ReDim streetNo(50000)
        ReDim postcode(50000)
        ReDim streetAddress(50000)
        ReDim code(50000)
        ReDim shortTxt(50000)
        ReDim task(50000)
        ReDim taskTxt(50000)
        ReDim createdOn(50000)

        ''extract information into arrays
        For Each singleFile In allFiles
            Console.WriteLine(singleFile.FullName)
            If InStr(singleFile.FullName, "$") = 0 Then
                'open excel workbook to extract info
                wb = New XLWorkbook(singleFile.FullName)
                ''wb = xl.Workbooks.Open(singleFile.FullName)
                Dim ws = wb.Worksheet(1)
                Dim lastrow = ws.RangeUsed

                'loop through rows of excel sheet
                Dim i As Integer
                Dim rowCount As Integer
                i = 0
                For Each row In ws.Rows
                    rowCount = row.RowNumber
                    notificationNo(i) = ws.Cell(rowCount, 1).GetString
                    notificationDesc(i) = ws.Cell(rowCount, 3).GetString
                    account(i) = ws.Cell(rowCount, 4).GetString
                    functionalLoc(i) = ws.Cell(rowCount, 5).GetString
                    township(i) = ws.Cell(rowCount, 6).GetString
                    notificationDate(i) = ws.Cell(rowCount, 8).GetString
                    name(i) = ws.Cell(rowCount, 9).GetString
                    customerTown(i) = ws.Cell(rowCount, 10).GetString
                    'streetNo(i) = ws.Cell(rowCount, 14).GetString
                    streetAddress(i) = ws.Cell(rowCount, 11).GetString
                    sgID(i) = ws.Cell(rowCount, 12).GetString
                    postcode(i) = ws.Cell(rowCount, 13).GetString
                    code(i) = ws.Cell(rowCount, 14).GetString
                    shortTxt(i) = ws.Cell(rowCount, 15).GetString
                    task(i) = ws.Cell(rowCount, 16).GetString
                    taskTxt(i) = ws.Cell(rowCount, 18).GetString
                    createdOn(i) = ws.Cell(rowCount, 19).GetString
                    i = i + 1
                Next
            End If
            singleFile.Delete()
        Next

        ''write data over to bizagi
        For i = 1 To code.Length
            'check which code it is
            If code(i - 1) = "X1" Or code(i - 1) = "P02" Then
                'write values to xml
                newCase = "<BizAgiWSParam>"
                newCase += "<domain>City Power</domain>"
                newCase += "<userName>RlaCock</userName>"
                newCase += "<Cases>"
                newCase += "<Case>"
                newCase += "<Process>ServiceConnectionApplicati</Process>"
                newCase += "<Entities>"
                newCase += "<ServiceConnectionApp>"
                If InStr(notificationNo(i - 1), "&") > 0 Then
                    notificationNo(i - 1) = Replace(notificationNo(i - 1), "&", "and")
                End If
                newCase += "<NotificationNo>" & notificationNo(i - 1) & "</NotificationNo>"
                newCase += "<NotificationDate>" & notificationDate(i - 1) & "</NotificationDate>"
                If InStr(notificationDesc(i - 1), "&") > 0 Then
                    notificationDesc(i - 1) = Replace(notificationDesc(i - 1), "&", "and")
                End If
                newCase += "<ApplicationDescription>" & notificationDesc(i - 1) & "</ApplicationDescription>"
                If InStr(township(i - 1), "&") > 0 Then
                    township(i - 1) = Replace(township(i - 1), "&", "and")
                End If
                newCase += "<SCLocation.Depot>" & township(i - 1) & "</SCLocation.Depot>"
                If InStr(functionalLoc(i - 1), "&") > 0 Then
                    functionalLoc(i - 1) = Replace(functionalLoc(i - 1), "&", "and")
                End If
                newCase += "<SCLocation.FunctionalLocation>" & functionalLoc(i - 1) & "</SCLocation.FunctionalLocation>"
                If InStr(account(i - 1), "&") > 0 Then
                    account(i - 1) = Replace(account(i - 1), "&", "and")
                End If
                newCase += "<Customer.CustomerNo>" & account(i - 1) & "</Customer.CustomerNo>"
                If InStr(name(i - 1), "&") > 0 Then
                    name(i - 1) = Replace(name(i - 1), "&", "and")
                End If
                newCase += "<Customer.CustomerNames>" & name(i - 1) & "</Customer.CustomerNames>"
                If InStr(customerTown(i - 1), "&") > 0 Then
                    customerTown(i - 1) = Replace(customerTown(i - 1), "&", "and")
                End If
                newCase += "<Customer.CustomerTownship>" & customerTown(i - 1) & "</Customer.CustomerTownship>"
                newCase += "<Customer.CustomerStreetNo>" & streetAddress(i - 1) & "</Customer.CustomerStreetNo>"
                newCase += "<Customer.CustomerPostCode>" & postcode(i - 1) & "</Customer.CustomerPostCode>"
                newCase += "<SCLocation.SGCode>" & sgID(i - 1) & "</SCLocation.SGCode>"
                newCase += "<LastTaskUpdate>" & createdOn(i - 1) & "</LastTaskUpdate>"
                'newCase += "<SAPUpdates.NotificationNo>" & notificationNo(i - 1) & "</SAPUpdates.NotificationNo>"
                'newCase += "<SAPUpdates.TaskCode>" & code(i - 1) & "</SAPUpdates.TaskCode>"
                'newCase += "<SAPUpdates.TaskCodeText>" & shortTxt(i - 1) & "</SAPUpdates.TaskCodeText>"
                'newCase += "<SAPUpdates.TaskOrder>" & task(i - 1) & "</SAPUpdates.TaskOrder>"
                'newCase += "<SAPUpdates.TaskDetails>" & taskTxt(i - 1) & "</SAPUpdates.TaskDetails>"
                'newCase += "<SAPUpdates.TaskDate>" & createdOn(i - 1) & "</SAPUpdates.TaskDate>"
                newCase += "</ServiceConnectionApp>"
                newCase += "</Entities>"
                newCase += "</Case>"
                newCase += "</Cases>"
                newCase += "</BizAgiWSParam>"

                'case to query on
                queryCase = "<BizAgiWSParam>"
                queryCase += "<EntityData>"
                queryCase += "<EntityName>ServiceConnectionApp</EntityName>"
                queryCase += "  <Filters>" & _
                 "<![CDATA[NotificationNo = " & notificationNo(i - 1) & "]]>" & _
                "</Filters>"
                queryCase += "</EntityData>"
                queryCase += "</BizAgiWSParam>"
                xmldoc.LoadXml(service.getEntitiesAsString(queryCase))
                ''xmldoc.Save("C:\Users\bkruger\Desktop\Doc.xml")

                'loop xml files.
                Dim elemList As XmlNodeList = xmldoc.GetElementsByTagName("NotificationNo")
                Dim counter As Integer
                Dim exist As String
                For counter = 0 To elemList.Count - 1
                    exist = elemList(counter).InnerXml
                    Console.WriteLine(elemList(counter).InnerXml)
                Next counter
                If exist = "" Then
                    Console.WriteLine(insertService.createCasesAsString(newCase))
                End If
                'xmldoc.LoadXml(newCase)
                newCase = ""
                exist = ""
                ''xmldoc.Load("C:\Users\bkruger\Desktop\testCase.xml")
                ''service.createCasesAsString(newCase)               
            ElseIf code(i - 1) = "P07" Or code(i - 1) = "P03" Or code(i - 1) = "P04" Or code(i - 1) = "P05" Or code(i - 1) = "P07" Or code(i - 1) = "P08" Or code(i - 1) = "P011" Or code(i - 1) = "P012" Or code(i - 1) = "P013" Or code(i - 1) = "P015" Or code(i - 1) = "P016" Or code(i - 1) = "P017" Or code(i - 1) = "P018" Or code(i - 1) = "P019" Or code(i - 1) = "P020" Or code(i - 1) = "P021" Or code(i - 1) = "P022" Or code(i - 1) = "P023" Or code(i - 1) = "X10" Or code(i - 1) = "X11" Or code(i - 1) = "X12" Or code(i - 1) = "X13" Or code(i - 1) = "X14" Or code(i - 1) = "X15" Or code(i - 1) = "X16" Or code(i - 1) = "X17" Or code(i - 1) = "X18" Or code(i - 1) = "X19" Or code(i - 1) = "X2" Or code(i - 1) = "X20" Or code(i - 1) = "X21" Or code(i - 1) = "X22" Or code(i - 1) = "X23" Or code(i - 1) = "X24" Or code(i - 1) = "X25" Or code(i - 1) = "X26" Or code(i - 1) = "X3" Or code(i - 1) = "X4" Or code(i - 1) = "X5" Or code(i - 1) = "X6" Or code(i - 1) = "X7" Or code(i - 1) = "X8" Or code(i - 1) = "X9" Then
                'case to query on
                queryCase = "<BizAgiWSParam>"
                queryCase += "<EntityData>"
                queryCase += "<EntityName>ServiceConnectionApp</EntityName>"
                queryCase += "  <Filters>" & _
                 "<![CDATA[NotificationNo = " & notificationNo(i - 1) & "]]>" & _
                "</Filters>"
                queryCase += "</EntityData>"
                queryCase += "</BizAgiWSParam>"
                xmldoc.LoadXml(service.getEntitiesAsString(queryCase))
                ''xmldoc.Save("C:\Users\bkruger\Desktop\Doc.xml")

                'loop xml files.
                Dim elemList As XmlNodeList = xmldoc.GetElementsByTagName("CPReference")
                Dim counter As Integer
                Dim exist As String
                For counter = 0 To elemList.Count - 1
                    exist = elemList(counter).InnerXml
                    exist = Right(exist, Len(exist) - InStr(exist, "_"))
                Next counter
                'write values to xml
                updateCase = "<BizAgiWSParam>"
                updateCase += "<Process>ServiceConnectionApplicati</Process>"
                updateCase += "<Entities idCase=""" & exist & """ > "
                updateCase += "<ServiceConnectionApp>"
                updateCase += "<SAPUpdates>"
                updateCase += "<NotificationNo>" & notificationNo(i - 1) & "</NotificationNo>"
                updateCase += "<TaskCode>" & code(i - 1) & "</TaskCode>"
                If InStr(shortTxt(i - 1), "&") > 0 Then
                    shortTxt(i - 1) = Replace(shortTxt(i - 1), "&", "and")
                End If
                updateCase += "<TaskCodeText>" & shortTxt(i - 1) & "</TaskCodeText>"
                updateCase += "<TaskOrder>" & task(i - 1) & "</TaskOrder>"
                If InStr(taskTxt(i - 1), "&") > 0 Then
                    taskTxt(i - 1) = Replace(taskTxt(i - 1), "&", "and")
                End If
                updateCase += "<TaskDetails>" & taskTxt(i - 1) & "</TaskDetails>"
                updateCase += "<TaskDate>" & createdOn(i - 1) & "</TaskDate>"
                'updateCase += "<LastTaskUpdate>" & createdOn(i - 1) & "</LastTaskUpdate>"
                updateCase += "</SAPUpdates>"
                updateCase += "</ServiceConnectionApp>"
                updateCase += "</Entities>"
                updateCase += "</BizAgiWSParam>"
                xmldoc.LoadXml(updateCase)
                ''xmldoc.Save("C:\Users\bkruger\Desktop\Doc.xml")
                Console.WriteLine(service.saveEntityAsString(updateCase))
                'xmldoc.Load()
                updateCase = ""
            End If
        Next

        Console.WriteLine("Please Press any key to continue!")
        Console.ReadKey()


    End Sub

End Module
