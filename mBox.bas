Attribute VB_Name = "mBox"
Sub createfolder(token As String)

'define XML and HTTP components
Dim zipResult As New MSXML2.DOMDocument
Dim zipService As New MSXML2.XMLHTTP

Dim xmlInput As String 'this is the JSON that will form the body of the http request
Dim query As String 'this is the URL, including any parameters

'write the JSON that will go in the body of the http request
xmlInput = "{""name"": ""New Folder"", ""parent"": {""id"": ""0""}}"

'write the query string
query = "https://api.box.com/2.0/folders"

'create HTTP request to query URL
zipService.Open "POST", query, False

'set HTTP request header
zipService.setRequestHeader "Authorization:", "Bearer " & token

'send HTTP request
zipService.send xmlInput

End Sub

Sub main2()

GetFolderContentsBOX "vW7S0DqKNrRwz2x2f0O5q1Em5JZxlpMC", "8479106597"

End Sub

Sub main()

DownloadFile "vW7S0DqKNrRwz2x2f0O5q1Em5JZxlpMC", "72956081037"

End Sub

Sub main3()

createfolder "W9aKuiW64r1EXqSZrxB1uYPicEIa5Ran"

End Sub

Sub UploadBoxFile(ByVal sToken As String)

Dim curlInput As XMLHTTP60
Dim sQuery As String
Dim sXMLInput As String
' need to imitate a form in cURL to complete this one (thats what I managed to get from the cURL documentation on the web
'a google research sugested to use XMLHTTP.6.0. instead of XMLHTTP.

Set curlInput = CreateObject("MSXML2.XMLHTTP.6.0")

' a different location is used for uploading

sQuery = "https://upload.box.com/api/2.0/files/content"

' I think that the mistake is in this string I need to build it a bit differently or somehow indicated the -F parameter when sending

sXMLInput = "attributes={name: ""FullDump.xlsx"", ""parent"": {""id"": ""8479106597""}}" & vbNewLine & "file=D:\Reporting\NewDashboard\Dashboard2.0.xlsx"

curlInput.Open "POST", sQuery, False

curlInput.setRequestHeader "Authorization:", "Bearer " & sToken
curlInput.send sXMLInput

End Sub

Sub GetFolderContentsBOX(ByVal sToken As String, ByVal sFolderID As String)

Dim curlInput As MSXML2.XMLHTTP
Dim curlOutput As MSXML2.DOMDocument
Dim test As Variant
Dim sQuery As String
Dim sResult As String




sQuery = "https://api.box.com/2.0/folders/" & sFolderID & "/items?limit=100&offset=0"

Set curlInput = CreateObject("MSXML2.XMLHTTP")

curlInput.Open "GET", sQuery, False

curlInput.setRequestHeader "Authorization:", "Bearer " & sToken
curlInput.send
Dim doc As Object
Set curlOutput = curlInput.responseXML
sResult = curlInput.responseText

ThisWorkbook.Sheets(1).Cells(1, 1).Value = sResult




End Sub


Sub DownloadFile(ByVal sToken As String, ByVal sFileID As String)


Dim wbBook As Workbook
Dim curlInput As MSXML2.XMLHTTP
Dim curlOutput As MSXML2.DOMDocument
Dim test As Variant
Dim sQuery As String
Dim sResult As String




sQuery = "https://api.box.com/2.0/files/" & sFileID & "/content"

Set curlInput = CreateObject("MSXML2.XMLHTTP")

curlInput.Open "GET", sQuery, False

curlInput.setRequestHeader "Authorization:", "Bearer " & sToken
curlInput.send
sResult = curlInput.responseText

ThisWorkbook.Sheets(1).Cells(1, 1).Value = sResult


End Sub


