VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   1095
      Left            =   4800
      TabIndex        =   3
      Top             =   3360
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   975
      Left            =   5160
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1335
      Left            =   840
      TabIndex        =   1
      Top             =   2880
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create JSON"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xmlHttp As New MSXML2.XMLHTTP60
Dim url As String
Dim postData As String
Dim response As String
Private Sub Command1_Click()
Dim jsonFile As Integer
jsonFile = FreeFile

'Dim xmlHttp As New MSXML2.XMLHTTP60  'As New 'MSXML2.xmlHttp
'Dim url As String
'Dim postData As String

url = "https://ewaybillgst.gov.in/BillGeneration/BillGenerationService.svc/BillGeneration"

'Print ewayBill.json
Open App.Path & "\ewaybill.json" For Output As jsonFile
Print #jsonFile, "{"
Print #jsonFile, "  ""user"": {"
Print #jsonFile, "    ""username"": ""JohnDoe"","
Print #jsonFile, "    ""password"": ""s3cr3t"""
Print #jsonFile, "  },"
Print #jsonFile, "  ""product"": {"
Print #jsonFile, "    ""name"": ""Widget"","
Print #jsonFile, "    ""description"": ""A great widget"","
Print #jsonFile, "    ""quantity"": 1,"
Print #jsonFile, "    ""price"": 10.99"
Print #jsonFile, "  },"
Print #jsonFile, "  ""customer"": {"
Print #jsonFile, "    ""name"": ""JaneDoe"","
Print #jsonFile, "    ""address"": ""123 Main St"","
Print #jsonFile, "    ""city"": ""Anytown"","
Print #jsonFile, "    ""state"": ""CA"","
Print #jsonFile, "    ""zip"": ""12345"""
Print #jsonFile, "  }"
Print #jsonFile, "}"
Close jsonFile

xmlHttp.Open "POST", url, False
xmlHttp.setRequestHeader "Content-Type", "application/json"
xmlHttp.send postData

'Dim response As String
response = xmlHttp.responseText

'parse the response to get the e-way bill number'
End Sub

Private Sub Command2_Click()
Dim json As New Scripting.Dictionary
Dim jsonString As String

' Read ewaybill.json file contents into a string variable
Open "path/to/ewaybill.json" For Input As #1
jsonString = Input(LOF(1), #1)
Close #1

' Parse the JSON string and store the result in a dictionary object
Set json = JsonConverter.ParseJson(jsonString)

End Sub

Private Sub Command3_Click()
Dim json As New Scripting.Dictionary
Dim jsonString As String

' Add the required data to the JSON dictionary
json("UserName") = "your_username"
json("Password") = "your_password"
json("TransMode") = "1"
json("TransDocNo") = "TR1001"
json("TransDocDate") = "01-01-2022"
json("VehicleNo") = "KA-1234"
json("FromStateCode") = "29"
json("ToStateCode") = "33"
json("DocType") = "INV"
json("DocNo") = "INV1001"
json("DocDate") = "01-01-2022"
json("Qty") = 10
json("Value") = 1000
json("TaxableAmount") = 900
json("IGST") = 180
json("CGST") = 81
json("SGST") = 81

' Encode the JSON dictionary as a string
jsonString = JsonConverter.ConvertToJson(json)


Dim xmlHttp As New MSXML2.XMLHTTP60
Dim url As String

' Build the URL for the POST request
url = "https://ewaybillgst.gov.in/BillGenerationAPI/BillGeneration/BillGenerate"

' Set the authentication headers
xmlHttp.setRequestHeader "Authorization", "Basic " & Base64Encode("your_username" & ":" & "your_password")

' Set the request headers
xmlHttp.setRequestHeader "Content-Type", "application/json"

' Send the POST request with the JSON request file as the payload
xmlHttp.Open "POST", url, False
xmlHttp.send jsonString

' Get the response from the GST portal API
Dim response As String
response = xmlHttp.responseText





Dim responseJson As New Scripting.Dictionary

' Parse the JSON response and store the result in a dictionary object
Set responseJson = JsonConverter.ParseJson(response)

' Extract the e-way bill number from the response dictionary
Dim ewayBillNumber As String
ewayBillNumber = responseJson("Data")("EWBNo")



End Sub

Private Sub Command4_Click()
url = "https://ewaybillgst.gov.in/BillGeneration/BillGenerationService.svc/BillGeneration"
postData = "{  ""UserName"": ""your username"",  ""Password"": ""your password"",  ""Bill"": {    ""TransactionType"": 1,    ""SupplyType"": 1,    ""DocType"": 1,    ""SubSupplyType"": 1,    ""DocNo"": ""doc number"",    ""DocDate"": ""doc date"",    ""FromGSTIN"": ""from GSTIN"",    ""FromName"": ""from name"",    ""FromAddress"": ""from address"",    ""FromPincode"": ""from pincode"",    ""FromStateCode"": ""from state code"",    ""ToGSTIN"": ""to GSTIN"",    ""ToName"": ""to name"",    ""ToAddress"": ""to address"",    ""ToPincode"": ""to pincode"",    ""ToStateCode"": ""to state code"",    ""TotalValue"": ""total value"",    ""CGST"": ""cgst value"",    ""SGST"": ""sgst value"",    ""IGST"": ""igst value"",    ""CESS"": ""cess value"",    ""TransportMode"": 1,    ""VehicleNo"": ""vehicle number"",    ""Distance"": ""distance"",    ""TransporterID"": ""transporter ID"",    ""TransporterName"": ""transporter name"",    ""TransporterDocNo"": ""transporter document number"", " & _
    """TransporterDocDate"": ""transporter document date""  }}"

xmlHttp.Open "POST", url, False
xmlHttp.setRequestHeader "Content-Type", "application/json"
xmlHttp.send postData

response = xmlHttp.responseText
Debug.Print xmlHttp.responseText

'parse the response to get the e-way bill number'

End Sub
