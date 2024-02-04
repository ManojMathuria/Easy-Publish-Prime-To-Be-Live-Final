Attribute VB_Name = "ModUserDefinedFunctions"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Function FixAPIString(ByVal strInput As String) As String
    ' strips Trailing Nulls From strings returned by API Function
    Dim intPosition As Integer
    
    intPosition = InStr(1, strInput, Chr(0))
    If intPosition Then
        FixAPIString = Left(strInput, intPosition - 1)
    Else
        FixAPIString = strInput     ' No Nulls Found
    End If
End Function

Public Function ReadFromFile(ByVal xKeyName As String) As String
    Dim sReturn As String * 255
    GetPrivateProfileString "ewaybill", xKeyName, "", sReturn, 255, App.Path + "\" + IIf(CheckEmpty(Command$, False), "ewaybill", Command$) + ".json"
    sReturn = FixAPIString(sReturn)
    ReadFromFile = RTrim(sReturn)
End Function
Public Sub WriteToFile(ByVal xKeyName As String, ByVal xString As String)
    WritePrivateProfileString "ewaybill", xKeyName, xString, App.Path + "\" + IIf(CheckEmpty(Command$, False), "ewaybill", Command$) + ".json"
End Sub
Public Function CheckEmpty(ByVal strExpression As Variant, ByVal xDspMsg As Boolean) As Boolean
    If LTrim(RTrim(strExpression)) = "" Or IsNull(strExpression) Then
       If xDspMsg Then DisplayError ("Mandatory Field")
       CheckEmpty = True
    End If
End Function
Public Function Base64Encode()
Dim xmlDoc As New MSXML2.DOMDocument60
Dim encodedString As String
Dim binaryData() As Byte

' Convert a string to Base64
encodedString = xmlDoc.createElement("b64").nodeTypedValue
encodedString = Base64Encode("Hello World!")

' Convert binary data to Base64
binaryData = file.ReadAllBytes("image.jpg")
encodedString = xmlDoc.createElement("b64").nodeTypedValue
encodedString = Base64Encode(binaryData)
End Function
Public Sub DisplayError(ByVal strErrorMsg As String)
    On Error Resume Next
    Beep
    MsgBox RTrim(LTrim(strErrorMsg)) & " !!!", vbExclamation, "Error !"
    Err.Clear
End Sub

