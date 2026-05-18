Attribute VB_Name = "modMidasCore"
' ==========================================
' MIDAS API Module
' ==========================================
Option Explicit
Private Const BASE_URL As String = "https://moa-engineers.midasit.com:443/civil"

' ==========================================
' Get MAPI Key from Registry
' ==========================================
Private Function GetMapiKey() As String
    On Error GoTo ErrorHandler
    
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    GetMapiKey = shell.RegRead("HKCU\Software\MIDAS\CVLwNX_KR\CONNECTION\key")
    Exit Function
    
ErrorHandler:
    Err.Raise vbObjectError + 1001, "GetMapiKey", _
              "Failed to read MAPI Key from registry: " & Err.Description
End Function

' ==========================================
' Core HTTP Request
' ==========================================
Private Function ApiRequest(METHOD As String, endpoint As String, _
                            Optional payload As Variant = Null) As Object
    On Error GoTo ErrorHandler
    
    Dim bodyStr As String
    If IsNull(payload) Then
        bodyStr = "{}"
    ElseIf IsObject(payload) Then
        bodyStr = JsonConverter.ConvertToJson(payload)
    Else
        bodyStr = "{}"
    End If
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    
    http.Open UCase(METHOD), BASE_URL & endpoint, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "MAPI-Key", GetMapiKey()
    http.Send bodyStr
    
    Debug.Print "Status: " & http.Status
    Debug.Print "Response: " & http.responseText
    
    Set ApiRequest = JsonConverter.ParseJson(http.responseText)
    
    Set http = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise vbObjectError + 1002, "ApiRequest", _
              "API request failed: " & Err.Description
End Function

' ==========================================
' HTTP Methods
' ==========================================
Private Function ApiGet(endpoint As String) As Object
    Set ApiGet = ApiRequest("GET", endpoint)
End Function

Private Function ApiPost(endpoint As String, Optional payload As Variant = Null) As Object
    Set ApiPost = ApiRequest("POST", endpoint, payload)
End Function

Private Function ApiPut(endpoint As String, Optional payload As Variant = Null) As Object
    Set ApiPut = ApiRequest("PUT", endpoint, payload)
End Function

Private Function ApiDelete(endpoint As String, Optional payload As Variant = Null) As Object
    Set ApiDelete = ApiRequest("DELETE", endpoint, payload)
End Function

' ==========================================
' Public Functions
' ==========================================

' ------------------------------------------
' DOC
' ------------------------------------------

' doc/NEW ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function NewFile() As Object
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload.Add "Argument", CreateObject("Scripting.Dictionary")
    
    Set NewFile = ApiPost("/doc/NEW", payload)
End Function

' doc/OPEN ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function OpenFile(filePath As String) As Object
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload.Add "Argument", filePath
    
    Set OpenFile = ApiPost("/doc/OPEN", payload)
End Function

' doc/OPEN ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function CloseFile() As Object
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload.Add "Argument", CreateObject("Scripting.Dictionary")
    
    Set CloseFile = ApiPost("/doc/CLOSE", payload)
End Function

' doc/SAVE ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function SaveFile() As Object
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload.Add "Argument", CreateObject("Scripting.Dictionary")
    Set SaveFile = ApiPost("/doc/SAVE", payload)
End Function

' doc/SAVEAS ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function SaveAsFile(argument As String) As Object
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload.Add "Argument", argument
    Set SaveAsFile = ApiPost("/doc/SAVEAS", payload)
End Function

' doc/STAGAS ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function SaveCurrentStageAs() As Object
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload.Add "Argument", CreateObject("Scripting.Dictionary")
    Set SaveCurrentStageAs = ApiPost("/doc/STAGAS", payload)
End Function

' doc/IMPORT ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function ImportFile(argument As String) As Object
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload.Add "Argument", argument
    Set ImportFile = ApiPost("/doc/IMPORT", payload)
End Function

' doc/IMPORTMXT ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function ImportMxtFile(argument As String) As Object
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload.Add "Argument", argument
    Set ImportMxtFile = ApiPost("/doc/IMPORTMXT", payload)
End Function

' doc/EXPORT ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function ExportFile(argument As String) As Object
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload.Add "Argument", argument
    Set ExportFile = ApiPost("/doc/EXPORT", payload)
End Function

' doc/EXPORTMXT ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function ExportMxtFile(argument As String) As Object
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload.Add "Argument", argument
    Set ExportMxtFile = ApiPost("/doc/EXPORTMXT", payload)
End Function

' doc/ANAL ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function RunAnalysis(Optional analysisType As String = "") As Object
    Dim payload As Object, arg As Object
    Set payload = CreateObject("Scripting.Dictionary")
    Set arg = CreateObject("Scripting.Dictionary")
    arg.Add "TYPE", analysisType
    payload.Add "Argument", arg
    Set RunAnalysis = ApiPost("/doc/ANAL", payload)
End Function




