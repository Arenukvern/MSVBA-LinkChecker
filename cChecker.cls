VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const strHttp As String = "http://"
Const strHttps As String = "https://"
Public dicBrokenLink As Scripting.Dictionary
Public arrLink As Variant
Public dicUsedLink As Scripting.Dictionary

Private Sub Class_Initialize()
  Set dicUsedLink = New Scripting.Dictionary
  Set dicBrokenLink = New Scripting.Dictionary
End Sub

Public Sub CheckStatus()
On Error GoTo ErrorHandler
  Dim val As Variant
  Dim i As String
  
  Me.AddHttpToLink
  
  For Each val In dicUsedLink
  
    i = dicUsedLink.item(val)
    
    Select Case URLResponse(val)
      Case False
        dicBrokenLink.Add val, i
    End Select
  Next val


Exit Sub
ErrorHandler:
    MsgBox prompt:="AddHttpToLink" & Err.Description & " " & Err.Number
End Sub
Public Sub AddHttpToLink()
On Error GoTo ErrorHandler
  Dim val As Variant
  Dim posHttp As Integer
  Dim posHttps As Integer
  Dim arrJoin As Variant
  Dim i As String
    
    Me.ConvertArrToDic
    
    For Each val In dicUsedLink
      i = dicUsedLink.item(val)
      posHttps = InStr(1, val, strHttps, vbTextCompare)
      posHttp = InStr(1, val, strHttp, vbTextCompare)
      If posHttp > 0 Or posHttps > 0 Then GoTo NextUnit
    
      arrJoin = Array(strHttps, val)
      arrJoin = Join(arrJoin, "")
      If dicUsedLink.Exists(arrJoin) Then GoTo NextUnit
      dicUsedLink.Remove val
      dicUsedLink.Add arrJoin, i
NextUnit:
    Next val
    
Exit Sub
ErrorHandler:
    MsgBox prompt:="AddHttpToLink" & Err.Description & " " & Err.Number
End Sub
Public Sub ConvertArrToDic()
On Error GoTo ErrorHandler
  
  Dim val As Variant
  
  For Each val In arrLink
     If Not (dicUsedLink.Exists(val)) Then
        dicUsedLink.Add val, val
     End If
     
  Next val
  
Exit Sub
ErrorHandler:
    MsgBox prompt:="ConvertArrToDic" & Err.Description & " " & Err.Number
End Sub
Private Function URLResponse(ByVal strUrl As String) As Boolean
  On Error GoTo ErrHandler
    
  Dim ErrorCode As Boolean
  Dim codStatus As String
  Dim objRequest As Object
  Set objRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
  With objRequest
    .Open "GET", strUrl
    .Send
    codStatus = .status
  End With
    
  If codStatus = "200" Then
    ErrorCode = True
    GoTo ExitHandler
  End If
  
ErrHandler:
  ErrorCode = False

ExitHandler:
  Set objRequest = Nothing
  URLResponse = ErrorCode
  
End Function

