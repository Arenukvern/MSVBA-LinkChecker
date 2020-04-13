Attribute VB_Name = "LinkChecker"
Option Explicit

Private Enum enmURLstatus
  [_first] = 1
    eURLstate_Correct = 1
    eURLstate_Wrong
  [_last]
End Enum
Private Enum enmURLSelectMethod
  [_first] = 1
    eURLSelectMethod_Original = 1
    eURLSelectMethod_ReplaceSizes
  [_last]
End Enum


Private URLStatus As enmURLstatus
Private URLSelectMethod As enmURLSelectMethod

Public Sub CheckURLs()
  
  ShowProgressBar
  
  'Variables for worksheets
  Dim shtSource As Worksheet
  Set shtSource = ThisWorkbook.Sheets("Source")
  
  Dim lastrow As Long
  Dim firstrow As Integer
  Dim UsedColumn As Long
  firstrow = 4
  UsedColumn = 34
  'Searching for last row
  lastrow = shtSource.Range("C1").CurrentRegion.Rows.Count
  'Select method of URL correction
  URLSelectMethod = enmURLSelectMethod.eURLSelectMethod_Original

  Dim rngURLs As Range
  Dim arrOldURLs As Variant

  Set rngURLs = shtSource.Range(Cells(firstrow, UsedColumn), Cells(lastrow, UsedColumn))
  arrOldURLs = rngURLs.Value2 'load an array
  Dim i As Long
  Dim newURL As String
  For i = LBound(arrOldURLs) To UBound(arrOldURLs)
    UpdateProgressBar i, arrOldURLs ' upadting progress bar
    newURL = arrOldURLs(i, 1)
    newURL = correctURLs(newURL, i, firstrow, UsedColumn, shtSource)
    arrOldURLs(i, 1) = newURL
  Next i
  
  rngURLs.Value2 = arrOldURLs
  
  CloseProgressBar

  MsgBox prompt:="Done!"

End Sub

Private Function correctURLs( _
  ByVal oldURLs As String, _
  ByVal pos As Long, _
  ByVal firstrow As Long, _
  ByVal UsedColumn As Long, _
  ByRef shtSource As Worksheet) As String
  Const separator As String = "|"

  'First we check is there is separator
  
  Select Case InStr(1, oldURLs, separator)
    Case Is > 0 'Array case
      
      'we need to split string into array
      Dim arrURLs As Variant
      arrURLs = Split(oldURLs, separator, , vbTextCompare)
      Dim URLTemp As String
      Dim newURLs As String
      Dim i As Long
      For i = LBound(arrURLs) To UBound(arrURLs)
        URLTemp = vbNullString
        URLTemp = SwitchMethodOfURLSelecting(arrURLs(i))
        arrURLs(i) = URLTemp
      Next i
      newURLs = Join(arrURLs, separator)
      CloserForURLCorrection pos + firstrow - 1, UsedColumn, shtSource
    Case Else ' String case
      
      newURLs = SwitchMethodOfURLSelecting(oldURLs)
      
      CloserForURLCorrection pos + firstrow - 1, UsedColumn, shtSource
  End Select
  
  
  correctURLs = newURLs
  
End Function

Private Function SwitchMethodOfURLSelecting( _
  ByVal strUrl As String) As String
  
  Select Case URLSelectMethod
    Case enmURLSelectMethod.eURLSelectMethod_Original
      SwitchMethodOfURLSelecting = GetCorrectURL(strUrl)
    Case enmURLSelectMethod.eURLSelectMethod_ReplaceSizes
      SwitchMethodOfURLSelecting = correctSize(strUrl)
  End Select
  
End Function


Private Function correctSize( _
  ByVal strUrl As String) As String
  
  Dim arrSizes As Variant
  Dim SizeForReplace As String
  arrSizes = Array("/2048/", "/500/")
  SizeForReplace = "/280/"
  Dim newURL As String
  Dim i As Long
  Dim tempURL As String
  For i = LBound(arrSizes) To UBound(arrSizes)
    URLStatus = enmURLstatus.eURLstate_Wrong
    newURL = Replace(strUrl, SizeForReplace, arrSizes(i), , , vbTextCompare)
    correctSize = GetCorrectURL(newURL, strUrl)
    'Check is correctSize filled
    Select Case URLStatus
      Case enmURLstatus.eURLstate_Correct
        Exit Function
    End Select
  Next i
  correctSize = strUrl
End Function

Private Function GetCorrectURL( _
  ByVal newURL As String, _
  Optional ByVal oldURL As String) As String
  
  getURLStatus newURL 'URL validation
  
  Select Case URLStatus
    Case enmURLstatus.eURLstate_Correct
      GetCorrectURL = newURL
    Case enmURLstatus.eURLstate_Wrong
      Select Case LenB(oldURL) > 0
        Case True
          GetCorrectURL = oldURL
        Case False
          GetCorrectURL = newURL
      End Select
      
  End Select
  
End Function

Private Sub getURLStatus( _
  ByVal strUrl As String)
  'Checking cases
  'First check - is it image?
  Select Case URLExtensionIMG(strUrl)
    Case True
      URLStatus = enmURLstatus.eURLstate_Correct
    Case False
      URLStatus = enmURLstatus.eURLstate_Wrong
      Exit Sub
  End Select
  'Second check - server response
  Select Case URLResponse(strUrl)
    Case "200"
      URLStatus = enmURLstatus.eURLstate_Correct
    Case Else
      URLStatus = enmURLstatus.eURLstate_Wrong
      Exit Sub
  End Select
  
End Sub

Private Sub CloserForURLCorrection( _
  ByVal row As Long, _
  ByVal column As Long, _
  ByRef shtSource As Worksheet)
  Dim rngURLCell As Range
  Set rngURLCell = shtSource.Range( _
    Cells(row, column), Cells(row, column))
  Select Case URLStatus
    Case enmURLstatus.eURLstate_Wrong
      Select Case LenB(rngURLCell.Value2) > 0
        Case True
          'Mark as red and exit function
          rngURLCell.Interior.Color = RGB(221, 110, 135)
        Case False
          rngURLCell.Interior.Color = xlNone
      End Select
    Case enmURLstatus.eURLstate_Correct
      rngURLCell.Interior.Color = xlNone
  End Select
End Sub

Private Function URLExtensionIMG(ByVal strUrl As String) As Boolean
  
  URLExtensionIMG = False
  
  Dim objRequest As Object
  Dim extensions As Variant
  extensions = Array(".gif", ".png", ".jpg", ".jpeg", ".tiff")
  Dim i As Integer
  Dim pos As Integer
  For i = LBound(extensions) To UBound(extensions)
    pos = InStr(1, strUrl, extensions(i), vbTextCompare)
    Select Case pos
      Case Is > 0
        URLExtensionIMG = True
        Exit Function
    End Select
  Next i
  
End Function


Private Function URLResponse(ByVal strUrl As String) As String

  On Error GoTo ErrHandler
    
  Dim ErrorCode As String
  Dim objRequest As Object
  Set objRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
  With objRequest
    .Open "GET", strUrl
    .Send
    ErrorCode = .status
  End With
    
  GoTo ExitHandler

ErrHandler:
  ErrorCode = Err.Number

ExitHandler:
  Set objRequest = Nothing
  URLResponse = ErrorCode
  
End Function

