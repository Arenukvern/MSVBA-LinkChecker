Attribute VB_Name = "mNewLinkChecker"
Option Explicit

Public Sub LinkChecker()
On Error GoTo ErrorHandler

    Dim rngLinkDataCell As Range
    Dim rngLinkData As Range
    Dim arrLinkData As Variant
    Dim shtActiveSheet As Worksheet
    
    Set shtActiveSheet = ActiveWorkbook.ActiveSheet
    Set rngLinkDataCell = Application.InputBox("Please select first cell of link", Type:=8)
    
    Set rngLinkData = GetRangeLinkData(rngLinkDataCell, shtActiveSheet)
    arrLinkData = rngLinkData.value
    
    Dim clsCheckerLink As cChecker
    Set clsCheckerLink = New cChecker
    
    clsCheckerLink.arrLink = arrLinkData
    
    clsCheckerLink.CheckStatus
    
    rngLinkData.Interior.Color = xlNone
    
    MarksBrokenLinks rngLinkDataCell, clsCheckerLink
    
 
    
  Exit Sub
ErrorHandler:
  MsgBox prompt:="LinkChecker" & Err.Description & " " & Err.Number
End Sub

Private Function GetRangeLinkData(rngData As Range, shtActiveSht) As Range
  On Error GoTo ErrorHandler
  
    Dim FirstRow As Long
    Dim FirstColumn As Long
    Dim LastRow As Long
   
    With rngData
        FirstRow = .row
        FirstColumn = .column
        LastRow = .CurrentRegion.Rows.Count
    End With
    
    With shtActiveSht
        Set GetRangeLinkData = .Range(.Cells(FirstRow, FirstColumn), _
                                 .Cells(LastRow + FirstRow - 1, FirstColumn))
    End With
    
    
  Exit Function
ErrorHandler:
  MsgBox prompt:="GetRangeLinkData" & Err.Description & " " & Err.Number
End Function

Private Function MarksBrokenLinks(rngLink As Range, _
                  cls As cChecker)
  Dim i As Variant
  Dim s As Long
  Dim j As Long
  
    For Each i In cls.dicBrokenLink
        For j = LBound(cls.arrLink) To UBound(cls.arrLink)
            s = j - 1
            If cls.arrLink(j, 1) = cls.dicBrokenLink.item(i) Then
              rngLink.offset(s, 0).Interior.Color = RGB(221, 110, 135)
            End If
        Next j
    Next i
End Function
