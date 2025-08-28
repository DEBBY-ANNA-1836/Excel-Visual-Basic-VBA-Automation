VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImg 
   Caption         =   "UserForm1"
   ClientHeight    =   11920
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   19380
   OleObjectBlob   =   "frmImg.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public locationToHighlight As String
Private Const DEFAULT_HIGHLIGHT_COLOR As Long = vbYellow

' UserForm Activate event: highlight location on form load
Private Sub UserForm_Activate()
    HighlightLabels locationToHighlight
End Sub

' Highlight a specific control with color (default yellow)
Private Sub HighlightControl(ctrlName As String, Optional color As Variant)
    On Error Resume Next
    If IsMissing(color) Then color = DEFAULT_HIGHLIGHT_COLOR
    Dim ctrl As Control
    Set ctrl = Me.Controls(ctrlName)
    If Not ctrl Is Nothing Then
        ctrl.BackColor = color
        ctrl.ForeColor = RGB(0, 0, 0)
    End If
End Sub

' Highlight the rack, row, and bin labels based on location string
Private Sub HighlightLabels(location As String)
    On Error GoTo ErrorHandler
    
    ' Clear previous highlights first
    ResetRackLabels
    
    ' Validate location format
    If location = "" Or InStr(location, "_") = 0 Or InStr(location, ".") = 0 Then
        Debug.Print "Invalid location format: " & location
        Exit Sub
    End If
    
    Dim rackRow As String
    Dim binPart As String
    Dim rack As String
    Dim rowNum As String
    Dim binNum As String
    
    ' Split into rack-row and bin parts
    rackRow = Split(location, "_")(0)  ' e.g. "R12.1"
    binPart = Split(location, "_")(1)  ' e.g. "B251.1"
    
    ' Extract rack and row number
    If InStr(rackRow, ".") > 0 Then
        rack = Split(rackRow, ".")(0)  ' "R12"
        rowNum = Split(rackRow, ".")(1) ' "1"
    Else
        rack = rackRow
        rowNum = "0"
    End If
    
    ' Extract bin number (remove decimal)
    If InStr(binPart, ".") > 0 Then
        binNum = Split(binPart, ".")(0) ' "B251"
    Else
        binNum = binPart
    End If
    
    Debug.Print "Parsed location: " & location
    Debug.Print "Rack: " & rack
    Debug.Print "RowNum: " & rowNum
    Debug.Print "Bin: " & binNum
    
    ' Highlight rack
    If ControlExists(rack) Then
        Me.Controls(rack).BackColor = RGB(255, 255, 0) ' Yellow
    Else
        Debug.Print rack & " label not found!"
    End If
    
    ' Highlight row label (rack + rowNum)
    Dim rowLabel As String
    rowLabel = rack & rowNum
    
    If ControlExists(rowLabel) Then
        Me.Controls(rowLabel).BackColor = RGB(255, 255, 0) ' Yellow
    Else
        Debug.Print rowLabel & " label not found!"
    End If
    
    ' Highlight bin
    If ControlExists(binNum) Then
        Me.Controls(binNum).BackColor = RGB(255, 255, 0) ' Yellow
    Else
        Debug.Print binNum & " label not found!"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in HighlightLabels: " & Err.Description
End Sub

' Check if control exists on the form
Private Function ControlExists(ctrlName As String) As Boolean
    On Error Resume Next
    Dim testCtrl As Control
    Set testCtrl = Me.Controls(ctrlName)
    ControlExists = (Err.Number = 0)
    On Error GoTo 0
End Function

' Highlight a range of bins, highlight target bin in red, others yellow
Private Sub HighlightBinRange(startBin As Integer, endBin As Integer, targetBin As String)
    Dim i As Integer
    For i = startBin To endBin
        Dim labelName As String
        labelName = "B" & i
        If labelName = targetBin Then
            HighlightControl labelName, RGB(255, 0, 0) ' Red
        Else
            HighlightControl labelName, RGB(255, 255, 0) ' Yellow
        End If
    Next i
End Sub

' Reset all rack labels to default colors
Public Sub ResetRackLabels()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "Label" Then
            ctrl.BackColor = &H8000000F ' Default system background
            ctrl.ForeColor = &H80000012 ' Default system foreground
        End If
    Next ctrl
End Sub

' Submit button clicked: validate, update sheets, and reset form
Private Sub cmdSubmit_Click()
    If Not frmTest.ValidateInputs Then Exit Sub
    
    Dim wsOut As Worksheet, wsMat As Worksheet
    Set wsOut = ThisWorkbook.Sheets("Outbound List")
    Set wsMat = ThisWorkbook.Sheets("Material List")
    
    Dim matDesc As String: matDesc = LCase(Trim(frmTest.cboMaterialDescription.Value))
    Dim foundRow As Long: foundRow = 0
    Dim lastRow As Long: lastRow = wsMat.Cells(wsMat.Rows.Count, "B").End(xlUp).Row
    Dim i As Long
    
    For i = 2 To lastRow
        If LCase(Trim(wsMat.Cells(i, 2).Value)) = matDesc Then
            foundRow = i
            Exit For
        End If
    Next i
    
    If foundRow = 0 Then
        MsgBox "Material not found!", vbCritical
        Exit Sub
    End If
    
    Dim qtyTaken As Double: qtyTaken = val(frmTest.txtQtyTaken.Value)
    wsMat.Cells(foundRow, 6).Value = wsMat.Cells(foundRow, 6).Value - qtyTaken
    
    Dim outRow As Long: outRow = wsOut.Cells(wsOut.Rows.Count, "B").End(xlUp).Row + 1
    
    With wsOut
        .Cells(outRow, 2).Value = frmTest.cboMaterialDescription.Value
        .Cells(outRow, 3).Value = frmTest.cboLine.Value
        .Cells(outRow, 4).Value = frmTest.txtStation.Value
        .Cells(outRow, 5).Value = frmTest.txtRowNo.Value
        .Cells(outRow, 6).Value = frmTest.txtLocation.Value
        .Cells(outRow, 7).Value = frmTest.txtQtyTaken.Value
        .Cells(outRow, 8).Value = frmTest.txtEmpName.Value
        .Cells(outRow, 9).Value = frmTest.txtEmpID.Value
        .Cells(outRow, 10).Value = frmTest.txtCost.Value
        .Cells(outRow, 11).Value = frmTest.txtDate.Value
        .Cells(outRow, 12).Value = frmTest.txtTime.Value
    End With
    
    MsgBox "Transaction recorded!", vbInformation
    frmTest.ClearForm
    Unload Me
End Sub

' Cancel button closes the form
Private Sub cmdCancel_Click()
    Unload Me
End Sub
