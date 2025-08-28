VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTest 
   Caption         =   "UserForm1"
   ClientHeight    =   12540
   ClientLeft      =   96
   ClientTop       =   336
   ClientWidth     =   18972
   OleObjectBlob   =   "frmTest.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub UserForm_Initialize()
    txtDate.Value = Format(Date, "dd-mm-yyyy")
    txtTime.Value = Format(Time, "hh:mm:ss AM/PM")

    With cboLine
        .Clear
        .AddItem "FRL"
        .AddItem "APM"
        .AddItem "PMI"
        .AddItem "TFW"
        .AddItem "DG6 L1&L2"
        .AddItem "DG6 L3"
        .AddItem "FSM"
        .AddItem "EMM"
        .AddItem "DSS1"
        .AddItem "DSS2"
        .AddItem "RKLE1"
        .AddItem "RKLE2"
        .AddItem "QMM"
        .AddItem "TEF"
        .AddItem "R.F.R"
    End With

    LoadMaterialDescriptions
    ClearForm
End Sub

Private Sub LoadMaterialDescriptions()
    Dim wsMatList As Worksheet
    Set wsMatList = ThisWorkbook.Sheets("Material List")

    With cboMaterialDescription
        .Clear
        Dim lastRow As Long: lastRow = wsMatList.Cells(wsMatList.Rows.Count, "B").End(xlUp).Row
        Dim i As Long
        For i = 2 To lastRow
            .AddItem wsMatList.Cells(i, 2).Value
        Next i
    End With
End Sub

Public Sub ClearForm()
    txtStation.Value = ""
    txtLocation.Value = ""
    txtCost.Value = ""
    txtQtyAvailable.Value = ""
    txtQtyTaken.Value = ""
    txtEmpName.Value = ""
    txtEmpID.Value = ""
    txtRowNo.Value = ""
    ResetRackLabels
End Sub

Private Sub ResetRackLabels()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "Label" Then
            ctrl.BackColor = &H8000000F
            ctrl.ForeColor = &H80000012
        End If
    Next ctrl
End Sub

Public Sub HighlightRackLocation(location As String)
    If location = "" Or InStr(location, ".") = 0 Or InStr(location, "_") = 0 Then Exit Sub

    ResetRackLabels

    Dim parsed() As String
    parsed = Split(location, ".")
    If UBound(parsed) < 1 Then Exit Sub

    Dim rack As String: rack = parsed(0)
    Dim rowBin() As String: rowBin = Split(parsed(1), "_")
    If UBound(rowBin) < 1 Then Exit Sub

    Dim rowNum As String: rowNum = rowBin(0)
    Dim rowLabel As String: rowLabel = rack & rowNum

    Dim ctrl As Control
    On Error Resume Next
    Set ctrl = Me.Controls(rack): If Not ctrl Is Nothing Then ctrl.BackColor = vbRed
    Set ctrl = Me.Controls(rowLabel): If Not ctrl Is Nothing Then ctrl.BackColor = vbYellow
End Sub

Private Sub cboMaterialDescription_Change()
    If Trim(cboMaterialDescription.Value) = "" Then
        ClearForm
        Exit Sub
    End If

    Dim wsMatList As Worksheet: Set wsMatList = ThisWorkbook.Sheets("Material List")
    Dim matDesc As String: matDesc = LCase(Trim(cboMaterialDescription.Value))
    Dim foundRow As Long: foundRow = 0

    Dim lastRow As Long: lastRow = wsMatList.Cells(wsMatList.Rows.Count, "B").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If LCase(Trim(wsMatList.Cells(i, 2).Value)) = matDesc Then
            foundRow = i
            Exit For
        End If
    Next i

    If foundRow > 0 Then
        txtRowNo.Value = wsMatList.Cells(foundRow, 4).Value
        txtLocation.Value = wsMatList.Cells(foundRow, 5).Value
        txtCost.Value = wsMatList.Cells(foundRow, 7).Value
        txtQtyAvailable.Value = wsMatList.Cells(foundRow, 6).Value

        Dim lineUsed As String: lineUsed = Trim(wsMatList.Cells(foundRow, 3).Value)
        Dim idx As Long
        For idx = 0 To cboLine.ListCount - 1
            If LCase(cboLine.List(idx)) = LCase(lineUsed) Then
                cboLine.ListIndex = idx
                Exit For
            End If
        Next idx
    Else
        MsgBox "Material not found!", vbExclamation
        ClearForm
    End If
End Sub

Private Sub cboLine_Change()
    If Trim(txtLocation.Value) <> "" Then
        HighlightRackLocation txtLocation.Value
    End If
End Sub
Private Sub cmdLocate_Click()
    If Trim(txtLocation.Value) = "" Then
        MsgBox "Please select a material with a valid location first.", vbExclamation
        Exit Sub
    End If

    ' Show location value in a message box for debugging
    MsgBox "Location value in frmTest: " & txtLocation.Value, vbInformation

    Load frmImg
    frmImg.locationToHighlight = txtLocation.Value

    ' Show what was passed to frmImg
    MsgBox "Passed to frmImg: " & frmImg.locationToHighlight, vbInformation

    frmImg.Show
End Sub


Public Function ValidateInputs() As Boolean
    ValidateInputs = True

    If Trim(cboMaterialDescription.Value) = "" Then
        MsgBox "Select a Material Description", vbExclamation
        cboMaterialDescription.SetFocus: ValidateInputs = False: Exit Function
    End If

    If Trim(cboLine.Value) = "" Then
        MsgBox "Select a Line Used", vbExclamation
        cboLine.SetFocus: ValidateInputs = False: Exit Function
    End If

    If Trim(txtStation.Value) = "" Then
        MsgBox "Enter Station", vbExclamation
        txtStation.SetFocus: ValidateInputs = False: Exit Function
    End If

    If Trim(txtLocation.Value) = "" Then
        MsgBox "Location is required", vbExclamation
        txtLocation.SetFocus: ValidateInputs = False: Exit Function
    End If

    If Not IsNumeric(txtQtyTaken.Value) Or val(txtQtyTaken.Value) <= 0 Then
        MsgBox "Enter valid Quantity Taken", vbExclamation
        txtQtyTaken.SetFocus: ValidateInputs = False: Exit Function
    End If

    If val(txtQtyTaken.Value) > val(txtQtyAvailable.Value) Then
        MsgBox "Quantity taken exceeds available stock", vbCritical
        txtQtyTaken.SetFocus: ValidateInputs = False: Exit Function
    End If

    If Not IsNumeric(txtCost.Value) Or val(txtCost.Value) < 0 Then
        MsgBox "Enter valid Cost", vbExclamation
        txtCost.SetFocus: ValidateInputs = False: Exit Function
    End If
End Function

Private Sub txtQtyTaken_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0: Beep
    End If
End Sub

Private Sub txtCost_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
        KeyAscii = 0: Beep
    End If
End Sub

