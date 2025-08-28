VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInbound 
   Caption         =   "UserForm1"
   ClientHeight    =   12792
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   25212
   OleObjectBlob   =   "frmInbound.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInbound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Module-level declarations
Private wsMaterialList As Worksheet
Private wsInbound As Worksheet

' Initialize the UserForm
Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler

    ' Maximize form size
    With Me
        .StartUpPosition = 0
        .Left = 0
        .Top = 0
        .Width = Application.Width
        .Height = Application.Height
    End With

    ' Set worksheet references
    Set wsMaterialList = ThisWorkbook.Sheets("Material List")
    Set wsInbound = ThisWorkbook.Sheets("Inbound List")

    ' Set date and time
    txtDate.Value = Format(Date, "dd-mm-yyyy")
    txtTime.Value = Format(Time, "hh:mm:ss AM/PM")

    ' Populate line combo box
    With cboLineUsed
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

    ' Populate Material Description combo box
    Dim lastRow As Long, i As Long, val As String
    lastRow = wsMaterialList.Cells(wsMaterialList.Rows.Count, "B").End(xlUp).Row

    cboMaterialDesc.Clear
    For i = 2 To lastRow
        val = Trim(wsMaterialList.Cells(i, "B").Value)
        If val <> "" Then cboMaterialDesc.AddItem val
    Next i

    Exit Sub

ErrHandler:
    MsgBox "Error in UserForm_Initialize: " & Err.Description, vbCritical
End Sub

' Optional - no autofill on material selection
Private Sub cboMaterialDesc_Change()
End Sub

' Submit button
Private Sub cmdSubmitInbound_Click()
    On Error GoTo ErrHandler

    Dim rackLabel As MSForms.Label, labelName As String
    Dim rackPrefix As String, rackRow As String, extraLocation As String, fullRackLocation As String
    Dim nextRow As Long, lastUsedRow As Long
    Dim matLastRow As Long, foundRow As Long, i As Long
    Dim ctrl As Control

    ' Validation
    If Trim(cboMaterialDesc.Value) = "" Then
        MsgBox "Please select or enter a Material Description.", vbExclamation: Exit Sub
    End If
    If Trim(cboLineUsed.Value) = "" Then
        MsgBox "Please select a Line Used.", vbExclamation: Exit Sub
    End If
    If Not IsNumeric(txtQtyAvailable.Value) Or val(txtQtyAvailable.Value) <= 0 Then
        MsgBox "Please enter a valid quantity.", vbExclamation: Exit Sub
    End If

    ' Line mapping
    Select Case cboLineUsed.Value
        Case "FRL": labelName = "R20": rackPrefix = "R20"
        Case "R.F.R": labelName = "RFR": rackPrefix = "RFR"
        Case "PMI": labelName = "R21": rackPrefix = "R21"
        Case "TFW": labelName = "R13": rackPrefix = "R13"
        Case "DG6 L1&L2": labelName = "R14": rackPrefix = "R14"
        Case "DG6 L3", "EMM", "FSM": labelName = "R15": rackPrefix = "R15"
        Case "DSS1": labelName = "R16": rackPrefix = "R16"
        Case "DSS2": labelName = "R17": rackPrefix = "R17"
        Case "DSS3": labelName = "R18": rackPrefix = "R18"
        Case "DSS4": labelName = "R19": rackPrefix = "R19"
        Case "RKLE1": labelName = "R12": rackPrefix = "R12"
        Case "RKLE2": labelName = "R11": rackPrefix = "R11"
        Case Else: labelName = "": rackPrefix = "R00"
    End Select

    ' Reset label colors
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "Label" And Left(ctrl.Name, 1) = "R" Then
            ctrl.BackColor = &H8000000F
            ctrl.ForeColor = &H80000012
        End If
    Next ctrl

    ' Highlight selected rack
    If labelName <> "" Then
        On Error Resume Next
        Set rackLabel = Me.Controls(labelName)
        If Not rackLabel Is Nothing Then
            rackLabel.BackColor = RGB(255, 255, 0)
            rackLabel.ForeColor = RGB(0, 0, 0)
        End If
        On Error GoTo 0
    End If

    ' Prompt for rack row
    rackRow = InputBox("Enter the rack row number (1 to 8):", "Rack Row")
    If rackRow = "" Or Not IsNumeric(rackRow) Or val(rackRow) < 1 Or val(rackRow) > 8 Then
        MsgBox "Invalid row number. Please enter 1 to 8.", vbExclamation: Exit Sub
    End If

    ' Prompt for extra location
    extraLocation = InputBox("Enter additional location (optional):", "Additional Location")
    fullRackLocation = rackPrefix & "." & rackRow & "_" & extraLocation

    ' Find next row in Inbound List
    lastUsedRow = wsInbound.Cells(wsInbound.Rows.Count, "B").End(xlUp).Row
    nextRow = IIf(lastUsedRow < 2, 2, lastUsedRow + 1)

    ' Write to Inbound sheet
    With wsInbound
        .Cells(nextRow, 2).Value = txtPartNum.Value
        .Cells(nextRow, 3).Value = cboMake.Value
        .Cells(nextRow, 4).Value = cboMaterialDesc.Value
        .Cells(nextRow, 5).Value = txtPONumber.Value
        .Cells(nextRow, 6).Value = WorkOn.Value
        .Cells(nextRow, 7).Value = cboLineUsed.Value
        .Cells(nextRow, 8).Value = rackRow
        .Cells(nextRow, 9).Value = fullRackLocation
        .Cells(nextRow, 10).Value = val(txtQtyAvailable.Value)
        .Cells(nextRow, 11).Value = val(txtCost.Value)
        .Cells(nextRow, 12).Value = txtEmpName.Value
        .Cells(nextRow, 13).Value = txtEmpID.Value
        .Cells(nextRow, 14).Value = txtDate.Value
        .Cells(nextRow, 15).Value = txtTime.Value
    End With

    ' Update Material List quantity or add new item
    matLastRow = wsMaterialList.Cells(wsMaterialList.Rows.Count, "B").End(xlUp).Row
    foundRow = 0
    For i = 2 To matLastRow
        If LCase(Trim(wsMaterialList.Cells(i, 2).Value)) = LCase(Trim(cboMaterialDesc.Value)) Then
            foundRow = i: Exit For
        End If
    Next i

    If foundRow > 0 Then
        wsMaterialList.Cells(foundRow, 6).Value = wsMaterialList.Cells(foundRow, 6).Value + val(txtQtyAvailable.Value)
    Else
        With wsMaterialList
            .Cells(matLastRow + 1, 2).Value = cboMaterialDesc.Value
            .Cells(matLastRow + 1, 3).Value = cboLineUsed.Value
            .Cells(matLastRow + 1, 4).Value = rackRow
            .Cells(matLastRow + 1, 5).Value = fullRackLocation
            .Cells(matLastRow + 1, 6).Value = val(txtQtyAvailable.Value)
            .Cells(matLastRow + 1, 7).Value = val(txtCost.Value)
        End With
    End If

    ' Clear form
    cboMaterialDesc.Value = ""
    txtPONumber.Value = ""
    WorkOn.Value = ""
    cboLineUsed.Value = ""
    txtQtyAvailable.Value = ""
    txtCost.Value = ""
    txtEmpName.Value = ""
    txtEmpID.Value = ""

    ' Reset label colors
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "Label" And Left(ctrl.Name, 1) = "R" Then
            ctrl.BackColor = &H8000000F
            ctrl.ForeColor = &H80000012
        End If
    Next ctrl

    ' Close the form
    Unload Me
    Exit Sub

ErrHandler:
    MsgBox "Error submitting data: " & Err.Description, vbCritical
End Sub

' Cancel button
Private Sub cmdCancel_Click()
    Unload Me
End Sub

