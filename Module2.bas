Attribute VB_Name = "Module2"
Sub CheckSheetNames()
    Dim ws As Worksheet
    Dim hasMaterialList As Boolean
    Dim hasInboundList As Boolean
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "Material List" Then hasMaterialList = True
        If ws.Name = "Inbound List" Then hasInboundList = True
    Next ws
    
    If hasMaterialList Then
        MsgBox "? 'Material List' found"
    Else
        MsgBox "? 'Material List' sheet is MISSING", vbCritical
    End If
    
    If hasInboundList Then
        MsgBox "? 'Inbound List' found"
    Else
        MsgBox "? 'Inbound List' sheet is MISSING", vbCritical
    End If
End Sub

