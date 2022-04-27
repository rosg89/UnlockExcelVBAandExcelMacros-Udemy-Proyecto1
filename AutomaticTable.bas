Attribute VB_Name = "AutomaticTable"
Option Explicit

Sub auto_table_contents()

    Dim startCell As Range 'for input box to select range
    Dim endCell As Range 'for message box as info
    Dim sh As Worksheet
    Dim shName As String
    Dim MsgConfirm As VBA.VbMsgBoxResult 'for message box to confirm
    
    'en caso de que el usuario presione cancelar
    'va a haber un error y ponemos resume next
    On Error Resume Next
    
    'Le pide al usuario que seleccione una celda
    Set startCell = Excel.Application.InputBox("Where do you want to insert the table of contents?" _
    & vbNewLine & "Please, select a cell:", "Insert a Table of Contents", , , , , , 8)
    
    'si el error es de objeto, ponemos el error handling
    If Err.Number = 424 Then Exit Sub
    On Error GoTo Handle
    
    'se asegura que solo se seleccione una celda
    Set startCell = startCell.Cells(1, 1)
    
    'evitar overwrite
    Set endCell = startCell.Offset(Worksheets.Count - 2, 1)
    MsgConfirm = VBA.MsgBox("The values in cells:" & vbNewLine & startCell.Address & " to " & endCell.Address _
    & " could be overwritten." & vbNewLine & "Would you like to continue?", vbOKCancel + vbDefaultButton2, "Confirmation required")
    If MsgConfirm = vbCancel Then Exit Sub
    
    'loop a traves de las sheets
    'copia los titulos de todas las sheets en la activa
    For Each sh In Worksheets
        shName = sh.Name
        startCell = shName
        
    If ActiveSheet.Name <> shName Then
            If sh.Visible = xlSheetVisible Then
                ActiveSheet.Hyperlinks.Add Anchor:=startCell, Address:="", SubAddress:= _
                "'" & shName & "'!A1", TextToDisplay:=shName
                startCell.Offset(0, 1).Value = sh.Range("A1").Value
                Set startCell = startCell.Offset(1, 0)
            End If 'sheet is visible
        End If 'sheet is not activesheet
    Next sh
    Exit Sub
        
 
Handle:
MsgBox "Unfortunately, an error has ocurred!"

End Sub
