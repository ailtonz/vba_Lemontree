Option Compare Database

Private Sub Form_Open(Cancel As Integer)
    DoCmd.Maximize
End Sub
Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click


    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    MsgBox Err.Description
    Resume Exit_cmdFechar_Click
    
End Sub
