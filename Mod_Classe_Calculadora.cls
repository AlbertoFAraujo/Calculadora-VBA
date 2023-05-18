VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "visiCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents label As MSForms.label
Public txt As MSForms.TextBox

Private Sub label_Click()

On Error GoTo tratarErro

Select Case Left(label.Name, 6)

Case "btnLim"

    txt.Value = ""
    
Case "btnDel"
    
    txt.Value = Mid(txt.Value, 1, Len(txt.Value) - 1)
    
Case "btnIgu"

    If InStr(1, txt.Value, ",", 1) Then
    
        txt.Value = Replace(txt.Value, ",", ".")
        
        txt.Value = Application.Evaluate(txt.Value)
        
    Else
        
        txt.Value = Application.Evaluate(txt.Value)

    End If

Case "btnOpM"

    txt.Value = txt.Value & "*"

Case "btnNum"

    txt.Value = txt.Value & label.Caption
    
Case "btnOpe"

    txt.Value = txt.Value & label.Caption
    

End Select

tratarErro:

    If Err.Number = 0 Then
        
        Exit Sub
    
    Else
    
        MsgBox "Entrada Inválida!", vbCritical + vbOKOnly, "Entrada incorreta:"
        txt.Value = ""
        Exit Sub
    
    End If
    
End Sub

Private Sub label_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

If label.Name = "btnLimpar" Or label.Name = "btnIgual" Then

    label.BackColor = RGB(128, 0, 0)
    label.Font.Size = 24
Else

    label.BackColor = RGB(47, 47, 47)
    label.Font.Size = 24

End If
End Sub

Private Sub label_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

   
If label.Name = "btnLimpar" Or label.Name = "btnIgual" Then

    label.BackColor = RGB(192, 0, 0)
    label.Font.Size = 26
    
Else

    label.BackColor = RGB(64, 64, 64)
    label.Font.Size = 26

End If
End Sub

