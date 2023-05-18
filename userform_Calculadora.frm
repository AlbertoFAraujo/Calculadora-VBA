VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calculadora 
   Caption         =   "Calculadora"
   ClientHeight    =   5490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4260
   OleObjectBlob   =   "userform_Calculadora.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Calculadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public minhaColecao As New Collection

Private Sub Label1_Click()

Application.Visible = True
Calculadora.Hide

End Sub

Private Sub UserForm_Initialize()

Dim obj As Object
Dim btn As New visiCalc

For Each obj In Calculadora.Controls

    If TypeName(obj) = "Label" Then
    
        Set btn = New visiCalc
        Set btn.label = obj
        Set btn.txt = Me.visorCalculadora
        minhaColecao.Add btn
          
    End If

Next obj
End Sub

Private Sub UserForm_Terminate()
    
    Application.Quit
    
    
End Sub

Private Sub visorCalculadora_Change()

If Me.visorCalculadora.TextLength > 8 Then

    Me.visorCalculadora.Font.Size = 24
    Me.visorCalculadora.WordWrap = True

End If
End Sub
