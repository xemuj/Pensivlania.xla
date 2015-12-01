Rem Attribute VBA_ModuleType=VBAFormModule
Option VBASupport 1
Private Sub CommandButton1_Click()
If TextBox1.Text <> "" Then
TOT = Val(TextBox1.Text)
Call PENSIL
Unload Me
 Else
 MsgBox "Debe ingresar el No. de Estaciones"
 TextBox1.SetFocus
 Exit Sub
 End If

End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub CommandButton3_Click()
Call donar
Unload Me
End Sub
