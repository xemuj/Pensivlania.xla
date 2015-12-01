Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Public Sub PonerMenu()

Dim Pensilvania As Object

Set Pensilvania = CommandBars("Tools").Controls.Add(Type:=msoControlButton)
With Pensilvania
    .BeginGroup = True
    .Caption = "Pensilvania"
    .Tag = "Pensilvania"
    .OnAction = "iniciar_pen"
End With


Set Pensilvania = Nothing

End Sub

 Public Sub QuitarMenu()
   Dim Pensilvania As Object
 
   Set Pensilvania = CommandBars.FindControl(Type:=msoControlButton, Tag:="Pensilvania")
   If Not (Pensilvania Is Nothing) Then
     Pensilvania.Delete
   End If
 
 End Sub
