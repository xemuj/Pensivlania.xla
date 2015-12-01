Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
 Option Explicit
 
 Private Sub Workbook_AddinInstall()
 MsgBox "El comando para ejecutar Pensilvania " _
 & "en el menú Herramientas de Microsoft Excel."
  MsgBox "Creado Por: " _
 & "Sergio Estuardo Ovalle Pineda."

 End Sub
 
 Private Sub Workbook_BeforeClose(Cancel As Boolean)
   QuitarMenu
  End Sub
 
 Private Sub Workbook_Open()
   PonerMenu
 End Sub

