Attribute VB_Name = "Macro21"
' ******************************************************************************
' C:\Users\paula\AppData\Local\Temp\swx7044\Macro1.swb - macro recorded on 03/02/24 by paula
' ******************************************************************************
Dim swApp As Object

Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

Sub main()

Set swApp = Application.SldWorks

Set Part = swApp.ActiveDoc
Dim myModelView As Object
Set myModelView = Part.ActiveView
myModelView.FrameState = swWindowState_e.swWindowMaximized
End Sub
