Attribute VB_Name = "DragForm"
Option Explicit

Declare Sub ReleaseCapture Lib "USER32" ()
Public Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Public Sub FormDrag(Form1 As Form)  'ex.FormDrag Me
    ReleaseCapture
    Call SendMessage(Form1.hWnd, &HA1, 2, 0&)
End Sub

