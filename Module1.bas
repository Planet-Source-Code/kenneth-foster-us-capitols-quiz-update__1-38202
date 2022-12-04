Attribute VB_Name = "RoundCorners"
Option Explicit

' Used to set the shape of the form
Public Declare Function SetWindowRgn Lib "USER32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
' Used to create the rounded rectangle region
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Public Type PointAPI
    X As Long
    Y As Long
End Type

Dim Resizable As Integer
'===================================================================
Public Sub MakeWindow(Form1 As Form, IsResizable As Boolean)
DoTrans Form1
End Sub
'=====================================================================
Public Sub DoTrans(Form1 As Form)
   
    Dim FormWidthInPixels As Long
    Dim FormHeightInPixels As Long
    Dim a
    
 FormWidthInPixels = Form1.Width / Screen.TwipsPerPixelX
 FormHeightInPixels = Form1.Height / Screen.TwipsPerPixelY
    
    
a = CreateRoundRectRgn(0, 0, FormWidthInPixels, FormHeightInPixels, 30, 30)
a = SetWindowRgn(Form1.hWnd, a, True)

End Sub
'---------------------------------------------------------------------
Public Sub ResizeForm(Form1 As Form, OldCursorPos As PointAPI, NewCursorPos As PointAPI, ResizeMode As Integer)
On Error Resume Next
    
    
' Declare some variables
    Dim DifferenceX
    Dim DifferenceY
    
' Put the difference between the first cursor pos and the second into variables
    DifferenceX = (NewCursorPos.X - OldCursorPos.X) * Screen.TwipsPerPixelX
    DifferenceY = (NewCursorPos.Y - OldCursorPos.Y) * Screen.TwipsPerPixelY
    
' Determine which resizing mode (above) has been called and resize accordingly
    Select Case ResizeMode
    Case 0
       Form1.Move Form1.Left + DifferenceX, Form1.Top, Form1.Width - DifferenceX, Form1.Height
    Case 1
        Form1.Move Form1.Left, Form1.Top, Form1.Width + DifferenceX, Form1.Height
    Case 2
        Form1.Move Form1.Left, Form1.Top + DifferenceY, Form1.Width, Form1.Height - DifferenceY
    Case 3
        Form1.Move Form1.Left, Form1.Top, Form1.Width, Form1.Height + DifferenceY
    Case 4
        Form1.Move Form1.Left, Form1.Top, Form1.Width + DifferenceX, Form1.Height + DifferenceY
    Case 5
        Form1.Move Form1.Left + DifferenceX, Form1.Top, Form1.Width - DifferenceX, Form1.Height + DifferenceY
    Case 6
        Form1.Move Form1.Left, Form1.Top + DifferenceY, Form1.Width + DifferenceX, Form1.Height - DifferenceY
    Case 7
        Form1.Move Form1.Left + DifferenceX, Form1.Top + DifferenceY, Form1.Width - DifferenceX, Form1.Height - DifferenceY
    End Select
    
' Check to see if the form has been resized below the minimum size
    If Form1.Width < 57 * Screen.TwipsPerPixelX Then Form1.Width = 57 * Screen.TwipsPerPixelX
    If Form1.Height < 90 * Screen.TwipsPerPixelY Then Form1.Height = 90 * Screen.TwipsPerPixelY
    
' After resizing the form, make the form "rounded rectangle" shaped
    MakeWindow Form1, True
End Sub


'===============================================================
