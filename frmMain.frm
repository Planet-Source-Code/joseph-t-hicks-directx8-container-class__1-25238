VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_DblClick()
    'End the program
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Make sure that we can see the form before we actually do anythign
    If Me.Visible = False Then Exit Sub
    
    'Declare a variable that will NOT go out of scope
    Static sngX As Single
    
    'Rotate the camera on the Y axis of space (+Y is UP, remember?)
    '(so we're rotating AROUND the pyramid)
    If Button = 1 Then jDX.RotateCameraY (X - sngX) / 600   'I chose 600 because it's similar to the speed
                                                            'of the mouse movement... if you make it a smaller
                                                            'number, the pyramid will rotate faster
    'Store the current x position for next time. :)
    sngX = X
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Stop rendering.  Even if we aren't rendering anything, it won't hurt anything to stop
    jDX.EndRender
    Unload Me
End Sub


