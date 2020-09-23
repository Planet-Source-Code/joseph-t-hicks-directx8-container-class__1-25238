Attribute VB_Name = "modMain"
Option Explicit

'The only global variable you need for DirectX8... my very own jDXEngine
'
'This class was written by Joseph Hicks to help those that are having difficulties
'with DirectX8 graphics (god knows that I am!!!).  As of the time of this writing, it is
'very simplistic, but still gives the general idea of what needs to be created, initialized
'when creating a DirectX8 application.  Feel free to use this code as you see fit, but if you
'make any major modifications or feature enhancements, could you PLEASE be so kind as to
'email me a copy of the changes so I can learn more, too? (jhicks@hsadallas.com)
'I will continue to update this class as I learn more about DirectX8, and as others send
'me their contributions (assuming that they actually do... we're on the honor system here
'guys...) and will continue to post updates when they are warranted. :)  If you have any
'questions about this class, please don't hesitate to email me at any of the following email addys:
'jhicks@hsadallas.com (work)
'captainn64@mindspring.com (home)
'captainn64@yahoo.com (mobile)
'
'Thanks for your interest in my work! =þ

Public jDX As jDXEngine


Public Sub main()
    On Error GoTo error_h
    
    'Technically, you 'could' just define it as Public jDX = New jDXEngine in the
    '(Declarations) section, but i prefer late-binding, and that's not really the
    'topic in discussion here. :)
    '(Instantiate the class)
    Set jDX = New jDXEngine
    
    'You have to have a VISIBLE form before you can do any drawing, so let's get to it!
    frmMain.Show
    DoEvents
    
    'The .InitWindowed function will return true/false depending on if the init worked.
    'Parameters will be explained in the class definition
    If Not jDX.InitWindowed(frmMain.hWnd, True, jDX.MakeVector(0, 0, -5), vbBlack) Then
        MsgBox "Initialization failed.  Program will now terminate.", vbCritical
        Unload frmMain
        End
    End If
    
    With jDX
        'The .AddTriangle sub will store 3 points to create a triangle with a given color
        'You don't HAVE to use the .jRGB function to create your colors.  If you want to
        'use the plain old RGB() function, just switch the R and B values around.
        .AddTriangle -1, -1, -1, 0, 1, 0, 1, -1, -1, .jRGB(255, 0, 0), .jRGB(0, 255, 0), .jRGB(0, 0, 255)
        .AddTriangle -1, -1, 1, 0, 1, 0, -1, -1, -1, .jRGB(255, 255, 0), .jRGB(0, 255, 0), .jRGB(255, 0, 0)
        .AddTriangle 1, -1, 1, 0, 1, 0, -1, -1, 1, .jRGB(255, 255, 255), .jRGB(0, 255, 0), .jRGB(255, 255, 0)
        .AddTriangle 1, -1, -1, 0, 1, 0, 1, -1, 1, .jRGB(0, 0, 255), .jRGB(0, 255, 0), .jRGB(255, 255, 255)
        
        'This will start the loop that will continue drawing until the .EndRender method is called
        .Render
    End With
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "Main")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub


Public Function ErrMsg(ErrObj As ErrObject, strProc As String)
    'See comments on this function in the class definition
    
    Dim intFreeFile As Integer
    
    intFreeFile = FreeFile
    
    Open App.Path & "\error.log" For Append As #intFreeFile
        Print #intFreeFile, Date
        Print #intFreeFile, Time
        Print #intFreeFile, " "
        Print #intFreeFile, ErrObj.Number
        Print #intFreeFile, ErrObj.Description
        Print #intFreeFile, "(" & strProc & ")"
    Close intFreeFile
    
    Select Case MsgBox(ErrObj.Number & vbCrLf & ErrObj.Description, vbExclamation + vbAbortRetryIgnore, strProc)
        Case vbRetry
            ErrMsg = vbRetry
        Case vbIgnore
            ErrMsg = vbIgnore
        Case Else
            ErrMsg = vbAbort
    End Select
    
End Function


