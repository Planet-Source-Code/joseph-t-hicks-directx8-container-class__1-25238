VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "jDXEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Declare DirectX objects needed
Dim DX As DirectX8
Dim D3D As Direct3D8
Dim D3DX As D3DX8
Dim D3DDevice As Direct3DDevice8

'Local class variables for use in rendering
Dim bRunning As Boolean
Dim BackColor As Long

'Value of pi (4*atn(1))
Const mPI = 3.14159265358979

'Vertex type
Private Type TLVERTEX
    X As Single
    Y As Single
    z As Single
    rhw As Single
    color As Long
    specular As Long
    tu As Single
    tv As Single
End Type

'Vertex type - only one used
Private Type LITVERTEX
    X As Single
    Y As Single
    z As Single
    color As Long
    specular As Single
    tu As Single
    tv As Single
End Type
'Vertex type definition for DirectX
Const FVF_LVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1)

'Vertex Type
Private Type VERTEX
    X As Single
    Y As Single
    z As Single
    nx As Single
    ny As Single
    nz As Single
    tu As Single
    tv As Single
End Type

'Dynamic array of vertices to make triangles from
Dim Vertices() As LITVERTEX
'A number which points to how many vertices we are holding in the array
Dim VertexArrayIndex As Long

'Matrices
Dim matWorld As D3DMATRIX
Dim matProj As D3DMATRIX
Dim matView As D3DMATRIX
Dim matTemp As D3DMATRIX

'A vector to store the current position of the camera
Dim vtrCameraPos As D3DVECTOR




'x1-3,y1-3,and z1-3 are xyz coordinates for the 3 vertices of the triangle in 3D space
'color1/2/3 is the specific color value of each vertex
Public Sub AddTriangle(x1 As Single, y1 As Single, z1 As Single, x2 As Single, y2 As Single, z2 As Single, x3 As Single, y3 As Single, z3 As Single, color1 As Long, color2 As Long, color3 As Long)
    On Error GoTo error_h
    
    'Create 3 vertices
    Dim v1 As LITVERTEX, v2 As LITVERTEX, v3 As LITVERTEX
    
    'Copy the parameter values to the 3 vertices
    v1 = MakeLitVertex(x1, y1, z1, color1, 0, 0, 0)
    v2 = MakeLitVertex(x2, y2, z2, color2, 0, 0, 0)
    v3 = MakeLitVertex(x3, y3, z3, color3, 0, 0, 0)
    
    'Add all vertices to the main array of vertices
    AddVertex v1
    AddVertex v2
    AddVertex v3
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "jDXEngine.AddTriangle")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

'vert is a specific LITVERTEX to store for when we render
Private Sub AddVertex(vert As LITVERTEX)
    On Error GoTo error_h
    
    'Add another slot to hold an extra vertex
    ReDim Preserve Vertices(VertexArrayIndex)
    
    'Store the new vertex
    Vertices(VertexArrayIndex) = vert
    
    'Increment the counter (VertexArrayIndex should always be ONE more than UBound(Vertices))
    VertexArrayIndex = VertexArrayIndex + 1
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "jDXEngine.AddVertex")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

'r=red value(0-255),g=green,b=blue
Public Function jRGB(r As Integer, g As Integer, b As Integer) As Long
    On Error GoTo error_h
    
    'I've noticed when working with DX8 that the long color code values it uses
    'are exactly the same as VB's RGB() function except that the R and B values
    'are transposed, so I wrote this little helper function to help out with that.
    jRGB = RGB(b, g, r)
    
    Exit Function
error_h:
    Select Case ErrMsg(Err, "jDXEngine.DXRGB(" & r & "," & g & "," & b & ")")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Function
    End Select
End Function

Public Sub EndRender()
    On Error GoTo error_h
    
    'All we need to do to stop rendering is set this little variable here to false. :)
    bRunning = False
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "jDXEngine.EndRender")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

'hWnd = the handle to the window that you want to draw in (i THINK it can be anything with a hWnd property, but don't quote me on that)
'HardwareDevice = true if you want to use a 3d accelorator card... false if not
'vtrCamPos = a vector desribing the initial position of the camera
'lngBackColor = the background color when rendering
Public Function InitWindowed(hWnd As Long, HardwareDevice As Boolean, vtrCamPos As D3DVECTOR, lngBackColor As Long) As Boolean
    On Error GoTo error_h
    
    'Boolean values default to false, but just to be safe...
    InitWindowed = False
    
    'Declare variables to store display information
    Dim d3dpp As D3DPRESENT_PARAMETERS
    Dim DispMode As D3DDISPLAYMODE
    
    'Create the main 3 objects (D3DX8 is NOT one of the main 3... D3DDevice is the 3rd, but we must configure it first)
    Set DX = New DirectX8
    Set D3D = DX.Direct3DCreate
    
    'D3DX8 objects are NOT required for use in DX8, but it contains a LOT of helper functions you may need
    'so it doesn't hurt to declare it for when you need it
    Set D3DX = New D3DX8

    'Here we are getting the current display mode from the default diaplay adapter (your video card)
    'The current display information is stored in the second parameter "DispMode"
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    'D3DPRESENT_PARAMETERS objects show which options are (dis/en)abled and/or how they are configured.
    With d3dpp
        .BackBufferFormat = DispMode.Format     'The BackBuffer is what is actually drawn upon
        .Windowed = 1                           'We want to draw in a window as opposed to full screen
        .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC  'This says to re-draw when the monitor does
        .EnableAutoDepthStencil = 1             'Enable auto-depth perception
        .AutoDepthStencilFormat = D3DFMT_D16    '16bit depth
    End With
    
    'This creates the 3rd main DX8 object
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, IIf(HardwareDevice, D3DDEVTYPE_HAL, D3DDEVTYPE_REF), hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
    
    'Now we can configure the D3DDevice
    With D3DDevice
        .SetVertexShader FVF_LVERTEX                 'This is the vertex definition we're using
        .SetRenderState D3DRS_LIGHTING, 0            'Turn lighting off (off is default, but just to be sure)
                                                     '(We're turning the lights off because the vertices are
                                                     ' already lit)
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE 'Turn off culling
        .SetRenderState D3DRS_ZENABLE, 1             'Enable depth perception
    End With
    
    'I won't go into what a matrix is, but when you Identity one, it's like a 'reset' to all default values
    D3DXMatrixIdentity matWorld
    D3DDevice.SetTransform D3DTS_WORLD, matWorld    'If you were to rotate this matrix first, the entire
                                                    'drawing would be rotated as well
    
    'Store the camera position from what was passed
    vtrCameraPos = vtrCamPos
    
    'Create a matrix that sets where the camera is(eye), what it's looking at(at), and which direction is up(up)
    'Technically, you can make any direction you want to be up, but generally, it's (0,1,0) (Positive Y axis is up)
    D3DXMatrixLookAtLH matView, vtrCameraPos, MakeVector(0, 0, 0), MakeVector(0, 1, 0)
    D3DDevice.SetTransform D3DTS_VIEW, matView      'Set the camera position according to the matrix
    
    'This matrix is for the 'lens' of the camera
    D3DXMatrixPerspectiveFovLH matProj, mPI / 4, 1, 0.1, 500
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj
    
    'Just a long to store the user-chosen background color
    BackColor = lngBackColor
    
    'Everything went well
    InitWindowed = True
    
    'Since everything's configured, we can start to Render() whenever the programmer says so
    bRunning = True
    
    Exit Function
error_h:
    Select Case ErrMsg(Err, "jDXEngine.InitWindowed(" & hWnd & ")")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Function
    End Select
End Function

Private Function MakeLitVertex(X As Single, Y As Single, z As Single, color As Long, specular As Single, tu As Single, tv As Single) As LITVERTEX
    'Copy the values into a LITVERTEX data type
    With MakeLitVertex
        .X = X
        .Y = Y
        .z = z
        .color = color
        .specular = specular
        .tu = tu
        .tv = tv
    End With
End Function


'ErrObj = the actual error
'strProc = a string containing the name of the procedure the error occured in
'(sometimes i send along parameter information as well if applicable...)
Private Function ErrMsg(ErrObj As ErrObject, strProc As String)
    'This is something of my own invention.  I use it for debugging purposes, and you should see a call
    'to it in almost every routine save the simplest.  What it does is LOG the error (because users never
    'bother to remember what the error said, or even it's error number), and present a message box with
    'the error information, and the ROUTINE that the error occurred in! (VERY useful for debugging).  All
    'of this information is put into a .txt file called "error.log" which can be viewed in notepad.
    
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

Public Function MakeVector(X As Single, Y As Single, z As Single) As D3DVECTOR
    'Again, copy the supplied values to a D3DVECTOR structure
    With MakeVector
        .X = X
        .Y = Y
        .z = z
    End With
End Function


Property Get PI() As Single
    'In case the programmer needs to play with PI...
    PI = mPI
End Property


Public Sub Render()
    On Error GoTo error_h
    
    'Keep running this loop until the programmer says not to (bRunning=false will stop the loop)
    Do While bRunning
        'Clear the screen so we can redraw
        D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, BackColor, 1#, 0
        
        'Begin scene
        D3DDevice.BeginScene
            'Rendering begins here:
            '''''''''''''''''''''''''''''''''''''''''''''''''''
            'Here, just for kicks, i'm gonna rotate the camera
            'RotateCameraY 0.05
            
            'Draw our geometry...
            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, (UBound(Vertices) + 1) / 3, Vertices(0), Len(Vertices(0))
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''
            'End of rendering
        'End of scene
        D3DDevice.EndScene
        
        'Present the scene
        D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
        
        'Free up the processor so the computer don't choke! :)
        DoEvents
    Loop
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "jDXEngine.Renter")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub


Private Sub ResetCameraPos()
    On Error GoTo error_h
    
    'This will take the current camera position vector, recreate the matrix for it,
    'and reset the camera's matrix to the new one
    D3DXMatrixLookAtLH matView, vtrCameraPos, MakeVector(0, 0, 0), MakeVector(0, 1, 0)
    D3DDevice.SetTransform D3DTS_VIEW, matView
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "ResetCameraPos")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub

'sngAngle = the angle (in radians) that you want to rotate by
Public Sub RotateCameraY(sngAngle As Single)
    On Error GoTo error_h
    
    'Rotate the camera's current matrix along the Y axis
    D3DXMatrixRotationY matView, sngAngle
    
    'Now use the new matrix to translate the camera's current positioning vector
    D3DXVec3TransformCoord vtrCameraPos, vtrCameraPos, matView
    
    'We 'could' reset the camera position here, but I made a separate sub for that
    ResetCameraPos
    
    Exit Sub
error_h:
    Select Case ErrMsg(Err, "jDXEngine.RotateCameraY(" & sngAngle & ")")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
        Case Else
            Exit Sub
    End Select
End Sub


Private Sub Class_Initialize()
    'Long variables default to 0, but I like to be safe.
    VertexArrayIndex = 0
End Sub


Private Sub Class_Terminate()
    'Theoretically, VB will do all of this for you, BUT again, just to be safe...
    
    On Error Resume Next       'This is so if the class goes out of scope before the InitWindowed function is called
    Set D3DX = Nothing
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set DX = Nothing
End Sub


