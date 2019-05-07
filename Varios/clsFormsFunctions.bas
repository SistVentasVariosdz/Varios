Attribute VB_Name = "clsFormsFunctions"
Option Explicit

Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    
    'Get/Set WindowLong Constants (only those used)
    Private Const GWL_STYLE = (-16)
    Private Const GWL_EXSTYLE = (-20)
    'SetWindowPos Constants (only those used)
    Private Const SWP_FRAMECHANGED = &H20 'The frame changed: send WM_NCCALCSIZE
    Private Const SWP_NOMOVE = &H2
    Private Const SWP_NOSIZE = &H1
    'Dialog Styles (also present in the GWL_STYLE area)
    Private Const DS_ABSALIGN As Long = &H1
    Private Const DS_SYSMODAL As Long = &H2
    Private Const DS_3DLOOK As Long = &H4
    Private Const DS_FIXEDSYS As Long = &H8
    Private Const DS_NOFAILCREATE As Long = &H10
    Private Const DS_LOCALEDIT As Long = &H20 'Edit items Get Local storage.
    Private Const DS_SETFONT As Long = &H40 'User specified font For Dlg controls
    Private Const DS_MODALFRAME As Long = &H80 'Can be combined With WS_CAPTION
    Private Const DS_NOIDLEMSG As Long = &H100 'WM_ENTERIDLE message will Not be sent
    Private Const DS_SETFOREGROUND As Long = &H200 'not In win3.1
    Private Const DS_CONTROL As Long = &H400
    Private Const DS_CENTER As Long = &H800
    Private Const DS_CENTERMOUSE As Long = &H1000
    Private Const DS_CONTEXTHELP As Long = &H2000
    'Window Styles (GWL_STYLE area)
    Private Const WS_OVERLAPPED As Long = &H0
    Private Const WS_POPUP As Long = &H80000000
    Private Const WS_CHILD As Long = &H40000000
    Private Const WS_MINIMIZE As Long = &H20000000
    Private Const WS_VISIBLE As Long = &H10000000
    Private Const WS_DISABLED As Long = &H8000000
    Private Const WS_CLIPSIBLINGS As Long = &H4000000
    Private Const WS_CLIPCHILDREN As Long = &H2000000
    Private Const WS_MAXIMIZE As Long = &H1000000
    Private Const WS_CAPTION As Long = &HC00000 'WS_BORDER | WS_DLGFRAME
    Private Const WS_BORDER As Long = &H800000
    Private Const WS_DLGFRAME As Long = &H400000
    Private Const WS_VSCROLL As Long = &H200000
    Private Const WS_HSCROLL As Long = &H100000
    Private Const WS_SYSMENU As Long = &H80000
    Private Const WS_THICKFRAME As Long = &H40000
    Private Const WS_GROUP As Long = &H20000
    Private Const WS_TABSTOP As Long = &H10000
    Private Const WS_MINIMIZEBOX As Long = &H20000
    Private Const WS_MAXIMIZEBOX As Long = &H10000
    Private Const WS_TILED As Long = WS_OVERLAPPED
    Private Const WS_ICONIC As Long = WS_MINIMIZE
    Private Const WS_SIZEBOX As Long = WS_THICKFRAME
    'Extended Window Styles (GWL_EXSTYLE are)
    Private Const WS_EX_DLGMODALFRAME As Long = &H1
    Private Const WS_EX_NOPARENTNOTIFY As Long = &H4
    Private Const WS_EX_TOPMOST As Long = &H8
    Private Const WS_EX_ACCEPTFILES As Long = &H10
    Private Const WS_EX_TRANSPARENT As Long = &H20
    Private Const WS_EX_MDICHILD As Long = &H40
    Private Const WS_EX_TOOLWINDOW As Long = &H80
    Private Const WS_EX_WINDOWEDGE As Long = &H100
    Private Const WS_EX_CLIENTEDGE As Long = &H200
    Private Const WS_EX_CONTEXTHELP As Long = &H400
    Private Const WS_EX_RIGHT As Long = &H1000
    Private Const WS_EX_LEFT As Long = &H0
    Private Const WS_EX_RTLREADING As Long = &H2000
    Private Const WS_EX_LTRREADING As Long = &H0
    Private Const WS_EX_LEFTSCROLLBAR As Long = &H4000
    Private Const WS_EX_RIGHTSCROLLBAR As Long = &H0
    Private Const WS_EX_CONTROLPARENT As Long = &H10000
    Private Const WS_EX_STATICEDGE As Long = &H20000
    Private Const WS_EX_APPWINDOW As Long = &H40000
    
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'    Private Const WM_CLOSE = &H10

'Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

'This are global variables to store the messages
Public gstrMessage As String
Public gstrParameters As String


'This sub changes the form border at Runtime
'This is a Stephen Kent function published in PSCode (thank you)

Public Sub ChangeFormBorder(frmForm As Form, _
    ByVal eNewBorder As FormBorderStyleConstants, _
    Optional ByVal bClipControls As Boolean = True, _
    Optional ByVal bControlBox As Boolean = True, _
    Optional ByVal bMaxButton As Boolean = True, _
    Optional ByVal bMinButton As Boolean = True, _
    Optional ByVal bShowInTaskBar As Boolean = True, _
    Optional ByVal bWhatsThisButton As Boolean = False)

    Dim lRet As Long
    Dim lStyleFlags As Long
    Dim lStyleExFlags As Long
    
    'Initialize our flags
    lStyleFlags = 0
    lStyleExFlags = 0
    
    'If we want ClipControls then add that f
    '     lag and change the form property

    If bClipControls Then
        lStyleFlags = lStyleFlags Or WS_CLIPCHILDREN
        frmForm.ClipControls = True
    Else
        frmForm.ClipControls = False
    End If

    
    'If we want the control box then add the
    '     flag (property is read-only)
    If bControlBox Then lStyleFlags = lStyleFlags Or WS_SYSMENU
    
    'If we want the max button then add the
    '     flag (property is read-only)
    If bMaxButton Then lStyleFlags = lStyleFlags Or WS_MAXIMIZEBOX
    
    'If we want the min button then add the
    '     flag (property is read-only)
    If bMinButton Then lStyleFlags = lStyleFlags Or WS_MINIMIZEBOX
    
    'If we want the form to show in taskbar
    '     then add the flag (property is read-only
    '     )
    If bShowInTaskBar Then lStyleExFlags = lStyleExFlags Or WS_EX_APPWINDOW
    
    'If we want the what's this button then
    '     add the flag (property is read-only)
    If bWhatsThisButton Then lStyleExFlags = lStyleExFlags Or WS_EX_CONTEXTHELP
    
    'If the form is an MDI Child form then a
    '     dd the flag (Don't want to screw up the

    '     form)
        If frmForm.MDIChild Then lStyleExFlags = lStyleExFlags Or WS_EX_MDICHILD
        
        'Now we need to set the flags for the bo
        '     rder we are changing to

        Select Case eNewBorder
            Case vbBSNone
            lStyleFlags = lStyleFlags Or (WS_VISIBLE Or WS_CLIPSIBLINGS)
            'No change to extended style flags.
            Case vbFixedSingle
            lStyleFlags = lStyleFlags Or (WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CAPTION)
            lStyleExFlags = lStyleExFlags Or WS_EX_WINDOWEDGE
            Case vbSizable
            lStyleFlags = lStyleFlags Or (WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CAPTION Or WS_THICKFRAME)
            lStyleExFlags = lStyleExFlags Or WS_EX_WINDOWEDGE
            Case vbFixedDialog
            lStyleFlags = lStyleFlags Or (WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CAPTION Or DS_MODALFRAME)
            lStyleExFlags = lStyleExFlags Or (WS_EX_WINDOWEDGE Or WS_EX_DLGMODALFRAME)
            Case vbFixedToolWindow
            lStyleFlags = lStyleFlags Or (WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CAPTION)
            lStyleExFlags = lStyleExFlags Or (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW)
            Case vbSizableToolWindow
            lStyleFlags = lStyleFlags Or (WS_VISIBLE Or WS_CLIPSIBLINGS Or WS_CAPTION Or WS_THICKFRAME)
            lStyleExFlags = lStyleExFlags Or (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW)
        End Select

    'WS_VISIBLE makes sure the form is visib
    '     le
    'WS_CLIPSIBLINGS makes sure that when th
    '     ere are other windows with the same rela
    '     tive family that they do not draw over e
    '     ach other.
    'WS_CAPTION provides the form's caption
    'WS_THICKFRAME makes the form sizable
    'DS_MODALFRAME allows dialog forms to ha
    '     ve 3d effect
    'WS_EX_WINDOWEDGE is for the border arou
    '     nd the form
    'WS_EX_DLGMODALFRAME says the window has
    '     a double border and may or may not have
    '     a caption
    'WS_EX_TOOLWINDOW says we need a shorter
    '     caption and smaller font
    
    'Change our styles
    lRet = SetWindowLong(frmForm.hwnd, GWL_STYLE, lStyleFlags)
    lRet = SetWindowLong(frmForm.hwnd, GWL_EXSTYLE, lStyleExFlags)
    
    'Signal that the frame has changed
    lRet = SetWindowPos(frmForm.hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED)
    
    'Make that we've changed the border in t
    '     he form's property
    frmForm.BorderStyle = eNewBorder
    
End Sub

'The sub to show another VB6 child form from our VB6 project
Public Sub ShowForm(ByRef objForm As Form, ByVal blnMDIChild As Boolean)

    If Not blnMDIChild Then
        SendMessage "OpenForm", objForm.Name
    Else
        objForm.Show vbModal
    End If

End Sub

' We leave a message in the global variables waiting for the .NET container to read it
Public Sub SendMessage(ByVal strMessage As String, ByVal strParameters As String)

    gstrMessage = strMessage
    gstrParameters = strParameters

    Do Until gstrMessage = ""
    
        Sleep 100
        DoEvents

    Loop

End Sub


