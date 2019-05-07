Attribute VB_Name = "ModDesbloquea"

Option Explicit


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Declaraciones
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' funciones para instalar y quitar el Hook al teclado
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, _
                             ByVal lpfn As Long, _
                             ByVal hmod As Long, _
                             ByVal dwThreadId As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    pDest As Any, _
    pSource As Any, _
    ByVal cb As Long)

Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, _
   ByVal nCode As Long, _
   ByVal wParam As Long, _
   ByVal lParam As Long) As Long


' Función para retornar el Hwnd del Admin a partir del caption
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

'Declaración del Api SendMessage ( cierra el Administrador de tareas )
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Public Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, _
    y, ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long


'Declaraciones API para eliminar el titlebar en tiempo de ejecución
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
'() '                         (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
'() '                         (ByVal hwnd As Long, ByVal nIndex As Long, _
'                          ByVal dwNewLong As Long) As Long

'Constante para usar con GetWindowLong y SetWindowLong
Private Const GWL_STYLE = (-16)
'Constante Tipo de estilo, en este caso para quitar la barra
Private Const WS_CAPTION = &HC00000

'Constantes para SetWindowPos
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOZORDER = &H4


'Always on Top constants
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE


'Constante para sendMessage
Private Const SC_CLOSE = &HF060&
Private Const WM_SYSCOMMAND = &H112


Private Type KBDLLHOOKSTRUCT
  vkCode As Long
  scanCode As Long
  flags As Long
  time As Long
  dwExtraInfo As Long
End Type
Private Const HC_ACTION = 0
Private Const WH_MOUSE_LL As Long = 14

Private Const WH_KEYBOARD_LL = 13&
Public KeyboardHandle As Long
Public IdMouse As Long
Private Const CAPTION_TASKBAR As String = "Administrador de tareas de windows"


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Fin declaraciones
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''





' funciones
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' si IsHooked retorna False, deshabilita Todo
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsHooked(ByRef Hookstruct As KBDLLHOOKSTRUCT) As Boolean

  ' Hwnd de la la barra de tareas
  Dim Hwnd_TaskBar As Long

  ' busca el Hwnd del Task Manager
  Hwnd_TaskBar = FindWindow(vbNullString, CAPTION_TASKBAR)

  ' comprueba si el Admin de tareas está abierto
  If Hwnd_TaskBar <> 0 Then
      ' Lo cierra
      Call SendMessage(Hwnd_TaskBar, WM_SYSCOMMAND, SC_CLOSE, ByVal 0&)
  End If

  ' comprueba la tecla de retroceso
  If Hookstruct.vkCode = vbKeyBack Then
     IsHooked = False ' indica que se habilita esta tecla
     Exit Function
  End If

  ' comprueba la tecla Tab
  If Hookstruct.vkCode = vbKeyTab Then
     IsHooked = False
     Exit Function
  End If

  ' comprueba la tecla delete
  If Hookstruct.vkCode = vbKeyDelete Then
     IsHooked = False
     Exit Function
  End If

  ' comprueba la tecla Enter
  If Hookstruct.vkCode = vbKeyReturn Then
     IsHooked = False
     Exit Function
  End If

  '... otras

  ' Comprueba las teclas para los números y desde la a-z
  If (Hookstruct.vkCode < 48) Or Hookstruct.vkCode > 90 Then
     IsHooked = True
  End If

  Exit Function
End Function


Public Function KeyboardCallback(ByVal Code As Long, _
                                 ByVal wParam As Long, _
                                 ByVal lParam As Long) As Long

  Static Hookstruct As KBDLLHOOKSTRUCT

  If (Code = HC_ACTION) Then
    Call CopyMemory(Hookstruct, ByVal lParam, Len(Hookstruct))
    If (IsHooked(Hookstruct)) Then
      KeyboardCallback = 1
      Exit Function
    End If
  End If

  ' deshabilita el mouse y las teclas
  KeyboardCallback = CallNextHookEx(KeyboardHandle, Code, wParam, lParam)

End Function


'función que desactiva el Mouse
'''''''''''''''''''''''''''''''
Public Function WinProcMouse(ByVal nCode As Long, _
                             ByVal wParam As Long, _
                             ByVal lParam As Long) As Long


    WinProcMouse = -1


End Function

' recibe el Hwnd del formulario
''''''''''''''''''''''''''''''''''''''''''''
Public Sub IniciarHook(HwndForm As Long)

  ' hook para el teclado, pasándole el nombre de la función "KeyboardCallback"
  'KeyboardHandle = SetWindowsHookEx(WH_KEYBOARD_LL, _
                                    AddressOf KeyboardCallback, _
                                    App.hInstance, 0&)

 ' hook para el mouse ( lo deshabilita ) le pasa función WinProcMouse
 'IdMouse = SetWindowsHookEx(WH_MOUSE_LL, _
 '                           AddressOf WinProcMouse, _
 '                           App.hInstance, 0)

  'Le saca la barra de título al formulario
 ' Call Quitar_Barra_Titulo(HwndForm, False)
' Pone la ventana Always OnT op
  'Call SetWindowPos(HwndForm, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
End Sub

' función que elimina el Titlebar en tiempo de ejecución
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Quitar_Barra_Titulo(ByVal hwnd As Long, _
                                ByVal Quitar As Boolean)


Dim El_Estilo As Long

    El_Estilo = GetWindowLong(hwnd, GWL_STYLE)
    If Quitar Then
        El_Estilo = El_Estilo Or WS_CAPTION
    Else
        El_Estilo = El_Estilo And Not WS_CAPTION
    End If

    SetWindowLong hwnd, GWL_STYLE, El_Estilo

    SetWindowPos hwnd, 0, 0, 0, 0, 0, _
                      SWP_FRAMECHANGED Or _
                      SWP_NOMOVE Or _
                      SWP_NOSIZE Or _
                      SWP_NOZORDER
End Sub

' Función que usa UnhookWindowsHookEx, para saber si el Hook se ha inicializado
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Hooked()
  Hooked = KeyboardHandle <> 0
End Function

Public Sub FinalizarHook(HwndForm As Long)
  If (Hooked) Then
    ' termina el Hook
    Call UnhookWindowsHookEx(KeyboardHandle)
    Call UnhookWindowsHookEx(IdMouse)
    ' reestablece el estilo anterior de la ventana
     Quitar_Barra_Titulo HwndForm, True
    ' Quita la ventana Always OnT op
     SetWindowPos HwndForm, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
  End If
End Sub

