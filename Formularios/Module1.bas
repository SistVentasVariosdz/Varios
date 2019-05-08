Attribute VB_Name = "Module1"


Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
' Declaracion de API
Declare Function DeleteFile& Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String)
 

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long


Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

 
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
 
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Public Function EjecBatch(Fichero$) As Boolean
    Dim Valor As Long
    Dim Comienzo As STARTUPINFO
    Dim Proceso As PROCESS_INFORMATION
    Comienzo.cb = Len(Comienzo)
    Valor = CreateProcessA(0&, Fichero$, 0&, 0&, 1&, &H20&, 0&, 0&, Comienzo, Proceso)
    Valor = WaitForSingleObject(Proceso.hProcess, -1&)
    If Valor = -1 Then
        EjecBatch = False
        MsgBox "El proceso " & Fichero$ & " no ha podido ser lanzado con éxito." & vbCrLf & "Asegúrese que el programa existe o el path es correcto.", 16, "Ejecución Batch"
    Else
        Valor = CloseHandle(Proceso.hProcess)
        EjecBatch = True
    End If
End Function
 
 
Function letra(Numero)
Dim Texto
Dim Millones
Dim Miles
Dim Cientos
Dim Decimales
Dim cadena
Dim CadMillones
Dim CadMiles
Dim CadCientos
Dim caddecimales
Texto = Round(Numero, 2)
Texto = FormatNumber(Texto, 2)
Texto = Right(Space(14) & Texto, 14)
Millones = Mid(Texto, 1, 3)
Miles = Mid(Texto, 5, 3)
Cientos = Mid(Texto, 9, 3)
Decimales = Mid(Texto, 13, 2)
CadMillones = ConvierteCifra(Millones, False)
CadMiles = ConvierteCifra(Miles, False)
CadCientos = ConvierteCifra(Cientos, True)
caddecimales = ConvierteDecimal(Decimales)
If Trim(CadMillones) > "" Then
If Trim(CadMillones) = "UN" Then
cadena = CadMillones & " MILLON"
Else
cadena = CadMillones & " MILLONES"
End If
End If
If Trim(CadMiles) > "" Then
If Trim(CadMiles) = "UN" Then
CadMiles = ""
cadena = cadena & "" & CadMiles & "MIL"
CadMiles = "UN"
Else
cadena = cadena & " " & CadMiles & " MIL"
End If
End If

If Decimales = "00" Then
If Trim(CadMillones & CadMiles & CadCientos & caddecimales) = "UN" Then
cadena = cadena & "UNO "
Else
If Miles & Cientos = "000000" Then
cadena = cadena & " " & Trim(CadCientos)
Else
cadena = cadena & " " & Trim(CadCientos)
End If
letra = Trim(cadena)
End If
Else
If Trim(CadMillones & CadMiles & CadCientos & caddecimales) = "UN" Then
cadena = cadena & "UNO " & "CON " & Trim(caddecimales)
Else
If Miles & Cientos = "000000" Then
cadena = cadena & " " & Trim(CadCientos) & " CON " & Trim(Decimales)
Else
cadena = cadena & " " & Trim(CadCientos) & " CON " & Trim(Decimales)
End If
letra = Trim(cadena)
End If
End If

End Function

Function ConvierteCifra(Texto, IsCientos As Boolean)
Dim Centena
Dim Decena
Dim Unidad
Dim txtCentena
Dim txtDecena
Dim txtUnidad
Centena = Mid(Texto, 1, 1)
Decena = Mid(Texto, 2, 1)
Unidad = Mid(Texto, 3, 1)
Select Case Centena
Case "1"
txtCentena = "CIEN"
If Decena & Unidad <> "00" Then
txtCentena = "CIENTO"
End If
Case "2"
txtCentena = "DOSCIENTOS"
Case "3"
txtCentena = "TRESCIENTOS"
Case "4"
txtCentena = "CUATROCIENTOS"
Case "5"
txtCentena = "QUINIENTOS"
Case "6"
txtCentena = "SEISCIENTOS"
Case "7"
txtCentena = "SETECIENTOS"
Case "8"
txtCentena = "OCHOCIENTOS"
Case "9"
txtCentena = "NOVECIENTOS"
End Select

Select Case Decena
Case "1"
txtDecena = "DIEZ"
Select Case Unidad
Case "1"
txtDecena = "ONCE"
Case "2"
txtDecena = "DOCE"
Case "3"
txtDecena = "TRECE"
Case "4"
txtDecena = "CATORCE"
Case "5"
txtDecena = "QUINCE"
Case "6"
txtDecena = "DIECISEIS"
Case "7"
txtDecena = "DIECISIETE"
Case "8"
txtDecena = "DIECIOCHO"
Case "9"
txtDecena = "DIECINUEVE"
End Select
Case "2"
txtDecena = "VEINTE"
If Unidad <> "0" Then
txtDecena = "VEINTI"
End If
Case "3"
txtDecena = "TREINTA"
If Unidad <> "0" Then
txtDecena = "TREINTA Y "
End If
Case "4"
txtDecena = "CUARENTA"
If Unidad <> "0" Then
txtDecena = "CUARENTA Y "
End If
Case "5"
txtDecena = "CINCUENTA"
If Unidad <> "0" Then
txtDecena = "CINCUENTA Y "
End If
Case "6"
txtDecena = "SESENTA"

If Unidad <> "0" Then
txtDecena = "SESENTA Y "
End If
Case "7"
txtDecena = "SETENTA"
If Unidad <> "0" Then
txtDecena = "SETENTA Y "
End If
Case "8"
txtDecena = "OCHENTA"
If Unidad <> "0" Then
txtDecena = "OCHENTA Y "
End If
Case "9"
txtDecena = "NOVENTA"
If Unidad <> "0" Then
txtDecena = "NOVENTA Y "
End If
End Select

If Decena <> "1" Then
Select Case Unidad
Case "1"
If IsCientos = False Then
txtUnidad = "UN"
Else
txtUnidad = "UNO"
End If
Case "2"
txtUnidad = "DOS"
Case "3"
txtUnidad = "TRES"
Case "4"
txtUnidad = "CUATRO"
Case "5"
txtUnidad = "CINCO"
Case "6"
txtUnidad = "SEIS"
Case "7"
txtUnidad = "SIETE"
Case "8"
txtUnidad = "OCHO"
Case "9"
txtUnidad = "NUEVE"
End Select
End If
ConvierteCifra = txtCentena & " " & txtDecena & txtUnidad
End Function


Function ConvierteDecimal(Texto)
Dim Decenadecimal
Dim Unidaddecimal
Dim txtDecenadecimal
Dim txtUnidaddecimal
Decenadecimal = Mid(Texto, 1, 1)
Unidaddecimal = Mid(Texto, 2, 1)

Select Case Decenadecimal
Case "1"
txtDecenadecimal = "DIEZ"
Select Case Unidaddecimal
Case "1"
txtDecenadecimal = "ONCE"
Case "2"
txtDecenadecimal = "DOCE"
Case "3"
txtDecenadecimal = "TRECE"
Case "4"
txtDecenadecimal = "CATORCE"
Case "5"
txtDecenadecimal = "QUINCE"
Case "6"
txtDecenadecimal = "DIECISEIS"
Case "7"
txtDecenadecimal = "DIECISIETE"
Case "8"
txtDecenadecimal = "DIECIOCHO"
Case "9"
txtDecenadecimal = "DIECINUEVE"
End Select
Case "2"
txtDecenadecimal = "VEINTE"
If Unidaddecimal <> "0" Then
txtDecenadecimal = "VEINTI"
End If
Case "3"
txtDecenadecimal = "TREINTA"
If Unidaddecimal <> "0" Then
txtDecenadecimal = "TREINTA Y "
End If
Case "4"
txtDecenadecimal = "CUARENTA"
If Unidaddecimal <> "0" Then
txtDecenadecimal = "CUARENTA Y "
End If
Case "5"
txtDecenadecimal = "CINCUENTA"
If Unidaddecimal <> "0" Then
txtDecenadecimal = "CINCUENTA Y "
End If
Case "6"
txtDecenadecimal = "SESENTA"

If Unidaddecimal <> "0" Then
txtDecenadecimal = "SESENTA Y "
End If
Case "7"
txtDecenadecimal = "SETENTA"
If Unidaddecimal <> "0" Then
txtDecenadecimal = "SETENTA Y "
End If
Case "8"
txtDecenadecimal = "OCHENTA"
If Unidaddecimal <> "0" Then
txtDecenadecimal = "OCHENTA Y "
End If
Case "9"
txtDecenadecimal = "NOVENTA"
If Unidaddecimal <> "0" Then
txtDecenadecimal = "NOVENTA Y "
End If
End Select

If Decenadecimal <> "1" Then
Select Case Unidaddecimal
Case "1"
txtUnidaddecimal = "UNO"
Case "2"
txtUnidaddecimal = "DOS"
Case "3"
txtUnidaddecimal = "TRES"
Case "4"
txtUnidaddecimal = "CUATRO"
Case "5"
txtUnidaddecimal = "CINCO"
Case "6"
txtUnidaddecimal = "SEIS"
Case "7"
txtUnidaddecimal = "SIETE"
Case "8"
txtUnidaddecimal = "OCHO"
Case "9"
txtUnidaddecimal = "NUEVE"
End Select
End If
If Decenadecimal = 0 And Unidaddecimal = 0 Then
ConvierteDecimal = ""
Else
ConvierteDecimal = txtDecenadecimal & txtUnidaddecimal
End If
End Function



