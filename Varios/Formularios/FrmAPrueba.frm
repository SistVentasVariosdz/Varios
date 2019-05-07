VERSION 5.00
Begin VB.Form FrmAPrueba 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Muestra Session"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Abrir Excel"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Pruebas Bixlon Etiquetera"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
End
Attribute VB_Name = "FrmAPrueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public oPrint As LibraryVB.clsPrintFile
Public oParent As Object
Public prnPrinter As Object
Dim iLin As Integer


Private Sub Command1_Click()
Call IMPRIMIR_REPORTEFORMATOANTIGUOTXT
End Sub
Private Sub Command2_Click()
'imprime_tk
Call imprimebixolon270
End Sub
Private Sub imprimebixolon270()
Dim Posx As Integer
Dim Posy As Integer

For Each prnPrinter In Printers

If prnPrinter.DeviceName = "BIXOLON SRP-270" Then
    Set Printer = prnPrinter
    'Printer.ScaleWidth = 760
    'Printer.ScaleHeight = 2970
    'Printer.ScaleMode = 6
    'Set up the control font.
    'Print in Windows font
    
    Printer.FontName = "FontB2x2[Ext.]"
    Printer.FontSize = 15
    Printer.FontBold = True
    Printer.Print "FACONTEX S.A.C"
    
    'Posx = Printer.CurrentX
    'Posy = Printer.CurrentY
    'Printer.Print Posx
    'Printer.Print Posy
    'Printer.CurrentY = 12
    Printer.FontName = "Arial" '"Dotum" '"Consolas" ' "FontA11"
    Printer.FontSize = 9
    Printer.Print "      "

    Printer.FontName = "FontA1x1[Ext.]" '"FontA11[Ext.]" '"Dotum" '"Consolas" ' "FontA11"
    Printer.FontSize = 9
    Printer.Print "RUC:20516661195"
    
    Posx = Printer.CurrentX
    Posy = Printer.CurrentY
    'Printer.Print Posx
    'Printer.Print Posy
    
    'Printer.TextHeight ("Wy")
    Printer.FontName = "FontA1x1[Ext.]" ' "FontB1x2[Ext.]" '"Dotum" '"Consolas" '"Courier New"
    Printer.FontSize = 9
    Printer.Print "012345678901234567890123456789"
    
    Posx = Printer.CurrentX
    Posy = Printer.CurrentY
    'Printer.Print Posx
    'Printer.Print Posy

    'Printer.Circle (0, 1000), 300

'
'    Printer.FontSize = 9
'    Printer.FontName = "Arial"
'    Printer.Print "Arial; Test3"
'
'
'    'Print in printer font
'    Printer.FontSize = 9
'    Printer.FontName = "FontA1x1"
'    Printer.Print "FontA1x1Test"
'
'    Printer.Font = "Arial"
'    Printer.FontBold = True
'    Printer.FontSize = 16
'    Printer.FontItalic = False
    
    
    Printer.FontSize = 7
    Printer.FontName = "FontControl"
    Printer.Print "G"
    
    '‘Use special-function character to cut the paper
    '‘P: Partial cut
    'g: Partial cut without paper feeding
    Printer.EndDoc
    
    Exit For
End If

Next

End Sub

Private Sub CortaBixolon270()

For Each prnPrinter In Printers
    If prnPrinter.DeviceName = "BIXOLON SRP-270" Then
        Set Printer = prnPrinter
        
        Printer.FontSize = 7
        Printer.FontName = "FontControl"
        Printer.Print "G"
        Printer.EndDoc
        Exit For
    End If
Next

End Sub
Public Sub imprime_tk()

'Close #1
Open "USB001" For Output As #1
'Open "c:\GUIA.txt" For Output As #1

Print #1, Chr(27) & "r" & Chr(0) '0 Color Negro, 1 para el Rojo
Print #1, Chr(27) & "!" & Chr(48) 'Agrando el titulo principal
Print #1, "Titulo principal"
Print #1, Chr(27) & "!" & Chr(0) 'Vuelvo a la configuración inicial
Print #1, "Subtitulo"
Print #1, "Direccion"
Print #1, "Localidad"
Print #1, "Telefono"
Print #1, "---------------------------------"
Print #1, "Nombre de Cliente"
Print #1, Chr(27) & "r" & Chr(1) '1 Color Rojo
Print #1, "Texto"
Print #1, Chr(27) & "r" & Chr(0) '0 Color Negro
Print #1, "Texto"
Print #1, "Texto"
Print #1, "---------------------------------"
Print #1, "Fecha: " & Format(Now, "dd/mm/yyyy hh:ss")
Print #1, " Que tenga un buen día !"
Print #1, Chr(27) & "J" 'Avanzo una linea
Print #1, Chr(27) & "m" 'Realizo el corte parcial del papel
Close #1

'oPrint.SendPrint "c:\GUIA.txt"
'Set oPrint = Nothing

End Sub

Sub IMPRIMIR_REPORTEFORMATOANTIGUOTXT()
Dim RsPro As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim strCadena As String

iLin = 0
Set oPrint = New clsPrintFile

Close #1
Open "c:\GUIA.txt" For Output As #1
    
Plin Chr(15)
Plin "PRENDA1"
Plin "PRENDA2"
Plin "PRENDA3"
Plin "PRENDA4"
Plin "PRENDA9"
            
strCadena = "------------------------------"
Plin strCadena
strCadena = Space(3) & "Observacion:"
Plin strCadena

Plin " PRENDAX"
Plin " PRENDAY"

Plin Chr(12)

Close #1
oPrint.SendPrint "c:\GUIA.txt"
Set oPrint = Nothing

End Sub

Private Sub Command3_Click()
    Reporte vRuta, "po-215516", "rpt_fichaTecnica"
End Sub

Private Sub Reporte(sRuta As String, Nombre_Estilo As String, nombre_plantilla As String)
On Error GoTo fin
Dim elarchivo As String
Dim resultado As String
Dim oo As Object

elarchivo = sRuta & "\" & Nombre_Estilo & ".xls"
resultado = Dir$(elarchivo)

If resultado = "" Then
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open sRuta & "\" & nombre_plantilla & ".XLT"
    oo.DisplayAlerts = False
    oo.Visible = False
    oo.run "REPORTE", Nombre_Estilo
    
    oo.ActiveWorkbook.SaveAs (vRuta & "\" & Nombre_Estilo & ".xls")
    oo.Workbooks.Open sRuta & "\" & Nombre_Estilo & ".xls"
    oo.DisplayAlerts = False
    oo.Visible = True
    
Else
   
    Dim objExcel As Object
    Set objExcel = CreateObject("excel.application")
    objExcel.Workbooks.Open sRuta & "\" & Nombre_Estilo & ".xls"
    objExcel.DisplayAlerts = False
    objExcel.Visible = True
    
End If
         
         
Exit Sub
fin:
MsgBox "Problemas para mostrar reporte " + err.Description, vbInformation + vbOKOnly, "Mensaje del sistema"
End Sub

