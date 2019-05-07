VERSION 5.00
Begin VB.Form FrmImprimeEtiquetaInventario 
   Caption         =   "Imprime Etiqueta Inventarios"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImprimeEtiquetaUnaCol 
      Caption         =   "Imprime Etiqueta Invetario 1 Columna"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txtNroEtiquetas 
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimeEtiquetas 
      Caption         =   "Imprime Etiqueta Invetario 3 Columnas"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "FrmImprimeEtiquetaInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimeEtiquetas_Click()
 Dim I As Integer
 I = 1

If txtNroEtiquetas.Text = "" Then
  txtNroEtiquetas.Text = 1
End If
 
Do While I <= Trim(txtNroEtiquetas.Text)
    Call Imprime_ZEBRA_grande
    I = I + 1
Loop

End Sub

Private Sub cmdImprimeEtiquetaUnaCol_Click()

 Dim I As Integer
 I = 1

If txtNroEtiquetas.Text = "" Then
  txtNroEtiquetas.Text = 1
End If
 
Do While I <= Trim(txtNroEtiquetas.Text)
    Call Imprime_ZEBRA_UnaColumna
    I = I + 1
Loop

End Sub

Private Sub Form_Load()
txtNroEtiquetas.Text = 1
End Sub

Private Sub txtNroEtiquetas_KeyPress(KeyAscii As Integer)
Call SoloNumeros(txtNroEtiquetas, KeyAscii, False)
End Sub
Private Function Imprime_ZEBRA_UnaColumna()
On Error GoTo errx
Dim sSQL  As String, SBARRA As String, sempresa As String, sEmpresa1 As String
Dim mRs As ADODB.Recordset
Dim oPrint As clsPrintFile

sempresa = "TEXTILES JOC SRL"
Printer.Print " "
Printer.Print "^XA"
Printer.Print "^PRC"
Printer.Print "^LH0,0^FS"
Printer.Print "^LL1261"
Printer.Print "^MD0"
Printer.Print "^MNY"

'Printer.Print "^FO10,20^A0N,50,40^CI13^FR^FD" & "Importado Por: "; RTrim(sEmpresa) & "^FS"
'Printer.Print "^FO10,45^A0N,40,40^CI13^FR^FD" & "Rif:J-294891390" & "^FS"
'Printer.Print "^FO10,80^A0N,40,25^CI13^FR^FD" & "Jersey Viscoza Full Ly. 30/1" & "^FS"
'Printer.Print "^FO10,120^A0N,40,25^CI13^FR^FD" & "95% Viscoza 5% Spandex" & "^FS"
'Printer.Print "^FO10,160^A0N,40,25^CI13^FR^FD" & "Partida: "; RTrim(partida) & "^FS"
'Printer.Print "^FO10,205^A0N,40,25^CI13^FR^FD" & "Color: "; RTrim(Color) & "^FS"
'Printer.Print "^FO10,245^A0N,40,25^CI13^FR^FD" & "Peso: "; Format(Str(KIlos), "##.00") & "^FS"
'Printer.Print "^FO250,270^A0N,40,40^CI13^FR^FD" & "HECHO EN PERU " & "^FS"


Printer.Print "^FO20,80^A0N,50,100^CI13^FR^FD" & RTrim(sempresa) & "^FS"
Printer.Print "^FO240,150^A0N,40,40^CI13^FR^FD" & "INVENTARIO" & "^FS"
Printer.Print "^FO240,220^A0N,40,40^CI13^FR^FD" & "DICIEMBRE 2014" & "^FS"



Printer.Print "^PQ1,0, 0, n"
Printer.Print "^XZ"
Printer.Print "^FX End of job"
Printer.Print "^XA"
Printer.Print "^IDR:ID*.*"
Printer.Print "^XZ"
Printer.EndDoc


Exit Function
errx:
    Close #1
    errores Err.numer
End Function

Private Function Imprime_ZEBRA_grande()
On Error GoTo errx
Dim sSQL  As String, SBARRA As String, sempresa As String, sEmpresa1 As String
Dim mRs As ADODB.Recordset
Dim oPrint As clsPrintFile

sempresa = "TEXTILES"
sEmpresa1 = "JOC SRL"
Printer.Print " "
Printer.Print "^XA"
Printer.Print "^PRC"
Printer.Print "^LH0,0^FS"
Printer.Print "^LL1261"
Printer.Print "^MD0"
Printer.Print "^MNY"

Printer.Print "^FO25,25^A0N,12,50^CI13^FR^FD" & "INVENTARIO" & "^FS"
Printer.Print "^FO70,50^A0N,12,50^CI13^FR^FD" & "JULIO" & "^FS"
Printer.Print "^FO70,75^A0N,12,50^CI13^FR^FD" & "2014" & "^FS"
Printer.Print "^FO25,100^A0N,12,50^CI13^FR^FD" & RTrim(sempresa) & "^FS"
Printer.Print "^FO30,125^A0N,12,50^CI13^FR^FD" & RTrim(sEmpresa1) & "^FS"

Printer.Print "^FO290,25^A0N,12,50^CI13^FR^FD" & "INVENTARIO" & "^FS"
Printer.Print "^FO350,50^A0N,12,50^CI13^FR^FD" & "JULIO" & "^FS"
Printer.Print "^FO350,75^A0N,12,50^CI13^FR^FD" & "2014" & "^FS"
Printer.Print "^FO305,100^A0N,12,50^CI13^FR^FD" & RTrim(sempresa) & "^FS"
Printer.Print "^FO310,125^A0N,12,50^CI13^FR^FD" & RTrim(sEmpresa1) & "^FS"


Printer.Print "^FO565,25^A0N,12,50^CI13^FR^FD" & "INVENTARIO" & "^FS"
Printer.Print "^FO630,50^A0N,12,50^CI13^FR^FD" & "JULIO" & "^FS"
Printer.Print "^FO630,75^A0N,12,50^CI13^FR^FD" & "2014" & "^FS"
Printer.Print "^FO585,100^A0N,12,50^CI13^FR^FD" & RTrim(sempresa) & "^FS"
Printer.Print "^FO590,125^A0N,12,50^CI13^FR^FD" & RTrim(sEmpresa1) & "^FS"




Printer.Print "^PQ1,0, 0, n"
Printer.Print "^XZ"
Printer.Print "^FX End of job"
Printer.Print "^XA"
Printer.Print "^IDR:ID*.*"
Printer.Print "^XZ"
Printer.EndDoc


Exit Function
errx:
    Close #1
    errores Err.numer
End Function

