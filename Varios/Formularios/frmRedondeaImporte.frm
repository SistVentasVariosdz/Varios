VERSION 5.00
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmRedondeaImporte 
   Caption         =   "Redondea Importe"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtImporteNeto 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   290
      Width           =   1695
   End
   Begin VB.TextBox txtImporteIgv 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtImporteTotal 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtImporteGastosFinan 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtImporteDscto 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtImpTotalActual 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdSuma 
      Caption         =   "+"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmdResta 
      Caption         =   "-"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmdRecalcular 
      Caption         =   "R"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox txtImporteOtros 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Salir"
      Height          =   615
      Left            =   4200
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Deshacer"
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar &Redondeo"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   4080
      Width           =   2295
   End
   Begin NumBoxProject.NumBox txtValorRedondeo 
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   285
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      TypeVal         =   2
      Mask            =   "9,999,999,999.99"
      Formato         =   "#,###,###,###.##"
      AllowedMask     =   -1
      MaskLen         =   10
      Aling           =   3
      Text            =   "0.00"
      CanEmpty        =   -1
      ShowError       =   0
      Locked          =   0   'False
      Enabled         =   0   'False
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DecimalNumber   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "Im. Neto :"
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Imp. IGV :"
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Importe Total :"
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Valor Redondeo :"
      Height          =   375
      Left            =   4200
      TabIndex        =   19
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Imp. Gastos Finan.  :"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Imp. Descuento.  :"
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Importe Total Actual :"
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Imp. Otros :"
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   2400
      Width           =   1455
   End
End
Attribute VB_Name = "frmRedondeaImporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrSQL As String
Dim flag As Boolean
Private maxRedondeo As Double
Public porcIGV As Double
Public ValorActualIGV As String
Public ValorActualImporteTotal As String
Public ValorActualImporteNeto As String
Public ValorActualImporteTotalR As String
Public grilla As GridEx




Private Sub Check1_Click()
If Not flag Then
    txtValorRedondeo.Enabled = True
    cmdSuma.Enabled = True
    cmdResta.Enabled = True
    flag = True
Else
    txtValorRedondeo.Enabled = False
    cmdSuma.Enabled = False
    cmdResta.Enabled = False
    flag = False
End If
End Sub

Private Sub cmdRecalcular_Click()
Dim Valor As Double

txtValorRedondeo.Enabled = False
cmdSuma.Enabled = False
cmdResta.Enabled = False
Check1.Value = 0
If Check1.Value = 1 Then Check1_Click

Valor = Round(CDbl(txtImporteTotal) - CDbl(ValorActualImporteTotal), 2)
If Valor < 0 Then Valor = Valor * -1
If maxRedondeo > Valor Then
    If txtImporteDscto <> "0" Then
      If txtImporteIgv <> "0" Then
          Valor = CDbl(txtImporteTotal) + CDbl(txtImporteDscto)
          'txtImporteNeto = Round(valor / 1.19, 2)
          txtImporteNeto = Round(Valor / (1 + porcIGV / 100), 2)
      Else
          txtImporteNeto = Round(CDbl(txtImporteTotal) + CDbl(txtImporteDscto), 2)
          Exit Sub
      End If
    Else
      If txtImporteOtros <> "0" Then
          txtImporteNeto = Round(CDbl(txtImporteTotal) - CDbl(txtImporteOtros), 2)
          Exit Sub
      End If
      If txtImporteGastosFinan <> "0" Then
          txtImporteNeto = Round((CDbl(txtImporteTotal) - CDbl(txtImporteGastosFinan)) / (1 + porcIGV / 100), 2)
          txtImporteIgv = CDbl(txtImporteTotal) - CDbl(txtImporteNeto) - CDbl(txtImporteGastosFinan)
          Exit Sub
      End If
      
      txtImporteNeto = Round(CDbl(txtImporteTotal) / (1 + porcIGV / 100), 2)
      txtImporteIgv = CDbl(txtImporteTotal) - CDbl(txtImporteNeto)
    End If
    
Else
    MsgBox "El valor de Redondeo debe ser menor a : " & maxRedondeo, , "Redondear Importe"
    
End If
End Sub

Private Sub cmdResta_Click()

   
   If maxRedondeo > CDbl(txtValorRedondeo.Text) And (CDbl(txtValorRedondeo.Text)) > 0 Then
      'txtImporteNeto = CDbl(ValorActualImporteNeto) - CDbl(txtValorRedondeo.Text)
      txtImporteNeto = CDbl(txtImporteNeto) - CDbl(txtValorRedondeo.Text)
      txtImporteIgv = IIf(ValorActualIGV = 0, 0, Round(CDbl(txtImporteNeto.Text) * 0.19, 2))
      txtImporteTotal.Text = CDbl(txtImporteNeto.Text) + CDbl(IIf(txtImporteIgv.Text = "", 0, txtImporteIgv.Text)) + CDbl(IIf(txtImporteGastosFinan.Text = "", 0, txtImporteGastosFinan.Text)) - CDbl(IIf(txtImporteDscto.Text = "", 0, txtImporteDscto.Text))
      'ValorActualImporteNeto = txtImporteNeto
   Else
      MsgBox "El valor de Redondeo debe ser menor a : " & maxRedondeo, , "Redondear Importe"
   End If
End Sub

Private Sub cmdSuma_Click()

   If maxRedondeo > CDbl(txtValorRedondeo.Text) And (CDbl(txtValorRedondeo.Text)) > 0 Then
      'txtImporteNeto = CDbl(ValorActualImporteNeto) + CDbl(txtValorRedondeo.Text)
      txtImporteNeto = CDbl(txtImporteNeto) + CDbl(txtValorRedondeo.Text)
      txtImporteIgv = IIf(ValorActualIGV = 0, 0, Round(CDbl(txtImporteNeto.Text) * 0.19, 2))
      txtImporteTotal.Text = CDbl(txtImporteNeto.Text) + CDbl(IIf(txtImporteIgv.Text = "", 0, txtImporteIgv.Text)) + CDbl(IIf(txtImporteGastosFinan.Text = "", 0, txtImporteGastosFinan.Text)) - CDbl(IIf(txtImporteDscto.Text = "", 0, txtImporteDscto.Text))
   Else
       MsgBox "El valor de Redondeo debe ser menor a : " & maxRedondeo, , "Redondear Importe"
   End If
   'ValorActualImporteNeto = txtImporteNeto

End Sub

Private Sub Command1_Click()
If MsgBox("Desea Redondear el Documento " & grilla.Value(grilla.Columns("Cod_TipDoc").Index) & "-" & grilla.Value(grilla.Columns("Nro_Doc").Index), vbYesNo, "CONFIRMAR") = vbYes Then
    GuardarRedondeo
    Unload Me
End If
End Sub

Private Sub GuardarRedondeo()

    Dim Con As New ADODB.Connection
    Dim StrSQL As String
        
    On Error GoTo GuardarRedondeoErr

    Con.ConnectionString = cConnect
    Con.Open
    
        Con.BeginTrans

        
        StrSQL = "exec UP_MAN_VENTAS_REDONDEO '" & _
        grilla.Value(grilla.Columns("Cod_TipDoc").Index) & "','" & _
        grilla.Value(grilla.Columns("Serie").Index) & "','" & _
        grilla.Value(grilla.Columns("Nro_Doc").Index) & "'," & _
        txtImporteTotal & "," & _
        txtImporteNeto & "," & _
        txtImporteIgv & "," & _
        txtImpTotalActual - txtImporteTotal
        
        Con.Execute StrSQL

        Con.CommitTrans
        Dim amensaje As New clsMessages
        amensaje.CODIGO = CodeMsg.kMeSsaGe_INF_DATA_save
        Informa "", amensaje
        
        'GuardarRedondeo = True
        
    Exit Sub
GuardarRedondeoErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler err, "GuardarRedondeo"
End Sub

Public Sub Habilita()

End Sub

Private Sub Command2_Click()
      txtImporteNeto = ValorActualImporteNeto
      txtImporteIgv = ValorActualIGV
      txtImporteTotal.Text = ValorActualImporteTotal
      
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
flag = False
txtValorRedondeo.Enabled = False
StrSQL = "select max_redondeo from cn_control"
maxRedondeo = DevuelveCampo(StrSQL, cConnect)
End Sub

