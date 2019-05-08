VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmDatosPartidas 
   Caption         =   "Muestra QYC Pendientes |Partidas sin Fecha de Tinto"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "Partidas sin Fecha Proceso Tinto "
      Height          =   195
      Index           =   1
      Left            =   3480
      TabIndex        =   3
      Top             =   600
      Width           =   2775
   End
   Begin VB.OptionButton Option1 
      Caption         =   "QyC Pendientes Despacho"
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   1800
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dtpFecEmiIni 
      Height          =   315
      Left            =   2100
      TabIndex        =   4
      Top             =   1200
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      Format          =   71041025
      CurrentDate     =   37543
   End
   Begin MSComCtl2.DTPicker dtpFecEmiFin 
      Height          =   315
      Left            =   3870
      TabIndex        =   5
      Top             =   1200
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   556
      _Version        =   393216
      Format          =   71041025
      CurrentDate     =   37543
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   240
      Top             =   1800
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Rango de Fechas"
      Height          =   360
      Left            =   720
      TabIndex        =   6
      Top             =   1230
      Width           =   1395
   End
End
Attribute VB_Name = "FrmDatosPartidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public indice As Integer
Public strsql As String

Private Sub cmdImprimir_Click()
Call reporte
End Sub

Private Sub Form_Load()
    dtpFecEmiIni = Format(Date - 30, "dd/mm/YYYY")
    dtpFecEmiFin = Format(Date, "dd/mm/YYYY")
    indice = 1
End Sub

Private Sub Option1_Click(Index As Integer)
    indice = Index + 1
End Sub

Private Sub reporte()
On Error GoTo fin

Set oo = CreateObject("excel.application")

If indice = 1 Then
    oo.Workbooks.Open vRuta & "\Rpt_Qyc_Pendienes_Despacho.XLT"
End If
If indice = 2 Then
    oo.Workbooks.Open vRuta & "\Rpt_Partidas_Sin_Fecha_Tinto.XLT"
End If

    oo.DisplayAlerts = False
    oo.Visible = True
    oo.Run "REPORTE", indice, dtpFecEmiIni, dtpFecEmiFin, cConnect

Exit Sub
fin:
MsgBox "Problemas para mostrar reporte " + err.Description, vbInformation + vbOKOnly, "Mensaje del sistema"
End Sub
