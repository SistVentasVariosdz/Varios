VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReporteRollosCalidad3_4 
   Caption         =   "Rango de fechas"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Rango de Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   330
      TabIndex        =   0
      Top             =   180
      Width           =   3945
      Begin MSComCtl2.DTPicker DTPInicio 
         Height          =   255
         Left            =   2010
         TabIndex        =   1
         Top             =   435
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Format          =   74383361
         CurrentDate     =   38416
      End
      Begin MSComCtl2.DTPicker DTPFin 
         Height          =   255
         Left            =   2010
         TabIndex        =   2
         Top             =   795
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Format          =   74383361
         CurrentDate     =   38416
      End
      Begin FunctionsButtons.FunctButt FunctButt5 
         Height          =   510
         Left            =   720
         TabIndex        =   5
         Top             =   1320
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmReporteRollosCalidad3_4.frx":0000
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fin"
         Height          =   195
         Left            =   570
         TabIndex        =   4
         Top             =   870
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
         Height          =   195
         Left            =   570
         TabIndex        =   3
         Top             =   555
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmReporteRollosCalidad3_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String

Private Sub FunctButt5_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
        Call Reporte
Case "SALIR"
    Unload Me
End Select
End Sub


Sub Reporte()
Dim oo As Object

On Error GoTo AceptarErr

Screen.MousePointer = 11

Set oo = CreateObject("excel.application")
oo.Workbooks.Open vRuta & "\rptReporteRollosCalidad3_4.xlt"
oo.Visible = True
'oo.Run "Reporte", cCONNECT, "tj_muestra_rollos_calidades_3_4_por_rango_de_fechas '" & Format(DTPInicio.Value, "dd/mm/yyyy") & "','" & Format(DTPFin.Value, "dd/mm/yyyy") & "'", Format(DTPInicio.Value, "dd/mm/yyyy") & "," & Format(DTPFin.Value, "dd/mm/yyyy")

oo.Run "Reporte", cConnect, "tj_muestra_rollos_calidades_3_4_por_rango_de_fechas '" & DTPInicio.Value & "','" & DTPFin.Value & "'", DTPInicio.Value, DTPFin.Value



Screen.MousePointer = 0

oo.Visible = True
Set oo = Nothing

Exit Sub
AceptarErr:
    MsgBox err.Description, vbCritical
    Screen.MousePointer = 0
End Sub


