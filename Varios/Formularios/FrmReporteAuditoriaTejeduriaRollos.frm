VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmReporteAuditoriaTejeduriaRollos 
   Caption         =   "Reporte Tejeduria Rollos - Calidad 3"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   4260
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.OptionButton OptPorTejedor 
         Caption         =   "Agrupado por Tejedor"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   1440
         Width           =   2055
      End
      Begin VB.OptionButton OptPorDia 
         Caption         =   "Agrupado por Dia - Rollo"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   1080
         Value           =   -1  'True
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPInicio 
         Height          =   255
         Left            =   2040
         TabIndex        =   1
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Format          =   73859073
         CurrentDate     =   38416
      End
      Begin MSComCtl2.DTPicker DTPFin 
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Format          =   73859073
         CurrentDate     =   38416
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fin"
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   675
         Width           =   705
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   600
      TabIndex        =   7
      Top             =   1920
      Width           =   2505
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmReporteAuditoriaTejeduriaRollos.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmReporteAuditoriaTejeduriaRollos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
    If OptPorDia Then
        Call Reporte
    Else
        Call Reporte2
    End If
Case "SALIR"
    Unload Me
End Select
End Sub


Sub Reporte()
Dim rs As New ADODB.Recordset
On Error GoTo ErrorImpresion
Dim oo As Object


strSQL = "cc_reporte_auditoria_tejeduria_rollos_calidad3 '" & DTPInicio.Value & "','" & DTPFin.Value & "'"

Set rs = Nothing
rs.CursorLocation = adUseClient
rs.Open strSQL, cConnect

    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\RptRollosCalidad3.XLT"
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "reporte", rs.DataSource, "DEL: " & Format(DTPInicio.Value, "dd/mm/yyyy") & " AL: " & Format(DTPFin.Value, "dd/mm/yyyy")
    Set oo = Nothing
        
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub

Sub Reporte2()
Dim oo As Object

On Error GoTo ErrorImpresion

strSQL = "cc_muestra_evaluacion_rollos_3raCalidad_por_tejedor '" & DTPInicio.Value & "','" & DTPFin.Value & "'"

Set oo = CreateObject("excel.application")
oo.Workbooks.Open vRuta & "\RptRollos3ra_x_Tejedor.XLT"
oo.Visible = True
oo.DisplayAlerts = False
oo.Run "reporte", "Del " & Format(DTPInicio.Value, "dd/mm/yyyy") & " al " & Format(DTPFin.Value, "dd/mm/yyyy"), strSQL, cConnect
Set oo = Nothing
    
Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub

