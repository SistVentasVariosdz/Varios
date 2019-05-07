VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form FrmProduccionTermofijado 
   Caption         =   "Reporte de Produccion Termofijado"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   13545
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   13455
      Begin GridEX20.GridEX GridEX1 
         Height          =   5895
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   10398
         Version         =   "2.0"
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "FrmProduccionTermofijado.frx":0000
         Column(2)       =   "FrmProduccionTermofijado.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "FrmProduccionTermofijado.frx":016C
         FormatStyle(2)  =   "FrmProduccionTermofijado.frx":02A4
         FormatStyle(3)  =   "FrmProduccionTermofijado.frx":0354
         FormatStyle(4)  =   "FrmProduccionTermofijado.frx":0408
         FormatStyle(5)  =   "FrmProduccionTermofijado.frx":04E0
         FormatStyle(6)  =   "FrmProduccionTermofijado.frx":0598
         ImageCount      =   0
         PrinterProperties=   "FrmProduccionTermofijado.frx":0678
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rango"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13455
      Begin VB.CommandButton Cmd_Buscar 
         Caption         =   "&Buscar"
         Height          =   495
         Left            =   12120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   75235329
         CurrentDate     =   41317
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   300
         Left            =   4080
         TabIndex        =   3
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   75235329
         CurrentDate     =   41317
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Left            =   3000
         TabIndex        =   4
         Top             =   360
         Width           =   825
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   10920
      TabIndex        =   6
      Top             =   7080
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmProduccionTermofijado.frx":0850
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   2160
      Top             =   7200
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmProduccionTermofijado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Buscar_Click()
CARGA_GRID
End Sub

Sub CARGA_GRID()
On Error GoTo ErrCargaGrid
    strSQL = "EXEC  Ti_Muestra_ProduccionTermofijado '" & Format(DTPicker1, "dd/mm/yyyy") & "','" & Format(DTPicker2, "dd/mm/yyyy") & "'"
    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    CONFIGURARGRID
Exit Sub
ErrCargaGrid:
ErrorHandler Err, "Carga_Grid"
End Sub

Sub CONFIGURARGRID()

    GridEX1.Columns("cod_maquina_tinto").Width = 1000
    GridEX1.Columns("Fecha_Creacion").Width = 1200
    GridEX1.Columns("Fecha_Creacion_termo").Width = 1200
    GridEX1.Columns("Fec_Ult_Programacion").Width = 1200
    GridEX1.Columns("Cod_ordtra").Width = 800
    GridEX1.Columns("Nom_Cliente").Width = 2500
    GridEX1.Columns("Cod_Color").Width = 800
    GridEX1.Columns("Des_color").Width = 2000
    GridEX1.Columns("Des_Tela").Width = 2000
    GridEX1.Columns("kgs_asignados").Width = 1000
    GridEX1.Columns("kgs_termofijado").Width = 1000
    GridEX1.Columns("Guia").Width = 1000
    
    
End Sub

Private Sub Form_Load()
DTPicker1 = Date
DTPicker2 = Date
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "IMPRIMIR"
    Imprimir
Case "SALIR"
    Unload Me
End Select
End Sub


Private Sub Imprimir()
On Error GoTo Fin
Dim strSQL As String
Dim oo As Object, vRutaLogo As Variant
    
    Screen.MousePointer = 11
    strSQL = "SELECT Ruta_Logo From SEGURIDAD..SEG_EMPRESAS " & _
             "WHERE Cod_Empresa = '" & vemp & "'"
    vRutaLogo = DevuelveCampo(strSQL, cConnect)
    
    vRutaLogo = CStr(IIf(IsNull(vRutaLogo), "", vRutaLogo))
          Set oo = CreateObject("excel.application")
          oo.workbooks.Open vRuta & "\Ti_Termofijado.xlt"
          oo.DisplayAlerts = False
          oo.Visible = True
    
    oo.run "REPORTE", GridEX1.ADORecordset, cConnect
    
    Screen.MousePointer = vbNormal
    'oo.Workbooks.Close
    Set oo = Nothing
Exit Sub
Fin:
    Screen.MousePointer = vbNormal
    errores Err.Number
End Sub





