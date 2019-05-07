VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmVer_Cobros 
   Caption         =   "Ver Cobros"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   5280
      TabIndex        =   0
      Top             =   5640
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   900
      Custom          =   "0~0~SALIR~Verdadero~Verdadero~&Salir~0~0~1~~0~Falso~Falso~&Salir~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5655
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   9975
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      ImageCount      =   1
      ImagePicture1   =   "FrmVer_Cobros.frx":0000
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "FrmVer_Cobros.frx":0352
      Column(2)       =   "FrmVer_Cobros.frx":041A
      FormatStylesCount=   8
      FormatStyle(1)  =   "FrmVer_Cobros.frx":04BE
      FormatStyle(2)  =   "FrmVer_Cobros.frx":05F6
      FormatStyle(3)  =   "FrmVer_Cobros.frx":06A6
      FormatStyle(4)  =   "FrmVer_Cobros.frx":075A
      FormatStyle(5)  =   "FrmVer_Cobros.frx":0832
      FormatStyle(6)  =   "FrmVer_Cobros.frx":08EA
      FormatStyle(7)  =   "FrmVer_Cobros.frx":09CA
      FormatStyle(8)  =   "FrmVer_Cobros.frx":0E82
      ImageCount      =   1
      ImagePicture(1) =   "FrmVer_Cobros.frx":12CE
      PrinterProperties=   "FrmVer_Cobros.frx":1620
   End
End
Attribute VB_Name = "FrmVer_Cobros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strSQL As String, SNum_Corre As String

Public Function Buscar() As Boolean
On Error GoTo errores
    Dim vBookmark As Variant
    Dim colTemp As JSColumn
    
Set gridex1.ADORecordset = CargarRecordSetDesconectado(strSQL, cCONNECT)

    gridex1.ContinuousScroll = True


    gridex1.FrozenColumns = 2

    gridex1.GroupFooterStyle = jgexTotalsGroupFooter
    Set colTemp = gridex1.Columns("importe")
    colTemp.AggregateFunction = jgexSum
    colTemp.TotalRowPrefix = ""
    
Exit Function
errores:
    errores err.Number
End Function

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case Is = "SALIR"
  Unload Me
End Select
End Sub



