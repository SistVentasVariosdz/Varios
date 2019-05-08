VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmMuestraDetalleMovPrendas 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Muestra Detalle de Movimientos Prendas"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   14715
   StartUpPosition =   3  'Windows Default
   Begin GridEX20.GridEX GridEX1 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   12938
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      BackColorBkg    =   12648384
      ColumnHeaderHeight=   285
      FormatStylesCount=   7
      FormatStyle(1)  =   "FrmMuestraDetalleMovPrendas.frx":0000
      FormatStyle(2)  =   "FrmMuestraDetalleMovPrendas.frx":0138
      FormatStyle(3)  =   "FrmMuestraDetalleMovPrendas.frx":01E8
      FormatStyle(4)  =   "FrmMuestraDetalleMovPrendas.frx":029C
      FormatStyle(5)  =   "FrmMuestraDetalleMovPrendas.frx":0374
      FormatStyle(6)  =   "FrmMuestraDetalleMovPrendas.frx":042C
      FormatStyle(7)  =   "FrmMuestraDetalleMovPrendas.frx":050C
      ImageCount      =   0
      PrinterProperties=   "FrmMuestraDetalleMovPrendas.frx":052C
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   11880
      TabIndex        =   1
      Top             =   7440
      Width           =   2700
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmMuestraDetalleMovPrendas.frx":0704
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmMuestraDetalleMovPrendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private STRSQL As String
Public scod_almacen As String
Public scod_Estcli As String
Public scod_present As Integer
Public scod_talla  As String

Public Sub muestradatos()
On Error GoTo FIN
    
    STRSQL = " cf_muestra_detalle_movimiento_prendas '" & scod_almacen & "','" & scod_Estcli & "','" & scod_present & "','" & scod_talla & "' "
    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(STRSQL, cConnect)
    CONFIGURA_GRILLA
Exit Sub
FIN:
MsgBox "INCONVENIENTES" + err.Description, vbInformation + vbOKOnly, "IMPORTANTE"

End Sub


Private Sub CONFIGURA_GRILLA()
    On Error GoTo SALTO_ERROR
    Dim C As Integer
    With GridEX1
    
        For C = 1 To .Columns.Count
            .Columns(C).Visible = False
            .Columns(C).HeaderAlignment = jgexAlignCenter
            .Columns(C).TextAlignment = jgexAlignCenter
        Next C

        With .Columns("cod_almacen")
            .Visible = True
            .Width = 500
            .Caption = "Almacen"
            .TextAlignment = jgexAlignLeft
        End With
                        
        With .Columns("num_movstk")
            .Visible = True
            .Width = 800
            .Caption = "Mov"
            .TextAlignment = jgexAlignLeft
        End With
                        
        With .Columns("Fecha")
            .Visible = True
            .Width = 1000
            .Caption = "Fecha"
            .TextAlignment = jgexAlignLeft
        End With
                        
        With .Columns("Documento")
            .Visible = True
            .Width = 1500
            .Caption = "Documento"
            .TextAlignment = jgexAlignLeft
        End With
                        
        With .Columns("DES_TIPMOV")
            .Visible = True
            .Width = 2500
            .Caption = "Transaccion"
            .TextAlignment = jgexAlignLeft
        End With
   
   
        With .Columns("cod_estcli")
            .Visible = True
            .Width = 1000
            .Caption = "Codigo"
            .TextAlignment = jgexAlignLeft
        End With
                        
        With .Columns("DES_ESTCLI")
            .Visible = True
            .Width = 1500
            .Caption = "Estilo"
            .TextAlignment = jgexAlignLeft
        End With
        
        
        With .Columns("des_present")
            .Visible = True
            .Width = 1000
            .Caption = "Color"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("cod_talla")
            .Visible = True
            .Width = 500
            .Caption = "Talla"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("CAN_MOVIMIENTO")
            .Visible = True
            .Width = 1000
            .Caption = "Cantidad"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("tipo_mov")
            .Visible = True
            .Width = 1000
            .Caption = "Tipo"
            .TextAlignment = jgexAlignLeft
        End With
        
        End With
        
    Exit Sub
    
SALTO_ERROR:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub
Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim strCalidad As String
    Select Case ActionName
    Case "IMPRIMIR"
        Reporte
    Case "SALIR"
        Unload Me
    End Select
End Sub

Private Sub Reporte()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String

    Ruta = vRuta & "\RptDetalleMovimientos.xlt"
    Screen.MousePointer = 11
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    STRSQL = "SELECT Ruta_Logo FROM SEGURIDAD..SEG_EMPRESAS WHERE cod_EMPRESA ='" & _
             Trim(vemp1) & "'"
    oo.Run "reporte", GridEX1.ADORecordset
    Set oo = Nothing
    Screen.MousePointer = 0
Exit Sub
hand:
    Screen.MousePointer = 0
    ErrorHandler err, "Reporte"
    Set oo = Nothing
End Sub

