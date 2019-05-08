VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmCapturaMovimientoTejeduria 
   Caption         =   "Captura Movimiento de Tejeduria"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   15420
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   0
      TabIndex        =   3
      Top             =   7560
      Width           =   15435
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "IMPRIMIR ROLLOS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdVerRollos 
         Caption         =   "VER ROLLOS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "CANCELAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   11400
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdRecepcionar 
         Caption         =   "RECEPCIONAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   13440
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame FraBuscar 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Argumentos de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15435
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "BUSCAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   13440
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox Cbo_Almacen 
         Height          =   315
         Left            =   1170
         TabIndex        =   1
         Top             =   240
         Width           =   4320
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ALMACEN:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   6510
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   11483
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      FormatStylesCount=   8
      FormatStyle(1)  =   "FrmCapturaMovimientoTejeduria.frx":0000
      FormatStyle(2)  =   "FrmCapturaMovimientoTejeduria.frx":0138
      FormatStyle(3)  =   "FrmCapturaMovimientoTejeduria.frx":01E8
      FormatStyle(4)  =   "FrmCapturaMovimientoTejeduria.frx":029C
      FormatStyle(5)  =   "FrmCapturaMovimientoTejeduria.frx":0374
      FormatStyle(6)  =   "FrmCapturaMovimientoTejeduria.frx":042C
      FormatStyle(7)  =   "FrmCapturaMovimientoTejeduria.frx":050C
      FormatStyle(8)  =   "FrmCapturaMovimientoTejeduria.frx":052C
      ImageCount      =   0
      PrinterProperties=   "FrmCapturaMovimientoTejeduria.frx":0610
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   15360
      Top             =   7800
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmCapturaMovimientoTejeduria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrSQL As String
Public ordtra_tinto As String

Private Sub cmdBuscar_Click()
 Call buscarMovimientos
End Sub
Private Sub buscarMovimientos()
On Error GoTo fin

StrSQL = " lg_muestra_movimientos_tejeduria '" & Left(Cbo_Almacen, 2) & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)
Call Configurar_Grid

Exit Sub
fin:
MsgBox "Inconvientes para mostrar los movimientos ", vbCritical + vbInformation, "Mensaje"

End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub CmdImprimir_Click()
Call ReporteDetalleRollo
End Sub

Private Sub ReporteDetalleRollo()
On Error GoTo hand
    Dim oo As Object, sRuta As String
    Dim rs As New Recordset
    StrSQL = "EXEC TJ_SM_MUESTRA_MOV_TELA_CRUDA_ROLLOS_REPORTE '" & Trim(GridEX1.Value(GridEX1.Columns("cod_almacen").Index)) & "', '" & Trim(GridEX1.Value(GridEX1.Columns("num_movstk").Index)) & "',''"
    
    Set rs = CargarRecordSetDesconectado(StrSQL, cConnect)
    
    If rs.RecordCount <= 0 Then
      MsgBox "Movimiento no tiene Rollos, consultar a tejeduria", vbInformation + vbOKOnly, "Mensaje"
      Exit Sub
    End If
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\rptDetalleRollos.xlt"
    oo.Visible = True
    
    oo.Run "Reporte", rs
Exit Sub
hand:
ErrorHandler err, Me.Caption

End Sub

Private Sub cmdRecepcionar_Click()
Call capturaMovimiento

End Sub

Private Sub CmdVerRollos_Click()

Load FrmVerRollosTejeduria
FrmVerRollosTejeduria.scod_almacen = GridEX1.Value(GridEX1.Columns("cod_almacen").Index)
FrmVerRollosTejeduria.snum_movstk = GridEX1.Value(GridEX1.Columns("num_movstk").Index)
FrmVerRollosTejeduria.muestrarollos
FrmVerRollosTejeduria.Show 1

End Sub

Private Sub Form_Load()
Call FillAlmacen
End Sub
Private Sub FillAlmacen()

Dim rstAux As ADODB.Recordset
Dim StrSQL As String
    
StrSQL = "LG_MUESTRA_ALMACENES_TEJEDURIA"
         
Set rstAux = CargarRecordSetDesconectado(StrSQL, cConnect)
Cbo_Almacen.Clear
With rstAux
    If .RecordCount > 0 Then .MoveFirst
    Do Until .EOF
        Cbo_Almacen.AddItem !Cod_almacen & " " & !nom_almacen
        .MoveNext
    Loop
    .Close
End With
If Cbo_Almacen.ListCount > 0 Then Cbo_Almacen.ListIndex = 0
Set rstAux = Nothing
    
End Sub

Private Sub capturaMovimiento()
On Error GoTo fin

StrSQL = "TI_CAPTURA_TELA_CRUDA_TEJEDURIA '" & GridEX1.Value(GridEX1.Columns("COD_ALMACEN_REL").Index) & "','" & GridEX1.Value(GridEX1.Columns("cod_almacen").Index) & "','" & GridEX1.Value(GridEX1.Columns("num_movstk").Index) & "' "
Call ExecuteSQL(cConnect, StrSQL)

MsgBox "Se realizo con exito la recepcion de la tela cruda de tejeduria, Guia 007-" + GridEX1.Value(GridEX1.Columns("num_movstk").Index), vbInformation + vbOKOnly, "Mensaje"
Call buscarMovimientos

Exit Sub
fin:
MsgBox "Inconvenientes para realizar la recepcion: " + err.Description, vbInformation + vbOKOnly, "Mensaje"

End Sub

Public Sub Configurar_Grid()
    Dim C As Integer
    With GridEX1
    
        For C = 1 To .Columns.Count
            With .Columns(C)
                .Caption = UCase(.Caption)
                .HeaderAlignment = jgexAlignCenter
                .TextAlignment = jgexAlignCenter
                .Visible = False
            End With
        Next C


        
        With .Columns("cod_almacen")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "COD"
            
        End With

        With .Columns("num_movstk")
             .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Visible = True
            .Caption = "num_movstk"
        End With
        
        With .Columns("fec_movstk")
        .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "fecha"
            .Visible = True
        End With
        With .Columns("des_proveedor")
        .Visible = True
            .Width = 2000
            .TextAlignment = jgexAlignLeft
            .Caption = "Proveedor"
            .Visible = True
        End With
        
         'COD_TELA
        With .Columns("des_tipmov")
        .Visible = True
            .Width = 2000
            .TextAlignment = jgexAlignLeft
            .Caption = "Tipo Mov"
            .Visible = True
        End With
        
        With .Columns("des_tipmov")
        .Visible = True
            .Width = 2000
            .TextAlignment = jgexAlignLeft
            .Caption = "Tipo Mov"
            .Visible = True
        End With
        
        With .Columns("oc")
        .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "OF"
            .Visible = True
        End With

        With .Columns("ot")
        .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "Partida"
            .Visible = True
        End With

        With .Columns("observaciones")
        .Visible = True
            .Width = 4000
            .TextAlignment = jgexAlignLeft
            .Caption = "OBS."
            .Visible = True
        End With

        With .Columns("cod_usuario")
        .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "Usuario"
            .Visible = True
        End With
        
    End With
    Call setcolorcolumnas
End Sub
Private Sub setcolorcolumnas()
    GridEX1.Columns("ot").CellStyle = "partida"
    
End Sub

