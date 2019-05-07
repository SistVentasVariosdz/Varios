VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInspeccionOrdenesCompra 
   Caption         =   "INSPECCION DE ORDENES DE COMPRA"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18165
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   18165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&IMPRIMIR"
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
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8160
      Width           =   1785
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&CANCELAR"
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
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8160
      Width           =   1905
   End
   Begin VB.CommandButton cmdGuardarCambios 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&GUARDAR CAMBIOS"
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
      Left            =   16080
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8160
      Width           =   1785
   End
   Begin GridEX20.GridEX GridEX3 
      Height          =   2055
      Left            =   5520
      TabIndex        =   13
      Top             =   4200
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3625
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ScrollToolTipColumn=   ""
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "FrmInspeccionOrdenesCompra.frx":0000
      FormatStyle(2)  =   "FrmInspeccionOrdenesCompra.frx":0138
      FormatStyle(3)  =   "FrmInspeccionOrdenesCompra.frx":01E8
      FormatStyle(4)  =   "FrmInspeccionOrdenesCompra.frx":029C
      FormatStyle(5)  =   "FrmInspeccionOrdenesCompra.frx":0374
      FormatStyle(6)  =   "FrmInspeccionOrdenesCompra.frx":042C
      FormatStyle(7)  =   "FrmInspeccionOrdenesCompra.frx":050C
      ImageCount      =   0
      PrinterProperties=   "FrmInspeccionOrdenesCompra.frx":052C
   End
   Begin GridEX20.GridEX GridEX2 
      Height          =   2055
      Left            =   2400
      TabIndex        =   14
      Top             =   4200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3625
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ScrollToolTipColumn=   ""
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "FrmInspeccionOrdenesCompra.frx":0704
      FormatStyle(2)  =   "FrmInspeccionOrdenesCompra.frx":083C
      FormatStyle(3)  =   "FrmInspeccionOrdenesCompra.frx":08EC
      FormatStyle(4)  =   "FrmInspeccionOrdenesCompra.frx":09A0
      FormatStyle(5)  =   "FrmInspeccionOrdenesCompra.frx":0A78
      FormatStyle(6)  =   "FrmInspeccionOrdenesCompra.frx":0B30
      FormatStyle(7)  =   "FrmInspeccionOrdenesCompra.frx":0C10
      ImageCount      =   0
      PrinterProperties=   "FrmInspeccionOrdenesCompra.frx":0C30
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   18135
      Begin VB.TextBox txtcod_item 
         Height          =   285
         Left            =   1185
         MaxLength       =   8
         TabIndex        =   9
         Top             =   645
         Width           =   1005
      End
      Begin VB.TextBox txtdes_item 
         Height          =   315
         Left            =   2235
         TabIndex        =   8
         Top             =   630
         Width           =   4200
      End
      Begin VB.TextBox txtDesProveedor 
         Height          =   285
         Left            =   6030
         MaxLength       =   50
         TabIndex        =   7
         Top             =   285
         Width           =   4155
      End
      Begin VB.TextBox txtCodProveedor 
         Height          =   285
         Left            =   4665
         MaxLength       =   12
         TabIndex        =   6
         Top             =   285
         Width           =   1365
      End
      Begin VB.ComboBox CmbTipItem 
         Height          =   315
         Left            =   1185
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   2355
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&BUSCAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   16440
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1185
      End
      Begin MSComCtl2.DTPicker dtpFec_INI 
         Height          =   285
         Left            =   11280
         TabIndex        =   20
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   128
         Format          =   73007105
         CurrentDate     =   38182
      End
      Begin MSComCtl2.DTPicker dtpFec_FIN 
         Height          =   285
         Left            =   13800
         TabIndex        =   21
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   128
         Format          =   73007105
         CurrentDate     =   38182
      End
      Begin VB.Label Label4 
         Caption         =   "FIN REQ:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   13080
         TabIndex        =   19
         Top             =   480
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "INICIO REQ:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10320
         TabIndex        =   17
         Top             =   480
         Width           =   1050
      End
      Begin VB.Label Label6 
         Caption         =   "ITEM"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   690
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "PROVEEDOR:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3600
         TabIndex        =   11
         Top             =   315
         Width           =   1005
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "TIPO ITEM:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Tag             =   "Hilado :"
         Top             =   315
         Width           =   870
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Left            =   15360
         TabIndex        =   3
         Top             =   120
         Width           =   75
      End
   End
   Begin VB.CheckBox chkExpandir 
      BackColor       =   &H80000010&
      Caption         =   "EXPANDIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   15840
      TabIndex        =   0
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1125
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   6885
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   18135
      _ExtentX        =   31988
      _ExtentY        =   12144
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GridLineStyle   =   2
      RowHeight       =   20
      HeaderFontName  =   "Arial"
      HeaderFontBold  =   -1  'True
      HeaderFontWeight=   700
      FontName        =   "Tahoma"
      ColumnHeaderHeight=   300
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "FrmInspeccionOrdenesCompra.frx":0E08
      FormatStyle(2)  =   "FrmInspeccionOrdenesCompra.frx":0F30
      FormatStyle(3)  =   "FrmInspeccionOrdenesCompra.frx":0FE0
      FormatStyle(4)  =   "FrmInspeccionOrdenesCompra.frx":1094
      FormatStyle(5)  =   "FrmInspeccionOrdenesCompra.frx":116C
      FormatStyle(6)  =   "FrmInspeccionOrdenesCompra.frx":1224
      FormatStyle(7)  =   "FrmInspeccionOrdenesCompra.frx":1304
      ImageCount      =   0
      PrinterProperties=   "FrmInspeccionOrdenesCompra.frx":1324
   End
   Begin VB.Label Label2 
      Caption         =   "ORIGEN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   690
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   240
      Top             =   6480
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmInspeccionOrdenesCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public CODIGO As String, Descripcion As String
Private StrSQL As String
Private sOpcion As String
Private indice As Integer
Private tipo As String

Private Sub chkExpandir_Click()
    If GridEX1.RowCount = 0 Then Exit Sub
    With GridEX1
        Select Case CBool(chkExpandir.Value)
            Case True: .ExpandAll
            Case False: .CollapseAll
        End Select
        If .RowCount > 0 Then
            .Row = -1
            .Col = .Columns.Count - 1
        End If
    End With
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub
Private Sub cmdGuardarCambios_Click()
  If GridEX1.RowCount = 0 Then Exit Sub
  Call guadarCambios
End Sub

Private Sub guadarCambios()
Dim rs As New ADODB.Recordset

Dim MSX As String
Dim X As Long

Dim MSY As String
Dim Y As Long

Dim MSW As String
Dim W As Long

Dim mensaje As String

X = 0
On Error GoTo fin:
Set rs = Nothing
GridEX1.Update
Set rs = GridEX1.ADORecordset
rs.Update
rs.MoveFirst
Do While Not rs.EOF
    If rs!cambiar.Value = True Then
      If Trim(rs!Cod_almacen) <> "" Then
           If Trim(rs!INSPECCION) <> "" Then
                If Trim(rs!accion) <> "" Then
                    
                    StrSQL = "UP_MAN_LG_ORDCOMPITEM_INSPECCION 'I','" & rs!Ser_OrdComp & "','" & rs!Cod_OrdComp & "','" & rs!Sec_OrdComp & _
                    "','" & rs!Cod_almacen & "','" & rs!num_movstk & "','" & rs!num_secuencia & "','" & rs!COD_USUARIO_INSPEC & _
                    "' ,'" & rs!INSPECCION & "','" & rs!accion & "','" & rs!fecha_inspec & "','" & rs!OBS_NO_CONFOR & "'"
                    Call ExecuteCommandSQL(cConnect, StrSQL)
                 Else
                    If W < 3 Then
                       MSW = MSW + "[" + rs!Ser_OrdComp + "-" + rs!Cod_OrdComp + "-" + rs!Sec_OrdComp + "]"
                       W = W + 1
                    End If
                 End If
            Else
                    If Y < 3 Then
                       MSY = MSY + "[" + rs!Ser_OrdComp + "-" + rs!Cod_OrdComp + "-" + rs!Sec_OrdComp + "]"
                       Y = Y + 1
                    End If
            End If
       Else
        
            If X < 3 Then
               MSX = MSX + "[" + rs!Ser_OrdComp + "-" + rs!Cod_OrdComp + "-" + rs!Sec_OrdComp + "]"
               X = X + 1
            End If
         
        
       End If
    End If
rs.MoveNext
Loop

mensaje = ""
If X > 0 Then
  mensaje = mensaje + " Algunas ordenes no tienen Ingreso en el almacen: " + MSX
End If

If Y > 0 Then
  mensaje = mensaje + " Seleccione Una Inspeccion: " + MSY
End If

If W > 0 Then
  mensaje = mensaje + " Seleccione una Accion : " + MSW
End If

If mensaje <> "" Then
    MsgBox "Revisar: " + mensaje, vbInformation + vbOKOnly, "Importante"
End If

Call BUSCAR

Exit Sub
fin:
MsgBox "Inconvenientes para realizar cambios", vbInformation + vbOKOnly, "Mensaje"

End Sub
Private Sub txtCodProveedor_KeyPress(KeyAscii As Integer)
    txtDesProveedor.Text = ""
    If KeyAscii = 13 Then
        'If Trim(txtCodProveedor.Text) <> "" Then
            txtCodProveedor.Text = Right("000000000000" & txtCodProveedor.Text, 12)
            Call BUSCA_PROVEEDOR(1, 1)
            cmdBuscar.SetFocus
        'End If
    End If
End Sub
Private Sub txtDesProveedor_KeyPress(KeyAscii As Integer)
    txtCodProveedor.Text = ""
    If KeyAscii = 13 Then
        'If Trim(txtDesProveedor.Text) <> "" Then
            Call BUSCA_PROVEEDOR(2, 1)
            cmdBuscar.SetFocus
        'End If
    End If
End Sub
Sub BUSCA_PROVEEDOR(tipo As Integer, Ubic As Integer)
    Select Case tipo
        Case 1:
                If Ubic = 1 Then
                
                    StrSQL = "SELECT Des_Proveedor FROM LG_PROVEEDOR WHERE Cod_Proveedor = '" & txtCodProveedor.Text & "'"
                    txtDesProveedor.Text = Trim(DevuelveCampo(StrSQL, cConnect))
                    'FunctBuscar.SetFocus
                End If
        Case 2:
                Dim oTipo As New frmBusqGeneral
                Dim rs As New ADODB.Recordset
                Set oTipo.oParent = Me
                If Ubic = 1 Then
                    oTipo.SQuery = "SELECT Cod_Proveedor as Código, Des_Proveedor as Descripción FROM LG_PROVEEDOR WHERE Des_Proveedor like '%" & Trim(txtDesProveedor.Text) & "%' order by des_proveedor"
                End If
                oTipo.Cargar_Datos
                oTipo.Show 1
                If CODIGO <> "" Then
                    If Ubic = 1 Then
                        txtCodProveedor.Text = Trim(CODIGO)
                        txtDesProveedor.Text = Trim(Descripcion)
                        'FunctBuscar.SetFocus
                        CODIGO = ""
                        Descripcion = ""
                    End If
                End If
                Set oTipo = Nothing
                Set rs = Nothing
                
    End Select
End Sub

Private Sub cmdBuscar_Click()
Call BUSCAR
End Sub
Private Sub Form_Load()
    
    StrSQL = "select Des_TipItem +space(100)+Tip_Item from lg_tipitem order by 1"
    LlenaCombo CmbTipItem, StrSQL, cConnect
    
'    StrSQL = "SELECT des_origen + space(100) + cod_origen  FROM LG_Origen"
'    LlenaCombo cboCod_Origen, StrSQL, cConnect
'    cboCod_Origen.AddItem "AMBOS   "
    
    dtpFec_INI = Now - 7
    dtpFec_FIN = Now
     
    Set GridEX2.ADORecordset = CargarRecordSetDesconectado("SELECT COD_CONFORMIDAD_INSPEC , DES_NO_CONFORMIDAD FROM LG_CONFORMIDAD", cConnect)
    GridEX2.ColumnAutoResize = True
    GridEX2.ActAsDropDown = True
    GridEX2.BoundColumnIndex = 1
    GridEX2.ReplaceColumnIndex = 2
    GridEX2.Columns("COD_CONFORMIDAD_INSPEC").Visible = False
    GridEX2.MoveFirst
    
    Set GridEX3.ADORecordset = CargarRecordSetDesconectado("SELECT COD_ACCION ,DES_ACCION_INSPECCION FROM LG_ACCION_INSPECCION WHERE COD_CONFORMIDAD_INSPEC = '01' ", cConnect)
    GridEX3.ColumnAutoResize = True
    GridEX3.ActAsDropDown = True
    GridEX3.BoundColumnIndex = 1
    GridEX3.ReplaceColumnIndex = 2
    GridEX3.Columns("COD_ACCION").Visible = False
    GridEX3.MoveFirst
  
End Sub

Sub BUSCAR()
On Error GoTo fin
  
    StrSQL = "LG_MUESTRA_INSPECCION_PRODUCTOS '" & Right(CmbTipItem, 1) & "','" & Trim(txtCodProveedor.Text) & "','" & Trim(txtcod_item.Text) & "','','" & vusu & "','" & Format(dtpFec_INI, "DD/MM/YYYY") & "','" & Format(dtpFec_FIN, "DD/MM/YYYY") & "'"
    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)
    
    Call CONFIGURA_GRILLA

Exit Sub

fin:
MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption

End Sub
Public Sub CONFIGURA_GRILLA()
    Dim C As Integer
    Dim colTemp As JSColumn
    Dim fmtCon  As JSFmtCondition
    
    With GridEX1
    
        For C = 1 To .Columns.Count
            With .Columns(C)
                .Caption = UCase(.Caption)
                .HeaderAlignment = jgexAlignCenter
                .TextAlignment = jgexAlignCenter
                .Visible = False
            End With
        Next C

        
        With .Columns("area")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "AREA"
        End With

        With .Columns("proveedor")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "PROVEEDOR"
        End With

        With .Columns("FEC_REQ")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "FEC REQ."
        End With
        With .Columns("FEC_PROGRAMADA")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "FEC PROGR."
        End With
        
        With .Columns("FEC_ENTREGA")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "FEC ENTREGA."
        End With
        
        With .Columns("ORDEN")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "ORDEN"
        End With
        
        ''PROVEEDOR
        
        With .Columns("COD_ITEM")
            .Visible = True
            .Width = 800
            .TextAlignment = jgexAlignLeft
            .Caption = "COD_ITEM"
        End With
        
        With .Columns("DES_ITEM")
            .Visible = True
            .Width = 1500
            .TextAlignment = jgexAlignLeft
            .Caption = "DES_ITEM"
        End With
        
        With .Columns("NUM_GUIA")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "NUM_GUIA"
        End With
        
        With .Columns("FACTURA")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "FACTURA"
        End With

        With .Columns("COD_USUARIO_INSPEC")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "USUARIO INSP."
        End With
        
        With .Columns("INSPECCION")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "INSPECCION"
        End With
        
        With .Columns("OBS_NO_CONFOR")
            .Visible = True
            .Width = 2500
            .TextAlignment = jgexAlignLeft
            .Caption = "OBS NO CONFOR"
        End With
        
        With .Columns("ACCION")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "ACCION"
        End With
       
        With .Columns("FECHA_INSPEC")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "FECHA INSPEC"
        End With
    
        With .Columns("UM")
            .Visible = True
            .Width = 800
            .TextAlignment = jgexAlignLeft
            .Caption = "UM"
        End With
    
        With .Columns("cambiar")
            .Visible = True
            .Width = 800
            .TextAlignment = jgexAlignLeft
            .Caption = "CAMBIAR"
        End With
    
        With GridEX1.Columns("INSPECCION")
          .TextAlignment = jgexAlignLeft
          .EditType = jgexEditCombo
          Set .DropDownControl = GridEX2
        End With
        
        With GridEX1.Columns("ACCION")
          .TextAlignment = jgexAlignLeft
          .EditType = jgexEditCombo
          Set .DropDownControl = GridEX3
        End With
    
    
    Dim oGroup01 As GridEX20.JSGroup
    Dim oGroup02 As GridEX20.JSGroup
    Dim valorcant   As JSColumn
    Dim valorStock   As JSColumn

'      With GridEX1
'
'        Set oGroup01 = .Groups.Add(.Columns("PROVEEDOR").Index, jgexSortAscending)
'        .DefaultGroupMode = jgexDGMExpanded
'        .BackColorRowGroup = RGB(239, 235, 222)
'
'           .GroupFooterStyle = jgexTotalsGroupFooter
'
'      End With
  End With

    'Call colorGrupo
    
End Sub
Private Sub colorGrupo()
    Dim fmtCon As JSFmtCondition
    
    Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("PROVEEDOR").Index, jgexGreaterThan, 0)
    
'    With GridEX1.FmtConditions
'            .ApplyGroupCondition = True
'            .ShowGroupConditionCount = True
'            .GroupConditionCountTitle = "ORDENES "
'            Set fmtCon = .GroupCondition
'    End With
    
'    fmtCon.SetCondition GridEX1.Columns("PROVEEDOR").Index, jgexGreaterThan, 0
'    fmtCon.FormatStyle.FontBold = True
'    fmtCon.FormatStyle.BackColor = &HFFFFC0   '&HC0FFC0    ' &HC0E0FF    ' '&HC0FFFF

End Sub

Private Sub setcolorcolumnas()
    'GridEX1.Columns("ot").CellStyle = "partida"
End Sub
Private Sub CmdImprimir_Click()
    If GridEX1.RowCount <= 0 Then Exit Sub
    Call Reporte
    
End Sub
Private Sub TxtCod_Item_KeyPress(KeyAscii As Integer)
    txtdes_item.Text = ""
    
    Dim StrSQL As String
    If KeyAscii = 13 Then
        If Trim(txtcod_item.Text) = "" Then
            Call MsgBox("Sirvase ingresar un codigo de Item", vbInformation)
        Else
            txtcod_item.Text = CompletaCodigo(Trim(txtcod_item.Text), 8, 2)
            
            'Esta consulta es para obtener el Codigo de Cliente
            StrSQL = "SELECT Des_Item FROM LG_ITEM WHERE Cod_Item='" & txtcod_item.Text & "'"
            txtdes_item.Text = DevuelveCampo(StrSQL, cConnect)
        End If
       ' FunctBuscar.SetFocus
    End If
End Sub

Private Sub txtdes_item_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

        Dim StrSQL As String
        If Len(Trim(txtdes_item.Text)) < 3 Then
            Call MsgBox("El Texto Ingresado debe contar con un mínimo de 3 caracteres", vbExclamation)
            Exit Sub
        Else
            StrSQL = "SELECT Cod_Item as Código, Des_Item as Descripción  FROM LG_ITEM WHERE Des_Item LIKE '" & Trim(txtdes_item.Text) & "%'"
        End If
        
        Dim oTipo As New frmBusqGeneral
        Dim rs As New ADODB.Recordset
        Set oTipo.oParent = Me
        oTipo.SQuery = StrSQL
        oTipo.Cargar_Datos
        oTipo.Show 1
        If CODIGO <> "" Then
            txtcod_item.Text = CODIGO
            txtdes_item.Text = Descripcion
        End If
        Set oTipo = Nothing
        Set rs = Nothing
    
End If

End Sub

'''===============================================================inicio de edicion de grilla
Private Sub GridEX1_AfterColEdit(ByVal ColIndex As Integer)
  AfterColEdit_DETALLE_FACTURA (ColIndex)
End Sub
Sub AfterColEdit_DETALLE_FACTURA(ByVal ColIndex As Integer)
Dim sSQL As String
On Error GoTo Error_Handler

Dim oGroup As GridEX20.JSGroup
Select Case ColIndex

  Case Is = GridEX1.Columns("COD_USUARIO_INSPEC").Index
    ''
  Case Is = GridEX1.Columns("INSPECCION").Index
    
    GridEX1.Value(GridEX1.Columns("accion").Index) = ""
    If Trim(GridEX1.Value(GridEX1.Columns("INSPECCION").Index)) = "" Then
        MsgBox "Seleccione una inspeccion", vbInformation + vbOKOnly, "Importante"
        Exit Sub
    Else
    
    Set GridEX3.ADORecordset = CargarRecordSetDesconectado("SELECT COD_ACCION ,DES_ACCION_INSPECCION FROM LG_ACCION_INSPECCION WHERE COD_CONFORMIDAD_INSPEC = '" & GridEX1.Value(GridEX1.Columns("INSPECCION").Index) & "' ", cConnect)
    GridEX3.ColumnAutoResize = True
    GridEX3.ActAsDropDown = True
    GridEX3.BoundColumnIndex = 1
    GridEX3.ReplaceColumnIndex = 2
    GridEX3.Columns("COD_ACCION").Visible = False
  
    End If

  End Select
Exit Sub

Resume
Error_Handler:
errores err.Number
End Sub

Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
  Select Case ColIndex
'    Case Is = GridEX1.Columns("COD_USUARIO_INSPEC").Index
'      Cancel = False
    Case Is = GridEX1.Columns("INSPECCION").Index
      Cancel = False
    Case Is = GridEX1.Columns("ACCION").Index
      Cancel = False
'    Case Is = GridEX1.Columns("FECHA_INSPEC").Index
'      Cancel = False
    Case Is = GridEX1.Columns("OBS_NO_CONFOR").Index
      Cancel = False
    Case Is = GridEX1.Columns("cambiar").Index
      Cancel = False
      
    Case Else
      Cancel = True
  End Select
End Sub

'Private Sub GridEX1_Click()
'    Dim ColIndex As Long
'
'    Dim oRowData As JSRowData
'    Dim SGRUPO As String
'    Dim iRow As Long
'    Dim i As Long
'    Dim sCaptionGroup As String
'        If GridEX1.RowCount > 0 Then
'        ColIndex = GridEX1.Col
'         If ColIndex = 0 Then Exit Sub
'
'            If UCase(GridEX1.Columns(ColIndex).Key) = "ELI" Then
'                bClickColSelec = True
'                SendKeys "{ENTER}"
''            ElseIf UCase(grxDatos.Columns(ColIndex).Key) = "CANT" Then
''                If IsNumeric(grxDatos.Value(grxDatos.Columns("CANT").Index)) = False Then
''                    grxDatos.Value(grxDatos.Columns("CANT").Index) = 0
''                End If
'            End If
'    End If
'End Sub
'''===============================================================fin edicion de grilla

Private Sub Reporte()
    
    Dim oo As Object
    Dim sRutaLogo  As String
    Dim Ruta As String
    On Error GoTo errReporte
    
    Ruta = vRuta & "\RptMuestraInspeccionOrdenes.xlt"
    
    Set oo = CreateObject("excel.application")
    
    StrSQL = "SELECT Ruta_Logo = ISNULL(Ruta_Logo, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
    sRutaLogo = DevuelveCampo(StrSQL, cConnect)
        
    oo.Workbooks.Open Ruta
    oo.Visible = False
    oo.DisplayAlerts = False
    oo.Run "Reporte", sRutaLogo, GridEX1.ADORecordset, txtDesProveedor.Text, CmbTipItem, "DESDE: " + Format(dtpFec_INI, "DD/MM/YYYY") + " HASTA: " + Format(dtpFec_FIN, "DD/MM/YYYY")
    oo.Visible = True
            
    Set oo = Nothing
    
    Exit Sub
errReporte:
        MsgBox "Hubo error en la impresion del Reporte de Tiempos " & err.Description, vbCritical, "Impresion"
    
End Sub


Private Sub GridEX1_GroupByBoxHeaderClick(ByVal Group As JSGroup)
    Group.SortOrder = -Group.SortOrder
End Sub


