VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmPackingListExpo 
   Caption         =   "Packing List Exportacion Textil"
   ClientHeight    =   9645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   ScaleHeight     =   9645
   ScaleWidth      =   13395
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   240
      TabIndex        =   36
      Top             =   2880
      Width           =   12975
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   435
         Left            =   5760
         TabIndex        =   40
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtPartida 
         Height          =   285
         Left            =   840
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox ChkTodo2 
         Caption         =   "Seleccionar Todo"
         Height          =   255
         Left            =   11160
         TabIndex        =   37
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Partidas"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton CmdEliminar 
      Height          =   495
      Left            =   6240
      Picture         =   "FrmPackingListExpo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6000
      Width           =   1245
   End
   Begin VB.CommandButton CmdAnadir 
      Height          =   495
      Left            =   4440
      Picture         =   "FrmPackingListExpo.frx":05CF
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6000
      Width           =   1365
   End
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   510
      Left            =   9840
      TabIndex        =   31
      Top             =   9000
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
   Begin VB.TextBox TxtContRollos 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   29
      Text            =   "0"
      Top             =   9000
      Width           =   975
   End
   Begin VB.CheckBox ChkTodo1 
      Caption         =   "Seleccionar Todo"
      Height          =   255
      Left            =   11520
      TabIndex        =   27
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Frame FrmFardo 
      Height          =   615
      Left            =   240
      TabIndex        =   19
      Top             =   2280
      Visible         =   0   'False
      Width           =   12975
      Begin VB.TextBox TxtFardos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   33
         Text            =   "0"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtKilosPorCargar 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   8640
         TabIndex        =   26
         Text            =   "0"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtKilosCargados 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   6000
         TabIndex        =   24
         Text            =   "0"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtKilosFardo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         TabIndex        =   22
         Text            =   "0"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Kilos por Cargar:"
         Height          =   255
         Left            =   7320
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label label100 
         Caption         =   "Kilos Cargados:"
         Height          =   255
         Left            =   4800
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Kilos por Fardo:"
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Fardo N°"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Tipo Traslado"
      Height          =   855
      Left            =   360
      TabIndex        =   17
      Top             =   1320
      Width           =   4215
      Begin VB.TextBox TxtNum_Fardos 
         Height          =   285
         Left            =   2760
         TabIndex        =   8
         Text            =   "0"
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton OptRollo 
         BackColor       =   &H0080FFFF&
         Caption         =   "Por Rollo"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton OptFardo 
         BackColor       =   &H0080FFFF&
         Caption         =   "Por Fardo"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblNum_fardos 
         BackColor       =   &H0080FFFF&
         Caption         =   "N° de Fardos:"
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Generar Packing List"
      Height          =   495
      Left            =   6000
      TabIndex        =   13
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   12975
      Begin MSComCtl2.DTPicker DTEmision 
         Height          =   255
         Left            =   8280
         TabIndex        =   14
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   70516737
         CurrentDate     =   41366
      End
      Begin VB.TextBox TXTFacturaProforma 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtSer_OrdComp 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   720
         Width           =   525
      End
      Begin VB.TextBox txtCod_OrdComp 
         Height          =   285
         Left            =   2340
         TabIndex        =   4
         Top             =   720
         Width           =   1155
      End
      Begin VB.TextBox txtNom_Cliente 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   300
         Width           =   3480
      End
      Begin VB.TextBox txtAbr_Cliente 
         Height          =   285
         Left            =   720
         TabIndex        =   1
         Top             =   300
         Width           =   735
      End
      Begin VB.Label LblPackingList 
         BackColor       =   &H0080FFFF&
         Caption         =   "Label8"
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
         Left            =   6600
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FFFF&
         Caption         =   "Packing List Nro"
         Height          =   255
         Left            =   5160
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080FFFF&
         Caption         =   "Fecha Emision"
         Height          =   255
         Left            =   7080
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Factura Proforma"
         Height          =   255
         Left            =   3600
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Cliente:"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Orden de Pedido:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Height          =   2355
      Left            =   240
      TabIndex        =   28
      Top             =   3600
      Width           =   12960
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   12
      AllowColumnShrinking=   0   'False
      SelectTypeRow   =   1
      BackColorOdd    =   13434879
      RowHeight       =   423
      ExtraHeight     =   53
      Columns.Count   =   12
      Columns(0).Width=   529
      Columns(0).Name =   "Check"
      Columns(0).Alignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Style=   2
      Columns(1).Width=   1111
      Columns(1).Caption=   "Partida"
      Columns(1).Name =   "Partida"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   2408
      Columns(2).Caption=   "Almacen"
      Columns(2).Name =   "Almacen"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   1720
      Columns(3).Caption=   "Movimiento"
      Columns(3).Name =   "Movimiento"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   1535
      Columns(4).Caption=   "Sec. Mov."
      Columns(4).Name =   "Secuencia_General"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   1640
      Columns(5).Caption=   "Sec. Rollo Mov."
      Columns(5).Name =   "Secuencia"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(5).Locked=   -1  'True
      Columns(6).Width=   3519
      Columns(6).Caption=   "Tela"
      Columns(6).Name =   "Tela"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(6).Locked=   -1  'True
      Columns(7).Width=   3598
      Columns(7).Caption=   "Color"
      Columns(7).Name =   "Color"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(7).Locked=   -1  'True
      Columns(8).Width=   1535
      Columns(8).Caption=   "Cod_Rollo"
      Columns(8).Name =   "Cod_Rollo"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(8).Locked=   -1  'True
      Columns(9).Width=   1376
      Columns(9).Caption=   "Peso"
      Columns(9).Name =   "Peso"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(9).Locked=   -1  'True
      Columns(10).Width=   1429
      Columns(10).Caption=   "PesoNeto"
      Columns(10).Name=   "PesoNeto"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   3200
      Columns(11).Caption=   "IdRolloUnico"
      Columns(11).Name=   "IdRolloUnico"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      _ExtentX        =   22860
      _ExtentY        =   4154
      _StockProps     =   79
      Caption         =   "Rollos Disponibles del Almacen"
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid2 
      Height          =   2355
      Left            =   240
      TabIndex        =   32
      Top             =   6480
      Width           =   12945
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   13
      AllowColumnShrinking=   0   'False
      SelectTypeRow   =   1
      BackColorOdd    =   13434879
      RowHeight       =   423
      ExtraHeight     =   53
      Columns.Count   =   13
      Columns(0).Width=   529
      Columns(0).Name =   "Check"
      Columns(0).Alignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Style=   2
      Columns(1).Width=   1191
      Columns(1).Caption=   "Sec.PL"
      Columns(1).Name =   "Secuencia"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   1111
      Columns(2).Caption=   "Partida"
      Columns(2).Name =   "Partida"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   2408
      Columns(3).Caption=   "Almacen"
      Columns(3).Name =   "Almacen"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   1720
      Columns(4).Caption=   "Movimiento"
      Columns(4).Name =   "Movimiento"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   1535
      Columns(5).Caption=   "Sec. Mov."
      Columns(5).Name =   "Secuencia_General"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(5).Locked=   -1  'True
      Columns(6).Width=   3519
      Columns(6).Caption=   "Tela"
      Columns(6).Name =   "Tela"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(6).Locked=   -1  'True
      Columns(7).Width=   3757
      Columns(7).Caption=   "Color"
      Columns(7).Name =   "Color"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(7).Locked=   -1  'True
      Columns(8).Width=   1561
      Columns(8).Caption=   "Sec_Rollo_Fardo"
      Columns(8).Name =   "Secuencia_Rollo"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(8).Locked=   -1  'True
      Columns(9).Width=   1535
      Columns(9).Caption=   "Cod_Rollo"
      Columns(9).Name =   "Cod_Rollo"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(9).Locked=   -1  'True
      Columns(10).Width=   1376
      Columns(10).Caption=   "Peso"
      Columns(10).Name=   "Peso"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(10).Locked=   -1  'True
      Columns(11).Width=   1429
      Columns(11).Caption=   "PesoNeto"
      Columns(11).Name=   "PesoNeto"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   3200
      Columns(12).Caption=   "IdRolloUnico"
      Columns(12).Name=   "IdRolloUnico"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      _ExtentX        =   22834
      _ExtentY        =   4154
      _StockProps     =   79
      Caption         =   "Rollos del Packing List"
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label11 
      Caption         =   "Total Rollo(s)"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   9120
      Width           =   975
   End
End
Attribute VB_Name = "FrmPackingListExpo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Opcion As String
Public CODIGO As String
Public descripcion As String
Dim Codigo_Cliente As String
Public frm_Opcion As String
Public tipo_Traslado As String
Public cod_PackingList As String
Dim rsx1 As New ADODB.Recordset
Dim rsx2 As New ADODB.Recordset

Private Sub ChkTodo1_Click()
Dim contador As Integer
If ChkTodo1 Then
        
    If SSDBGrid2.Rows < 1 Then
        Exit Sub
    End If
    For contador = 0 To SSDBGrid2.Rows - 1
       
        SSDBGrid2.Bookmark = contador
        SSDBGrid2.Columns("check").Value = True
    Next
Else
    If SSDBGrid2.Rows < 1 Then
        Exit Sub
    End If
    For contador = 0 To SSDBGrid2.Rows - 1
        SSDBGrid2.Bookmark = contador
        SSDBGrid2.Columns("check").Value = False
    Next '(opcionProv - 1)
End If
End Sub

Private Sub ChkTodo2_Click()
Dim contador As Integer
If ChkTodo2 Then
        
    If SSDBGrid1.Rows < 1 Then
        Exit Sub
    End If
    For contador = 0 To SSDBGrid1.Rows - 1
       
        SSDBGrid1.Bookmark = contador
        SSDBGrid1.Columns("check").Value = True
    Next
Else
    If SSDBGrid1.Rows < 1 Then
        Exit Sub
    End If
    For contador = 0 To SSDBGrid1.Rows - 1
        SSDBGrid1.Bookmark = contador
        SSDBGrid1.Columns("check").Value = False
    Next '(opcionProv - 1)
End If
'End Sub




'Private Sub CmbFardos_Change()
'On Error GoTo xerror:
'Carga_GridRolloFardos
'Exit Sub
'xerror:
'Errores err.Number
'Exit Sub

End Sub

Private Sub CmbFardos_Click()
'MsgBox CmbFardos.Text
On Error GoTo xerror:
    Carga_GridRolloFardos
    Cargar_Totales_x_FArdo
    Exit Sub
xerror:
   errores err.Number
   Exit Sub
End Sub

Private Sub CmbFardos_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    'MsgBox CmbFardos.Text
    'Call SoloNumeros(CmbFardos, KeyAscii, True, 2, 5)
End If
End Sub

Private Sub CmdAnadir_Click()
On Error GoTo xerror:
Dim vMessage As Variant
Dim vResp As String, sTit As String
Dim filas As Integer
Dim StrSQLMant As String
 Dim i As Integer
If FrmFardo.Visible = True Then

    If TxtKilosFardo = 0 Then
        vMessage = (MsgBox("¿Desea Ingresar la cantidad de Kilos al Fardo N° " & TxtFardos.Text & "?", vbQuestion + vbYesNo, sTit))
        If vMessage = vbYes Then
        TxtKilosFardo.SetFocus
        With TxtKilosFardo
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
            Exit Sub
        End If
    End If
    vMessage = (MsgBox("¿Desea Añadir este Detalle al Fardo N° " & TxtFardos.Text, vbQuestion + vbYesNo, sTit))
    If vMessage = vbYes Then
           ' Dim i As Integer
            For i = 0 To SSDBGrid1.Rows - 1
            SSDBGrid1.Bookmark = i
             If SSDBGrid1.Columns("check").Value = True Then
             
             StrSQLMant = "Exec Ti_Mant_Detalle_PackingList 'I','" & txtSer_OrdComp & "','" & txtCod_OrdComp & "','" & TxtAbr_Cliente & "','" & _
                           LblPackingList.Caption & "','" & TXTFacturaProforma.Text & "','" & tipo_Traslado & "','" & _
                           0 & "','" & DTEmision.Value & "',0,'" & TxtFardos.Text & "',0,'" & Mid(SSDBGrid1.Columns("Almacen").Value, 1, 2) & _
                           "','" & SSDBGrid1.Columns("Movimiento").Value & "','" & SSDBGrid1.Columns("Secuencia_General").Value & _
                           "','" & SSDBGrid1.Columns("Cod_rollo").Value & "','" & SSDBGrid1.Columns("Partida").Value & "','" & _
                           Mid(SSDBGrid1.Columns("Tela").Value, 1, 8) & "','" & Mid(SSDBGrid1.Columns("Color").Value, 1, 6) & "','" & TxtKilosFardo & "','" & _
                           SSDBGrid1.Columns("Peso").Value & "','" & SSDBGrid1.Columns("Secuencia_General").Value & "','" & SSDBGrid1.Columns("PesoNeto").Value & "','" & IdRolloUnico & "'"
                    filas = ExecuteSQL(cConnect, StrSQLMant)
                    
            End If
            Next
            carga_GRidRollosAlmacen
            Carga_GridRolloFardos
            Cargar_Totales_x_FArdo
    End If
Else
    vMessage = (MsgBox("¿Desea Añadir este Detalle como Rollos individuales?", vbQuestion + vbYesNo, sTit))
    If vMessage = vbYes Then
        For i = 0 To SSDBGrid1.Rows - 1
            SSDBGrid1.Bookmark = i
             If SSDBGrid1.Columns("check").Value = True Then
             
             StrSQLMant = "Exec Ti_Mant_Detalle_PackingList 'I','" & txtSer_OrdComp & "','" & txtCod_OrdComp & "','" & TxtAbr_Cliente & "','" & _
                           LblPackingList.Caption & "','" & TXTFacturaProforma.Text & "','" & tipo_Traslado & "','" & _
                           0 & "','" & DTEmision.Value & "',0,'" & TxtFardos.Text & "',0,'" & Mid(SSDBGrid1.Columns("Almacen").Value, 1, 2) & _
                           "','" & SSDBGrid1.Columns("Movimiento").Value & "','" & SSDBGrid1.Columns("Secuencia_General").Value & _
                           "','" & SSDBGrid1.Columns("Cod_rollo").Value & "','" & SSDBGrid1.Columns("Partida").Value & "','" & _
                           Mid(SSDBGrid1.Columns("Tela").Value, 1, 8) & "','" & Mid(SSDBGrid1.Columns("Color").Value, 1, 6) & "','" & TxtKilosFardo & "','" & _
                           SSDBGrid1.Columns("Peso").Value & "','" & SSDBGrid1.Columns("Secuencia_General").Value & "','" & SSDBGrid1.Columns("PesoNeto").Value & "','" & SSDBGrid1.Columns("IdRolloUnico").Value & "'"
                    filas = ExecuteSQL(cConnect, StrSQLMant)
                    
            End If
            Next
            carga_GRidRollosAlmacen
            Carga_GridRolloFardos
    End If
End If
Exit Sub
xerror:
errores err.Number
carga_GRidRollosAlmacen
Carga_GridRolloFardos
If OptFardo = True Or frm_Opcion = "F" Then
Cargar_Totales_x_FArdo
End If
Exit Sub
End Sub

Private Sub CmdBuscar_Click()
Call carga_GRidRollosAlmacen
End Sub

Private Sub CmdEliminar_Click()
On Error GoTo xerror:
Dim vMessage As Variant
Dim vResp As String, sTit As String
Dim StrSQLMant As String
Dim i As Integer
Dim filas As Integer
If FrmFardo.Visible = True Then
    
    vMessage = (MsgBox("¿Desea Eliminar este Detalle del Fardo N° " & TxtFardos.Text, vbQuestion + vbYesNo, sTit))
    If vMessage = vbYes Then

            For i = 0 To SSDBGrid2.Rows - 1
            SSDBGrid2.Bookmark = i
            StrSQLMant = ""
             If SSDBGrid2.Columns("check").Value = True Then
       StrSQLMant = "Exec Ti_Mant_Detalle_PackingList 'D','" & txtSer_OrdComp & "','" & txtCod_OrdComp & "','" & TxtAbr_Cliente & "','" & _
                           LblPackingList.Caption & "','" & TXTFacturaProforma.Text & "','" & tipo_Traslado & "','" & _
                           SSDBGrid2.Columns("Secuencia_rollo").Value & "','" & DTEmision.Value & "',0,'" & TxtFardos.Text & "',0,'" & _
                           Mid(SSDBGrid2.Columns("Almacen").Value, 1, 2) & _
                           "','" & SSDBGrid2.Columns("Movimiento").Value & "','" & SSDBGrid2.Columns("Secuencia_General").Value & _
                           "','" & SSDBGrid2.Columns("Cod_rollo").Value & "','" & SSDBGrid2.Columns("Partida").Value & "','" & _
                           Mid(SSDBGrid2.Columns("Tela").Value, 1, 8) & "','" & Mid(SSDBGrid2.Columns("Color").Value, 1, 6) & "','" & TxtKilosFardo & "','" & _
                           SSDBGrid2.Columns("Peso").Value & "','" & SSDBGrid2.Columns("Secuencia_rollo").Value & "','0','" & SSDBGrid2.Columns("IdRolloUnico").Value & "'"
            End If
            If StrSQLMant <> "" Then
                filas = ExecuteSQL(cConnect, StrSQLMant)
            End If
            Next
       'Dim filas As Integer
        carga_GRidRollosAlmacen
       Carga_GridRolloFardos
       Cargar_Totales_x_FArdo

    End If
Else
    vMessage = (MsgBox("¿Desea Eliminar este Detalle?", vbQuestion + vbYesNo, sTit))
    If vMessage = vbYes Then
       For i = 0 To SSDBGrid2.Rows - 1
            SSDBGrid2.Bookmark = i
            StrSQLMant = ""
             If SSDBGrid2.Columns("check").Value = True Then
       StrSQLMant = "Exec Ti_Mant_Detalle_PackingList 'D','" & txtSer_OrdComp & "','" & txtCod_OrdComp & "','" & TxtAbr_Cliente & "','" & _
                           LblPackingList.Caption & "','" & TXTFacturaProforma.Text & "','" & tipo_Traslado & "','" & _
                           SSDBGrid2.Columns("Secuencia").Value & "','" & DTEmision.Value & "',0,'" & TxtFardos.Text & "',0,'" & _
                           Mid(SSDBGrid2.Columns("Almacen").Value, 1, 2) & _
                           "','" & SSDBGrid2.Columns("Movimiento").Value & "','" & SSDBGrid2.Columns("Secuencia_General").Value & _
                           "','" & SSDBGrid2.Columns("Cod_rollo").Value & "','" & SSDBGrid2.Columns("Partida").Value & "','" & _
                           Mid(SSDBGrid2.Columns("Tela").Value, 1, 8) & "','" & Mid(SSDBGrid2.Columns("Color").Value, 1, 6) & "','" & TxtKilosFardo & "','" & _
                           SSDBGrid2.Columns("Peso").Value & "','" & SSDBGrid2.Columns("Secuencia_rollo").Value & "','0','" & SSDBGrid2.Columns("IdRolloUnico").Value & "'"
            End If
            If StrSQLMant <> "" Then
                filas = ExecuteSQL(cConnect, StrSQLMant)
            End If
            Next
      
    End If
     carga_GRidRollosAlmacen
       Carga_GridRolloFardos
      ' Cargar_Totales_x_FArdo
End If
      
Exit Sub
xerror:
errores err.Number
Exit Sub
End Sub

Sub Cargar_Totales_x_FArdo()
Dim Total_Kilos_Fardo As Double
Dim StrsqlX As String
Dim rsx_Y As New ADODB.Recordset
StrsqlX = "exec Lista_Totales_Cabecera_FArdo'" & TxtAbr_Cliente & "','" & txtSer_OrdComp & "','" & txtCod_OrdComp & _
         "','" & LblPackingList.Caption & "'," & TxtFardos.Text & ""
Set rsx_Y = CargarRecordSetDesconectado(StrsqlX, cConnect)
TxtKilosFardo = rsx_Y.Fields("Kilos_x_Fardo").Value
txtKilosCargados = rsx_Y.Fields("Kilos_Cargados").Value
TxtKilosPorCargar = rsx_Y.Fields("Kilos_Restantes").Value

If TxtKilosFardo.Text < txtKilosCargados Then
    TxtKilosFardo.ForeColor = &HFF&
Else
    TxtKilosFardo.ForeColor = &H80000008
End If

If TxtKilosPorCargar < 0 Then
    TxtKilosPorCargar.ForeColor = &HFF&
Else
    TxtKilosPorCargar.ForeColor = &H80000008
End If

End Sub


Private Sub Command1_Click()
On Error GoTo xerror:
Dim StrsqlX As String

Dim Cod_PackingList_Generado As String
Dim Cantidad_Filas As Integer

    If Trim(TxtAbr_Cliente.Text = "") Or Trim(TxtNom_Cliente.Text) = "" Then
        MsgBox "Debe seleccionar un Cliente", vbCritical, "Mensaje"
        TxtAbr_Cliente.SetFocus
        Exit Sub
    End If
    If Trim(txtSer_OrdComp) = "" Or Trim(txtCod_OrdComp) = "" Then
        MsgBox "Debe Ingresar una Orden de Compra Valida", vbCritical, "Mensaje"
        txtSer_OrdComp.SetFocus
        Exit Sub
    End If
    
    If Trim(TXTFacturaProforma) = "" Then
        MsgBox "Debe Ingresar la FActura Proforma", vbCritical, "Mensaje"
        TXTFacturaProforma.SetFocus
        Exit Sub
    End If
    
'    If OptFardo.Value = True Then
'        If Trim(TxtNum_Fardos.Text) <> "" Then
'            If (TxtNum_Fardos.Text) < 1 Then
'                MsgBox "Debe Ingresar la cantidad de Fardos", vbCritical, "Mensaje"
'            TxtNum_Fardos.SetFocus
'            Exit Sub
'            End If
'        Else
'            MsgBox "Debe Ingresar la cantidad de Fardos", vbCritical, "Mensaje"
'            TxtNum_Fardos.SetFocus
'            Exit Sub
'        End If
'    End If

If OptFardo.Value = False And OptRollo.Value = False Then
    MsgBox "Debe Seleccionar una forma de Traslado de los rollos", vbCritical, "Mensaje"
    OptFardo.SetFocus
    Exit Sub
End If

If OptFardo.Value = True Then
    tipo_Traslado = "F"
Else
    tipo_Traslado = "R"
End If
    'Procedimiento para Registrar Packing Cabecera
      StrsqlX = "Exec USP_Mant_PAckingList 'I','" & Trim(TxtAbr_Cliente) & "','" & Trim(txtSer_OrdComp) & "','" & _
              Trim(txtCod_OrdComp) & "','" & TXTFacturaProforma & "','','" & DTEmision.Value & "','','" & tipo_Traslado & _
              "',0," & TxtNum_Fardos & ",0,0,0,'" & vusu & "'"
              
     LblPackingList.Visible = True
     Cantidad_Filas = ExecuteSQL(cConnect, StrsqlX)
     If Cantidad_Filas > 0 Then
     LblPackingList.Caption = DevuelveCampo("select max(Cod_PackingList) from TI_PACKING_LIST_EXPO_CAB ", cConnect)
     MsgBox "Se genero satisfactoriamente el packing List", vbInformation, "MEnsaje"
     Deshabilita_Campos
     CmdAnadir.Enabled = True
     CmdEliminar.Enabled = True
     carga_GRidRollosAlmacen
     
     If OptFardo.Value = True Then
        'CArgar_Combo
        TxtFardos.Enabled = True
     Else
        FrmFardo.Visible = False
     End If
     
      End If
    Exit Sub
xerror:
errores err.Number
Exit Sub
End Sub

Sub carga_GRidRollosAlmacen()
    Set rsx1 = New ADODB.Recordset
    Dim rs_Prov1 As ADODB.Recordset
    rsx1.CursorLocation = adUseClient
    rsx1.CursorType = adOpenStatic
    rsx1.ActiveConnection = cConnect
    If rsx1.State <> 0 Then rsx1.Close
    rsx1.Open "Exec Lista_Rollos_Disponibles_Packing_List '" & TxtAbr_Cliente & "','" & txtSer_OrdComp & "','" & txtCod_OrdComp & "','" & Trim(txtPartida.Text) & "'"
    Set rs_Prov1 = rsx1.Clone
    If rsx1.RecordCount >= 0 Then
    
        Me.SSDBGrid1.Redraw = False
        SSDBGridSetGrid Me.SSDBGrid1
        ADODBToSSDBGridOC rs_Prov1, SSDBGrid1
        SSDBGrid1.ActiveRowStyleSet = "RowActive"
        SSDBGrid1.SelectTypeRow = ssSelectionTypeMultiSelectRange
        Me.SSDBGrid1.Visible = True
        'me.SSDBGrid1.ColumnHeaders("").
       'Me.SSDBGrid1.Columns("Secuencia_General").Visible = False

   End If
End Sub

Sub Carga_GridRolloFardos()
Dim rs_Prov2 As ADODB.Recordset
 Set rsx2 = New ADODB.Recordset
 Set rs_Prov2 = New ADODB.Recordset
     
     If rsx2.State <> 0 Then rsx2.Close
     If rs_Prov2.State <> 0 Then rs_Prov2.Close
    rsx2.CursorLocation = adUseClient
    rsx2.CursorType = adOpenStatic
    rsx2.ActiveConnection = cConnect
    
         rsx2.Open "Exec Listar_Detalle_Rollos_PackingList '" & TxtAbr_Cliente & "','" & txtSer_OrdComp & "','" & txtCod_OrdComp & "','" & _
                 LblPackingList.Caption & "','" & tipo_Traslado & "','" & TxtFardos.Text & "'"
     'If rsx2.RecordCount Then
         Set rs_Prov2 = rsx2.Clone
         'If rsx2.RecordCount > 0 Then
             Me.SSDBGrid2.Redraw = False
             SSDBGridSetGrid Me.SSDBGrid2
             ADODBToSSDBGridOC rs_Prov2, SSDBGrid2
             SSDBGrid2.ActiveRowStyleSet = "RowActive"
             SSDBGrid2.SelectTypeRow = ssSelectionTypeMultiSelectRange
             Me.SSDBGrid2.Visible = True
       '     Me.SSDBGrid2.Columns("Secuencia_General").Visible = False
        'End If
   'End If
   TxtContRollos.Text = SSDBGrid2.Rows
End Sub

Sub Deshabilita_Campos()
    DTEmision.Enabled = False
    TxtAbr_Cliente.Enabled = False
    TxtNom_Cliente.Enabled = False
    TXTFacturaProforma.Enabled = False
    OptFardo.Enabled = False
    OptRollo.Enabled = False
    TxtNum_Fardos.Enabled = False
    Command1.Enabled = False
    'TxtKilosFardo.Enabled = False
    txtCod_OrdComp.Enabled = False
    txtSer_OrdComp.Enabled = False
    CmdAnadir.Enabled = False
    CmdEliminar.Enabled = False
End Sub

Sub CArgar_Combo()
'Dim i As Integer
'CmbFardos.Enabled = True
'CmbFardos.Clear
'For i = 1 To CInt(TxtNum_Fardos.Text)
'    CmbFardos.AddItem i
'Next
'CmbFardos.ListIndex = 0
End Sub

Private Sub Form_Load()
LblPackingList.Visible = False
DTEmision.Value = Date
If cod_PackingList <> "" And frm_Opcion = "U" Then
LblPackingList = cod_PackingList
LblPackingList.Visible = True
Else
LblPackingList = ""
LblPackingList.Visible = False
End If

OptFardo_Click
    If Opcion = "I" Then
        LblPackingList.Visible = False
    End If
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
 Select Case ActionName
        Case "IMPRIMIR"
            
       
        Case "SALIR"
            Unload Me
    End Select
End Sub

Private Sub OptFardo_Click()
TxtNum_Fardos.Visible = True
lblNum_fardos.Visible = True
FrmFardo.Visible = True
End Sub

Private Sub OptRollo_Click()
TxtNum_Fardos.Visible = False
lblNum_fardos.Visible = False
FrmFardo.Visible = False
End Sub



Private Sub txtAbr_Cliente_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        If Trim(TxtAbr_Cliente.Text) = "" Then
            Call BUSCA_CLIENTE(3)
        Else
            Call BUSCA_CLIENTE(1)
        End If
    End If
End Sub

Private Sub txtCod_OrdComp_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        'TxtFacturaProforma.SetFocus
        txtCod_OrdComp.Text = Format(Trim(txtCod_OrdComp.Text), "000000")
        If Len(Trim(txtCod_OrdComp)) = 6 And Len(Trim(txtSer_OrdComp)) = 3 Then
        Call Busca_Facturas_Proformas(1)
        End If
    Else
        Call SoloNumeros(txtCod_OrdComp, KeyAscii, False, 0, 6)
    End If
End Sub
Private Sub txtCod_OrdComp_LostFocus()
    txtCod_OrdComp.Text = Format(Trim(txtCod_OrdComp.Text), "000000")
End Sub

Private Sub TxtFacturaProforma_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TXTFacturaProforma.Text = Format(Trim(TXTFacturaProforma.Text), "00000000")
    Call Busca_Facturas_Proformas(2)
    End If
End Sub

Private Sub TxtFacturaProforma_LostFocus()
    TXTFacturaProforma.Text = Format(Trim(TXTFacturaProforma.Text), "00000000")
End Sub

Sub Busca_Facturas_Proformas(Tipo As Integer)
Dim STRSQL As String
    Select Case Tipo
        Case 1:
            Dim rsFAct As New ADODB.Record
                    STRSQL = "EXEC Lista_Facturas_Proforma_Packing_List '" & Trim(Me.TxtAbr_Cliente.Text) & "','" & Trim(Me.txtSer_OrdComp.Text) & "','" & Trim(Me.txtCod_OrdComp) & "'"
                    TXTFacturaProforma.Text = Trim(DevuelveCampo(STRSQL, cConnect))
                    
                    'Me.txtAbr_Cliente.Text = UCase(Me.txtAbr_Cliente.Text)
                    'If Trim(txtNom_Cliente.Text) <> "" Then CARGA_GRID
        Case 2:
                    Dim oTipo As New frmBusGeneral6
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.SQuery = "EXEC Lista_Facturas_Proforma_Packing_List '" & Trim(Me.TxtAbr_Cliente.Text) & "','" & Trim(Me.txtSer_OrdComp.Text) & "','" & Trim(Me.txtCod_OrdComp) & "'"
                    'Else
                     '   oTipo.SQuery = "EXEC TI_BUSCA_CLIENTE 3,'','','" & vusu & "'"
                    End If
                    
                    oTipo.CARGAR_DATOS
                   ' oTipo.DGridLista.Columns(2).Width = 3500
                    oTipo.Show 1
                    If CODIGO <> "" Then
                         Me.TXTFacturaProforma.Text = Trim(CODIGO)
                         'Me.txtNom_Cliente.Text = Trim(descripcion)
'                         OptCliPend.SetFocus
                         CODIGO = "": descripcion = ""
                        ' CARGA_GRID
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
    End Select
    
End Sub

Private Sub TxtFardos_Change()
On Error GoTo xerror:
    Carga_GridRolloFardos
    Cargar_Totales_x_FArdo
    Exit Sub
xerror:
   errores err.Number
   Exit Sub
End Sub

Private Sub TxtFardos_KeyPress(KeyAscii As Integer)
       If KeyAscii <> 13 Then
            Call SoloNumeros(TxtKilosFardo, KeyAscii, True, 2, 5)
       End If
End Sub

Private Sub TxtKilosFardo_KeyPress(KeyAscii As Integer)
'On Error GoTo xerror:
       
       
       If KeyAscii = 13 Then
            Dim strsql_UP As String
            Dim vMessage As Variant
            Dim fila As Integer
            'If CmbFardos.Enabled = True And TxtKilosFardo.Text <> "" And TxtKilosFardo.Text <> "0" Then
            If TxtKilosFardo.Text <> "" And TxtKilosFardo.Text <> "0" Then
                vMessage = (MsgBox("¿Desea Actualizar el peso Total del Fardo N° " & TxtFardos.Text, vbQuestion + vbYesNo, "Actualizar"))
                If vMessage = vbYes Then
                       strsql_UP = "Exec Actualiza_Total_Por_FArdo '" & TxtAbr_Cliente & "','" & txtSer_OrdComp & "','" & txtCod_OrdComp & _
                                    "','" & LblPackingList.Caption & "','" & TXTFacturaProforma & "','" & TxtFardos.Text & "','" & Me.TxtKilosFardo & "'"
                                    
                fila = ExecuteSQL(cConnect, strsql_UP)
                MsgBox "Se actualizo correctamente el total de Kilos del Fardo N° " & TxtFardos.Text, vbInformation
                Cargar_Totales_x_FArdo
                End If
            End If
       Else
               Call SoloNumeros(TxtKilosFardo, KeyAscii, True, 2, 5)
       End If
       
'xerror:
'       Errores err.Number
 '      Exit Sub
End Sub

Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Trim(TxtNom_Cliente.Text) = "" Then
            Call BUSCA_CLIENTE(3)
        Else
            Call BUSCA_CLIENTE(2)
        End If
    End If
End Sub
Public Sub BUSCA_CLIENTE(Tipo As Integer)
Dim STRSQL As String
    Select Case Tipo
        Case 1:
                    STRSQL = "EXEC TI_BUSCA_CLIENTE 1,'" & Trim(Me.TxtAbr_Cliente.Text) & "','','" & vusu & "'"
                    Me.TxtNom_Cliente.Text = Trim(DevuelveCampo(STRSQL, cConnect))
                    Me.TxtAbr_Cliente.Text = UCase(Me.TxtAbr_Cliente.Text)
                    'If Trim(txtNom_Cliente.Text) <> "" Then CARGA_GRID
        Case 2, 3:
                    Dim oTipo As New frmBusGeneral6
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.SQuery = "EXEC TI_BUSCA_CLIENTE 2,'','" & Trim(TxtNom_Cliente.Text) & "','" & vusu & "'"
                    Else
                        oTipo.SQuery = "EXEC TI_BUSCA_CLIENTE 3,'','','" & vusu & "'"
                    End If
                    
                    oTipo.CARGAR_DATOS
                    oTipo.DGridLista.Columns(2).Width = 3500
                    oTipo.Show 1
                    If CODIGO <> "" Then
                         Me.TxtAbr_Cliente.Text = Trim(CODIGO)
                         Me.TxtNom_Cliente.Text = Trim(descripcion)
'                         OptCliPend.SetFocus
                         CODIGO = "": descripcion = ""
                        ' CARGA_GRID
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
    End Select
    
End Sub
Private Sub TxtNum_Fardos_KeyPress(KeyAscii As Integer)
        Call SoloNumeros(TxtNum_Fardos, KeyAscii, False, 0, 3)
End Sub
Private Sub txtPartida_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmdBuscar.SetFocus
    End If
End Sub
Private Sub txtPartida_LostFocus()
    txtPartida.Text = Format(txtPartida.Text, "00000")
End Sub
Private Sub txtSer_OrdComp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        txtCod_OrdComp.SetFocus
            txtSer_OrdComp.Text = Format(Trim(txtSer_OrdComp.Text), "000")
        If Len(Trim(txtCod_OrdComp)) = 8 And Len(Trim(txtSer_OrdComp)) = 3 Then

        Call Busca_Facturas_Proformas(1)
        End If
    Else
        Call SoloNumeros(txtSer_OrdComp, KeyAscii, False, 0, 3)
    End If
End Sub
Private Sub txtSer_OrdComp_LostFocus()
    txtSer_OrdComp.Text = Format(Trim(txtSer_OrdComp.Text), "000")
End Sub
