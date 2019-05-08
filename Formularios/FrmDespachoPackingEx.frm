VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmDespachoPackingEx 
   Caption         =   "Despacho de Packing"
   ClientHeight    =   9210
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   9210
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChkTodo1 
      Caption         =   "Seleccionar Todo"
      Height          =   255
      Left            =   13320
      TabIndex        =   7
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton CmdAnadir 
      Height          =   495
      Left            =   3240
      Picture         =   "FrmDespachoPackingEx.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3270
      Width           =   1365
   End
   Begin VB.CommandButton CmdEliminar 
      Height          =   495
      Left            =   4680
      Picture         =   "FrmDespachoPackingEx.frx":05BE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14895
      Begin VB.TextBox Txt_NumPacking 
         Height          =   285
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox ChkTodo2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Seleccionar Todo"
         Height          =   255
         Left            =   13080
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   510
         Left            =   3960
         TabIndex        =   3
         Top             =   120
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   900
         Custom          =   $"FrmDespachoPackingEx.frx":0B8D
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FFFF&
         Caption         =   "N° Packing List"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Height          =   2355
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   14880
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
      Columns(1).Width=   1693
      Columns(1).Caption=   "N° Fardo"
      Columns(1).Name =   "Partida"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   3200
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "S° Rollo"
      Columns(2).Name =   "Almacen"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   4207
      Columns(3).Caption=   "Tela"
      Columns(3).Name =   "Movimiento"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   3784
      Columns(4).Caption=   "Color"
      Columns(4).Name =   "Secuencia_General"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   1640
      Columns(5).Caption=   "Kilos Rollo"
      Columns(5).Name =   "Secuencia"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(5).Locked=   -1  'True
      Columns(6).Width=   1561
      Columns(6).Caption=   "Metros2"
      Columns(6).Name =   "Tela"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(6).Locked=   -1  'True
      Columns(7).Width=   1535
      Columns(7).Caption=   "Peso Sin Tara"
      Columns(7).Name =   "Color"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(7).Locked=   -1  'True
      Columns(8).Width=   1852
      Columns(8).Caption=   "Id Rollo"
      Columns(8).Name =   "Cod_Rollo"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(8).Locked=   -1  'True
      Columns(9).Width=   1376
      Columns(9).Caption=   "PesoBruto"
      Columns(9).Name =   "Peso"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(9).Locked=   -1  'True
      Columns(10).Width=   1693
      Columns(10).Caption=   "Partida"
      Columns(10).Name=   "PesoNeto"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   2143
      Columns(11).Caption=   "N° Mov"
      Columns(11).Name=   "IdRolloUnico"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   2117
      Columns(12).Caption=   "N° Secuencia"
      Columns(12).Name=   "Num_Secuencia"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      _ExtentX        =   26247
      _ExtentY        =   4154
      _StockProps     =   79
      Caption         =   "Packing Disponibles del Almacén"
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
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   14880
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
      Columns(1).Width=   1693
      Columns(1).Caption=   "N° Fardo"
      Columns(1).Name =   "Partida"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   3200
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "S° Rollo"
      Columns(2).Name =   "Almacen"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   4207
      Columns(3).Caption=   "Tela"
      Columns(3).Name =   "Movimiento"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   3784
      Columns(4).Caption=   "Color"
      Columns(4).Name =   "Secuencia_General"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   1640
      Columns(5).Caption=   "Kilos Rollo"
      Columns(5).Name =   "Secuencia"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(5).Locked=   -1  'True
      Columns(6).Width=   1561
      Columns(6).Caption=   "Metros2"
      Columns(6).Name =   "Tela"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(6).Locked=   -1  'True
      Columns(7).Width=   1535
      Columns(7).Caption=   "Peso Sin Tara"
      Columns(7).Name =   "Color"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(7).Locked=   -1  'True
      Columns(8).Width=   1852
      Columns(8).Caption=   "Id Rollo"
      Columns(8).Name =   "Cod_Rollo"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(8).Locked=   -1  'True
      Columns(9).Width=   1376
      Columns(9).Caption=   "PesoBruto"
      Columns(9).Name =   "Peso"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(9).Locked=   -1  'True
      Columns(10).Width=   1693
      Columns(10).Caption=   "Partida"
      Columns(10).Name=   "PesoNeto"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   2143
      Columns(11).Caption=   "N° Mov"
      Columns(11).Name=   "IdRolloUnico"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   2117
      Columns(12).Caption=   "N° Secuencia"
      Columns(12).Name=   "Num_Secuencia"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      _ExtentX        =   26247
      _ExtentY        =   4154
      _StockProps     =   79
      Caption         =   "Rollos Disponibles Del Packing"
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
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid3 
      Height          =   2715
      Left            =   120
      TabIndex        =   10
      Top             =   6360
      Width           =   14880
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   10
      AllowColumnShrinking=   0   'False
      SelectTypeRow   =   1
      BackColorOdd    =   13434879
      RowHeight       =   423
      ExtraHeight     =   53
      Columns.Count   =   10
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Name =   "Check"
      Columns(0).Alignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Style=   2
      Columns(1).Width=   1931
      Columns(1).Caption=   "N° Fardo"
      Columns(1).Name =   "Partida"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   3200
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "S° Rollo"
      Columns(2).Name =   "Almacen"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   6959
      Columns(3).Caption=   "Tela"
      Columns(3).Name =   "Movimiento"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   6403
      Columns(4).Caption=   "Color"
      Columns(4).Name =   "Secuencia_General"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   1640
      Columns(5).Caption=   "Kilos Rollo"
      Columns(5).Name =   "Secuencia"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(5).Locked=   -1  'True
      Columns(6).Width=   1561
      Columns(6).Caption=   "Metros2"
      Columns(6).Name =   "Tela"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(6).Locked=   -1  'True
      Columns(7).Width=   1535
      Columns(7).Caption=   "Peso Sin Tara"
      Columns(7).Name =   "Color"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(7).Locked=   -1  'True
      Columns(8).Width=   1376
      Columns(8).Caption=   "PesoBruto"
      Columns(8).Name =   "Peso"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(8).Locked=   -1  'True
      Columns(9).Width=   1693
      Columns(9).Caption=   "Partida"
      Columns(9).Name =   "PesoNeto"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      _ExtentX        =   26247
      _ExtentY        =   4789
      _StockProps     =   79
      Caption         =   "Resumen Telas a Despachar"
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
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   1920
      Top             =   9120
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmDespachoPackingEx"
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

    vMessage = (MsgBox("¿Desea agregar estos rollos para su despacho?", vbQuestion + vbYesNo, sTit))
    If vMessage = vbYes Then
           ' Dim i As Integer
            For i = 0 To SSDBGrid1.Rows - 1
            SSDBGrid1.Bookmark = i
             If SSDBGrid1.Columns("check").Value = True Then
             
                    StrSQLMant = "Exec Usp_TI_PACKING_LIST_EXPO_DET_Despacho 'U','" & Txt_NumPacking.Text & "','" & SSDBGrid1.Columns("Num_Secuencia").Value & "'"
                    
                    filas = ExecuteSQL(cConnect, StrSQLMant)
                    
            End If
            Next
            
            Call carga_GRidRollosAlmacen("1", Me.Txt_NumPacking.Text)
            Call carga_GRidRollosAlmacen("2", Me.Txt_NumPacking.Text)
            Call carga_GRidRollosAlmacen("3", Me.Txt_NumPacking.Text)
    End If
Exit Sub
xerror:
errores err.Number
Call carga_GRidRollosAlmacen("1", Me.Txt_NumPacking.Text)
Call carga_GRidRollosAlmacen("2", Me.Txt_NumPacking.Text)
Call carga_GRidRollosAlmacen("3", Me.Txt_NumPacking.Text)
Exit Sub
End Sub

Private Sub CmdEliminar_Click()
On Error GoTo xerror:
Dim vMessage As Variant
Dim vResp As String, sTit As String
Dim filas As Integer
Dim StrSQLMant As String
 Dim i As Integer

    vMessage = (MsgBox("¿Desea agregar estos rollos para su despacho?", vbQuestion + vbYesNo, sTit))
    If vMessage = vbYes Then
           ' Dim i As Integer
            For i = 0 To SSDBGrid2.Rows - 1
            SSDBGrid2.Bookmark = i
             If SSDBGrid2.Columns("check").Value = True Then
             
                    StrSQLMant = "Exec Usp_TI_PACKING_LIST_EXPO_DET_Despacho 'D','" & Txt_NumPacking.Text & "','" & SSDBGrid2.Columns("Num_Secuencia").Value & "'"
                    
                    filas = ExecuteSQL(cConnect, StrSQLMant)
                    
            End If
            Next
            
            Call carga_GRidRollosAlmacen("1", Me.Txt_NumPacking.Text)
            Call carga_GRidRollosAlmacen("2", Me.Txt_NumPacking.Text)
            Call carga_GRidRollosAlmacen("3", Me.Txt_NumPacking.Text)
    End If
Exit Sub
xerror:
errores err.Number
Call carga_GRidRollosAlmacen("1", Me.Txt_NumPacking.Text)
Call carga_GRidRollosAlmacen("2", Me.Txt_NumPacking.Text)
Call carga_GRidRollosAlmacen("3", Me.Txt_NumPacking.Text)
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



Private Sub ListaPacking(ByVal sNroPacking As String)
On Error GoTo err
Dim SQL As String


SQL = "Exec Lista_Deta_PackingList_Despacho '" & sNroPacking & "'"

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(SQL, cConnect)

GridEX1.Columns("Secuencia_Fardo").Width = 800
GridEX1.Columns("Tela").Width = 2000
GridEX1.Columns("Color").Width = 2000
GridEX1.Columns("Kilos_rollos").Width = 800
GridEX1.Columns("Metros2").Width = 800
GridEX1.Columns("PesoSinTara").Width = 800
GridEX1.Columns("ID_Rollo_X_Anio").Width = 1200
GridEX1.Columns("PesoBruto").Width = 500
GridEX1.Columns("Cod_OrdTra").Width = 500
GridEX1.Columns("Num_MovStk").Width = 1200
GridEX1.Columns("Num_Secuencia").Width = 1200

GridEX1.Columns("Secuencia_Fardo").Caption = "N° Fardo"
GridEX1.Columns("ID_Rollo_X_Anio").Caption = "Id Rollo"
GridEX1.Columns("Cod_Ordtra").Caption = "Partida"
GridEX1.Columns("Num_MovStk").Caption = "N° Movimiento"
GridEX1.Columns("Num_Secuencia").Caption = "N°Secuencia" '
GridEX1.Columns("Secuencia_Rollo").Visible = False


GridEX1.Columns("Item").ColumnType = jgexCheckBox



Exit Sub
err:
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Partidas Enviadas Hacia Calidad"

End Sub

Sub carga_GRidRollosAlmacen(ByVal sOpcion As String, ByVal sNroPacking As String)
On Error GoTo fin

If sOpcion = "1" Then

    Set rsx1 = New ADODB.Recordset
        
    Dim rs_Prov1 As ADODB.Recordset
    rsx1.CursorLocation = adUseClient
    rsx1.CursorType = adOpenStatic
    rsx1.ActiveConnection = cConnect
    
    If rsx1.State <> 0 Then rsx1.Close
    rsx1.Open "Exec Lista_Deta_PackingList_Despacho '" & sOpcion & "','" & sNroPacking & "'"
    Set rs_Prov1 = rsx1.Clone
        If rsx1.RecordCount >= 0 Then
        
            Me.SSDBGrid1.Redraw = False
            SSDBGridSetGrid Me.SSDBGrid1
            ADODBToSSDBGridOC rs_Prov1, SSDBGrid1
            SSDBGrid1.ActiveRowStyleSet = "RowActive"
            SSDBGrid1.SelectTypeRow = ssSelectionTypeMultiSelectRange
            Me.SSDBGrid1.Visible = True
            Me.SSDBGrid1.Caption = "Packing Disponibles del Almacén - N° Registros=" & rsx1.RecordCount
            
    
        End If
   
End If

If sOpcion = "2" Then

    Set rsx2 = New ADODB.Recordset
    
    Dim rs_Prov2 As ADODB.Recordset
    rsx2.CursorLocation = adUseClient
    rsx2.CursorType = adOpenStatic
    rsx2.ActiveConnection = cConnect
    
    If rsx2.State <> 0 Then rsx2.Close
    rsx2.Open "Exec Lista_Deta_PackingList_Despacho '" & sOpcion & "','" & sNroPacking & "'"
    Set rs_Prov2 = rsx2.Clone
        If rsx2.RecordCount >= 0 Then
        
            Me.SSDBGrid2.Redraw = False
            SSDBGridSetGrid Me.SSDBGrid2
            ADODBToSSDBGridOC rs_Prov2, SSDBGrid2
            SSDBGrid2.ActiveRowStyleSet = "RowActive"
            SSDBGrid2.SelectTypeRow = ssSelectionTypeMultiSelectRange
            Me.SSDBGrid2.Visible = True
            Me.SSDBGrid2.Caption = "Rollos Disponibles Del Packing - N° Registros=" & rsx2.RecordCount
    
        End If

End If
   
If sOpcion = "3" Then

    Set rsx3 = New ADODB.Recordset
    
    Dim rs_Prov3 As ADODB.Recordset
    rsx3.CursorLocation = adUseClient
    rsx3.CursorType = adOpenStatic
    rsx3.ActiveConnection = cConnect
    
    If rsx3.State <> 0 Then rsx3.Close
    rsx3.Open "Exec Lista_Deta_PackingList_Despacho '" & sOpcion & "','" & sNroPacking & "'"
    Set rs_Prov3 = rsx3.Clone
        If rsx3.RecordCount >= 0 Then
        
            Me.SSDBGrid3.Redraw = False
            SSDBGridSetGrid Me.SSDBGrid3
            ADODBToSSDBGridOC rs_Prov3, SSDBGrid3
            SSDBGrid3.ActiveRowStyleSet = "RowActive"
            SSDBGrid3.SelectTypeRow = ssSelectionTypeMultiSelectRange
            Me.SSDBGrid3.Visible = True
            Me.SSDBGrid3.Caption = "Resumen Telas a Despachar - N° Registros=" & rsx3.RecordCount
    
        End If

End If
Exit Sub
fin:
MsgBox "Inconvenientes para mostrar los packing " + err.Description, vbInformation + vbOKOnly, "Mensaje"
   
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


Private Sub Form_Load()
'FunctButt2.FunctionsUser = get_botones1(Me, vper, vemp1, Me.Name)

End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
 Select Case ActionName
        Case "BUSCAR"
            Call carga_GRidRollosAlmacen("1", Me.Txt_NumPacking.Text)
            Call carga_GRidRollosAlmacen("2", Me.Txt_NumPacking.Text)
            Call carga_GRidRollosAlmacen("3", Me.Txt_NumPacking.Text)
                Exit Sub
                
        Case "DESPACHAR"
            
            Call Grabar
            Call carga_GRidRollosAlmacen("1", Me.Txt_NumPacking.Text)
            Call carga_GRidRollosAlmacen("2", Me.Txt_NumPacking.Text)
            Call carga_GRidRollosAlmacen("3", Me.Txt_NumPacking.Text)
            
            
            
        Case "SALIR"
            Unload Me
    End Select
End Sub

Private Sub Txt_NumPacking_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_NumPacking.Text = Format(Trim(Txt_NumPacking.Text), "00000000")
        'Command1.SetFocus
        Call carga_GRidRollosAlmacen("1", Me.Txt_NumPacking.Text)
        Call carga_GRidRollosAlmacen("2", Me.Txt_NumPacking.Text)
        Call carga_GRidRollosAlmacen("3", Me.Txt_NumPacking.Text)
        
        Exit Sub
    Else
        Call SoloNumeros(Txt_NumPacking, KeyAscii, False, 0, 6)
    End If
End Sub

Private Sub Txt_NumPacking_LostFocus()
    Txt_NumPacking.Text = Format(Trim(Txt_NumPacking.Text), "00000000")
End Sub
Sub Grabar()
    On Error GoTo errGrabar

    bGrabando = True


'===========SE AGREGO LA SERIE Y NUMERO DE SERVICIO=================
    strsql = "EXEC USP_DESPACHA_PACKING_LIST '" & Txt_NumPacking.Text & _
                             "','" & vusu & "'"
                                                       
                            

    Call ExecuteSQL(cConnect, strsql)
    vOk = True
    bGrabando = False
    
    MsgBox "Se realizo el despacho satisfactoriamente", vbInformation, "Información"
    'Unload Me
    Exit Sub
    
errGrabar:
    bGrabando = False
    vOk = False
    ErrorHandler err, "Grabar"
End Sub




