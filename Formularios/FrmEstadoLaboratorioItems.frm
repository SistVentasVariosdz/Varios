VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmEstadoLaboratorioItems 
   Caption         =   "ESTADO LABORATORIO DE QUIMICOS"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_CodigoItem 
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txt_desitem 
      Height          =   285
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   5175
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&BUSCAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   9480
      TabIndex        =   3
      Top             =   0
      Width           =   1425
   End
   Begin VB.CheckBox chkRechazado 
      BackColor       =   &H80000004&
      Caption         =   "RECHAZADO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   180
      Left            =   9480
      TabIndex        =   2
      Top             =   600
      Width           =   1305
   End
   Begin VB.CheckBox chkAprobado 
      BackColor       =   &H80000004&
      Caption         =   "APROBADO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   180
      Left            =   7800
      TabIndex        =   1
      Top             =   600
      Width           =   1245
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "&GUARDAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   9480
      TabIndex        =   0
      Top             =   8520
      Width           =   1425
   End
   Begin GridEX20.GridEX grxRegistros 
      Height          =   7575
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   13361
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GridLineStyle   =   2
      RowHeight       =   20
      GroupByBoxVisible=   0   'False
      BackColorGBBox  =   8421504
      HeaderFontName  =   "Verdana"
      HeaderFontBold  =   -1  'True
      HeaderFontSize  =   6.75
      HeaderFontWeight=   700
      FontName        =   "Tahoma"
      ColumnHeaderHeight=   270
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "FrmEstadoLaboratorioItems.frx":0000
      Column(2)       =   "FrmEstadoLaboratorioItems.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "FrmEstadoLaboratorioItems.frx":016C
      FormatStyle(2)  =   "FrmEstadoLaboratorioItems.frx":0294
      FormatStyle(3)  =   "FrmEstadoLaboratorioItems.frx":0344
      FormatStyle(4)  =   "FrmEstadoLaboratorioItems.frx":03F8
      FormatStyle(5)  =   "FrmEstadoLaboratorioItems.frx":04D0
      FormatStyle(6)  =   "FrmEstadoLaboratorioItems.frx":0588
      ImageCount      =   0
      PrinterProperties=   "FrmEstadoLaboratorioItems.frx":0668
   End
   Begin VB.Label Label1 
      Caption         =   "CODIGO"
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
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "DESCRIPCION"
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
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   2520
      Top             =   8640
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmEstadoLaboratorioItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkAprobado_Click()
     If grxRegistros.RowCount = 0 Then Exit Sub
    
    Dim RS As New ADODB.Recordset
    Dim Valor As Boolean
    Dim I As Long

    If chkAprobado.Value = Checked Then
        Valor = True
    Else
        Valor = False
    End If

    grxRegistros.Update
    Set RS = grxRegistros.ADORecordset
    RS.MoveFirst
    Do While Not RS.EOF
        RS("APROBADO") = Valor
        If Valor = True Then
            RS("RECHAZADO") = Not Valor
        End If
        
        RS.MoveNext
    Loop
   
    RS.MoveFirst
    RS.Update
    Set grxRegistros.ADORecordset = RS
End Sub

Private Sub chkRechazado_Click()
     If grxRegistros.RowCount = 0 Then Exit Sub
    
    Dim RS As New ADODB.Recordset
    Dim Valor As Boolean
    Dim I As Long

    If chkRechazado.Value = Checked Then
        Valor = True
    Else
        Valor = False
    End If

    grxRegistros.Update
    Set RS = grxRegistros.ADORecordset
    RS.MoveFirst
    Do While Not RS.EOF
        RS("RECHAZADO") = Valor
        If Valor = True Then
            RS("APROBADO") = Not Valor
        End If
        RS.MoveNext
    Loop
   
    RS.MoveFirst
    RS.Update
    Set grxRegistros.ADORecordset = RS
    
End Sub

Private Sub cmdBuscar_Click()
Call cargaGrillaItems
End Sub

Private Sub cargaGrillaItems()
   Dim RSX As New ADODB.Recordset
   
   STRSQL = " EXEC LG_MUESTRA_ITEMS_STATUS_LABORATORIO '" & txt_CodigoItem.Text & "','" & Trim(txt_desitem.Text) & "'"
   Set grxRegistros.ADORecordset = CargarRecordSetDesconectado(STRSQL, cConnect)

End Sub

Private Sub Command1_Click()

End Sub

Private Sub CmdGuardar_Click()
If grxRegistros.RowCount <= 0 Then Exit Sub
If MsgBox("Esta Seguro  Realizar los Cambios", vbYesNo, "Confirmar Cambios") = vbYes Then
  Call SALVA_ESTADO_LABORATORIO
End If

End Sub

'''*******************EVENTOS POR COLUMNA **********************************************************
Private Sub grxRegistros_AfterColEdit(ByVal ColIndex As Integer)
  AfterColEdit_DETALLE_FACTURA (ColIndex)
End Sub

Sub AfterColEdit_DETALLE_FACTURA(ByVal ColIndex As Integer)

Dim sSQL As String
On Error GoTo Error_Handler

Dim oGroup As GridEX20.JSGroup
Select Case ColIndex
    
  Case Is = grxRegistros.Columns("APROBADO").Index
     If grxRegistros.Value(grxRegistros.Columns("rechazado").Index) = True Then
         grxRegistros.Value(grxRegistros.Columns("rechazado").Index) = False
     End If
  Case Is = grxRegistros.Columns("rechazado").Index
     If grxRegistros.Value(grxRegistros.Columns("aprobado").Index) = True Then
         grxRegistros.Value(grxRegistros.Columns("aprobado").Index) = False
     End If
     
  End Select
Exit Sub

Resume
Error_Handler:
errores Err.Number
End Sub
''''******************************HABILITA LA EDICION SOLO DE ALGUNAS COLUMNAS LAS TIENEN CANCEL=FALE***********************
Private Sub grxRegistros_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
  Select Case ColIndex
    Case Is = grxRegistros.Columns("aprobado").Index
      Cancel = False
    Case Is = grxRegistros.Columns("rechazado").Index
      Cancel = False
    Case Else
      Cancel = True
  End Select
End Sub

Private Sub grxRegistros_Click()

Dim ColIndex As Long

If grxRegistros.RowCount > 0 Then
    ColIndex = grxRegistros.Col
    If UCase(grxRegistros.Columns(ColIndex).Key) = "APROBADO" Or UCase(grxRegistros.Columns(ColIndex).Key) = "RECHAZADO" Then
        SendKeys "{ENTER}"
    End If
End If
        
End Sub
Private Sub SALVA_ESTADO_LABORATORIO()

Dim STRSQL As String
Dim RSAUX As ADODB.Recordset
Dim ESTADO As String
Dim I As Integer
On Error GoTo Error_Handler
I = 1
grxRegistros.Update
Set RSAUX = grxRegistros.ADORecordset
RSAUX.MoveFirst
Do While I <= RSAUX.RecordCount
 
        ESTADO = "0"
        If RSAUX("APROBADO") = True Then
          ESTADO = "1"
        End If
        If RSAUX("RECHAZADO") = True Then
          ESTADO = "2"
        End If
        STRSQL = "LG_SALVA_DATOS_ITEMS_ESTADO_LABORATORIO '" & RSAUX("COD_ITEM") & "','" & RSAUX("LOTE") & "','" & ESTADO & "'"
        Call ExecuteSQL(cConnect, STRSQL)
       RSAUX.MoveNext
I = I + 1
Loop

Call MsgBox("Cambios Guardados con exito", vbExclamation, "Mensaje")
Call cargaGrillaItems
Exit Sub

Resume
Error_Handler:
errores Err.Number
End Sub
Private Sub txt_CodigoItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmdBuscar.SetFocus
    End If
End Sub
Private Sub txt_desitem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdBuscar.SetFocus
    End If
End Sub


