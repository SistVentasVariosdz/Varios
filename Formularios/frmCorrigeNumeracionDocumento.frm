VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Begin VB.Form frmCorrigeNumeracionDocumento 
   Caption         =   "CAMBIO DE CORRELATIVO DOCUMENTOS"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&CANCELAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3120
      Width           =   1005
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&GUARDAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3120
      Width           =   1005
   End
   Begin VB.TextBox txtNro_DocumNuevo 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   4560
      MaxLength       =   8
      TabIndex        =   9
      Top             =   2415
      Width           =   2020
   End
   Begin VB.TextBox txtNum_Docum_Actual 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   8
      Top             =   2415
      Width           =   2020
   End
   Begin VB.ComboBox cmdSerie 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ComboBox cmdDocumento 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Width           =   5415
   End
   Begin VB.ComboBox cmdCaja 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   5415
   End
   Begin VB.ComboBox cmdTienda 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   5415
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   840
      Top             =   3240
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "NRO NUEVO:"
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
      Top             =   2520
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "NRO ACTUAL:"
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
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   1050
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "SERIE:"
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
      Left            =   480
      TabIndex        =   7
      Top             =   1920
      Width           =   510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "DOCUMENTO:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CAJA:"
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
      Left            =   660
      TabIndex        =   3
      Top             =   840
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "TIENDA:"
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
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   645
   End
End
Attribute VB_Name = "frmCorrigeNumeracionDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancelar_Click()
Unload Me

End Sub

Private Sub cmdDocumento_Click()
  Call FillSerie
  Call buscaNumeroActual
End Sub
Private Sub cmdCaja_Click()
  Call FillDocumento
  Call FillSerie
  Call buscaNumeroActual
End Sub

Private Sub GuardaNumeroNuevo()
  If txtNro_DocumNuevo.Text = "" Then
   Call MsgBox("Ingrese un Numero Valiado", vbCritical, "Mensaje")
    Exit Sub
  End If
  If Len(txtNro_DocumNuevo.Text) <> 8 Then
   Call MsgBox("Ingrese un Numero Valiado", vbCritical, "Mensaje")
    Exit Sub
  End If

If MsgBox("¡¡¡Esta apunto de Cambiar el numero Correlativo del documento de venta!!!:" & Chr(13) & Chr(10) & ":::::> " & Trim(Right(cmdDocumento, Len(cmdDocumento) - 2)) & Chr(13) & Chr(10) & "¿Son los datos correctos?", vbYesNo, "CONFIRMAR") = vbYes Then
    strSQL = "CN_VENTAS_CAJAS_CORRIGE_CORRELATIVO  '" & Left(cmdTienda, 3) & "','" & Left(cmdCaja, 2) & "','" & Left(cmdDocumento, 2) & "','" & Trim(cmdSerie) & "','" & Trim(txtNro_DocumNuevo.Text) & "'"
    Call ExecuteCommandSQL(cConnect, strSQL)
    Call MsgBox("los Cambios se Realizaron con exito", vbInformation, "Mensaje")
    Call buscaNumeroActual
End If

End Sub

Private Sub cmdGuardar_Click()
Call GuardaNumeroNuevo
End Sub

Private Sub cmdSerie_Click()
Call buscaNumeroActual
End Sub
Private Sub buscaNumeroActual()
On Error GoTo Fin
    Dim strSQL As String
    strSQL = "SELECT COR_NUMACTU FROM CN_VENTAS_CAJAS_DOCUMENTOS WHERE COD_TIENDA = '" & Left(cmdTienda, 3) & "' and cod_caja=" & Left(cmdCaja, 2) & " and COD_TIPDOC='" & Left(cmdDocumento, 2) & "'"
    txtNum_Docum_Actual.Text = DevuelveCampo(strSQL, cConnect)
    txtNro_DocumNuevo.Text = txtNum_Docum_Actual.Text
Exit Sub
Fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, sTit
End Sub
Private Sub cmdTienda_Click()
  Call FillCaja
  Call FillDocumento
  Call FillSerie
  Call buscaNumeroActual
End Sub

Private Sub Form_Load()
  Call FillTienda
  Call FillCaja
  Call FillDocumento
  Call FillSerie
  Call buscaNumeroActual
  
  txtNum_Docum_Actual.Text = Format(txtNum_Docum_Actual, "00000000")
  txtNro_DocumNuevo.Text = Format(txtNro_DocumNuevo, "00000000")
  
End Sub
Private Sub FillTienda()
On Error GoTo Fin
Dim sTit As String
Dim rstAux As New ADODB.Recordset

    sTit = "Cargar Tiendas"
    strSQL = "SELECT COD_TIENDA,DES_TIENDA FROM CN_VENTAS_TIENDAS"
    Set rstAux = CargarRecordSetDesconectado(strSQL, cConnect)
    cmdTienda.Clear
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
            cmdTienda.AddItem !COD_TIENDA & " " & !DES_TIENDA
            .MoveNext
        Loop
        .Close
    End With
    If cmdTienda.ListCount > 0 Then cmdTienda.ListIndex = 0
    Set rstAux = Nothing
Exit Sub
Fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, sTit
End Sub
Private Sub FillCaja()
On Error GoTo Fin
Dim sTit As String
Dim rstAux As New ADODB.Recordset
    
    sTit = "Cargar Cajas"
    strSQL = "SELECT  COD_CAJA,DESCRIPCION=  'CAJA'+ COD_CAJA  FROM CN_VENTAS_CAJAS WHERE COD_TIENDA = '" & Left(cmdTienda, 3) & "'"
    Set rstAux = CargarRecordSetDesconectado(strSQL, cConnect)
    cmdCaja.Clear
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
            cmdCaja.AddItem !COD_CAJA & " " & !Descripcion
            .MoveNext
        Loop
        .Close
    End With
    If cmdCaja.ListCount > 0 Then cmdCaja.ListIndex = 0
    Set rstAux = Nothing
Exit Sub
Fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, sTit
End Sub
Private Sub FillDocumento()
On Error GoTo Fin
Dim sTit As String
Dim rstAux As New ADODB.Recordset
    
    sTit = "Cargar Documentos"
    strSQL = "SELECT  A.COD_TIPDOC,DES_TIPDOC FROM CN_VENTAS_CAJAS_DOCUMENTOS A INNER JOIN  CN_TIPOSDOCUM B ON A.COD_TIPDOC=B.COD_TIPDOC WHERE COD_TIENDA = '" & Left(cmdTienda, 3) & "' and cod_caja=" & Left(cmdCaja, 2) & ""
    
    Set rstAux = CargarRecordSetDesconectado(strSQL, cConnect)
    cmdDocumento.Clear
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
            cmdDocumento.AddItem !Cod_TipDoc & " " & !DES_TIPDOC
            .MoveNext
        Loop
        .Close
    End With
    If cmdDocumento.ListCount > 0 Then cmdDocumento.ListIndex = 0
    Set rstAux = Nothing
Exit Sub
Fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, sTit
End Sub
Private Sub FillSerie()
On Error GoTo Fin
Dim sTit As String
Dim rstAux As New ADODB.Recordset
    
    sTit = "Cargar Serie"
    strSQL = "SELECT  COR_DOCSERIE  FROM CN_VENTAS_CAJAS_DOCUMENTOS WHERE COD_TIENDA = '" & Left(cmdTienda, 3) & "' and cod_caja=" & Left(cmdCaja, 2) & " and COD_TIPDOC='" & Left(cmdDocumento, 2) & "'"
    
    Set rstAux = CargarRecordSetDesconectado(strSQL, cConnect)
    cmdSerie.Clear
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
            cmdSerie.AddItem !COR_DOCSERIE
            .MoveNext
        Loop
        .Close
    End With
    If cmdSerie.ListCount > 0 Then cmdSerie.ListIndex = 0
    Set rstAux = Nothing
Exit Sub
Fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, sTit
End Sub
Private Sub txtNro_DocumNuevo_LostFocus()
  txtNro_DocumNuevo.Text = Format(txtNro_DocumNuevo, "00000000")
End Sub
