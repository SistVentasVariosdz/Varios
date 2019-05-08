VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmCierraCajas 
   Caption         =   "CIERRES DE CAJAS "
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   13230
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCajaActual 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   360
      Width           =   3375
   End
   Begin VB.CommandButton cmdDeshacerCierre 
      Caption         =   "DESHACER CIERRE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   5
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   4
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "CERRAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   3
      Top             =   6480
      Width           =   1455
   End
   Begin VB.TextBox txtTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "C I E R R E   D E   D I A   V E N T A"
      Top             =   0
      Width           =   13335
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   13215
      Begin GridEX20.GridEX grxDatos 
         Height          =   5235
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   9234
         Version         =   "2.0"
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         GridLineStyle   =   2
         HideSelection   =   2
         MethodHoldFields=   -1  'True
         GroupByBoxInfoText=   "Arrastra la cabecera de una columna aquí para agruparlo por esa misma columna"
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         HeaderFontName  =   "Verdana"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   6.75
         HeaderFontWeight=   700
         FontName        =   "Tahoma"
         ColumnHeaderHeight=   270
         IntProp1        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmCierreCajas.frx":0000
         Column(2)       =   "frmCierreCajas.frx":00C8
         FormatStylesCount=   9
         FormatStyle(1)  =   "frmCierreCajas.frx":016C
         FormatStyle(2)  =   "frmCierreCajas.frx":0294
         FormatStyle(3)  =   "frmCierreCajas.frx":0344
         FormatStyle(4)  =   "frmCierreCajas.frx":03F8
         FormatStyle(5)  =   "frmCierreCajas.frx":04D0
         FormatStyle(6)  =   "frmCierreCajas.frx":0588
         FormatStyle(7)  =   "frmCierreCajas.frx":0668
         FormatStyle(8)  =   "frmCierreCajas.frx":06F8
         FormatStyle(9)  =   "frmCierreCajas.frx":0830
         ImageCount      =   0
         PrinterProperties=   "frmCierreCajas.frx":0944
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   720
      Top             =   6360
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "CAJA ACTUAL"
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
      Left            =   8640
      TabIndex        =   7
      Top             =   480
      Width           =   1065
   End
End
Attribute VB_Name = "FrmCierraCajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strSQL As String
Private cod_fabrica_default As String
Private cod_tienda_default  As String
Private cod_caja_default As String
Private D_ULTI_CIER_DEFAULT As Date
Public PER_DESHACECIERRE As Byte
Private Sub cmdCancelar_Click()
Unload Me
End Sub
Private Sub cmdDeshacerCierre_Click()
    If validaDeshaceCierre = True Then
        If MsgBox("¡¡¡ADVERTENCIA!!! " & Chr(13) & " ¿ Esta seguro(a) de Deshacer el cierre de La Caja actual? ", vbInformation + vbYesNo, "ADVERTENCIA") = vbYes Then
            Call deshacecierreCaja
            Call muestraCajaCierre
        End If
    End If
End Sub
Private Sub deshacecierreCaja()
On Error GoTo fin
Dim I As Long

    strSQL = " CN_VENTAS_CAJAS_DESHACE_CIERRE  '" & grxDatos.Value(grxDatos.Columns("cod_fabrica_default").Index) & _
    "','" & grxDatos.Value(grxDatos.Columns("cod_tienda_default").Index) & _
    "','" & grxDatos.Value(grxDatos.Columns("cod_caja_default").Index) & _
    "','" & ComputerName & _
    "','" & usuario_windows & "','" & vusu & "','" & Now & "'"
    
    I = ExecuteSQL(cConnect, strSQL)
    
    Call MsgBox("Se deshizo con exito el cierre de la caja nro '" & grxDatos.Value(grxDatos.Columns("cod_caja_default").Index) & "' ", vbInformation + vbOKOnly, "Mensaje")

Exit Sub
fin:
MsgBox Err.Description, vbCritical + vbOKOnly, "Advertencia"

End Sub
Private Function validaDeshaceCierre() As Boolean
On Error GoTo fin

    validaDeshaceCierre = True

    If UCase(grxDatos.Value(grxDatos.Columns("flg_status_caja").Index)) = "ABIERTO" Then
        Call MsgBox("La Caja ya se encuentra Aperturada...no procede", vbCritical + vbOKOnly, "Advertencia")
        validaDeshaceCierre = False
        Exit Function
    End If

'    If cod_fabrica_default <> grxDatos.Value(grxDatos.Columns("cod_fabrica").Index) Then
'        Call MsgBox("Ud no esta autorizado para deshacer Cierre de esta caja", vbCritical + vbOKOnly, "Advertencia")
'        validaDeshaceCierre = False
'        Exit Function
'    End If
'
'    If cod_tienda_default <> grxDatos.Value(grxDatos.Columns("cod_tienda").Index) Then
'       Call MsgBox("Ud no esta autorizado para deshacer Cierre de esta caja", vbCritical + vbOKOnly, "Advertencia")
'       validaDeshaceCierre = False
'       Exit Function
'    End If
'
'    If cod_caja_default <> grxDatos.Value(grxDatos.Columns("cod_Caja").Index) Then
'      Call MsgBox("Ud no esta autorizado para deshacer Cierre de esta caja", vbCritical + vbOKOnly, "Advertencia")
'      validaDeshaceCierre = False
'      Exit Function
'    End If


Exit Function
fin:
MsgBox Err.Description, vbCritical + vbOKOnly, "Advertencia"

End Function

Private Sub Form_Load()
    cmdDeshacerCierre.Enabled = False
    Call muestraCajaCierre
    Call PERMISO_DESHACECIERRE
End Sub
Private Sub PERMISO_DESHACECIERRE()
On Error GoTo fin
Dim tit As String

PER_DESHACECIERRE = 0
strSQL = "SELECT ISNULL(COUNT(*),0) FROM SS_PERMISOS_USUARIOS WHERE COD_USUARIO = '" & vusu & "' AND COD_PERMISO  =  'DESHACE_CIERRE_CAJA' "
PER_DESHACECIERRE = DevuelveCampo(strSQL, cConnect)


Exit Sub
fin:
MsgBox Err.Description, vbCritical + vbOKOnly, tit

End Sub

Private Sub muestraCajaCierre()
On Error GoTo fin
Dim tit As String

    strSQL = "SM_MUESTRA_CN_VENTAS_CAJAS_CIERRE '" & ComputerName & "','" & usuario_windows & "'"
    Set grxDatos.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    
    cod_fabrica_default = grxDatos.Value(grxDatos.Columns("cod_fabrica_default").Index)
    cod_tienda_default = grxDatos.Value(grxDatos.Columns("Cod_tienda_default").Index)
    cod_caja_default = grxDatos.Value(grxDatos.Columns("cod_caja_default").Index)
    D_ULTI_CIER_DEFAULT = grxDatos.Value(grxDatos.Columns("D_ULTI_CIER_DEFAULT").Index)
    txtCajaActual.Text = "Caja Actual : " & cod_caja_default
    Call configuraGrilla

Exit Sub
fin:
MsgBox Err.Description, vbCritical + vbOKOnly, tit
End Sub

Private Sub configuraGrilla()
On Error GoTo fin

    With grxDatos
        
        For C = 1 To .Columns.Count
            .Columns(C).HeaderAlignment = jgexAlignCenter
            .Columns(C).TextAlignment = jgexAlignLeft
            .Columns(C).Visible = False
        Next C
        
        With .Columns("CAJ_CODIGO")
             .Visible = True
             .Width = 1500
             .Caption = "CAJA"
             .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("C_SERI_MAQU")
             .Visible = True
             .Width = 2500
             .Caption = "SM"
             .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("C_HOST_TRAB")
             .Visible = True
             .Width = 2500
             .Caption = "PC NOMBRE"
             .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("D_ULTI_APER")
             .Visible = True
             .Width = 1500
             .Caption = "ULT APER"
             .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("D_ULTI_CIER")
             .Visible = True
             .Width = 1500
             .Caption = "UTL CIERRE"
             .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("FLG_STATUS_CAJA")
             .Visible = True
             .Width = 1500
             .Caption = "ESTADO"
             .TextAlignment = jgexAlignLeft
        End With
     
     
     End With
     
    Dim fmtCon  As JSFmtCondition
    Set fmtCon = grxDatos.FmtConditions.Add(grxDatos.Columns("CAJA_ACTUAL").Index, jgexEqual, "S")
    fmtCon.FormatStyle.BackColor = &H80FFFF
     
Exit Sub
fin:
MsgBox Err.descriptio, vbCritical + vbOKOnly, "Advertencia"
End Sub
Private Sub cmdCerrar_Click()
If validacionCierreCaja = True Then
    If MsgBox("¡¡¡ADVERTENCIA!!! " & Chr(13) & "Al cerrar la caja no podra efectuar ventas hasta el siguiente dia" & Chr(13) & " ¿ Esta seguro(a) de cerrar La Caja actual? ", vbCritical + vbYesNo, "ADVERTENCIA") = vbYes Then
        Call cierrecaja
        Call muestraCajaCierre
    End If
End If
End Sub
Private Sub cierrecaja()
On Error GoTo fin
Dim I As Long

strSQL = " CN_VENTAS_CAJA_CIERRE '" & grxDatos.Value(grxDatos.Columns("cod_fabrica_default").Index) & _
"','" & grxDatos.Value(grxDatos.Columns("cod_tienda_default").Index) & _
"','" & grxDatos.Value(grxDatos.Columns("cod_caja_default").Index) & _
"','" & ComputerName & _
"','" & usuario_windows & "','" & vusu & "'"

I = ExecuteCommandSQL(cConnect, strSQL)

MsgBox "Se cerro la caja con exito", vbCritical + vbExclamation, "Advertencia"

Exit Sub
fin:
MsgBox Err.Description, vbCritical + vbOKOnly, "Advertencia"
End Sub
Private Function validacionCierreCaja() As Boolean
On Error GoTo fin
    validacionCierreCaja = True

    If UCase(grxDatos.Value(grxDatos.Columns("flg_status_caja").Index)) = "CERRADO" Then
        Call MsgBox("La Caja ya se encuentra cerrada...no procede", vbCritical + vbOKOnly, "Advertencia")
        validacionCierreCaja = False
        Exit Function
    End If

    If cod_fabrica_default <> grxDatos.Value(grxDatos.Columns("cod_fabrica").Index) Then
        Call MsgBox("Ud no esta autorizado para cerrar esta caja", vbCritical + vbOKOnly, "Advertencia")
        validacionCierreCaja = False
        Exit Function
    End If
    
    If cod_tienda_default <> grxDatos.Value(grxDatos.Columns("cod_tienda").Index) Then
       Call MsgBox("Ud no esta autorizado para cerrar esta caja", vbCritical + vbOKOnly, "Advertencia")
       validacionCierreCaja = False
       Exit Function
    End If
    
    If cod_caja_default <> grxDatos.Value(grxDatos.Columns("cod_Caja").Index) Then
      Call MsgBox("Ud no esta autorizado para cerrar esta caja", vbCritical + vbOKOnly, "Advertencia")
      validacionCierreCaja = False
      Exit Function
    End If

Exit Function
fin:
MsgBox Err.Description, vbCritical + vbOKOnly, "Advertencia"
End Function
Private Sub grxDatos_SelectionChange()
cmdDeshacerCierre.Enabled = False
If Format(Now, "dd/mm/yyyy") = Format(grxDatos.Value(grxDatos.Columns("D_ULTI_CIer").Index), "dd/mm/yyyy") And grxDatos.Value(grxDatos.Columns("flg_status_caja").Index) = "CERRADO" Then
  If PER_DESHACECIERRE = 1 Then
    cmdDeshacerCierre.Enabled = True
  End If
  
End If

End Sub
