VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMantFamTela 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Familia de Tela"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Componente Hilado"
   Begin VB.CommandButton cmdLast 
      Height          =   495
      Left            =   1560
      Picture         =   "frmMantFamTela.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Ultimo"
      Top             =   5160
      Width           =   495
   End
   Begin VB.CommandButton cmdFirst 
      Height          =   495
      Left            =   120
      Picture         =   "frmMantFamTela.frx":0172
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Primero"
      Top             =   5160
      Width           =   495
   End
   Begin VB.CommandButton cmdNext 
      Height          =   495
      Left            =   1080
      Picture         =   "frmMantFamTela.frx":02E4
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Siguiente"
      Top             =   5160
      Width           =   495
   End
   Begin VB.CommandButton cmdPrevious 
      Height          =   495
      Left            =   600
      Picture         =   "frmMantFamTela.frx":0456
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Anterior"
      Top             =   5160
      Width           =   495
   End
   Begin VB.Frame Fradetalle 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   120
      TabIndex        =   5
      Tag             =   "Detail"
      Top             =   3405
      Width           =   6300
      Begin VB.ComboBox cboCod_TipFamTela 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1120
         Width           =   2115
      End
      Begin VB.TextBox txtDes_FamTela 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         MaxLength       =   50
         TabIndex        =   1
         Top             =   360
         Width           =   4080
      End
      Begin VB.TextBox txtcod_ctacont 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         MaxLength       =   14
         TabIndex        =   2
         Top             =   735
         Width           =   2115
      End
      Begin VB.TextBox txtCod_FamTela 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   0
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Familia"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Cta Contable:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   210
         TabIndex        =   7
         Tag             =   "Mat. Prima :"
         Top             =   810
         Width           =   960
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Familia :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   210
         TabIndex        =   6
         Tag             =   "Hilado :"
         Top             =   435
         Width           =   570
      End
   End
   Begin VB.Frame Fralista 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Tag             =   "List"
      Top             =   75
      Width           =   6315
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2775
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   4895
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "cod_famTela"
            Caption         =   "Código"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "des_famTela"
            Caption         =   "Descripción"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cod_ctacont"
            Caption         =   "Cta.Contable"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Cod_TipFamTela"
            Caption         =   "Tipo Familia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1620.284
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1289.764
            EndProperty
         EndProperty
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2160
      TabIndex        =   4
      Top             =   5040
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantFamTela.frx":05C8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   6000
      Top             =   5040
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmMantFamTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public CODIGO, Descripcion As String
Dim sTipo As String
Dim StrSQL As String
Dim Rs_Carga As New ADODB.Recordset

Private Sub cmdFirst_Click()
    If Not Rs_Carga.BOF Then
        Rs_Carga.MoveFirst
    End If
End Sub
Private Sub cmdLast_Click()
    If Not Rs_Carga.EOF Then
        Rs_Carga.MoveLast
    End If
End Sub
Private Sub cmdNext_Click()
    If Not Rs_Carga.EOF Then
        Rs_Carga.MoveNext
    End If
End Sub
Private Sub cmdPrevious_Click()
    If Not Rs_Carga.BOF Then
        Rs_Carga.MovePrevious
    End If
End Sub
Private Sub Cargar_Datos()
    On Error GoTo Cargar_DatosErr
    StrSQL = "EXEC UP_SEL_FAMTELA"
    Set Rs_Carga = Nothing
    Rs_Carga.ActiveConnection = cConnect
    Rs_Carga.CursorType = adOpenStatic
    Rs_Carga.CursorLocation = adUseClient
    Rs_Carga.LockType = adLockReadOnly
    Rs_Carga.Open StrSQL
    Set DGridLista.DataSource = Rs_Carga
    DGridLista_RowColChange 0, 0
    If Rs_Carga.RecordCount > 0 Then
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
    Else
        LIMPIAR_DATOS
        DESHABILITA_DATOS
        HabilitaMant Me.MantFunc1, "ADICIONAR"
    End If
    Exit Sub
Cargar_DatosErr:
    Set Rs_Carga = Nothing
    ErrorHandler err, "Cargar_Datos"
End Sub


Private Sub CARGA_COMBOS()
    'Llena el combo con los tipos de Familia
    StrSQL = "SELECT Des_TipFamTela + SPACE(100) + Cod_TipFamTela FROM TX_TIPFAM"
    Call LlenaCombo(cboCod_TipFamTela, StrSQL, cConnect)
End Sub
Private Sub Form_Load()
    Call FormSet(Me)
    FormateaGrid Me.DGridLista
    Call Cargar_Datos
    Call CARGA_COMBOS
    'MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub

Private Sub SALVAR_DATOS()
 Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Dim StrSQL As String
    
    Con.ConnectionString = cConnect
    Con.Open
    
        Con.BeginTrans

        StrSQL = "EXEC UP_MAN_FAMTELA '" & _
        sTipo & "','" & _
        txtCod_FamTela.Text & "','" & _
        txtDes_FamTela.Text & "','" & _
        txtcod_ctacont.Text & "','" & _
        Right(cboCod_TipFamTela.Text, 1) & "'"

        Con.Execute StrSQL

        Con.CommitTrans
        Dim amensaje As New clsMessages
        amensaje.CODIGO = CodeMsg.kMeSsaGe_INF_DATA_save
        Informa "", amensaje
        
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler err, "Salvar_Datos"
End Sub
Private Sub ELIMINAR_DATOS()
 Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
   
    Con.ConnectionString = cConnect
    Con.Open
    Con.BeginTrans
       
        StrSQL = "EXEC UP_MAN_FAMTELA '" & _
        sTipo & "','" & _
        txtCod_FamTela.Text & "','" & _
        txtDes_FamTela.Text & "','" & _
        txtcod_ctacont.Text & "','" & _
        Right(cboCod_TipFamTela.Text, 1) & "'"
        
        Con.Execute StrSQL
        
    Con.CommitTrans
    
    Dim amensaje As New clsMessages
    amensaje.CODIGO = CodeMsg.kMeSsaGe_INF_DATA_DELETE
    Informa "", amensaje
    
Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler err, "Eliminar_Datos"

End Sub
Private Sub LIMPIAR_DATOS()
    txtCod_FamTela.Text = ""
    txtDes_FamTela.Text = ""
    txtcod_ctacont.Text = ""
    cboCod_TipFamTela.ListIndex = -1
End Sub
Private Sub DGridLista_Click()
    If Rs_Carga.State <> 1 Then
        Exit Sub
    End If
    If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
        txtCod_FamTela.Text = Trim(Rs_Carga("Cod_FamTela").Value)
        txtDes_FamTela.Text = Trim(Rs_Carga("Des_FamTela").Value)
        txtcod_ctacont.Text = Trim(Rs_Carga("Cod_CtaCont").Value)
        Call BuscaCombo(Rs_Carga("Cod_TipFamTela"), 2, cboCod_TipFamTela)
        DESHABILITA_DATOS
    End If
End Sub
Private Sub HABILITA_DATOS()
    txtCod_FamTela.Enabled = True
    txtDes_FamTela.Enabled = True
    txtcod_ctacont.Enabled = True
    'cboCod_TipFamTela.Enabled = True
End Sub
Private Sub DESHABILITA_DATOS()
    txtCod_FamTela.Enabled = False
    txtDes_FamTela.Enabled = False
    txtcod_ctacont.Enabled = False
    cboCod_TipFamTela.Enabled = False
End Sub
Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub
Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Rs_Carga.State <> 1 Then
        Exit Sub
    End If
    If Not Rs_Carga.EOF And Not Rs_Carga.BOF Then
        txtCod_FamTela.Text = Trim(Rs_Carga("Cod_FamTela").Value)
        txtDes_FamTela.Text = Trim(Rs_Carga("Des_FamTela").Value)
        txtcod_ctacont.Text = Trim(Rs_Carga("Cod_CtaCont").Value)
        Call BuscaCombo(Rs_Carga("Cod_TipFamTela"), 2, cboCod_TipFamTela)
        DESHABILITA_DATOS
    End If
End Sub
Private Sub RECARGAR_DATOS()
    Rs_Carga.Close
    Cargar_Datos
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Rs_Carga = Nothing
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub


Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Eliminar As Integer
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
           
           'Llena el combo con los tipos de Familia
            StrSQL = "SELECT Des_TipFamTela + SPACE(100) + Cod_TipFamTela FROM TX_TIPFAM WHERE flg_telas='S'"
            Call LlenaCombo(cboCod_TipFamTela, StrSQL, cConnect)

            LIMPIAR_DATOS
            HABILITA_DATOS
            cboCod_TipFamTela.Enabled = True
            txtCod_FamTela.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "MODIFICAR"
            sTipo = "U"
            HABILITA_DATOS
            txtCod_FamTela.Enabled = False
            txtDes_FamTela.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "ELIMINAR"
            sTipo = "D"
            If VALIDA_DATOS Then
                Eliminar = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Familia")
                If Eliminar = vbYes Then
                    ELIMINAR_DATOS
                    RECARGAR_DATOS
                    sTipo = " "
                End If
            End If
        Case "GRABAR"
            If VALIDA_DATOS Then
                SALVAR_DATOS
                RECARGAR_DATOS
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                DGridLista.Enabled = True
                'cmdBuscaMatPri.Enabled = False
                sTipo = ""
            End If
        Case "DESHACER"
            LIMPIAR_DATOS
            RECARGAR_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridLista.Enabled = True
            sTipo = ""
        Case "SALIR"
            Unload Me
    End Select
End Sub

Function VALIDA_DATOS() As Boolean
    Dim aMess(4)
    Dim amensaje As clsMessages
    Set amensaje = New clsMessages
    VALIDA_DATOS = True
    
    If sTipo = "I" Then
        If Trim(txtCod_FamTela.Text) = "" Then
            VALIDA_DATOS = False
            MsgBox ("Sirvase ingresar el código de familia")
            txtCod_FamTela.SetFocus
        Else
            StrSQL = "SELECT Cod_FamTela FROM TX_FAMTELA WHERE Cod_FamTela ='" & Trim(txtCod_FamTela.Text) & "'"
            If DevuelveCampo(StrSQL, cConnect) <> "" Then
                VALIDA_DATOS = False
                MsgBox ("El código ingresado ya se encuentra registrado. Sirvase ingresar otro")
                txtCod_FamTela.SetFocus
            End If
        End If
        If Trim(txtDes_FamTela.Text) = "" Then
            VALIDA_DATOS = False
            Call MsgBox("La descripción de la familia no puede estar vacia. Sirvase verificar", vbInformation)
            txtDes_FamTela.SetFocus
        End If
    End If
    
    If sTipo = "U" Then
        If Trim(txtDes_FamTela.Text) = "" Then
            VALIDA_DATOS = False
            Call MsgBox("La descripción de la familia no puede estar vacia. Sirvase verificar", vbInformation)
            txtDes_FamTela.SetFocus
        End If
    End If
    If sTipo = "D" Then
        StrSQL = "SELECT Cod_FamTela FROM TX_TELA WHERE Cod_FamTela ='" & Trim(txtCod_FamTela.Text) & "'"
        If DevuelveCampo(StrSQL, cConnect) <> "" Then
            VALIDA_DATOS = False
            MsgBox ("No se puede eliminar la familia, por que tiene telas relacionadas")
        End If
    End If
    
End Function

Private Sub txtcod_ctacont_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub

Private Sub txtPor_Mermacnf_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub

Private Sub txtPor_MermaLog_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub

Private Sub txtUlt_TelGen_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub

