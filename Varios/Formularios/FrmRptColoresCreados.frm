VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRptColoresCreados 
   Caption         =   "Muestra Colores Creados"
   ClientHeight    =   1350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
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
         Height          =   435
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   1185
      End
      Begin VB.TextBox txtNom_Cliente 
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   645
         Width           =   5175
      End
      Begin VB.TextBox txtAbr_Cliente 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   645
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPInicio 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Format          =   78315521
         CurrentDate     =   38182
      End
      Begin MSComCtl2.DTPicker DTPHasta 
         Height          =   270
         Left            =   3600
         TabIndex        =   5
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   476
         _Version        =   393216
         Format          =   78315521
         CurrentDate     =   38182
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Left            =   9000
         TabIndex        =   9
         Top             =   120
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CLIENTE:"
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
         TabIndex        =   8
         Top             =   645
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "DESDE:"
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
         Left            =   240
         TabIndex        =   7
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "HASTA:"
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
         Left            =   2880
         TabIndex        =   6
         Top             =   285
         Width           =   585
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   5640
      Top             =   1200
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmRptColoresCreados"
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
Private rs As New ADODB.Recordset

Private Sub cmdBuscar_Click()
Call BUSCAR
End Sub
Private Sub Form_Load()
    DTPInicio = Date - 7
    DTPHasta = Date
End Sub

Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtNom_Cliente.Text) = "" Then
            Call BUSCA_CLIENTE(3)
        Else
            Call BUSCA_CLIENTE(2)
        End If
    End If
End Sub

Private Sub txtAbr_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtAbr_Cliente.Text) = "" Then
            Call BUSCA_CLIENTE(3)
        Else
            Call BUSCA_CLIENTE(1)
        End If
    End If
End Sub



Public Sub BUSCA_CLIENTE(tipo As Integer)
    Select Case tipo
        Case 1:
                    StrSQL = "EXEC TI_BUSCA_CLIENTE 1,'" & Trim(Me.txtAbr_Cliente.Text) & "','','" & vusu & "'"
                    Me.txtNom_Cliente.Text = Trim(DevuelveCampo(StrSQL, cConnect))
                    If Trim(txtNom_Cliente.Text) <> "" Then BUSCAR
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If tipo = 2 Then
                        oTipo.SQuery = "EXEC TI_BUSCA_CLIENTE 2,'','" & Trim(txtNom_Cliente.Text) & "','" & vusu & "'"
                    Else
                        oTipo.SQuery = "EXEC TI_BUSCA_CLIENTE 3,'','','" & vusu & "'"
                    End If
                    
                    oTipo.Cargar_Datos
                    oTipo.gexList.Columns(2).Width = 4850
                    oTipo.Show 1
                    If CODIGO <> "" Then
                         Me.txtAbr_Cliente.Text = Trim(CODIGO)
                         Me.txtNom_Cliente.Text = Trim(Descripcion)
                         Me.txtAbr_Cliente.Tag = DevuelveCampo("select cod_cliente_tex from tx_cliente where abr_cliente ='" & Trim(CODIGO) & "'", cConnect)
'                         OptCliPend.SetFocus
                         CODIGO = "": Descripcion = ""
                         BUSCAR
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
    End Select
    
End Sub

Sub BUSCAR()
On Error GoTo fin
  

Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption

End Sub

Private Sub CmdImprimir_Click()

    Call Reporte

End Sub
Private Sub Reporte()
Dim oo As Object
Dim sRutaLogo  As String
Dim Ruta As String
On Error GoTo errReporte

Me.txtAbr_Cliente.Tag = DevuelveCampo("select cod_cliente_tex from tx_cliente where abr_cliente ='" & Trim(Me.txtAbr_Cliente.Text) & "'", cConnect)
StrSQL = "TI_MUESTRA_COLORES_CREADOS '" & DTPInicio.Value & "','" & DTPHasta.Value & "','" & Me.txtAbr_Cliente.Tag & "'"
Set rs = CargarRecordSetDesconectado(StrSQL, cConnect)


If rs.RecordCount <= 0 Then
  MsgBox "No se Encuentra datos para mostrar", vbInformation + vbOKOnly, "mensaje"
  Exit Sub
End If

Ruta = vRuta & "\RptMuestraColoresCreados.xlt"

Set oo = CreateObject("excel.application")

StrSQL = "SELECT Ruta_Logo = ISNULL(Ruta_Logo, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
sRutaLogo = DevuelveCampo(StrSQL, cConnect)
    
oo.Workbooks.Open Ruta
oo.Visible = False
oo.DisplayAlerts = False
oo.Run "Reporte", rs, txtNom_Cliente.Text, CStr(DTPInicio.Value) + " - " + CStr(DTPHasta.Value)
oo.Visible = True

Set oo = Nothing

Exit Sub
errReporte:
    MsgBox "Hubo error en la impresion del Reporte de Colores Creados " & err.Description, vbCritical, "Impresion"
End Sub

