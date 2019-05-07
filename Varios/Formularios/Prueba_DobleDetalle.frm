VERSION 5.00
Begin VB.Form Prueba_DobleDetalle 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "Prueba_DobleDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call Reporte

End Sub

Sub Reporte()
On Error GoTo ErrorImpresion
Dim oo As Object
Dim sMes As String
Dim sAno As String
Dim sSemana As String
Dim sFecha As String
Dim reg As ADODB.Recordset
Dim cCONNECT As String
Dim STRSQL As String

    STRSQL = " HI_MUESTRA_STOCKS_WARRANTS_ALMACEN 'G', '0'"
    
    Set reg = CargarRecordSetDesconectado(STRSQL, cCONNECT)
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open "C:\Tareas\RptWarrantsDisponibles.XLT"

    oo.Visible = True
    oo.DisplayAlerts = False
    'oo.Run "REPORTE", TxtCod_Fabrica.Text, txtAnhioSem.Text, txtSemana.Text, cCONNECT, TxtNom_Fabrica.Text, Txt_Tipo.Text
    'oo.Run "reporte", reg, RS.Fields("Ruta").Value, RS.Fields("Nom_File").Value, RS.Fields("Nom_plantillaSinMacro").Value
    oo.Run "reporte", reg, "", "", ""

    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    ErrorHandler err, "Reporte"
End Sub

'Sub Reporte()
'On Error GoTo hand
'Dim oo As Object
'Dim Ruta As String
'Dim Cadena
'    STRSQL = "SELECT Cod_Cliente_tex FROM Tx_CLIENTE WHERE Abr_Cliente='" & GridEX1.Value(GridEX1.Columns("Abr_Cliente").Index) & "'"
'    Cadena = "es_muestra_solicitudes_desarrollo_Local '" & DevuelveCampo(STRSQL, cCONNECT) & "','" & GridEX1.Value(GridEX1.Columns("Cod_TemCli").Index) & "'"
'    Ruta = vRuta & "\RptSolDesaColores_Local.xlt"
'    Set oo = CreateObject("excel.application")
'    oo.Workbooks.Open Ruta
'    oo.Visible = True
'    oo.DisplayAlerts = False
'    oo.Run "Reporte", GridEX1.Value(GridEX1.Columns("Abr_Cliente").Index) & "-" & GridEX1.Value(GridEX1.Columns("Nom_Cliente").Index), GridEX1.Value(GridEX1.Columns("Cod_TemCli").Index) & "-" & GridEX1.Value(GridEX1.Columns("Nom_TemCli").Index), Cadena, cCONNECT
'    Set oo = Nothing
'Exit Sub
'hand:
'    ErrorHandler err, "GeneraReportes"
'    Set oo = Nothing
'End Sub

