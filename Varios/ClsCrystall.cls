VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCrystall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Aplicacion  As Application
Private Reporte     As Report
Private botones     As PrintWindowOptions
Private crystalFormFieldDefns   As FormulaFieldDefinitions
Private crystalFormFieldDefn    As FormulaFieldDefinition
Private db          As Database
Private dbtables    As DatabaseTables
Private dbtable     As DatabaseTable
Private NameDatabase As String
Public Function Visualiza(ByVal parTitulo As String, ByVal sHasExportButton As Boolean, _
                            ByVal sHasPrintButton As Boolean, sHasGroupTree As Boolean) As Boolean
On Error GoTo errores

    Dim Ancho As Integer
    Dim Alto As Integer

    Alto = (Screen.Height / Screen.TwipsPerPixelY) - 25
    Ancho = (Screen.Width / Screen.TwipsPerPixelX)

    Set botones = Reporte.PrintWindowOptions
    botones.HasSearchButton = True
    botones.HasExportButton = sHasExportButton
    'Botones.HasExportButton = True
    botones.HasPrintButton = sHasPrintButton
    botones.HasPrintSetupButton = sHasPrintButton
    botones.HasGroupTree = sHasGroupTree
    
    Reporte.Preview parTitulo, 0, 0, Ancho, Alto   'Visualiza el reporte


Exit Function

errores:
MsgBox err.Description, vbCritical, "Visualiza"

End Function

Public Sub Criterio(ByVal strSQL As String)
On Error GoTo errores
    Reporte.SQLQueryString = strSQL
Exit Sub
errores:
MsgBox err.Description, vbCritical, "Criterio"

End Sub

Public Function Coneccion(ByVal parNombreReporte As String) As Boolean

    Dim NameReporte As String

On Error GoTo Abrir_Error


    NameReporte = App.Path & "\" & parNombreReporte


    If Aplicacion Is Nothing Then Set Aplicacion = CreateObject("Crystal.cRPE.Application")

    Set Reporte = Aplicacion.OpenReport(NameReporte)

    Set db = Reporte.Database
    Set dbtables = db.Tables
    Set dbtable = dbtables.Item(1)
    
    'dbtable.SetLogOnInfo gServidor, gBD, gUsuario, gClave
    'dbtable.SetLogOnInfo gDNS, gBD, gUsuarioDNS, gClaveDNS
'    cConnect
    NameDatabase = DevuelveCampo("Select Db_Name()", cConnect)
    
    'dbtable.SetLogOnInfo "serverdata", "precotex_textil", "sa", "1234567890"
    
    dbtable.SetLogOnInfo "serverdata", NameDatabase, "soporte", "soporte"
    Coneccion = True
    Exit Function

Abrir_Exit:
  Screen.MousePointer = vbDefault
  Exit Function

Abrir_Error:
  parNombreReporte = ""
  Coneccion = False
  If Not Reporte Is Nothing Then Set Reporte = Nothing
  If Aplicacion Is Nothing Then err.Description = "Unable to CreateObject(""Crystal.CRPE.Application"")"
  MsgBox err.Description, vbCritical, "Error"
  GoTo Abrir_Exit

End Function

Public Sub Formula(ByVal strValor As String, ByVal Nitem As Integer)
On Error GoTo errores
    Set crystalFormFieldDefns = Reporte.FormulaFields
    Set crystalFormFieldDefn = crystalFormFieldDefns.Item(Nitem)
    crystalFormFieldDefn.Text = strValor

Exit Sub
errores:
MsgBox err.Description, vbCritical, "Formula"


End Sub

Public Sub Imprime()

On Error GoTo Ctrl_Err

    Reporte.ProgressDialogEnabled = False
    Reporte.PrintOut False

    Exit Sub
Ctrl_Err:
    Exit Sub
End Sub

Public Sub Configura(ByVal pDriverName As String, ByVal pPrintername As String, ByVal pPortName As String)
On Error GoTo Ctrl_Err
    Reporte.SelectPrinter pDriverName, pPrintername, pPortName
    Exit Sub
Ctrl_Err:
    Exit Sub
End Sub





