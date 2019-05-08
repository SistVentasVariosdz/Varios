Attribute VB_Name = "modRutinas"

Public Const Deshabilitado = &H8000000A
Public Const TODOS = "<TODOS>"
Declare Function GetcomputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public oConexOo As String
Public UsuOo As String
Public PswOo As String

'Declaración de la API nombre session de windows
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long
'Definición de la función
'Obtiene el usuario actual del sistema

Public Function usuario_windows() As String
    'creamos variables
    Dim lpBuff As String * 25
    Dim ret As Long
    Dim UserName As String
    'Obtenemos el nombre de la api.
    ret = GetUserName(lpBuff, 25)
    UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
    ' Devolvemos el nombre de usuario
    usuario_windows = UserName
End Function

Public Function PreparaOpenOffice(Reporte As String, oo As Object) As Object
    
    Dim StrSQL As String
    StrSQL = "SELECT ConexRptOo FROM SEG_EMPRESAS WHERE cod_EMPRESA ='" & Trim(vemp) & "'"
    oConexOo = DevuelveCampo(StrSQL, cSEGURIDAD)
    StrSQL = "SELECT UsuRptOo FROM SEG_EMPRESAS WHERE cod_EMPRESA ='" & Trim(vemp) & "'"
    UsuOo = DevuelveCampo(StrSQL, cSEGURIDAD)
    StrSQL = "SELECT PswRptOo FROM SEG_EMPRESAS WHERE cod_EMPRESA ='" & Trim(vemp) & "'"
    PswOo = DevuelveCampo(StrSQL, cSEGURIDAD)
    Ruta = vRuta & "\" & Reporte
    
    Set oo = CreateObject("ooBusiness.Calc")
    oo.OfficeTemplateSheet = Ruta
    oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
    oo.MacroLibraryName = "Library1"
    oo.MacroModuleName = "Module1"
    oo.MacroName = "Reporte"
    
    Set PreparaOpenOffice = oo

End Function

Function ExisteArchivo(ByVal FileNa As String) As Boolean
    Dim FRes As String
    FRes = Dir$(FileNa)
    If FRes = "" Then ExisteArchivo = False Else ExisteArchivo = True
End Function

Function NumeroSemana(fecha As Date) As Integer
    Dim SFecha As String
    SFecha = "01/" & Month(fecha) & "/" & Year(fecha)
    getNumeroSemana = DatePart("ww", fecha) '- DatePart("ww", DateValue(sFecha)) + 1
End Function

Public Sub SelectionText(cltSel As Object)
 cltSel.SelStart = 0
 cltSel.SelLength = Len(cltSel.Text)
End Sub

Sub AVANZA(ByVal Tecla As Integer)
    Select Case Tecla
        Case 13, 40: SendKeys "{TAB}", True
        Case 38: SendKeys "+{TAB}", True
    End Select
End Sub

Sub Plin(ByVal Text)
    If IsNull(Text) Then
       Text = ""
    End If
     Print #1, Text
    iLin = iLin + 1
End Sub

Public Sub ADODBToSSDBGridOC(ByVal RsBuff As ADODB.Recordset, ByRef ssDBGrid As Object)  'As SSDataWidgets_B.ssDbGrid)
On Error Resume Next
Dim iContador As Long
Dim nCols As Integer
Dim iVerif As Integer
Dim Temp As String
Dim NVEZ As Boolean
Dim x%
Dim total1 As Long
Dim y%
Dim i As Long
Dim ic As Long
 
 ssDBGrid.FieldSeparator = "~"
 ssDBGrid.Redraw = False
 nCols = RsBuff.Fields.Count

 NVEZ = True


 x = 0
 Do While Not RsBuff.EOF
   Temp = ""
   For iContador = 0 To nCols - 1
      'ssDBGrid.Columns(iContador).Locked = True
      'ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
      'ssDBGrid.Columns(iContador).Style = 4 'ssStyleButton
      Temp = Temp & FixNulos(RsBuff(iContador).Value, vbString)
      If iContador < nCols - 1 Then
         Temp = Temp & "~"
      End If

      'If iContador >= FixNulos(ssDBGrid.TagVariant, vbLong) Then
      '      ssDBGrid.Columns(iContador).DataType = 5
      '      ssDBGrid.Columns(iContador).Alignment = 1
      'End If

      'ssDbgrid.Columns(iContador).DataType = 5
      'If ssDBGrid.Columns(iContador).DataType = 5 Or ssDBGrid.Columns(iContador).DataType = 6 Or ssDBGrid.Columns(iContador).DataType = 9 Or iContador > FixNulos(ssDBGrid.TagVariant, vbLong) Then
      '  If Val(FixNulos(RsBuff(iContador).Value, vbDouble)) > 0 Then
      '      ssDBGrid.Columns(iContador).TagVariant = Val(ssDBGrid.Columns(iContador).TagVariant) + FixNulos(RsBuff(iContador).Value, vbDouble)
      '  End If
      'End If
   Next
   NVEZ = False
   ssDBGrid.AddItem Temp
  RsBuff.MoveNext
  x = x + 1
  
  
 Loop
 
 ssDBGrid.AllowDragDrop = True
 ssDBGrid.RowHeight = 300 ' SSDBGrid.RowHeight * 1.25
 ssDBGrid.Refresh

 ssDBGrid.Redraw = True
 'RsBuff.Close
 'Set RsBuff = Nothing

End Sub
Function FixNulos(wtexto As Variant, wTipo As Integer)
   If IsNull(wtexto) Or Len(Trim(wtexto)) = 0 Then
      Select Case wTipo
        Case 2, 3, 4, 5
           wtexto = 0
        Case 7
           wtexto = Empty '(" Empty 'Format$("", "mm/dd/yyyy")
        Case 8
           wtexto = ""
        Case 11
           wtexto = False
      End Select
   End If
   FixNulos = wtexto
End Function




Public Sub SSDBGridSetGrid(ByRef ssDBGrid As Object)
    Dim i As Long
    Dim n As Long
    
    ssDBGrid.Col = 0
    ssDBGrid.SplitterPos = 0
    ssDBGrid.SplitterVisible = False
    ssDBGrid.RemoveAll
    ssDBGrid.Refresh
    ssDBGrid.Redraw = False
    n = ssDBGrid.Cols
    If Not IsEmpty(ssDBGrid.TagVariant) Then
        If n > ssDBGrid.TagVariant Then
            For i = n To ssDBGrid.TagVariant + 1 Step -1
                ssDBGrid.Columns.Remove ssDBGrid.Cols - 1
            Next
        End If
    End If
    ssDBGrid.Redraw = True
    ssDBGrid.Refresh
End Sub

Function Ceros(Texto As String) As String
On Error Resume Next
Ceros = Format(Texto, "0####")
End Function


Public Function ComputerName() As String
    Dim KeyName$
    Dim keylen&
    Dim iNull
            
    keylen& = 2000
    KeyName$ = String$(keylen, 0)
    
    GetcomputerName KeyName$, keylen&
    
    iNull = InStr(KeyName, Chr(0))
    ComputerName = Mid(KeyName$, 1, iNull - 1)
End Function

Function DevuelveFechaServidor() As Date
On Error GoTo hand
    DevuelveFechaServidor = DevuelveCampo("select getdate()", cConnect)
Exit Function
hand:
ErrorHandler err, "DevuelveFechaServidor"
End Function


Function ExisteCampo(pCampo As String, pTabla As String, pValor As Variant, Conexion As String, Optional pEsStringValor As Boolean = True) As Boolean
On Error GoTo hand

If pEsStringValor Then
    If DevuelveCampo("select count(" & pCampo & ") from " & pTabla & " where " & pCampo & " = '" & pValor & "'", Conexion) > 0 Then
        ExisteCampo = True
    Else
        ExisteCampo = False
    End If
Else
    If DevuelveCampo("select count(" & pCampo & ") from " & pTabla & " where " & pCampo & " = " & pValor, Conexion) > 0 Then
        ExisteCampo = True
    Else
        ExisteCampo = False
    End If
End If
Exit Function
hand:
ErrorHandler err, "ExisteCampo"
ExisteCampo = False
End Function

Public Sub FormSet(ByRef FormMe As Form)
    Dim oControl As Control
    Dim oDiccionario As Object
    Dim vbuff As Variant
    Dim sUserActions As String
    Set oDiccionario = Nothing
    Conecta
End Sub



Sub Conecta()
Set CadConn = Nothing
CadConn.ConnectionString = cConnect
CadConn.Open
'CadConn.ConnectionString = "Provider=sqloledb;Server=SERVIDOR;Database=lives;UID=sa;pwd=;"
'CadConn.Open
'cCONNECT = "Provider=sqloledb;Server=SERVIDOR;Database=lives;UID=sa;pwd=;"
'cSEGURIDAD = "Provider=sqloledb;Server=servidor;Database=seguridad;UID=sa;pwd=;"
'cCONNECT = "Provider=sqloledb;Server=SERVIDOR;Database=RGS;UID=sa;pwd=;"
End Sub

Public Function DevuelveMes(ByRef pMes As String, pIdioma As String) As Variant
On Error GoTo hand
DevuelveMes = DevuelveCampo("select dbo.uf_nombre_mes('" & Format(CInt(pMes), "0#") & "','" & pIdioma & "')", cConnect)
Exit Function
hand:
ErrorHandler err, "DevuelveMes"
End Function

Public Function DevuelveCampo(pQuerySql As String, pConexion As String) As Variant
On Error GoTo DevuelveCampoError
    Dim rstBuscaCampo As New ADODB.Recordset

   ' Set rstBuscaCampo.ActiveConnection = pConexion
    rstBuscaCampo.CursorLocation = adUseClient
    rstBuscaCampo.Open pQuerySql, pConexion, adOpenKeyset, adLockOptimistic

    If rstBuscaCampo.RecordCount > 0 Then
        DevuelveCampo = rstBuscaCampo(0)
    Else
        DevuelveCampo = ""
    End If
    Set rstBuscaCampo = Nothing
Exit Function
DevuelveCampoError:
    ErrorHandler err, "Funcion DevuelveCampo"
    err.Clear
    DevuelveCampo = ""
    Set rstBuscaCampo = Nothing
End Function

'-------------------------------------------------------------
' Function  : EjecutarQuery()
' Propósito : Ejecutar una sentencia SQL Query
' Input     : pQuery: SQL Query
'             pCursorType: ADO Cursor Type
' Output    : ADO Recordset obtenido
'-------------------------------------------------------------
Public Function EjecutarQuery(ByVal pQuery As String, _
                              ByVal pCursorType As ADODB.CursorTypeEnum) _
                              As ADODB.Recordset
   Dim adoRs As ADODB.Recordset
   Dim adoRsUltimo As ADODB.Recordset

   Set adoRs = New ADODB.Recordset
   With adoRs
      .ActiveConnection = g_cnnConexion
      .CursorLocation = adUseClient
      .CursorType = pCursorType
      .LockType = adLockOptimistic
      .Open pQuery
   End With

   ' Se obtiene el ultimo resultado de Recordset
   Do While Not (adoRs Is Nothing)
      Set adoRsUltimo = adoRs
      Set adoRs = adoRsUltimo.NextRecordset
   Loop
   Set adoRs = adoRsUltimo

   Set EjecutarQuery = adoRs

End Function


'Sub FormateaGrid(pGrid As MSDataGridLib.DataGrid)
'On Error GoTo hand
'        pGrid.MarqueeStyle = dbgHighlightRow
'        pGrid.HeadFont.Bold = True
'        pGrid.Refresh
'        pGrid.BackColor = -2147483624
'Exit Sub
'hand:
'ErrorHandler Err, "FormateaGrid"
'End Sub

Sub LlenaCombo(objObjeto As Object, strQuery As String, Conexion As String)
On Error GoTo LlenaComboError
    Dim rstBuscaCampo As New ADODB.Recordset
    
    rstBuscaCampo.CursorLocation = adUseClient
    rstBuscaCampo.Open strQuery, Conexion, adOpenDynamic, adLockOptimistic
        objObjeto.Clear
    If rstBuscaCampo.RecordCount > 0 Then
        rstBuscaCampo.MoveFirst
        With rstBuscaCampo
            If rstBuscaCampo.Fields.Count = 2 Then
                Do While Not .EOF
                    objObjeto.AddItem IIf(IsNull(rstBuscaCampo(0)), "", rstBuscaCampo(0)) & Space(3) & IIf(IsNull(rstBuscaCampo(1)), "", rstBuscaCampo(1))
                    .MoveNext
                Loop
            Else
                
                Do While Not .EOF
                    objObjeto.AddItem IIf(IsNull(rstBuscaCampo(0)), "", rstBuscaCampo(0))
                    .MoveNext
                Loop
            End If
        End With
    End If
Set rstBuscaCampo = Nothing
Exit Sub
LlenaComboError:
    ErrorHandler err, "Procedimiento LlenaCombo"
    err.Clear
    Set rstBuscaCampo = Nothing
End Sub


Sub BuscaCombo(strTexto As String, intPos As Integer, combo As ComboBox)
    Dim intCont As Integer
    Dim Encontro As Boolean
    Encontro = False
    combo.ListIndex = -1
    If Trim(strTexto) = "" Then Exit Sub
    If intPos = 1 Then
        For intCont = 0 To combo.ListCount - 1
            If strTexto = Mid(combo.List(intCont), 1, Len(strTexto)) Then
                combo.ListIndex = intCont
                Encontro = True
                Exit For
            End If
        Next
    Else
        For intCont = 0 To combo.ListCount - 1
            If strTexto = Right(combo.List(intCont), Len(strTexto)) Then
                combo.ListIndex = intCont
                Encontro = True
                Exit For
            End If
        Next
    End If
    If Encontro = False Then
        combo.ListIndex = -1
    End If
End Sub

Sub BuscaCampo(pRs_Lista As ADODB.Recordset, pCampo As String, pValor As String)
On Error GoTo hand
    Dim pIndice As Integer
    Dim pRs_Prov As New ADODB.Recordset
    
    If Not pRs_Lista.EOF And Not pRs_Lista.BOF Then
        Set pRs_Prov = pRs_Lista.Clone
        pIndice = 0
        pRs_Prov.MoveFirst
        While Not pRs_Prov.EOF
            If Mid(pRs_Prov(pCampo).Value, 1, Len(pValor)) = pValor Then
                pRs_Lista.MoveFirst
                pRs_Lista.Move (pIndice)
                
                pRs_Prov.Close
                Set pRs_Prov = Nothing
                
                Exit Sub
            End If
            pIndice = pIndice + 1
            pRs_Prov.MoveNext
        Wend
    End If
    pRs_Prov.Close
    Set pRs_Prov = Nothing

Exit Sub
hand:
ErrorHandler err, "BuscaCampo"
pRs_Prov.Close
Set pRs_Prov = Nothing
End Sub


Sub Main()
'FrmMantAlmacen.Show
'frmClaOrdComp.Show
'FrmMantTipMOv.Show
'FrmMovAlmacen.Show
End Sub

'-------------------------------------------------------------
' Procedure : SoloNumeros()
' Propósito : Funcion que permite el ingreso de solo numeros
'             sobre un control Textbox
' Input     : pTextbox: Control Textbox,
'             pKeyAscii: La tecla ingresada,
'             pConDecimales: Si se usa o no decimales,
'             pNumDecimales: Numero de Decimales permitidos,
'             pNumEntero: Numero de Enteros permitidos
'-------------------------------------------------------------
Public Sub SoloNumeros(ByVal pTextbox As TextBox, _
                       ByRef pKeyAscii As Integer, _
                       Optional ByVal pConDecimales As Boolean, _
                       Optional ByVal pNumDecimales As Integer, _
                       Optional ByVal pNumEnteros As Integer)
   If pNumEnteros = 0 Then pNumEnteros = 10
   If pKeyAscii = 8 Then
      If pConDecimales And pTextbox.SelStart > 0 Then
         If Mid(pTextbox, pTextbox.SelStart, 1) = "." Then
            If Len(Mid(pTextbox, 1, pTextbox.SelStart - 1)) >= pNumEnteros And Len(Mid(pTextbox, pTextbox.SelStart + 1)) > 0 Then pKeyAscii = 0
         End If
      End If
      Exit Sub
   End If
   If pKeyAscii = 46 Then
      If pConDecimales Then
         If InStr(1, pTextbox, ".") > 0 Then
            pKeyAscii = 0
         Else
            If Len(Mid(pTextbox, pTextbox.SelStart + 1)) > pNumDecimales Then pKeyAscii = 0
            If pTextbox.SelStart > 0 Then If Len(Mid(pTextbox, 1, pTextbox.SelStart - 1)) >= pNumEnteros Then pKeyAscii = 0
         End If
      Else
         pKeyAscii = 0
      End If
   Else
      If Not (pKeyAscii >= 48 And pKeyAscii <= 57) Then pKeyAscii = 0
      If pKeyAscii = 39 Or pKeyAscii = 13 Then
         pKeyAscii = 0
      End If
      
      Dim iPos As Integer
      iPos = InStr(1, pTextbox, ".")
      If iPos > 0 And pConDecimales Then _
         If Len(Mid(pTextbox, iPos)) > pNumDecimales Then _
            If InStr(pTextbox.SelStart + 1, pTextbox, ".") = 0 Then pKeyAscii = 0
            
      If pTextbox.SelStart < iPos Or iPos = 0 Then
         If pNumEnteros > 0 Then
            If InStr(pTextbox.SelStart + 1, pTextbox, ".") > 0 Then
               If Len(Mid(pTextbox, 1, InStr(pTextbox.SelStart + 1, pTextbox, ".") - 1)) >= pNumEnteros Then pKeyAscii = 0
            Else
               If Len(pTextbox) >= pNumEnteros Then pKeyAscii = 0
            End If
         End If
      End If
   End If
End Sub


'-------------------------------------------------------------
' Procedure : ErrorHandler()
' Propósito : Manejo de Excepciones Genérico
' Input     : pErr: Objeto Error VB,
'             pProcedure: Nombre del Procedimiento
'-------------------------------------------------------------
Public Sub ErrorHandler(ByVal pErr As ErrObject, ByVal pProcedure As String)
   Dim sMsg As String
   
   Screen.MousePointer = vbDefault
   sMsg = pProcedure & " : " & Chr(13) & _
          "Descripción : " & pErr.Description
          err.Clear
   MsgBox sMsg, vbExclamation, App.Title
End Sub

Public Sub ADODBToSSDBGrid(ByVal RsBuff As ADODB.Recordset, ByRef ssDBGrid As Object)  'As SSDataWidgets_B.ssDbGrid)
On Error Resume Next
Dim iContador As Long
Dim nCols As Integer
Dim iVerif As Integer
Dim Temp As String
Dim NVEZ As Boolean
Dim x%
Dim total1 As Long
Dim y%
Dim i As Long
Dim ic As Long
 
 ssDBGrid.FieldSeparator = "~"
 'Set rsBuff = New RBS.clsRecordSet
 'Set rsBuff.refObject = oData

 'rsBuff.Buffer = pBuff
 ssDBGrid.Redraw = False
 
 'nCols = RsBuff.Count
 nCols = RsBuff.Fields.Count

' ic = ssDBGrid.Cols
' If ssDBGrid.Cols < nCols Then
'    For i = nCols To ic + 1 Step -1
'       ssDBGrid.Columns.Add ssDBGrid.Cols    ' "Column" & i, 500, False, Nothing, "Column" & i
'       ssDBGrid.Columns(ssDBGrid.Cols - 1).Name = rsBuff(ssDBGrid.Cols).Name
'       ssDBGrid.Columns(ssDBGrid.Cols - 1).Caption = rsBuff(ssDBGrid.Cols).Name
'    Next i
' End If
'
' For y = 0 To ssDBGrid.Cols - 1
'   If ssDBGrid.Columns(y).DataType = 5 Or ssDBGrid.Columns(y).DataType = 6 Or ssDBGrid.Columns(y).DataType = 9 Then
'      ssDBGrid.Columns(y).TagVariant = 0
'   End If
' Next

 NVEZ = True


 x = 0
 Do While Not RsBuff.EOF
   Temp = ""
   For iContador = 0 To nCols - 1
      ssDBGrid.Columns(iContador).Locked = True
      ssDBGrid.Columns(iContador).CaptionAlignment = 0 'ssColCapAlignLeftJustify
      ssDBGrid.Columns(iContador).Style = 4 'ssStyleButton
      Temp = Temp & FixNulos(RsBuff(iContador).Value, vbString)
      If iContador < nCols - 1 Then
         Temp = Temp & "~"
      End If

      If iContador >= FixNulos(ssDBGrid.TagVariant, vbLong) Then
            ssDBGrid.Columns(iContador).DataType = 5
            ssDBGrid.Columns(iContador).Alignment = 1
      End If

      'ssDbgrid.Columns(iContador).DataType = 5
      If ssDBGrid.Columns(iContador).DataType = 5 Or ssDBGrid.Columns(iContador).DataType = 6 Or ssDBGrid.Columns(iContador).DataType = 9 Or iContador > FixNulos(ssDBGrid.TagVariant, vbLong) Then
        If Val(FixNulos(RsBuff(iContador).Value, vbDouble)) > 0 Then
            ssDBGrid.Columns(iContador).TagVariant = Val(ssDBGrid.Columns(iContador).TagVariant) + FixNulos(RsBuff(iContador).Value, vbDouble)
        End If
      End If
   Next
   NVEZ = False
   ssDBGrid.AddItem Temp
  RsBuff.MoveNext
  x = x + 1
 Loop
 ssDBGrid.AllowDragDrop = True
 ssDBGrid.RowHeight = 300 ' SSDBGrid.RowHeight * 1.25
 ssDBGrid.Refresh

 ssDBGrid.Redraw = True
 'RsBuff.Close
 'Set RsBuff = Nothing

End Sub

Public Sub SSDBGridSetGrid0(ByRef ssDBGrid As Object)
ssDBGrid.TagVariant = ssDBGrid.Cols
End Sub


'Funciones para el Grid del Janus
'-----------------------------------
Public Sub RefreshGridEx(ByRef prmGridEx As GridEx)
    prmGridEx.Rebind
    prmGridEx.HoldFields
End Sub

Public Sub InitGridEx(ByRef prmGridEx As GridEx, sSQL As String)
On Error GoTo Err_InitGridEx
'    prmGridEx.LoadLayout App.Path & "\" & prmGridEx.Name & ".txt"
    prmGridEx.DatabaseName = cConnect
    prmGridEx.RecordSource = sSQL
    RefreshGridEx prmGridEx
    Exit Sub
Err_InitGridEx:
    Resume Next
End Sub

Public Sub ReleaseGridEx(ByRef prmGridEx As GridEx)
    prmGridEx.SaveLayout App.Path & "\" & prmGridEx.Name & ".txt"
End Sub

Function FixData(wtexto As Variant, ofield As ADODB.FIELD)
   If IsNull(wtexto) Or Len(Trim(wtexto)) = 0 Then

      Select Case ofield.Type
        Case adBigInt, adCurrency, adDecimal, adDouble, adInteger, adNumeric, adSingle
            wtexto = 0
        Case adBoolean
            wtexto = False
        Case adDate
            wtexto = Empty
        Case adChar, adVarChar
            wtexto = ""
      End Select
   End If
   FixData = wtexto
End Function

Public Function CargarRecordSetDesconectado(ByVal sSQL As String, ByVal cConnect As String) As ADODB.Recordset
Dim rsBD As ADODB.Recordset
Dim rsGridEx As ADODB.Recordset
Dim ofield As Object
Dim oCon As ADODB.Connection

    Set oCon = New ADODB.Connection

    oCon.CursorLocation = adUseClient
    oCon.Open cConnect
    oCon.CommandTimeout = 900

    Set rsBD = New ADODB.Recordset
    Set rsBD.ActiveConnection = oCon

    rsBD.CursorLocation = adUseClient
    rsBD.CursorType = adOpenStatic

    rsBD.Open sSQL
    
    
    Set rsGridEx = New ADODB.Recordset
    rsGridEx.CursorLocation = adUseClient
    Set rsGridEx.ActiveConnection = Nothing

    For Each ofield In rsBD.Fields
        rsGridEx.Fields.Append ofield.Name, ofield.Type, ofield.DefinedSize, adFldIsNullable
        rsGridEx.Fields(ofield.Name).NumericScale = rsBD.Fields(ofield.Name).NumericScale
        rsGridEx.Fields(ofield.Name).DefinedSize = rsBD.Fields(ofield.Name).DefinedSize
        rsGridEx.Fields(ofield.Name).Precision = rsBD.Fields(ofield.Name).Precision
    Next
    rsGridEx.Open

    If rsBD.RecordCount Then
        rsBD.MoveFirst
        Do While Not rsBD.EOF
            rsGridEx.AddNew
            For Each ofield In rsBD.Fields
                rsGridEx.Fields(ofield.Name).Value = FixData(rsBD.Fields(ofield.Name).Value, rsBD.Fields(ofield.Name))
            Next
            rsGridEx.Update
            rsBD.MoveNext
        Loop
    End If

    Set CargarRecordSetDesconectado = rsGridEx

End Function

Public Function SetGeneralGridEX(ByRef GridEx As GridEX20.GridEx, ByVal iFixsCols As Integer, ByVal iTipoColorBack As Integer)

    If iFixsCols > 0 Then
        GridEx.FrozenColumns = iFixsCols
    End If

    If iTipoColorBack = 1 Then
        GridEx.BackColor = &H80000018
        GridEx.BackColorBkg = &H80000018
        GridEx.GridLines = jgexGLVertical
        GridEx.GridLineStyle = jgexGLSSmallDots
    Else
        GridEx.BackColor = &H80000005
        GridEx.BackColorBkg = &H80000005
        GridEx.GridLines = jgexGLBoth
        GridEx.GridLineStyle = jgexGLSSmallDots
    End If

End Function

Public Function ExecuteSQL(ByVal Connect As String, ByVal SQL As String) As Long
  'this function executes and SQL string and returns the number of records affected
  On Error GoTo ehExecuteSQL
  Dim objAdoCn As Object ' New ADODB.Connection

  Set objAdoCn = CreateObject("ADODB.Connection")    'ADO must be registered locally ' New ADODB.Connection  '
  objAdoCn.Open Connect                 'open connection
  objAdoCn.CommandTimeout = 900

  objAdoCn.Execute SQL, ExecuteSQL, 128  'recordsetAffected is returned
  objAdoCn.Close
  Set objAdoCn = Nothing

Exit Function
ehExecuteSQL:
 'MsgBox Err.Description
  'if transactions is not committed, it will be rolled back
  ExecuteSQL = -2                         '-2 indicates error condition
  err.Raise err.Number, "ExecuteSQL", err.Description
End Function

Public Sub ComboBoxToComboBox(ByRef lstOrigen As Object, ByRef lstDestino As Object, ByVal iModal As Integer)
'iModal 0 Pasa item actual
'       1 Pasa todos los items
Dim i As Long
Dim j As Long

If iModal = 0 Then
    If lstOrigen.ListIndex <> -1 Then
        lstDestino.AddItem ""
        For i = 0 To 0 ' lstOrigen.ColumnCount - 1

            lstDestino.List(lstDestino.ListCount - 1) = lstOrigen.List(lstOrigen.ListIndex)
        Next
        lstOrigen.RemoveItem lstOrigen.ListIndex
    End If
Else
    For j = 0 To lstOrigen.ListCount - 1
        If RTrim(lstOrigen.List(j)) <> "" Then
            lstDestino.AddItem ""
            For i = 0 To 0  ' lstOrigen.ColumnCount - 1
                lstDestino.List(lstDestino.ListCount - 1) = lstOrigen.List(j)
            Next
        End If
    Next

    For j = lstOrigen.ListCount - 1 To 0 Step -1
        lstOrigen.RemoveItem j
    Next
End If
End Sub

Public Function CompletaCodigo(CodOrigen As String, longcodfinal As Integer, PosfinalCod As Integer) As String
' CodOrigen     = Es el codigo que sera pasado por parametro
' LongCodFinal  = Es el tamaño del Codigo a devolver
' PosFinalCod   = Es la posicion de la 1era parte del codigo
    Dim Contador As Integer
    CompletaCodigo = Mid(CodOrigen, 1, PosfinalCod)
    For Contador = 1 To longcodfinal - Len(CodOrigen)
        CompletaCodigo = CompletaCodigo & "0"
    Next
    CompletaCodigo = CompletaCodigo & Right(CodOrigen, Len(CodOrigen) - PosfinalCod)
End Function

Public Function AlExp(varExp As Variant, xPos As Integer, yPos As Integer) As String
Dim sAux As String, sAux2 As String
    sAux = "": sAux2 = ""
    Select Case VarType(varExp)
    Case vbString
        sAux = Trim(Left(CStr(varExp), xPos))
        If Len(sAux) < xPos Then
            sAux = sAux & Space(xPos - Len(sAux))
        End If
    Case vbDouble, vbDecimal, vbInteger, vbLong
        sAux = Int(CDbl(varExp))
        sAux = Trim(sAux)
        If Len(sAux) < xPos Then
            sAux = String(xPos - Len(sAux), "0") & sAux
        End If
        sAux2 = Round(CDbl(varExp) - Int(CDbl(varExp)), 2)
        sAux2 = Trim(Mid(sAux2, 3, yPos))
        If Len(sAux2) < yPos Then
            sAux2 = sAux2 & String(yPos - Len(sAux2), "0")
        End If
        sAux = sAux & sAux2
    End Select
    AlExp = sAux
End Function

Public Function RPad(ByVal InString As Variant, _
                        ByVal iNumChar As Integer, _
                        Optional Char As Variant) As String
                        
    Dim WithThisChar As String
    Dim StringChar As String
    Dim iIndex As Integer
    
    If IsNull(InString) Then
        InString = ""
    Else
        InString = CStr(InString)
    End If
    
    StringChar = ""
    WithThisChar = IIf(IsMissing(Char), Space$(1), Char)
    
    For iIndex = 1 To iNumChar - Len(InString)
        StringChar = StringChar + WithThisChar
    Next
    
    RPad = Left(InString + StringChar, iNumChar)

End Function

Public Function VBsprintf(ByRef InString As String, ParamArray aInValues()) As String
Dim OutString As String
Dim ThisChar As String
Dim IndexString As Integer
Dim IndexValues As Integer
Dim iNotchar As Integer
Dim vValor As Variant
Dim strCadena As String

OutString = ""
IndexValues = 0


For IndexString = 1 To Len(InString)
 ThisChar = Mid(InString, IndexString, 1)
 
 
' If Asc(ThisChar) = 39 Then
'    MsgBox "llego "
' End If
 If ThisChar <> "$" Then
    OutString = OutString & ThisChar
 Else
   If VarType(aInValues(IndexValues)) = vbString Then
        vValor = aInValues(IndexValues)
        If Len(vValor) >= 2 Then
            If Mid(vValor, 1, 1) <> Chr(39) Then
                vValor = NotChar(vValor)
            End If
           '09/02/2000 2:08 pm
           strCadena = Mid(vValor, 2, Len(vValor) - 2)
           If InStr(strCadena, Chr(34)) Or InStr(strCadena, Chr(39)) Then
              strCadena = NotChar(strCadena)
              vValor = Chr(39) & strCadena & Chr(39)
           End If
        End If
   Else
        vValor = CStr(aInValues(IndexValues))
        vValor = NotChar(vValor)
   End If
   
   OutString = OutString + vValor
   IndexValues = IndexValues + 1
 End If
Next

VBsprintf = OutString

End Function

Private Function NotChar(ByVal vValor As String) As String
Dim i As Integer
Dim sReturn As String

If InStr(vValor, Chr(34)) Or InStr(vValor, Chr(39)) Then
    For i = 1 To Len(vValor)
        If Asc(Mid(vValor, i, 1)) <> 39 And Asc(Mid(vValor, i, 1)) <> 34 Then
            sReturn = sReturn + Mid(vValor, i, 1)
        End If
    Next
Else
   sReturn = vValor
End If
NotChar = sReturn
End Function


Sub FormateaGrid(pGrid As MSDataGridLib.DataGrid)
On Error GoTo hand
        pGrid.MarqueeStyle = dbgHighlightRow
        pGrid.HeadFont.Bold = True
      pGrid.BackColor = -2147483624
        pGrid.Refresh
Exit Sub
hand:
ErrorHandler err, "FormateaGrid"
End Sub


Public Function ExecuteCommandSQL(ByVal Connect As String, ByVal SQL As String) As Long
  On Error GoTo errorx
  Dim oCn As Object

  Set oCn = CreateObject("ADODB.Connection")
  oCn.Open Connect
  oCn.CommandTimeout = 1800
  
  oCn.Execute SQL, ExecuteCommandSQL, 128
  oCn.Close
  Set oCn = Nothing

Exit Function
errorx:

  ExecuteCommandSQL = -2
  err.Raise err.Number, "ExecuteCommandSQL", err.Description
End Function
Public Function GetDataSet(ByVal Connect As String, ByVal SQL As String) As Object 'ADOR.Recordset
  On Error GoTo errorx
  Dim oRs As Object
  Dim oCn As Object
  
 
  Set oRs = CreateObject("ADODB.Recordset")
  Set oCn = CreateObject("ADODB.Connection")
  oCn.CursorLocation = 3
  oCn.Open Connect
  oCn.CommandTimeout = 900
  oRs.Open SQL, oCn, 3, 4
  Set GetDataSet = oRs
  Set GetDataSet.ActiveConnection = oCn
  Set oRs.ActiveConnection = Nothing
  oCn.Close
  Set oCn = Nothing
 
Exit Function
errorx:
  err.Raise err.Number, err.Source, err.Description
  MsgBox err.Description
End Function

Public Function TomaDerechaDesdeCaracter(dato As String, caracter As String) As String
Dim Total As Integer
Dim cantidad As Integer
Total = Len(dato)
cantidad = InStr(1, dato, caracter, vbTextCompare)
TomaDerechaDesdeCaracter = Trim(Right(dato, Val(Total - cantidad)))

End Function
Function StrZero(nDato As Long, nZeros As Integer)
   Dim wdato As String, wAncho As Integer, wDatoOk As String
   Dim i As Integer
   wdato = Trim(Str(nDato))
   wAncho = Len(wdato)
   If wAncho < nZeros Then
      For i = 1 To nZeros - wAncho
      wDatoOk = wDatoOk + "0"
      Next i
      wDatoOk = wDatoOk + wdato
   Else
      wDatoOk = wdato
   End If
   StrZero = wDatoOk
End Function

Public Sub Mensaje(ByVal oCodeMsg As CodeMsg)
Dim aMess(4)
Dim amensaje As clsMessages
Set amensaje = New clsMessages

amensaje.CODIGO = oCodeMsg
'LoadMessage amess, amensaje.Codigo
amensaje.ShowMesage (iLanguage)

End Sub

Sub BuscaCombo1(strTexto As String, intPos As Integer, combo As ComboBox)
    Dim intCont As Integer
    Dim Encontro As Boolean
    Encontro = False
    If intPos = 1 Then
        For intCont = 0 To combo.ListCount - 1
            If strTexto = Mid(combo.List(intCont), 1, Len(strTexto)) Then
                combo.ListIndex = intCont
                Encontro = True
                Exit For
            End If
        Next
    Else
        For intCont = 0 To combo.ListCount - 1
            If strTexto = Right(combo.List(intCont), Len(strTexto)) Then
                combo.ListIndex = intCont
                Encontro = True
                Exit For
            End If
        Next
    End If
    If Encontro = False Then
        combo.ListIndex = -1
    End If
End Sub

Public Function EnLetras(Numero As String) As String

    Dim b, paso As Integer

    Dim expresion, entero, deci, flag As String

    EnLetras = letra(Numero)
    
    If EnLetras <> "" Then Exit Function

    flag = "N"

    For paso = 1 To Len(Numero)

        If Mid(Numero, paso, 1) = "." Then

            flag = "S"

        Else

            If flag = "N" Then

                entero = entero + Mid(Numero, paso, 1) 'Extae la parte entera del numero

            Else

                deci = deci + Mid(Numero, paso, 1) 'Extrae la parte decimal del numero

            End If

        End If

    Next paso

    

    If Len(deci) = 1 Then

        deci = deci & "0"

    End If

    

    flag = "N"

    If Val(Numero) >= -999999999 And Val(Numero) <= 999999999 Then 'si el numero esta dentro de 0 a 999.999.999

        For paso = Len(entero) To 1 Step -1

            b = Len(entero) - (paso - 1)

            Select Case paso

            Case 3, 6, 9

                Select Case Mid(entero, b, 1)

                    Case "1"

                        If Mid(entero, b + 1, 1) = "0" And Mid(entero, b + 2, 1) = "0" Then

                            expresion = expresion & "cien "

                        Else

                            expresion = expresion & "ciento "

                        End If

                    Case "2"

                        expresion = expresion & "doscientos "

                    Case "3"

                        expresion = expresion & "trescientos "

                    Case "4"

                        expresion = expresion & "cuatrocientos "

                    Case "5"

                        expresion = expresion & "quinientos "

                    Case "6"

                        expresion = expresion & "seiscientos "

                    Case "7"

                        expresion = expresion & "setecientos "

                    Case "8"

                        expresion = expresion & "ochocientos "

                    Case "9"

                        expresion = expresion & "novecientos "

                End Select

                

            Case 2, 5, 8

                Select Case Mid(entero, b, 1)

                    Case "1"

                        If Mid(entero, b + 1, 1) = "0" Then

                            flag = "S"

                            expresion = expresion & "diez "

                        End If

                        If Mid(entero, b + 1, 1) = "1" Then

                            flag = "S"

                            expresion = expresion & "once "

                        End If

                        If Mid(entero, b + 1, 1) = "2" Then

                            flag = "S"

                            expresion = expresion & "doce "

                        End If

                        If Mid(entero, b + 1, 1) = "3" Then

                            flag = "S"

                            expresion = expresion & "trece "

                        End If

                        If Mid(entero, b + 1, 1) = "4" Then

                            flag = "S"

                            expresion = expresion & "catorce "

                        End If

                        If Mid(entero, b + 1, 1) = "5" Then

                            flag = "S"

                            expresion = expresion & "quince "

                        End If

                        If Mid(entero, b + 1, 1) > "5" Then

                            flag = "N"

                            expresion = expresion & "dieci"

                        End If

                

                    Case "2"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "veinte "

                            flag = "S"

                        Else

                            expresion = expresion & "veinti"

                            flag = "N"

                        End If

                    

                    Case "3"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "treinta "

                            flag = "S"

                        Else

                            expresion = expresion & "treinta y "

                            flag = "N"

                        End If

                

                    Case "4"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "cuarenta "

                            flag = "S"

                        Else

                            expresion = expresion & "cuarenta y "

                            flag = "N"

                        End If

                

                    Case "5"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "cincuenta "

                            flag = "S"

                        Else

                            expresion = expresion & "cincuenta y "

                            flag = "N"

                        End If

                

                    Case "6"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "sesenta "

                            flag = "S"

                        Else

                            expresion = expresion & "sesenta y "

                            flag = "N"

                        End If

                

                    Case "7"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "setenta "

                            flag = "S"

                        Else

                            expresion = expresion & "setenta y "

                            flag = "N"

                        End If

                

                    Case "8"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "ochenta "

                            flag = "S"

                        Else

                            expresion = expresion & "ochenta y "

                            flag = "N"

                        End If

                

                    Case "9"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "noventa "

                            flag = "S"

                        Else

                            expresion = expresion & "noventa y "

                            flag = "N"

                        End If

                End Select

                

            Case 1, 4, 7

                Select Case Mid(entero, b, 1)

                    Case "1"

                        If flag = "N" Then

                            If paso = 1 Then

                                expresion = expresion & "uno "

                            Else

                                expresion = expresion & "un "

                            End If

                        End If

                    Case "2"

                        If flag = "N" Then

                            expresion = expresion & "dos "

                        End If

                    Case "3"

                        If flag = "N" Then

                            expresion = expresion & "tres "

                        End If

                    Case "4"

                        If flag = "N" Then

                            expresion = expresion & "cuatro "

                        End If

                    Case "5"

                        If flag = "N" Then

                            expresion = expresion & "cinco "

                        End If

                    Case "6"

                        If flag = "N" Then

                            expresion = expresion & "seis "

                        End If

                    Case "7"

                        If flag = "N" Then

                            expresion = expresion & "siete "

                        End If

                    Case "8"

                        If flag = "N" Then

                            expresion = expresion & "ocho "

                        End If

                    Case "9"

                        If flag = "N" Then

                            expresion = expresion & "nueve "

                        End If

                End Select

            End Select

            If paso = 4 Then

                If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or _
                  (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And _
                   Len(entero) <= 6) Then

                    expresion = expresion & "mil "

                End If

            End If

            If paso = 7 Then

                If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then

                    expresion = expresion & "millón "

                Else

                    expresion = expresion & "millones "

                End If

            End If

        Next paso

        

        If deci <> "" Then

            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo

                EnLetras = "menos " & expresion & "con " & deci ' & "/100"

            Else

                EnLetras = expresion & "con " & deci ' & "/100"

            End If

        Else

            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo

                EnLetras = "menos " & expresion

            Else

                EnLetras = expresion

            End If

        End If

    Else 'si el numero a convertir esta fuera del rango superior e inferior

        EnLetras = ""

    End If

End Function


'Sub Busca_Opcion_Anexo1(txtTipo As TextBox, txtCod As TextBox, txtDes As TextBox, frmME As Form, Opcion As String)
''Sub Busca_Opcion_Anexo1(txtCod As TextBox, txtDes As TextBox, frmME As Form)
'
'On Error GoTo Fin
'
'Dim rstAux As Object, strSQL As String
'Set rstAux = CreateObject("ADODB.Recordset")
'    'strSQL = "select cod_tipanex as Tipo,Cod_Anxo as Cod,Des_Anexo as Nombre,Num_Ruc as Ruc from cn_anexoscontables where cod_tipanex like '%" & StrTabla & "%' and "
'    strSQL = "CN_MUESTRA_ANEXOS_CLIENTES '" & txtTipo.Text & "','" & txtCod.Text & "','" & txtDes.Text & "'"
'    'StrSql = "Select " & strCampo1 & " AS Cod," & strCampo2 & " as Descripcion from " & StrTabla
'
'    txtCod = Trim(txtCod)
'    txtDes = Trim(txtDes)
''    Select Case Opcion
''    Case 1: strSQL = strSQL & strCampo1 & " like '%" & txtCod & "%'"
''    Case 2: strSQL = strSQL & strCampo2 & " like '%" & txtDes & "%'"
''    End Select
'
'    'txtCod = ""
'    'txtDes = ""
'    'txttipo = ""
'
'    With FrmBusqClientes
'        Set .oParent = frmME
'        .sQuery = strSQL
'        .Cargar_Datos
'
'        codigo = ""
'        .DGridLista.Columns("Cod").Visible = False
'        .DGridLista.Columns("Tipo").Width = 800
'        .DGridLista.Columns("Nombre").Width = 4075
'        .DGridLista.Columns("RUC").Width = 1200
'        Set rstAux = .DGridLista.ADORecordset
'        If Opcion = "I" Then
'            If rstAux.RecordCount > 1 Then
'              .Show vbModal
'            Else
'              frmME.codigo = ".."
'            End If
'
'        If frmME.codigo <> "" And rstAux.RecordCount > 0 Then
'            frmME.strCod_Anxo = Trim(rstAux!Cod)
'
'            txtDes = Trim(rstAux!Nombre)
'            txtCod = Trim(rstAux!ruc)
'            txtTipo = Trim(rstAux!Tipo)
'            Select Case Opcion
'            Case 1: SendKeys "{TAB}"
'            Case 2: SendKeys "{TAB}"
'            End Select
'        Else
'            SendKeys "{TAB}"
'        End If
'            Unload FrmBusqClientes
'        End If
'    End With
'
'    Set FrmBusqClientes = Nothing
'    rstAux.Close
'    Set rstAux = Nothing
'Exit Sub
'Fin:
'On Error Resume Next
'    Unload FrmBusqClientes
'    Set FrmBusqClientes = Nothing
'    rstAux.Close
'    Set rstAux = Nothing
'    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
'    "Búsqueda de Descuento (" & Opcion & ")"
'End Sub

Sub Busca_Opcion_Anexo1(strCampo1 As String, strCampo2 As String, strTabla As String, txtCod As TextBox, txtDes As TextBox, txttipo As TextBox, Opcion As Integer, frmME As Form)

On Error GoTo fin

Dim rstAux As Object, StrSQL As String
Set rstAux = CreateObject("ADODB.Recordset")
    StrSQL = "select cod_tipanex as Tipo,Cod_Anxo as Cod,Des_Anexo as Nombre,Num_Ruc as Ruc from cn_anexoscontables where cod_tipanex like '%" & strTabla & "%' and "
    
    'StrSql = "Select " & strCampo1 & " AS Cod," & strCampo2 & " as Descripcion from " & StrTabla

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    Select Case Opcion
    Case 1: StrSQL = StrSQL & strCampo1 & " like '%" & txtCod & "%'"
    Case 2: StrSQL = StrSQL & strCampo2 & " like '%" & txtDes & "%'"
    End Select
    txtCod = ""
    txtDes = ""
    txttipo = ""
    
    With frmBusqGeneral
        Set .oParent = frmME
        .SQuery = StrSQL
        .Cargar_Datos
        
        CODIGO = ""
        .gexList.Columns("Cod").Visible = False
        .gexList.Columns("Tipo").Width = 800
        .gexList.Columns("Nombre").Width = 4075
        .gexList.Columns("RUC").Width = 1200
        Set rstAux = .gexList.ADORecordset
        
        If rstAux.RecordCount > 1 Then
          .Show vbModal
        Else
          frmME.CODIGO = ".."
        End If
        
        If frmME.CODIGO <> "" And rstAux.RecordCount > 0 Then
            frmME.strCod_Anxo = Trim(rstAux!cod)
    
            txtDes = Trim(rstAux!Nombre)
            txtCod = Trim(rstAux!ruc)
            txttipo = Trim(rstAux!tipo)
            Select Case Opcion
            Case 1: SendKeys "{TAB}"
            Case 2: SendKeys "{TAB}"
            End Select
        Else
            SendKeys "{TAB}"
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & Opcion & ")"
End Sub

Sub Busca_Opcion_Anexo(strCampo1 As String, strCampo2 As String, strTabla As String, txtCod As TextBox, txtDes As TextBox, Opcion As Integer, frmME As Form)

On Error GoTo fin

Dim rstAux As Object, StrSQL As String
Set rstAux = CreateObject("ADODB.Recordset")
    StrSQL = "select Cod_Anxo as Cod,Des_Anexo as Nombre,Num_Ruc as Ruc from " & strTabla
    
    'StrSql = "Select " & strCampo1 & " AS Cod," & strCampo2 & " as Descripcion from " & StrTabla

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    Select Case Opcion
    Case 1: StrSQL = StrSQL & strCampo1 & " like '%" & txtCod & "%'"
    Case 2: StrSQL = StrSQL & strCampo2 & " like '%" & txtDes & "%'"
    End Select
    txtCod = ""
    txtDes = ""
    scodclienteAne = ""
    frmME.strCod_Anxo = ""
    With frmBusqGeneral
        Set .oParent = frmME
        .SQuery = StrSQL
        .Cargar_Datos
        
        CODIGO = ""
        .gexList.Columns("Cod").Visible = False
        .gexList.Columns("Nombre").Width = 4575
        .gexList.Columns("RUC").Width = 1695
        Set rstAux = .gexList.ADORecordset
        
        If rstAux.RecordCount > 1 Then
          .Show vbModal
        Else
          frmME.CODIGO = ".."
        End If
        If frmME.CODIGO <> "" And rstAux.RecordCount > 0 Then
            frmME.strCod_Anxo = Trim(rstAux!cod)
            txtDes = Trim(rstAux!Nombre)
            txtCod = Trim(rstAux!ruc)
            scodclienteAne = rstAux!cod
            Select Case Opcion
            Case 1: SendKeys "{TAB}"
            Case 2: SendKeys "{TAB}"
            End Select
        Else
            SendKeys "{TAB}"
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & Opcion & ")"
End Sub


'Sub Busca_Opcion_Anexo(strCampo1 As String, strCampo2 As String, strTabla As String, txtCod As TextBox, txtDes As TextBox, Opcion As Integer, frmME As Form)
'
'On Error GoTo Fin
'
'Dim rstAux As Object, STRSQL As String
'Set rstAux = CreateObject("ADODB.Recordset")
'    STRSQL = "select Cod_Anxo as Cod,Des_Anexo as Nombre,Num_Ruc as Ruc from " & strTabla
'
'    'StrSql = "Select " & strCampo1 & " AS Cod," & strCampo2 & " as Descripcion from " & StrTabla
'
'    txtCod = Trim(txtCod)
'    txtDes = Trim(txtDes)
'    Select Case Opcion
'    Case 1: STRSQL = STRSQL & strCampo1 & " like '%" & txtCod & "%'"
'    Case 2: STRSQL = STRSQL & strCampo2 & " like '%" & txtDes & "%'"
'    End Select
'    txtCod = ""
'    txtDes = ""
'    scodclienteAne = ""
'    frmME.strCod_Anxo = ""
'    With frmBusqGeneral
'        Set .oParent = frmME
'        .sQuery = STRSQL
'        .Cargar_Datos
'
'        codigo = ""
'        .DGridLista.Columns("Cod").Visible = False
'        .DGridLista.Columns("Nombre").Width = 4575
'        .DGridLista.Columns("RUC").Width = 1695
'        Set rstAux = .DGridLista.ADORecordset
'
'        If rstAux.RecordCount > 1 Then
'          .Show vbModal
'        Else
'          frmME.codigo = ".."
'        End If
'        If frmME.codigo <> "" And rstAux.RecordCount > 0 Then
'            frmME.strCod_Anxo = Trim(rstAux!Cod)
'            txtDes = Trim(rstAux!Nombre)
'            txtCod = Trim(rstAux!Ruc)
'            scodclienteAne = rstAux!Cod
'            Select Case Opcion
'            Case 1: SendKeys "{TAB}"
'            Case 2: SendKeys "{TAB}"
'            End Select
'        Else
'            SendKeys "{TAB}"
'        End If
'
'    End With
'    Unload frmBusqGeneral
'    Set frmBusqGeneral = Nothing
'    rstAux.Close
'    Set rstAux = Nothing
'Exit Sub
'Fin:
'On Error Resume Next
'    Unload frmBusqGeneral
'    Set frmBusqGeneral = Nothing
'    rstAux.Close
'    Set rstAux = Nothing
'    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
'    "Búsqueda de Descuento (" & Opcion & ")"
'End Sub

Sub Informa(ByVal Mens As String, Optional ByVal amensaje As clsMessages)
If Mens <> "" Then
    Dim rpta As Byte
    rpta = MsgBox(Mens, vbInformation, "Informa")
    Exit Sub
End If
Dim aMess(4)
LoadMessage aMess, amensaje.CODIGO
amensaje.ShowMesage (iLanguage)
End Sub
Public Sub LoadMessage(ByRef aMess As Variant, ByVal iIndex As Integer)
aMess(0) = aMessage(iIndex).tipo
aMess(1) = aMessage(iIndex).Code
If iLanguage = 1 Then
        aMess(2) = aMessage(iIndex).Description
    Else
        aMess(2) = aMessage(iIndex).Description
    End If
aMess(3) = aMessage(iIndex).HelpID
aMess(4) = aMessage(iIndex).Tag
End Sub
Public Function Des_Apos(strDes As String) As String
 Des_Apos = Replace(strDes, "'", "''")
End Function
