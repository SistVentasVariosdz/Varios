Attribute VB_Name = "ModGeneral"

Global conn As New ADODB.Connection

Public Sub HabilitaMant(ctl As Object, botones As String)
    ctl.FunctionsUser = botones
    'HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
End Sub


Public Sub Busca_Opcion(strCampo1 As String, strCampo2 As String, strTabla As String, txtCod As TextBox, txtDes As TextBox, Opcion As Integer, frmME As Form)
    On Error GoTo Fin
    Dim rstAux As ADODB.Recordset, strSQL As String

    strSQL = "Select " & strCampo1 & " AS Cod," & strCampo2 & " as Descripcion from " & strTabla

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    Select Case Opcion
        Case 1: strSQL = strSQL & strCampo1 & " like '%" & txtCod & "%'"
        Case 2: strSQL = strSQL & strCampo2 & " like '%" & txtDes & "%'"
    End Select
    txtCod = ""
    txtDes = ""
    With frmBusqGeneral
        Set .oParent = frmME
        .sQuery = strSQL
        .CARGAR_DATOS
        
        frmME.CODIGO = ""
        Set rstAux = .gexList.ADORecordset
        If .gexList.RowCount > 0 Then
          .Show vbModal
        Else
          frmME.CODIGO = ".."
        End If
        
        If frmME.CODIGO <> "" And rstAux.RecordCount > 0 Then
            txtCod = frmME.CODIGO 'Trim(rstAux!Cod)
            txtDes = frmME.DESCRIPCION 'Trim(rstAux!DESCRIPCION)
            Select Case Opcion
            Case 1: SendKeys "{TAB}": SendKeys "{TAB}"
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
Resume
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    'rstAux.Close
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & Opcion & ")"
End Sub

Public Sub Busca_Opcion_Cuenta(strCampo1 As String, strCampo2 As String, strTabla As String, txtCod As TextBox, txtDes As TextBox, Opcion As Integer, frmME As Form)

On Error GoTo Fin

Dim rstAux As ADODB.Recordset, strSQL As String

    strSQL = "select Cod = Sec_Cuenta_Banco,Descripcion = Cod_Moneda + '-' + Cod_Cuenta from " & strTabla

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    Select Case Opcion
    Case 1: strSQL = strSQL & strCampo1 & " like '%" & txtCod & "%'"
    Case 2: strSQL = strSQL & strCampo2 & " like '%" & txtDes & "%'"
    End Select
    txtCod = ""
    txtDes = ""
    With frmBusqGeneral
        Set .oParent = frmME
        .sQuery = strSQL
        .CARGAR_DATOS
        
        frmME.CODIGO = ""
        Set rstAux = .gexList.ADORecordset
        If rstAux.RecordCount > 1 Then
          .Show vbModal
        Else
          frmME.CODIGO = ".."
        End If
        
        If frmME.CODIGO <> "" And rstAux.RecordCount > 0 Then
            txtCod = Trim(frmME.CODIGO) 'Trim(rstAux!Cod)
            txtDes = Trim(frmME.DESCRIPCION)  ' Trim(rstAux!DESCRIPCION)
            Select Case Opcion
            Case 1: SendKeys "{TAB}": SendKeys "{TAB}"
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
Resume
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & Opcion & ")"
End Sub

Public Function fecha(ByVal mes As String)
    If iLanguage <> "1" Then
        If mes = "01" Then
            fecha = "January"
        ElseIf mes = "02" Then
            fecha = "February"
        ElseIf mes = "03" Then
            fecha = "March"
        ElseIf mes = "04" Then
            fecha = "April"
        ElseIf mes = "05" Then
            fecha = "May"
        ElseIf mes = "06" Then
            fecha = "June"
        ElseIf mes = "07" Then
            fecha = "July"
        ElseIf mes = "08" Then
            fecha = "August"
        ElseIf mes = "09" Then
            fecha = "September"
        ElseIf mes = "10" Then
            fecha = "October"
        ElseIf mes = "11" Then
            fecha = "November"
        ElseIf mes = "12" Then
            fecha = "December"
        End If
    Else
        If mes = "01" Then
            fecha = "Enero"
        ElseIf mes = "02" Then
            fecha = "Febrero"
        ElseIf mes = "03" Then
            fecha = "Marzo"
        ElseIf mes = "04" Then
            fecha = "Abril"
        ElseIf mes = "05" Then
            fecha = "Mayo"
        ElseIf mes = "06" Then
            fecha = "Junio"
        ElseIf mes = "07" Then
            fecha = "Julio"
        ElseIf mes = "08" Then
            fecha = "Agosto"
        ElseIf mes = "09" Then
            fecha = "Setiembre"
        ElseIf mes = "10" Then
            fecha = "Octubre"
        ElseIf mes = "11" Then
            fecha = "Noviembre"
        ElseIf mes = "12" Then
            fecha = "Diciembre"
        End If
    End If
End Function


