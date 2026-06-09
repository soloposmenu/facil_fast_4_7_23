Attribute VB_Name = "Fast_Domicilio"
Option Explicit

Public HAS_Domicilio As Boolean    'ACTIVA BASE DE DOMICILIO
Public nDomicilio As Integer       'MESA INICIAL DE DOMICILIO
Public msConnDomi As ADODB.Connection
'INFO: 24OCT2017
Public nTransDomicilio As Boolean

Public Function OpenDBDomicilio() As Boolean

Set msConnDomi = New ADODB.Connection

On Error GoTo ErrAdm:

' ~~~~~ msConnDomi.CursorLocation = adUseClient
' ~~~~~ msConnDomi.Open "Provider=MSDASQL.1;Driver={Microsoft Access Driver (*.mdb)};Extended Properties" & "=DBQ=" & DATA_PATH & "DOMICILIO.MDB;DriverId=25;FIL=MSAccess;MaxBufferSize=2048;" & "PageTimeout=5;UID=admin;PWD=;"

msConnDomi.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & GetFromINI("General", "DirectorioDatos", App.Path & "\soloini.ini") & "DOMICILIO.mdb;"
OpenDBDomicilio = True
On Error GoTo 0
Exit Function

ErrAdm:
    OpenDBDomicilio = False
    ShowMsg "LA CONFIGURACION DEL SISTEMA ESTA DEFINIDO CON LA OPCION DE DOMICILIO, PERO LA BASE DE DATOS NO EXISTE" & vbCrLf & Err.Number & " - " & Err.description, vbYellow, vbRed
End Function

Public Function CloseDBDomicilio() As Boolean

msConnDomi.Close
Set msConnDomi = Nothing

End Function
Public Function GetMotorizado(nID As Long) As String
'INFO: 23ENE2011
Dim rsMotorizado As ADODB.Recordset
Dim cSQL As String

Set rsMotorizado = New ADODB.Recordset
cSQL = "SELECT NOMBRE, APELLIDO FROM MOTO WHERE ID_MOTO = " & nID
rsMotorizado.Open cSQL, msConnDomi, adOpenStatic, adLockOptimistic
If rsMotorizado.EOF Then
    GetMotorizado = ""
Else
    GetMotorizado = rsMotorizado!nombre & " " & rsMotorizado!APELLIDO
End If
rsMotorizado.Close
Set rsMotorizado = Nothing

End Function
Public Function UpdateDOMINotas(cNotas As String) As Boolean
'UPDATE DATOS DEL PEDIDO, DEBE SER ANTES DE ENVIAR A CHEF
On Error GoTo ErrAdm:
msConnDomi.BeginTrans
msConnDomi.Execute "UPDATE MESA_CLIENTE SET NOTA_ENTREGA1 = '" & Left(UCase(cNotas), 50) & "' WHERE MESA = " & nMesa
msConnDomi.CommitTrans
On Error GoTo 0
Exit Function

ErrAdm:
EscribeLog Err.Number & "-" & Err.description
'msConnDomi.RollbackTrans
On Error GoTo 0
End Function
Public Function MesaAssigned() As String
Dim rsMesaClientes As ADODB.Recordset
'INFO: VERIFICA SI LA MESA/PEDIDO DE DOMICILIO A SIDO ASIGNADA.
'Y REGRESA EL NUMERO DE TELEFONO
DoEvents
Set rsMesaClientes = New ADODB.Recordset

rsMesaClientes.Open "SELECT TELEFONO FROM MESA_CLIENTE WHERE MESA = " & nMesa, msConnDomi, adOpenStatic, adLockReadOnly
If rsMesaClientes.EOF Then
    'MESA NO ASIGNADA
    MesaAssigned = ""
Else
    MesaAssigned = Trim(rsMesaClientes!TELEFONO)
End If
rsMesaClientes.Close
Set rsMesaClientes = Nothing

End Function
Public Function GetDomicilioCliente(nNumMesa As Integer) As String
'INFO: REGRESA TELEFONO, EXTENSION Y NOMBRE DEL CLIENTE
Dim rsDOMI As ADODB.Recordset
Dim cSQL As String
Dim cTempo As String

DoEvents
Set rsDOMI = New ADODB.Recordset

cSQL = "SELECT A.TELEFONO, A.EXTENSION,  B.NOMBRE + ' ' + B.APELLIDO AS CLIENTE_N"
cSQL = cSQL & " FROM MESA_CLIENTE AS A, CLIENTES AS B "
cSQL = cSQL & " WHERE A.MESA = " & nNumMesa
cSQL = cSQL & " AND A.TELEFONO = B.TELEFONO "
cSQL = cSQL & " AND A.EXTENSION = B.EXTENSION"

rsDOMI.Open cSQL, msConnDomi, adOpenStatic, adLockOptimistic
If Not rsDOMI.EOF Then
    cTempo = Format(rsDOMI!TELEFONO, "####-####") & " - " & rsDOMI!EXTENSION & " " & rsDOMI!CLIENTE_N
    cTempo = Left(cTempo, 30)
    GetDomicilioCliente = cTempo
Else
    GetDomicilioCliente = " - - "
End If
rsDOMI.Close
Set rsDOMI = Nothing
End Function
Public Function GetDomicilioInfo() As Variant
'INFO: DOMICILIO, BUSCA CLIENTE Y REGRESA ARRAY DE DATOS DE 2 DIMENSIONES (x,y)
'DONDE y SIEMPRE ES CERO(0)
Dim rsDOMI As ADODB.Recordset
Dim cSQL As String
Dim mEmptyMatrix(14, 0) As Variant

DoEvents
Set rsDOMI = New ADODB.Recordset

cSQL = "SELECT A.MESA AS PEDIDO, B.TELEFONO, B.EXTENSION, "
cSQL = cSQL & " B.NOMBRE, B.APELLIDO, "
cSQL = cSQL & " B.EMAIL, B.EMPRESA, B.ZONA, B.DIRECCION1, "
cSQL = cSQL & " B.DIRECCION2, B.DIRECCION3, A.NOTA_ENTREGA1,"
cSQL = cSQL & " C.DESCRIPCION_CORTA AS ZONA, A.FECHA, A.HORA, "
cSQL = cSQL & " A.ID_MOTO "
cSQL = cSQL & " FROM MESA_CLIENTE AS A, CLIENTES AS B, ZONAS AS C "
cSQL = cSQL & " WHERE A.MESA = " & nMesa
cSQL = cSQL & " AND A.TELEFONO = B.TELEFONO "
'INFO: 24SEP2012
cSQL = cSQL & " AND A.EXTENSION = B.EXTENSION "
cSQL = cSQL & " AND B.ZONA = C.ZONA"

'' =======================================
' ============   MATRIX INFO  ============
'' =======================================
''mEmptyMatrix (0, 0) = MESA o PEDIDO
''mEmptyMatrix (1, 0) = TELEFONO
''mEmptyMatrix (2, 0) = EXTENSION
''mEmptyMatrix (3, 0) = NOMBRE
''mEmptyMatrix (4, 0) = APELLIDO
''mEmptyMatrix (5, 0) = EMAIL
''mEmptyMatrix (6, 0) = EMPRESA
''mEmptyMatrix (7, 0) = ZONA_ID
''mEmptyMatrix (8, 0) = DIRECCION1
''mEmptyMatrix (9, 0) = DIRECCION2
''mEmptyMatrix(10, 0) = DIRECCION3
''mEmptyMatrix(11, 0) = NOTA_ENTREGA1
''mEmptyMatrix(12, 0) = ZONA_DESCRIPCION
''mEmptyMatrix(13, 0) = FECHA
''mEmptyMatrix(14, 0) = HORA
''mEmptyMatrix(15, 0) = ID_MOTORIZADO
'' =======================================

rsDOMI.Open cSQL, msConnDomi, adOpenStatic, adLockReadOnly
If rsDOMI.EOF Then
    'CLIENTES NO ESTA
    GetDomicilioInfo = mEmptyMatrix()
Else
    GetDomicilioInfo = rsDOMI.GetRows
End If
rsDOMI.Close
Set rsDOMI = Nothing

End Function

'---------------------------------------------------------------------------------------
' Procedure : GetDomicilioInfo_B
' Author    : hsequeira
' Date      : 29/09/2012
' Purpose   : 'INFO: DOMICILIO, BUSCA CLIENTE Y REGRESA ARRAY DE DATOS DE 2 DIMENSIONES (x,y)
' DONDE y SIEMPRE ES CERO(0)
' EN VEZ DE BUSCAR EN LA TABLA TEMPORAL BUSCA EN LA TABLA HISTORICA (CLIENTES_TRANS)
'---------------------------------------------------------------------------------------
'
Public Function GetDomicilioInfo_B() As Variant

Dim rsDOMI As ADODB.Recordset
Dim cSQL As String
Dim mEmptyMatrix(14, 0) As Variant

DoEvents
Set rsDOMI = New ADODB.Recordset

cSQL = "SELECT A.MESA AS PEDIDO, B.TELEFONO, B.EXTENSION, "
cSQL = cSQL & " B.NOMBRE, B.APELLIDO, "
cSQL = cSQL & " B.EMAIL, B.EMPRESA, B.ZONA, B.DIRECCION1, "
cSQL = cSQL & " B.DIRECCION2, B.DIRECCION3, A.NOTA_ENTREGA1,"
cSQL = cSQL & " C.DESCRIPCION_CORTA AS ZONA, A.FECHA, A.HORA, "
cSQL = cSQL & " A.ID_MOTO "
cSQL = cSQL & " FROM CLIENTES_TRANS AS A, CLIENTES AS B, ZONAS AS C "
'INFO: 29SEP2012
cSQL = cSQL & " WHERE A.FECHA = '" & Format(Date, "YYYYMMDD") & "'"
cSQL = cSQL & " AND A.MESA = " & nMesa
cSQL = cSQL & " AND A.TELEFONO = B.TELEFONO "
cSQL = cSQL & " AND A.EXTENSION = B.EXTENSION "
cSQL = cSQL & " AND B.ZONA = C.ZONA"

'' =======================================
' ============   MATRIX INFO  ============
'' =======================================
''mEmptyMatrix (0, 0) = MESA o PEDIDO
''mEmptyMatrix (1, 0) = TELEFONO
''mEmptyMatrix (2, 0) = EXTENSION
''mEmptyMatrix (3, 0) = NOMBRE
''mEmptyMatrix (4, 0) = APELLIDO
''mEmptyMatrix (5, 0) = EMAIL
''mEmptyMatrix (6, 0) = EMPRESA
''mEmptyMatrix (7, 0) = ZONA_ID
''mEmptyMatrix (8, 0) = DIRECCION1
''mEmptyMatrix (9, 0) = DIRECCION2
''mEmptyMatrix(10, 0) = DIRECCION3
''mEmptyMatrix(11, 0) = NOTA_ENTREGA1
''mEmptyMatrix(12, 0) = ZONA_DESCRIPCION
''mEmptyMatrix(13, 0) = FECHA
''mEmptyMatrix(14, 0) = HORA
''mEmptyMatrix(15, 0) = ID_MOTORIZADO
'' =======================================

rsDOMI.Open cSQL, msConnDomi, adOpenStatic, adLockReadOnly
If rsDOMI.EOF Then
    'CLIENTES NO ESTA
    GetDomicilioInfo = mEmptyMatrix()
Else
    GetDomicilioInfo = rsDOMI.GetRows
End If
rsDOMI.Close
Set rsDOMI = Nothing

End Function

Public Function UpdateDomiTrans(nTrans As Long, sngTotal As Single) As Boolean
Dim cSQL As String
Dim aDomiInfo() As Variant
Dim nTag As Byte
'INFO: ACTUALZIA EL VALOR PAGADO DE ESTA TRANSACCION.

aDomiInfo = GetDomicilioInfo()

On Error GoTo ErrAdm:

nTag = 0
cSQL = "INSERT INTO CLIENTES_TRANS VALUES ('" & aDomiInfo(1, 0) & "','" & aDomiInfo(2, 0) & "',"
cSQL = cSQL & nTrans & ",'" & Format(Date, "YYYYMMDD") & "','" & Format(Time, "HH:MM") & "'," & nMesa & ","
cSQL = cSQL & sngTotal & ",0, 0, "  'FALSE
cSQL = cSQL & IIf(IsNull(aDomiInfo(15, 0)), 0, aDomiInfo(15, 0)) & ")"

msConnDomi.BeginTrans
msConnDomi.Execute cSQL
msConnDomi.CommitTrans

'DELETE CURRENT MESA
msConnDomi.BeginTrans
nTag = 1
msConnDomi.Execute "DELETE FROM MESA_CLIENTE WHERE MESA = " & nMesa
msConnDomi.CommitTrans
On Error GoTo 0
Exit Function

ErrAdm:
'Resume
EscribeLog Err.Number & " - " & Err.description
EscribeLog cSQL
If nTag = 0 Then
    'msConnDomi.RollbackTrans
    Resume Next
End If
End Function

Public Function DropDomiMesa(nMesaToClose As Integer) As Boolean

'DELETE MESA QUE SE ESTA CERRANDO. 18MAY2011
msConnDomi.BeginTrans
msConnDomi.Execute "DELETE FROM MESA_CLIENTE WHERE MESA = " & nMesaToClose
msConnDomi.CommitTrans

End Function

Public Function FormatPhone(cTelefono As Variant) As String

Select Case Len(cTelefono)
    Case 7
        FormatPhone = Format(cTelefono, "000-0000")
    Case 8
        FormatPhone = Format(cTelefono, "0000-0000")
    Case 10
        FormatPhone = Format(cTelefono, "000-000-0000")
    Case Else
        FormatPhone = cTelefono
End Select
End Function

Public Function GetDomiSales(cReporte As String, Optional ZCounter As Long) As Variant
'INFO: DOMICILIO, BUSCA VENTAS Y MONTO. REGRESA ARRAY DE DATOS DE 2 DIMENSIONES (x,y)
'DONDE (y) SIEMPRE ES CERO(0)
'REGRESA LA CANTIDAD DE VECES QUE SE MARCO DOMICILIO y LOS MONTOS REGISTRADOS DESDE LA ULTIMA Z
Dim rsDOMI As ADODB.Recordset
Dim cSQL As String
Dim mEmptyMatrix(0, 0) As Variant

DoEvents
Set rsDOMI = New ADODB.Recordset

cSQL = "SELECT COUNT(*) AS CANT, SUM(MONTO) AS VENTAS "
cSQL = cSQL & " FROM CLIENTES_TRANS "
cSQL = cSQL & " WHERE IS_PRINTED = 0 "
'cSQL = cSQL & " HAVING SUM(MONTO) > 0 "

rsDOMI.Open cSQL, msConnDomi, adOpenStatic, adLockReadOnly
'On Error Resume Next
'rsDOMI.MoveFirst
'On Error GoTo 0
If rsDOMI.EOF Then
    'CLIENTES NO ESTA
    GetDomiSales = mEmptyMatrix
Else
    GetDomiSales = rsDOMI.GetRows
End If
rsDOMI.Close
Set rsDOMI = Nothing

If cReporte = "Z" Then
    'INFO: SI ES EL REPORTE Z, LIMPIA LAS TRANSACCIONES
    msConnDomi.BeginTrans
    msConnDomi.Execute "UPDATE CLIENTES_TRANS SET IS_PRINTED = -1, Z_COUNTER = " & ZCounter & " WHERE IS_PRINTED = 0"
    msConnDomi.CommitTrans
End If

End Function

Private Function GetFromINI(Section As String, Key As String, Directory As String) As String
Dim strBuffer As String

On Error GoTo FileError:
    strBuffer = String(750, Chr(0))
    Key$ = LCase$(Key$)
    GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
On Error GoTo 0
Exit Function

FileError:
    MsgBox Err.Number + ": NO SE ENCUENTRA ARCHIVO DE INICIALIZACION", vbCritical, "ERROR AL INICIAR"
    Resume Next
End Function


'---------------------------------------------------------------------------------------
' Procedure : PutDomicilioInfo
' Author    : hsequeira
' Date      : 18/01/2017
' Purpose   : CONTABILIZA LAS VENTAS DE DOMICILIO
' Movido desde pagos.
'---------------------------------------------------------------------------------------
'
Public Function PutDomicilioInfo(dNUM_TRANS As Long, dNVALOR_PAGO As Single) As Boolean
Dim cSQL As String
Dim rsDOMI As ADODB.Recordset

Set rsDOMI = New ADODB.Recordset
cSQL = "SELECT SUM(CANT) AS dCANT, SUM(PRECIO) AS dVENTAS FROM TMP_TRANS WHERE DEPTO = " & nHayDomicilio
rsDOMI.Open cSQL, msConn, adOpenStatic, adLockReadOnly

If rsDOMI.EOF Then
    rsDOMI.Close
    Set rsDOMI = Nothing
    PutDomicilioInfo = True
    Exit Function
End If

If IsNull(rsDOMI!dCANT) Then
    rsDOMI.Close
    Set rsDOMI = Nothing
    PutDomicilioInfo = True
    Exit Function
End If

cSQL = "INSERT INTO DOMICILIO (FECHA, HORA, REPORTEZ, CANT, VALOR_DOMICILIO, VALOR_FACTURA, DOC_SOLO) "
cSQL = cSQL & " VALUES ('"
cSQL = cSQL & Format(Date, "YYYYMMDD") & "','" & Format(Time, "HH:MM") & "', 'N', " & rsDOMI!dCANT & "," & rsDOMI!dVENTAS & ","
cSQL = cSQL & dNVALOR_PAGO & "," & dNUM_TRANS & ")"

msConn.Execute cSQL

rsDOMI.Close
Set rsDOMI = Nothing
PutDomicilioInfo = True
Exit Function
    
End Function

