Attribute VB_Name = "FAST00"
Option Explicit

Public Type DOCINFO
    pDocName As String
    pOutputFile As String
    pDatatype As String
End Type

Public Const MUNBYNSeconds = 5

Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Public Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As DOCINFO) As Long
Public Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long
Public lhPrinter As Long
Public rc As Long
Public Const GW_HWNDPREV = 3
Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public FILE_OK As Boolean
Public iISC As Single
Public iISCTransaccion As Single
Public rsISC As New ADODB.Recordset
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpstring As Any, ByVal lpFileName As String) As Long
Public IMPRIME_TICKET As Boolean    'IMPRIME TICKECT AL CLIENTE CON NUMERO DE TRANSACCION
Public IMPRIME_PEDIDO As Boolean    'IMPRIME EL PEDIDO EN LA COCINA
Public DUPLICA_PEDIDO As Boolean    'DUPLICA EL PEDIDO DE LA COCINA EN LA IMPRESORA DE FACTURACION
Public IsFastPrintServer As Boolean     'USA EL PRINT SERVER PARA MULTICAJA (AGO2009)

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)      '(AGO2009)
Public IsInTransaction As Boolean     'INDICA SI EL SISTEMA ESTA EN UNA TRANSACCION SQL
'INFO: DIC2009
Public nNUMTrans As Long        'NUMERO DE TRANSACCION
'INFO: FEB2010
Public Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CLOSE = &H10
Public HayDecimales As Boolean  'INFO: 4OCT2017
Public rsAcomPLU As ADODB.Recordset

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'NUEVO MESSAGEBOX AGOSTO 2010
Public ButtonNumber As Integer

Public Enum msgResult
'     vbtimedout = -1
'     vbCancel = 0
'     vbOK = 1
'     vbRetry = 2
     vbYes = 6
     vbNo = 7
'     vbHelp = 5
'     vbAbort = 6
'     vbIgnore = 7
 End Enum
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 'INFO: 22JUN2012 / 23JUN2012
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public nPLUPassed As Long
Public nDeptoPassed As Long
Public nClienteLlamada As Long
Public nClienteCamina As Long


'---------------------------------------------------------------------------------------
' Procedure : VerificaApp
' Author    : hsequeira
' Date      : 22/06/2012 / 23/06/2012
' Purpose   : EJECUTA EL COMANDO ENVIADO SOBRE LA APLICACION SOLICITADA
'---------------------------------------------------------------------------------------
'
Public Function VerificaApp(cAppSolicitada As String, cComando As String) As Boolean
Dim hWindow As Long
Dim lngResult As Long
Dim lngReturnValue As Long
Dim lngProcess As Long
Dim lngProcessID As Long

'API Constants
Const WM_CLOSE = &H10
Const INFINITE = &HFFFFFFFF
Const SMTO_BLOCK = &H1
Const SMTO_ABORTIFHUNG = &H2
Const WM_NULL = &H0
Const PROCESS_ALL_ACCESS = &H1F0FFF

 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
hWindow = FindWindow(vbNullString, cAppSolicitada)
If hWindow = 0 Then Exit Function   'NO SE HACE NINGUNA VERIFICACION
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

lngReturnValue = PostMessage(hWindow, WM_CLOSE, vbNull, vbNull)
'lngReturnValue = PostMessage(hWindow, WM_DESTROY, vbNull, vbNull)
'ESPERA 2 SEGUNDOS MAXIMO
'lngResult = WaitForSingleObject(hWindow, INFINITE)
lngResult = WaitForSingleObject(hWindow, 3000)

'Does the handle still exist?
DoEvents
hWindow = FindWindow(vbNullString, cAppSolicitada)
If IsWindow(hWindow) = 1 Then
   'The handle still exists. Use the TerminateProcess function
   'to close all related processes to this handle. See the
   'article for more information.
   'MsgBox "Handle still exists."
    lngReturnValue = GetWindowThreadProcessId(hWindow, lngProcessID)
    lngProcess = OpenProcess(PROCESS_ALL_ACCESS, 0&, lngProcessID)
    lngReturnValue = TerminateProcess(lngProcess, 0&)
   
    If (lngReturnValue = 0) Then
       ShowMsg "HAY QUE CERRAR SESION DE WINDOWS PARA OPERAR NORMALMENTE. ERROR DE DLL: " & Err.LastDllError, vbBlue, vbYellow
       ShowMsg "CIERRE LA APLICACION DE VENTAS y LUEGO CIERRE LA SESION DE WINDOWS." & vbCrLf & "VUELVA A ENTRAR NORMALMENTE", vbBlue, vbYellow
       Exit Function
    Else
       'frmTerm02.List1.AddItem "Terminated Application"
    End If
Else
   'Handle does not exist. O YA SE CERRO
   'ShowMsg "LIBRERIA FISCAL ESTA ENCENDIDA, SE PROCEDE A APAGARLA y ENCENDERLA NUEVAMENTE", vbBlue, vbYellow
End If

End Function


Public Function GetPrinterCounter(iPrinter As Long) As String
'Public Function GetPrinterCounter(iPrinter As Integer) As String
Dim rsPrinterCounter As ADODB.Recordset
Dim cSQL As String, CadenaSql As String
Dim nIntentos As Byte
Dim TempKeyValue As Long

TempKeyValue = -1

On Error GoTo FieldNotDefined:
Set rsPrinterCounter = New ADODB.Recordset

cSQL = "SELECT ID, NOMBRE, DESCRIP, CONTADOR FROM IMPRESORES WHERE ID = " & iPrinter

rsPrinterCounter.Open cSQL, msConn, adOpenStatic, adLockOptimistic

If rsPrinterCounter.EOF Then
    EscribeLog ("IMPRESORA (" & (iPrinter - 2) & ") ESTA DEFINIDA EN EL TABLA PLU, PERO NO EXISTE EN LA LISTA DE IMPRESORAS")
    ShowMsg "IMPRESORA (" & (iPrinter - 2) & ") ESTA DEFINIDA EN EL TABLA PLU, PERO NO EXISTE EN LA LISTA DE IMPRESORAS" & ". SE MODIFICO PEDIDO ENVIANDOLO A LA IMPRESORA DE PRE-CUENTA/PEDIDOS"
    EscribeLog ("SE MODIFICO PEDIDO ENVIANDOLO A LA IMPRESORA DE PRE-CUENTA/PEDIDOS")
    ShowMsg "SE MODIFICO PEDIDO ENVIANDOLO A LA IMPRESORA DE PRE-CUENTA/PEDIDOS" & vbCrLf & vbCrLf & "CONTACTE A SOLO POS MENU", vbYellow, vbRed
    cSQL = "SELECT ID, NOMBRE, DESCRIP, CONTADOR FROM IMPRESORES WHERE ID = 2"
    rsPrinterCounter.Close
    rsPrinterCounter.Open cSQL, msConn, adOpenDynamic, adLockOptimistic
End If

'Debug.Print "(" & iPrinter & ") Contador: " & rsPrinterCounter!Contador
Select Case rsPrinterCounter!Contador
    Case "", " "
        'BLANK, NEED TO FIX CONTADOR EN TABLA DE IMPRESORAS
        TempKeyValue = -1
    Case Is = Null
        'NULL, NEED TO FIX CONTADOR EN TABLA DE IMPRESORAS
        TempKeyValue = -1
    Case Else
        If IsNumeric(rsPrinterCounter!Contador) Then
            'OK
        ElseIf IsNull(rsPrinterCounter!Contador) Then
            'BLANK, NEED TO FIX
            TempKeyValue = -1
        Else
            TempKeyValue = -1
        End If
End Select

On Error Resume Next
TempKeyValue = rsPrinterCounter!Contador
On Error GoTo 0

On Error GoTo FieldNotDefined:

Call SOLOTrans("BEGIN")
    CadenaSql = "UPDATE IMPRESORES SET CONTADOR =  " & TempKeyValue + 1 & " WHERE ID = " & iPrinter
    msConn.Execute CadenaSql
Call SOLOTrans("COMMIT")

GetPrinterCounter = Trim(Str(TempKeyValue))
rsPrinterCounter.Close
Set rsPrinterCounter = Nothing

46 On Error GoTo 0
47 Exit Function

FieldNotDefined:
If Err.Number = 3001 Then
    MsgBox ("Ventas. Campo de Contador de Impresoras no esta definido: GetPrinterCounter: " & Err.Number & " - " & Err.description)
    Set rsPrinterCounter = Nothing
    GetPrinterCounter = ""
Else
    If Err.Number = 94 Then
        'INFO: INTENTO DE AUTOCORRECION DEL PROGRAMA
        '26/10/2007
    End If
    If nIntentos < 2 Then
        nIntentos = nIntentos + 1
        Resume
    Else
    End If
End If
If IsInTransaction Then
    Call SOLOTrans("ROLLBACK")
End If
End Function

Public Sub GetNewNumTrans()
'=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/
'                       INFO: DIC2009
'=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/

'INCREMENTA EL NUMERO DE TRANSACCION EN 1
msConn.BeginTrans
msConn.Execute "UPDATE ORGANIZACION SET TRANS_FAST = TRANS_FAST + 1"
msConn.CommitTrans

'LEE LA INFORMACION DE TRANSACCION
rs00.Requery
nNUMTrans = rs00!TRANS_FAST
'=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/
'=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/

End Sub

Public Function GetBoolean(boolValue As Boolean) As Integer
'FOR SQL SERVER THAT DOES NOT SUPPORT THE BOOLEAN / YES/NO DATA TYPE
Select Case boolValue
    Case True
        GetBoolean = -1
    Case False
        GetBoolean = 0
    Case 0
        GetBoolean = 0
    Case -1
        GetBoolean = -1
    Case Else
        GetBoolean = 0
End Select
End Function
Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
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
Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
On Error GoTo FileError:
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
On Error GoTo 0
Exit Sub

FileError:
    MsgBox Err.Number + ": NO SE ENCUENTRA ARCHIVO DE INICIALIZACION", vbCritical, "ERROR AL INICIAR"
    Resume Next
End Sub
Public Function ImprimeTicket(nTrans As Long) As Boolean
Dim i As Byte
Dim cTicket As String
'INFO: AGO2009
'Se anexara al inicio el # de Caja.
'Caja 1 = 1 - 11
'Caja 2 = 2 - 12
'Caja 1 = 1 - 13
'Caja 2 = 2 - 14

cTicket = CStr(nTrans)
cTicket = cCaja & " - " & Right(cTicket, 2)

ImprimeTicket = False

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "BEGIN")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev rs00!DESCRIP
Print2_OPOS_Dev "========================="
Print2_OPOS_Dev Chr(27) & "|3C" & "PEDIDO # " & cTicket
Print2_OPOS_Dev "========================="
Print2_OPOS_Dev Date & Space(2) & Time
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev Space(2)

'DoEvents
For i = 1 To 10
    Print2_OPOS_Dev Space(2)
Next
'DoEvents

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "END")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sleep MUNBYNSeconds
Sys_Pos.Coptr1.CutPaper 100
ImprimeTicket = True
End Function
Public Function SendToPrintServer(nTrans As Long) As Boolean
'INFO: ENVIA DATOS AL PRINT SERVER (AGO2009)
'REEMPLAZA FUNCION ImprimeEnCocina
'OCURRE ANTES DEL ACTUALIZADOR DE TRANSACCIONES(Call Pagos.Actualizador(nCopiaiISCTransaccion))
'DESDE LAS TABLAS TMP_TRANS A TRANSAC
Dim rsConsolidado As ADODB.Recordset
Dim cSQL As String
Dim cConsecutivo As String
Dim nPreviousPrinter As Long
Dim MatrixPrinters(10, 2) As Variant    'MATRIZ FIJA
Dim ArrayPrinters() As Variant      'MARTIZ TEMPORAL
Dim iArray As Long
Dim i As Integer
Dim cError As String
Dim ccLTabla As String

Set rsConsolidado = New ADODB.Recordset

ccLTabla = "TMP_TRANS."
cSQL = "SELECT * FROM TMP_TRANS "
cSQL = cSQL & " WHERE MESA = " & nMesa
cSQL = cSQL & " AND CAJERO = " & npNumCaj
cSQL = cSQL & " ORDER BY IMPRESORA, LIN"

rsConsolidado.Open cSQL, msConn, adOpenStatic, adLockOptimistic

On Error Resume Next

'===---===---===---===---===---===---===
'===---===---===---===---===---===---===
'INFO: DIC2009
'''' ArrayPrinters = NewGetPrinterCounter(rsConsolidado!IMPRESORA + 2)
'USANDO NUM_TRANS COMO INDICE DE IMPRESION
ccLTabla = "Sub.GetPCounter1."
'If nMesa = 0 Then
If nMesa = 120 Then
    ArrayPrinters = MAIN_TRANS_NewGetPrinterCounter(rsConsolidado!IMPRESORA)
Else
    ArrayPrinters = MAIN_TRANS_NewGetPrinterCounter(rsConsolidado!IMPRESORA + 2)
End If
'===---===---===---===---===---===---===
'===---===---===---===---===---===---===

nPreviousPrinter = rsConsolidado!IMPRESORA
MatrixPrinters(iArray, 0) = ArrayPrinters(0, 0)     'TEXT COUNTER
MatrixPrinters(iArray, 1) = ArrayPrinters(0, 1)     'NUMERIC COUNTER
MatrixPrinters(iArray, 2) = ArrayPrinters(0, 2)     'PRINTER NUMBER
cConsecutivo = MatrixPrinters(iArray, 1)
On Error GoTo 0

On Error GoTo ErrAdm:

'INFO: 07OCT2011
'=====> msPED.BeginTrans
'MesasPED "OPEN"

'INFO: 06DIC2011
'AGREGANDO EL NUMERO DE TRANSACCION/FACTURA PARA QUE EL
'FASTPRINTSERRVR LO IMPRIMA EN EL TICKET QUE SE IMPRIME EN LA COCINA
ccLTabla = "TASA_CAMBIO.# TRANS: " & nTrans
'INFO: REACTIVANDO. 19MARZO2015
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'UPDATE: 24JUN2024.
'cSQL = "UPDATE TASA_CAMBIO SET RESPONSABLE = " & nTrans & " WHERE PAIS = 'PANAMA'"
'msPED.Execute cSQL

Do Until rsConsolidado.EOF

    If nPreviousPrinter = rsConsolidado!IMPRESORA Then
        'DO NOTHING. PRINTER IS THE SAME
    Else
        'GET NEW CONSECUTIVE
        '===---===---===---===---===---===---===
        '===---===---===---===---===---===---===
        'INFO: DIC2009
        '''' ArrayPrinters = NewGetPrinterCounter(rsConsolidado!IMPRESORA + 2)
        'USANDO NUM_TRANS COMO INDICE DE IMPRESION
        ccLTabla = "Sub.GetPCounter2."
        'If nMesa = 0 Then
        If nMesa = 120 Then
            ArrayPrinters = MAIN_TRANS_NewGetPrinterCounter(rsConsolidado!IMPRESORA)
        Else
            ArrayPrinters = MAIN_TRANS_NewGetPrinterCounter(rsConsolidado!IMPRESORA + 2)
        End If
        '===---===---===---===---===---===---===
        '===---===---===---===---===---===---===
        nPreviousPrinter = rsConsolidado!IMPRESORA
        iArray = iArray + 1
        MatrixPrinters(iArray, 0) = ArrayPrinters(0, 0)     'TEXT COUNTER
        MatrixPrinters(iArray, 1) = ArrayPrinters(0, 1)     'NUMERIC COUNTER
        MatrixPrinters(iArray, 2) = ArrayPrinters(0, 2)     'PRINTER NUMBER
        cConsecutivo = MatrixPrinters(iArray, 1)
        'ArrayPrinters(iArray) = cConsecutivo
    End If
    '~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~
    'Debug.Print "impresora: " & rsConsolidado!IMPRESORA & " Item: " & rsConsolidado!DESCRIP
    'INFO: 16DIC2014
    If rsConsolidado!IMPRESORA = 0 Then
        'NO HACER NADA, NO ENVIAR EL PEDIDO YA QUE LO IMPRIME EN EL SERVER Y YA LO ESTA IMPRIMIENDO EN EL CONSOLIDADO
    '~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~
    Else
        ccLTabla = "PEDIDO_DETALLE"
        cSQL = "INSERT INTO PEDIDO_DETALLE (IMPRESORA_ORDEN, LIN, CANT, DESCRIPCION) "
        cSQL = cSQL & " VALUES ('"
        cSQL = cSQL & MatrixPrinters(iArray, 0) & "'," & rsConsolidado!LIN & ","
        
        If Mid(LTrim(rsConsolidado!DESCRIP), 1, 2) = "@@" Then
            cSQL = cSQL & rsConsolidado!CANT & ",'" & Space(3) & Mid(rsConsolidado!DESCRIP, 1, 26) & "')"
        Else
            cSQL = cSQL & rsConsolidado!CANT & ",'" & Mid(rsConsolidado!DESCRIP, 1, 26) & "')"
        End If
        'INFO: 7JUL2019. NO SE PASAN AL PEDIDO_DETALLE
        'ANULADOS, ANULACIONES, DESCUENTOS MARCADOS, CORRECION, CORREGIDOS
        Select Case Left(rsConsolidado!Tipo, 3)
            Case "BVO", "VO-", "DC-", "EC-", "BEC"
                'NO HACER NADA
            Case Else
                msPED.Execute cSQL
        End Select
        Rem Debug.Print cSQL
    End If
    '~~~~~~~~~~~~~~~~~~~~~~~~
    rsConsolidado.MoveNext
    '~~~~~~~~~~~~~~~~~~~~~~~~
Loop
'INFO: 07OCT2011
'=====> msPED.CommitTrans
'MesasPED "CLOSE"

rsConsolidado.Close

If cConsecutivo <> "" Then

    For i = 0 To iArray
        'Debug.Print "Impresora: " & MatrixPrinters(i, 2) & " Orden Impresora: " & MatrixPrinters(i, 0)
        '~~~~~~~~~~~~~~~~~~~~~~~~
        '~~~~~~~~~~~~~~~~~~~~~~~~
        'INFO: 16DIC2014
        If MatrixPrinters(i, 2) = 0 Then
            'NO HACER NADA, NO ENVIAR EL PEDIDO YA QUE LO IMPRIME EN EL SERVER Y YA LO ESTA IMPRIMIENDO EN EL CONSOLIDADO
        '~~~~~~~~~~~~~~~~~~~~~~~~
        '~~~~~~~~~~~~~~~~~~~~~~~~
        Else
            ccLTabla = "PEDIDO_MAIN"
            cSQL = "INSERT INTO PEDIDO_MAIN "
            'cSQL = cSQL & "(IMPRESORA_ORDEN, FECHA, HORA, NUM_ORDEN, MESA, MESERO, IS_PRINTED, IMPRESORA) "
            'INFO: 24JUN2024. CAMBIANDO DE UBICACION EL NUMERO DE FACTURA, DE LA TABLE TASA_CAMBIO A PEDIDO_MAIN
            cSQL = cSQL & "(IMPRESORA_ORDEN, FECHA, HORA, NUM_ORDEN, MESA, MESERO, IS_PRINTED, IMPRESORA, FACTURA) "
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            cSQL = cSQL & " VALUES ('"
            cSQL = cSQL & MatrixPrinters(i, 0) & "','" & Format(Date, "YYYYMMDD") & "','" & Format(Time, "HH:MM:SS") & "','"
            'cSQL = cSQL & MatrixPrinters(i, 1) & "'," & nMesa & ",'" & "CAJA - " & cCaja & "'," & 0 & "," & MatrixPrinters(i, 2) & " )"
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            cSQL = cSQL & MatrixPrinters(i, 1) & "'," & nMesa & ",'" & "CAJA - " & cCaja & "'," & 0 & "," & MatrixPrinters(i, 2) & "," & nTrans & " )"
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            '(0) ES LA IMPESORA DE FACTURACION, YA QUE ESTE ES EL CONSOLIDADO
            
            ' =====>> msPED.BeginTrans
            'MesasPED "OPEN"
            msPED.Execute cSQL
            'MesasPED "CLOSE"
            ' =====>> msPED.CommitTrans
            Rem Debug.Print cSQL
        End If
    Next
End If
'MesasPED "CLOSE"
On Error GoTo 0
Exit Function

ErrAdm:
''''''Err.Number
''''''-2147217833
''''''Print Err.Description
''''''El campo es demasiado pequeño para aceptar la cantidad de datos que intenta agregar. Intente insertar o pegar menos datos.
    cError = Err.Number & " - " & Err.description & Space(1) & ccLTabla & Space(1) & cSQL
    EscribeLog "Ventas. Error. SendToPrintServer: " & cError
    EscribeLog "Ventas. Error. SendToPrintServer: " & cSQL
    Debug.Print cError
    ShowMsg "Error. SendToPrintServer " & vbCrLf & MatrixPrinters(i, 0) & cError & vbCrLf & ccLTabla & vbCrLf & "CONTACTE A SOLO POS MENU" & vbCrLf & LCase(cSQL), vbRed, vbYellow
     Resume Next
End Function
Public Function DuplicaPedido(nTrans As Long) As Boolean
Dim cSQL As String
Dim rsCocina As ADODB.Recordset
Dim i As Byte


DuplicaPedido = False

On Error GoTo ErrHandler:

Set rsCocina = New ADODB.Recordset
cSQL = "SELECT * FROM TMP_TRANS WHERE MESA = " & nMesa
cSQL = cSQL & " AND CAJERO = " & npNumCaj
cSQL = cSQL & " ORDER BY LIN"

rsCocina.Open cSQL, msConn, adOpenStatic, adLockOptimistic

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "BEGIN")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Print2_OPOS_Dev Space(2)
'INFO: AGO2009
'Print2_OPOS_Dev "PEDIDO # " & nTrans
'Print2_OPOS_Dev "PEDIDO # " & cCaja & " - " & nTrans
'info: 15NOV2017
Print2_OPOS_Dev "PEDIDO # " & cCaja & " - " & Right(Trim(Str(nNUMTrans)), 2)

If nTransDomicilio Then
    Dim aDomiInfo() As Variant
    aDomiInfo = GetDomicilioInfo()
    Print2_OPOS_Dev String(33, "~")
    Print2_OPOS_Dev "= DOMICILIO = "
    Print2_OPOS_Dev GetDomicilioCliente(nMesa)
    Print2_OPOS_Dev "ZONA: " & aDomiInfo(12, 0)
    Print2_OPOS_Dev "DIR: " & aDomiInfo(8, 0)
    Rem Print2_OPOS_Dev aDomiInfo(9, 0)
    Rem Print2_OPOS_Dev aDomiInfo(10, 0)
    Print2_OPOS_Dev String(33, "~")
End If

'Print2_OPOS_Dev Format(Date, "LONG DATE")
Print2_OPOS_Dev Format(Date, "dddd, dd/mmm/yyyy")
Print2_OPOS_Dev Time
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev "SOLICITUD DE PLATOS"
Print2_OPOS_Dev "------------------------------"

Do While Not rsCocina.EOF
Rem Do Until rsCocina.EOF
    If Mid(LTrim(rsCocina!DESCRIP), 1, 2) = "@@" Then
        Print2_OPOS_Dev Space(3) & Mid(rsCocina!DESCRIP, 1, 26)
    Else
        If Sys_Pos.Coptr1.CapRecBold = True Then
            Print2_OPOS_Dev Chr(27) & "|3C" & Format(rsCocina!CANT, "##") & Space(2) & Chr(&H1B) + "|bC" + Mid(rsCocina!DESCRIP, 1, 26)
        Else
            Print2_OPOS_Dev Chr(27) & "|3C" & Format(rsCocina!CANT, "##") & Space(2) & Mid(rsCocina!DESCRIP, 1, 26)
        End If
    End If
    rsCocina.MoveNext
Loop

For i = 1 To 10
    Print2_OPOS_Dev Space(2)
Next

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "END")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sleep MUNBYNSeconds
Sys_Pos.Coptr1.CutPaper (100)


rsCocina.Close
Set rsCocina = Nothing

DuplicaPedido = True
Exit Function

ErrHandler:
    ShowMsg "ERROR GRAVE (TOME FOTO) " & vbCrLf & Err.Number & " - " & Err.description & vbCrLf & "EN  Function DuplicaPedido()", vbYellow, vbRed
    ShowMsg "SALGA DEL PROGRAMA Y VUELVA A ENTRAR", vbYellow, vbRed

On Error GoTo 0
End Function
Public Sub ActivatePrevInstance()

Dim OldTitle As String
Dim PrevHndl As Long
Dim result As Long

'Save the title of the application.
OldTitle = App.Title

'Rename the title of this application so FindWindow
'will not find this application instance.
App.Title = "Instancia de App No Deseada"

'Attempt to get window handle using VB4 class name.
PrevHndl = FindWindow("ThunderRTMain", OldTitle)

'Check for no success.
If PrevHndl = 0 Then
   'Attempt to get window handle using VB5 class name.
   PrevHndl = FindWindow("ThunderRT5Main", OldTitle)
End If

'Check if found
If PrevHndl = 0 Then
    'Attempt to get window handle using VB6 class name
    PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
End If

'Check if found
If PrevHndl = 0 Then
   'No previous instance found.
   Exit Sub
End If

'Get handle to previous window.
PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)

'Restore the program.
result = OpenIcon(PrevHndl)

'Activate the application.
result = SetForegroundWindow(PrevHndl)

'End the application.
End
End Sub
Public Sub MuestraProductosMarcados()
Dim cSQL As String
Dim nVarTemp As Single

'INFO: CHANGES TO SQL SERVER
'ENE/2010

'rs07.Close

If rs07.State = adStateOpen Then
    rs07.Close
End If

cSQL = "SELECT A.LIN,A.DESCRIP,A.CANT,"
cSQL = cSQL & " FORMAT(PRECIO_UNIT,'##0.00') AS MPRECIO_UNIT,"
'CSQL = CSQL & " CAST(PRECIO_UNIT AS VARCHAR(12)),"
cSQL = cSQL & " FORMAT(PRECIO,'##0.00') AS MPRECIO,"
'CSQL = CSQL & " CAST(PRECIO AS VARCHAR(12)),"
cSQL = cSQL & " A.CAJERO,A.MESERO,A.DEPTO,A.PLU,A.MESA,A.VALID, "
cSQL = cSQL & " A.ENVASE,A.FECHA,A.HORA,A.TIPO,A.DESCUENTO,A.CUENTA, "
cSQL = cSQL & " A.CAJA "
cSQL = cSQL & " FROM TMP_TRANS AS A "
cSQL = cSQL & " WHERE A.MESA = " & nMesa
cSQL = cSQL & " AND A.CAJERO = " & npNumCaj
cSQL = cSQL & " ORDER BY A.LIN"

rs07.Open cSQL, msConn, adOpenStatic, adLockOptimistic

Set PLU.PlatosMesa.DataSource = rs07
CajLin = rs07.RecordCount
If PLU.PlatosMesa.Rows <> 0 Then
    PLU.PlatosMesa.TopRow = (PLU.PlatosMesa.Rows - 1)
End If
PLU.SetupPantalla
rs07.Close

'rs07.Open "SELECT sum(a.precio) as precio FROM TMP_TRANS as a  WHERE a.MESA = " & nMesa & " AND a.CAJERO = " & npNumCaj, msConn, adOpenStatic, adLockReadOnly
'INFO: JULIO 2010
'////////////////////////////////////////////////////////////////
'ABAJO ESTA TOMANDO UN IMPUESTO GENERAL y SE LO APLICA A TODOS LOS PRODUCTOS
'EL IMPUESTO DEBE SER CALCULADO x PRODUCTO
cSQL = "SELECT sum(a.precio) AS PRECIO, "
cSQL = cSQL & " sum(a.precio * (a.con_tax/100)) as tax"
cSQL = cSQL & " FROM TMP_TRANS as a  "
cSQL = cSQL & " WHERE a.MESA = " & nMesa & " AND a.CAJERO = " & npNumCaj

rs07.Open cSQL, msConn, adOpenStatic, adLockReadOnly
'////////////////////////////////////////////////////////////////

On Error Resume Next
PLU.SubTot = FormatCurrency(IIf(IsNull(rs07!precio), 0#, rs07!precio), 2)

'INFO: JULIO 2010
'////////////////////////////////////////////////////////////////
'PLU.SubTot = FormatCURRENCY((PLU.SubTot + (rs07!precio * iISC)), 2)
'INFO: UPDATE 1MAY2018. AGREGANDO LOS MONTOS DEL SUB TOTAL e IMPUESTO
PLU.lbNOTax = Format(PLU.SubTot, "STANDARD")
nVarTemp = rs07!tax
PLU.SubTot = FormatCurrency((PLU.SubTot + rs07!tax), 2)
'iISCTransaccion = rs07!precio * iISC
iISCTransaccion = rs07!tax
PLU.lbTAX.Caption = Format(iISCTransaccion, "STANDARD")
'////////////////////////////////////////////////////////////////

SBTot = Format(PLU.SubTot, "STANDARD")
On Error GoTo 0
'24NOV2011
'INFO: AVECES MANDA ERRORES. DICE QUE YA ESTA CERRADO
'' =====>> REMOVED DoEvents
rs07.Close
End Sub

''''Public Sub PutISC(nISC As Single)
''''On Error Resume Next
''''rsISC.Open "SELECT DIARIO FROM ISC", msConn, adOpenDynamic, adLockOptimistic
''''rsISC!DIARIO = rsISC!DIARIO + nISC
''''rsISC.Update
''''rsISC.Close
''''On Error GoTo 0
''''End Sub
Public Sub PutISC(nISC As Single)
'INFO: ACTUALIZACION PARA QUE EL IMPUESTO SEA CALCULADO ANUALMENTE, SIEMPRE CON LA FECHA DE LA CAJA
'INFO: UPDATE 24NOV2011
Dim cSQL As String

On Error Resume Next
'rsISC.Open "SELECT DIARIO FROM ISC ", msConn, adOpenDynamic, adLockOptimistic
cSQL = "SELECT DIARIO FROM ISC WHERE ISC_YEAR = " & Year(Date)
rsISC.Open cSQL, msConn, adOpenDynamic, adLockOptimistic
rsISC!DIARIO = rsISC!DIARIO + nISC
rsISC.Update
''Debug.Print "Trans: " & rs00!TRANS + 1 & " - ISC In: " & nISC & ", ISC Acum: " & rsISC!DIARIO
''Debug.Print "============================================"
rsISC.Close
On Error GoTo 0
End Sub

Private Function MAIN_TRANS_NewGetPrinterCounter(iPrinter As Long) As Variant
Dim cSQL As String
Dim TempKeyValue As Long
Dim cPrinterPREFIX As String
Dim LocalArrayPrinters(0, 2) As Variant
Dim bFixPrinterCounter As Boolean
Dim nIntentos As Byte

On Error GoTo ErrAdm:

Select Case iPrinter
    'INFO: IMPRESORA ADDED. 16DIC2014
    Case 1
        cPrinterPREFIX = Trim(Str(cCaja)) & " - FACTURA1_"
    Case 2
        cPrinterPREFIX = Trim(Str(cCaja)) & " - FACTURACION_"
    Case 3
        cPrinterPREFIX = Trim(Str(cCaja)) & " - COCINA1_"
    Case 4
        cPrinterPREFIX = Trim(Str(cCaja)) & " - BAR1_"
    Case 5
        cPrinterPREFIX = Trim(Str(cCaja)) & " - COCINA2_"
    Case 6
        cPrinterPREFIX = Trim(Str(cCaja)) & " - COCINA3_"
    Case Else
        cPrinterPREFIX = Trim(Str(cCaja)) & " - EXTRA_"
End Select

TempKeyValue = GetPrinterCounter(iPrinter)

'INFO: AGO(2009)
LocalArrayPrinters(0, 0) = cPrinterPREFIX & Trim(Str(TempKeyValue))
LocalArrayPrinters(0, 1) = Trim(Str(TempKeyValue))
LocalArrayPrinters(0, 2) = iPrinter
MAIN_TRANS_NewGetPrinterCounter = LocalArrayPrinters

On Error GoTo 0
Exit Function

ErrAdm:
MsgBox Err.Number & " - " & Err.description, vbCritical, "INFO"
End Function



Private Function RandomNumber(nInicio As Long, nFinal As Long) As Long
Randomize   ' Initialize random-number generator.
RandomNumber = Int((nFinal * Rnd) + nInicio)   ' Generate random value between 150-350
End Function

Public Function SOLOTrans(cTipo As String) As Boolean
'Public Function SOLOTrans(cDBMS As String, cTipo As String) As Boolean
'04/03/2007
'INFO: cTipo es la instruccion para las transacciones
'Activa la variable InTransaction y le da valor
'me sirve para hacer Rollback
'PARA SOLO.MDB, reemplaza msconn.begintransaction y Call SOLOTrans("COMMIT")
On Error GoTo ErrAdm:
    Select Case cTipo
        Case "BEGIN"
            msConn.BeginTrans
            IsInTransaction = True
        Case "COMMIT"
            msConn.CommitTrans
            IsInTransaction = False
        Case "ROLLBACK"
            msConn.RollbackTrans
            IsInTransaction = False
    End Select

On Error GoTo 0
Exit Function

ErrAdm:
    EscribeLog ("ERROR EN SOLOTrans, INSTRUCCION: " & cTipo)
End Function

Public Function ShowMsg(cmsg As String, Optional oFontColor As Long, Optional oBackColor As Long, Optional Botones As Integer) As msgResult
'MENSAJE_DEL_SISTEMA = cMsg
' BOTONES: vbYes Y vbYesNo
Dim TheForm As Form

Load Mensaje

Select Case Botones
    Case vbOKOnly
        Mensaje.cmdAceptar.Enabled = True
    Case vbYesNo
        Mensaje.cmdAceptar.Enabled = True
        Mensaje.cmdCancelar.Enabled = True
        Mensaje.cmdAceptar.Visible = True
        Mensaje.cmdCancelar.Visible = True
    Case Else
        Mensaje.cmdAceptar.Enabled = True
End Select

If oBackColor = 0 Then oBackColor = Mensaje.BackColor
If oBackColor = oFontColor Then oBackColor = Mensaje.BackColor
8
If oFontColor = vbWhite Or oFontColor = 0 Then
Else
    Mensaje.BackColor = oBackColor
    Mensaje.lbMensaje.BackColor = oBackColor
    Mensaje.lbMensaje.ForeColor = oFontColor
End If

Mensaje.lbMensaje.Caption = cmsg
Mensaje.Show 1

ShowMsg = ButtonNumber
Unload Mensaje
End Function

Public Function GetEnteroFromString(cCadena As String) As String
Dim cVariable As String
Dim nLargo As Integer, i As Integer
Dim bDataFound As Boolean
'24 jul 2010
'EXTRAE EL NUMERO ENTERO DE UNA CADENA, RECORRE TODA LA CADENA

nLargo = Len(cCadena)
For i = 1 To nLargo
    Select Case Mid(cCadena, i, 1)
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            cVariable = cVariable & Mid(cCadena, i, 1)
            bDataFound = True
    End Select
Next
If bDataFound Then
    GetEnteroFromString = cVariable
Else
    GetEnteroFromString = ""
End If
End Function

'ACTUALIZACION EN MAYO2014 PARA IMPLENTAR NUEVAS FUNCIONES DEL OPOS
'---------------------------------------------------------------------------------------
' Procedure : PUT_AsyncMode
' Author    : hsequeira
' Date      : 08/12/2012
' Purpose   : Para la impresora de Precuenta.
' Si es una Impresora Paralela, el BUFFER de impresion es mayor que el de las seriales.
' Si es PARALELA poner AsyncMode = True, de lo contrario AsyncMode = false
' 06JUN2013: Agregando SRP-350plus termica
'
' 11NOV2013:
' By setting the AsyncMode property to FALSE and running the PrintNormal method,
' data is printed as it is sent and will not release control of the printer until the data is completely sent.
'
' 11NOV2013:
' Asynchronous output, especially when sending large amounts of data, is more efficient.
' Asynchronous output allows printing to happen in the background and thus has less effect on
' keyboard and mouse operations in a POS application. By using asynchronous printing, more
' user-friendly POS applications can be built.
' The PrintNormal method is used to send data to the printer. By setting the AsyncMode property
' to TRUE and running the PrintNormal method, data is printed asynchronously. Printing data is
' stored in a buffer and control is returned to the application immediately.
'
' 2MAY2014, IMPLEMENTANDO EN EL SISTEMA FAST
'---------------------------------------------------------------------------------------
'
Public Function PUT_AsyncMode() As Boolean

Select Case OPOS_DevName
    Case "SRP-350plus"      'INFO: 23SEP2013
        'SI SE PONE TRUE, CAUSA UN FALLO GENERAL
        Sys_Pos.Coptr1.AsyncMode = False
    'INFO: 21ABRIL2014. UPDATE PARA BIXOLON OPOS NUEVO
    'INFO: 28ABRIL2014. PARA LAS T-20 MAS NUEVAS."T20-42C" y "T-20II-42C"
    Case "LR3000", "TM-U220B", "TM-U200B", "SRP270", "SRP270P", "MP4200TH", "TM-T20E", "TM-T20U", "TM-U220B", "SEMOPOS.SO.SERIAL.POSPrinter", "SRP-275", "SRP-275P", "SRP-270", "SRP-270P", "TM-T20-42CU", "TM-T20-42CE", "TM-T20II-42CU", "TM-T20II-42CE"
        Sys_Pos.Coptr1.AsyncMode = True
    Case "TM-U950P", "TM-U950"
        Sys_Pos.Coptr1.AsyncMode = False
    Case "HKA80"    'INFO: 6ABR2018
        Sys_Pos.Coptr1.AsyncMode = False
    Case "SRP-E300"    'INFO: 5JUN2019
        Sys_Pos.Coptr1.AsyncMode = True
End Select

If Sys_Pos.Coptr1.ResultCode <> 0 Then
    EscribeLog "NO SE PUEDE ESTABLECER (AsyncMode) SE DEJARA EN FALSO"
    Sys_Pos.Coptr1.AsyncMode = False
    PUT_AsyncMode = False
Else
    PUT_AsyncMode = True
End If


''''Select Case OPOS_DevName
''''    Case "LR3000", "TM-U950", "TM-U220B", "TM-U200B", "SRP270", "MP4200TH"
''''        Sys_Pos.Coptr1.AsyncMode = True
''''    Case "SEMOPOS.SO.SERIAL.POSPrinter"
''''        If Sys_Pos.Coptr1.DeviceDescription = "SEMOPOS POS Printer [PARALLEL]" Then
''''            Sys_Pos.Coptr1.AsyncMode = True
''''        Else
''''            Sys_Pos.Coptr1.AsyncMode = False
''''        End If
''''    Case "TM-U950P", "TM-U220BP", "TM-U200BP", "SRP270P"
''''        Sys_Pos.Coptr1.AsyncMode = True
''''    Case Else
''''        ' AQUI VAN LAS IMPRESORAS MAS RAPIDAS
''''         Sys_Pos.Coptr1.AsyncMode = True
''''End Select

End Function


'---------------------------------------------------------------------------------------
' Procedimiento : OPOSTransactionPrint
' Autor       : hsequeira
' Fecha       : 24/10/2013
' Proposito   : ACTIVA LA TRANSACCION DE IMPRESION
' SI LA IMPRESORA TIENE LA CAPACIDAD DE IMPRIMIR EN TRANSACCION
'(QUE ACUMULA TODO EN EL BUFFER E IMPRIME CUANDO SE ENVIA EL TRANSACTION END)
'---------------------------------------------------------------------------------------
'
Public Function OPOSTransactionPrint(ccPrinterName As String, cOption As String, Optional DoJournal As Integer) As Boolean
Dim cDoJournal As String
If DoJournal = 0 Then cDoJournal = "NO" Else cDoJournal = "SI"
'Sys_Pos.Coptr1.ClearOutput
Select Case ccPrinterName
    Case "Coptr1"
        If Sys_Pos.Coptr1.CapTransaction Then
            If cOption = "BEGIN" Then
                If cDoJournal = "SI" Then
                    Sys_Pos.Coptr1.TransactionPrint PtrSJournal, PtrTpTransaction
                Else
                    Sys_Pos.Coptr1.TransactionPrint PtrSReceipt, PtrTpTransaction
                End If
            Else
                If cDoJournal = "SI" Then
                    Sys_Pos.Coptr1.TransactionPrint PtrSReceipt, PtrTpNormal
                Else
                    Sys_Pos.Coptr1.TransactionPrint PtrSReceipt, PtrTpNormal
                End If
            End If
        End If
'    Case "ImprBarra"
'        If Sys_Pos.ImprBarra.CapTransaction Then
'            If cOption = "BEGIN" Then
'                Sys_Pos.ImprBarra.TransactionPrint PtrSReceipt, PtrTpTransaction
'            Else
'                Sys_Pos.ImprBarra.TransactionPrint PtrSReceipt, PtrTpNormal
'            End If
'        End If
'    Case "ImprCocina"
'        If Sys_Pos.ImprCocina.CapTransaction Then
'            If cOption = "BEGIN" Then
'                Sys_Pos.ImprCocina.TransactionPrint PtrSReceipt, PtrTpTransaction
'            Else
'                Sys_Pos.ImprCocina.TransactionPrint PtrSReceipt, PtrTpNormal
'            End If
'        End If
'    Case "ImprCocina02"
'        If Sys_Pos.ImprCocina02.CapTransaction Then
'            If cOption = "BEGIN" Then
'                Sys_Pos.ImprCocina02.TransactionPrint PtrSReceipt, PtrTpTransaction
'            Else
'                Sys_Pos.ImprCocina02.TransactionPrint PtrSReceipt, PtrTpNormal
'            End If
'        End If
'    Case "ImprCocina03"
'        If Sys_Pos.ImprCocina03.CapTransaction Then
'            If cOption = "BEGIN" Then
'                Sys_Pos.ImprCocina03.TransactionPrint PtrSReceipt, PtrTpTransaction
'            Else
'                Sys_Pos.ImprCocina03.TransactionPrint PtrSReceipt, PtrTpNormal
'            End If
'        End If
End Select

End Function



'---------------------------------------------------------------------------------------
' Procedure : PrintImage
' Author    : hsequeira
' Date      : 21/03/2024
' Purpose   : IMPRIME IMAGEN HORIZONTAL CENTRADA (BMP)
'---------------------------------------------------------------------------------------
'
Public Sub PrintImage(ByVal imagePath As String)
    
    
    ' Load the image
    Dim Image As Object
    Set Image = LoadPicture(imagePath)
    
    ' Calculate the width of the image
    Dim imageWidth As Integer
    imageWidth = Image.Width / Screen.TwipsPerPixelX  ' Convert from Twips to pixels
    
    ' Calculate horizontal position to center the image
    Dim centerX As Integer
    Dim receiptWidth As Integer
    receiptWidth = Sys_Pos.Coptr1.RecLineWidth
    centerX = (receiptWidth - imageWidth) / 3
    
    ' Print the image
    'OPOSPOSPrinter1.TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL
    ''OPOSPOSPrinter1.PrintNormal PTR_S_RECEIPT, Chr(27) & "|L" & centerX & "C"
    ''OPOSPOSPrinter1.PrintBitmap PTR_S_RECEIPT, imagePath, centerX, PTR_BM_CENTER
    
    'OPOSPOSPrinter1.PrintNormal PTR_S_RECEIPT, Chr(27) & "|L" & centerX & "C"  ' Move to centerX position
    Sys_Pos.Coptr1.SetBitmap 1, PtrSReceipt, imagePath, Sys_Pos.Coptr1.RecLineWidth / 2, PtrBmCenter
    'Sys_Pos.Coptr1.PrintBitmap PtrSReceipt, imagePath, PtrBmAsis&, PtrBcCenter
    
    '.PrintBarCode PtrSReceipt, FE_QRData, PTR_BCS_QRCODE, .RecLineHeight * 0.6, .RecLineWidth * 0.6, PtrBcCenter, PtrBcTextNone&
    
    'OPOSPOSPrinter1.PrintNormal PTR_S_RECEIPT, Chr(27) & "|L" & 0 & "C"  '
    
    
    
    'OPOSPOSPrinter1.TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL
    
    If Err.Number <> 0 Then
        'ShowMsg "Error AL IMPRIMIR LOGO: " & Err.description, vbYellow, vbRed
        EscribeLog "Error AL IMPRIMIR LOGO: " & Err.description
    End If
End Sub

Public Function GetSpace(cTexto As String) As String
Dim spaces As String

If lineWidth > Len(cTexto) Then
    spaces = Space((lineWidth - Len(cTexto)) \ 2)
Else
    spaces = ""
End If

GetSpace = spaces
End Function
