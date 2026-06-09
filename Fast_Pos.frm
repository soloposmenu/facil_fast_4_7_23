VERSION 5.00
Object = "{CCB90040-B81E-11D2-AB74-0040054C3719}#1.0#0"; "OPOSCashDrawer.ocx"
Object = "{CCB90150-B81E-11D2-AB74-0040054C3719}#1.0#0"; "OPOSPOSPrinter.ocx"
Begin VB.Form Sys_Pos 
   BackColor       =   &H00EA6E67&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrada Cajeros"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   ClipControls    =   0   'False
   Icon            =   "Fast_Pos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   3675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2400
      Top             =   1440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   6
      Top             =   3780
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3050
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   6154
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   0
      Top             =   2450
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   3780
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1680
      Left            =   720
      Picture         =   "Fast_Pos.frx":0442
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2115
   End
   Begin OposCashDrawer_CCOCtl.OPOSCashDrawer Cocash1 
      Left            =   1920
      OleObjectBlob   =   "Fast_Pos.frx":B7634
      Top             =   4320
   End
   Begin OposPOSPrinter_CCOCtl.OPOSPOSPrinter ImprDevice 
      Left            =   1200
      OleObjectBlob   =   "Fast_Pos.frx":B7658
      Top             =   4320
   End
   Begin OposPOSPrinter_CCOCtl.OPOSPOSPrinter Coptr1 
      Left            =   240
      OleObjectBlob   =   "Fast_Pos.frx":B767C
      Top             =   4320
   End
   Begin VB.Label lbVersion 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00B39665&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   60
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00EA6E67&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   2
      Left            =   0
      TabIndex        =   7
      Top             =   4575
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00EA6E67&
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00EA6E67&
      Caption         =   "Número Cajero"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   2680
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00EA6E67&
      Caption         =   "Introduzca su número de cajero y contraseña"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   390
      TabIndex        =   2
      Top             =   1875
      Width           =   2775
   End
End
Attribute VB_Name = "Sys_Pos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsUsr As Recordset
Private nClosingTries As Byte       '25nov2014 = INTENTOS DE CIERRE CON IMPRESORA BLOQUEADA
'Private Declare Function GetTickCount Lib "kernel32" () As Long
'Dim nCounterTime As Long


Private Sub VerificaCierre()
Dim nFile As Integer

On Error GoTo ErrAdm:
nFile = FreeFile()
Open App.Path & "\INOUTLOG.TXT" For Input As #nFile
Do Until EOF(1)
   Line Input #nFile, a$
Loop
Close #nFile
If a$ = "OK" Then
    'CERRO BIEN
Else
    nFile = FreeFile()
    Open App.Path & "\VENTALOG.TXT" For Append As #nFile
        Print #nFile, "-- SOLO POS NO CERRO BIEN --" & Date & " " & Time
    Close #nFile
End If
nFile = FreeFile()
Open App.Path & "\INOUTLOG.TXT" For Output As #nFile
    Print #nFile, "NUNCA BORRAR ESTE ARCHIVO"
Close #nFile
On Error GoTo 0
Exit Sub

ErrAdm:
    MsgBox Err.Number & " " & Err.description, vbCritical, "Sys_Pos.Error en VerificaCierre"
End Sub

Private Sub AbrirFile()
'VERIFICA SI ES NECESARIO BORRAR TRANS LOCAL
Dim FecHost As Variant
Dim FecLoc As Variant
Dim RSLOC01 As Recordset
Dim nUpdateFlag As Integer
Dim iInt As Integer
'
'Set rsLoc00 = New Recordset
Set RSLOC01 = New Recordset
nUpdateFlag = 0 'CERO, NO HAY QUE ACTUALIZAR
iInt = 0

On Error GoTo ErrorAdm:
   ' Open DATA_PATH + "SOLOLINE.TXT" For Input As #1
   ' Close #1
On Error GoTo 0

If ON_LINE = True Then
    On Error GoTo ErrDBMSOpen:
    Sys_Pos.Caption = Sys_Pos.Caption & ".ON LINE"
    DoEvents

    DATA_PATH = GetFromINI("General", "DirectorioDatos", App.Path & "\Soloini.ini")
    

    cDataPath = GetFromINI("General", "ProveedorDatos", App.Path & "\Soloini.ini")

    cDataPath = cDataPath & ";Jet OLEDB:Database Password=master24"
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    'INFO: SQL SERVER EXPRESS
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
'    cDataPath = "Provider=SQLOLEDB.1;"
'    cDataPath = cDataPath & "Persist Security Info=True;User ID=cajeros;Password=solocajeros;"
'    cDataPath = cDataPath & "Initial Catalog=ACCESS_SQL;Data Source=CAJA1-SQL2005EX\SQLEXPRESS"

    'msConn.CursorLocation = adUseClient
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    
    msConn.Open cDataPath

    'ABRIENDO BASE DE DATOS DE PEDIDOS (AGO2009)
    'D:\ACCESS\Sushialbrook\
    'C:\ACCESS\Sushialbrook\
    Set msPED = New ADODB.Connection

    'msPED.Mode = adModeShareDenyNone
    'INFO: DIC2009
    
    msPED.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DATA_PATH & "MESASPED.MDB;"
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    'INFO: SQL SERVER EXPRESS
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    'msPED.CursorLocation = adUseClient
    
    'msPED.Open "Provider=SQLOLEDB.1;" & "Persist Security Info=True;User ID=cajeros;Password=solocajeros;" & "Initial Catalog=MESASPED_SQL;Data Source=CAJA1-SQL2005EX\SQLEXPRESS"
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    On Error GoTo 0
Else
End If

RSLOC01.Open "SELECT * FROM OPCIONES", msConn, adOpenStatic, adLockOptimistic

If RSLOC01!CHECK_UP = "Null" Then
    TipoApplicacion = " ESTE PRODUCTO ES UNA DEMOSTRACION."
Else
    TipoApplicacion = ""
End If
Me.Caption = TipoApplicacion + Me.Caption
SLIP_OK = False
If Not RSLOC01.EOF Then
    SLIP_OK = RSLOC01!SLIP_PRINTER
    REPCAJAX_OK = RSLOC01!REPORTX_OK
    MAX_DESCUENTO = RSLOC01!MAX_DESC
    OPC_SOLOFAST = RSLOC01!SOLOFAST
    TICKET_OK = RSLOC01!TICKET_OK
    FILE_OK = RSLOC01!FILE_OK
End If
RSLOC01.Close
Exit Sub

ErrorAdm:
ON_LINE = False
Resume Next

ErrDBMSOpen:
'Error grave NO SE ABRE DBMS
Dim OBJERR As Error
ShowMsg Err.Number & " - " & Err.description, vbYellow
For Each OBJERR In msConn.Errors
     'MsgBox OBJERR.Number & " <-> " & OBJERR.Description, vbCritical, "Error Grave. ANOTE EL NUMERO"
     ShowMsg OBJERR.Number & " <-> " & OBJERR.description, vbYellow, vbRed
Next
Unload Me
End
End Sub


Private Sub cmdacomp_Click()
nPLUPassed = 41
nDeptoPassed = 5
Acompanantes.Show
End Sub

Private Sub Command1_Click()
'INFO: 6FEB2015
Dim nNewTransNumber As Long

If Len(Text1) < 1 Or Len(Text2) < 1 Then Exit Sub
If Not IsNumeric(Text1) Then Exit Sub

On Error GoTo ErrorADO:
Set rs = New Recordset
Set rsUsr = New Recordset

rs.Open "SELECT numero,nombre FROM cajeros WHERE numero = " & Text1 & " and clave = " & "'" & Text2 & "'", msConn, adOpenForwardOnly, adLockReadOnly

rsUsr.Open "SELECT numero,nombre FROM USUARIOS WHERE numero = " & Text1 & " and clave = " & "'" & Text2 & "'", msConn, adOpenForwardOnly, adLockReadOnly

Set rs00 = New Recordset
rs00.Open "SELECT * FROM ORGANIZACION ", msConn, adOpenForwardOnly, adLockReadOnly
'TICKET_NUM = rs00!DIA_COUNT
nNewTransNumber = rs00!TRANS

rs00.Close

Call VerificaTabla_III("ORGANIZACION", "TRANS_FAST", nNewTransNumber + 1)

rs00.Open "SELECT * FROM ORGANIZACION ", msConn, adOpenForwardOnly, adLockReadOnly
TICKET_NUM = rs00!DIA_COUNT

If rsUsr.EOF Then   'SI NO ES ADMINISTRADOR BUSCA EN CAJEROS
    If rs.EOF Then
        MsgBox "Informacion es INCORRECTA, Intente de Nuevo", vbInformation, BoxTit
        Text1.SetFocus
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
        Exit Sub
    End If
Else
    Me.Text1 = ""
    Me.Text2 = ""
    
    RptCajas.Show 1
    
    Me.Text1.SetFocus
    
    Rem If Not Check_FE_Website("SILENCIO") Then
    If cLogMode = "SI" Then
        FE_Facil_ONLINE = False
    Else
        Call LoadPagoElectronico
        Call GetFacilRegistryValues
        Label2(2).ForeColor = &HFF0000
        Label2(2).BackColor = &HEA6E67
        Label2(2).FontSize = 10
        Label2(2).FontBold = False
        Label2(2) = "FACTURA ELECTRONICA OK"
        EscribeLog Label2(2)
        FE_Facil_ONLINE = True
    End If
    
    Exit Sub
End If

'INFO: SQL SERVER 2005 (DIC2009)
rsUsr.Close
Set rsUsr = Nothing

npNumCaj = rs!numero
cNomCaj = rs!nombre
nDesc01 = rs00!desc_01
nDesc02 = rs00!desc_02
nMesaBarra = rs00!mesa_barra
Text1 = "": Text2 = ""
On Error GoTo 0

'INFO: DEFINE LA ENTRADA DE LOS CAJEROS AL SISTEMA
EscribeLog ("Entrada a FASTPOS (" & App.Major & "." & App.Minor & "." & App.Revision & ") ==> " & cNomCaj)

'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
'''29SEP2011
If cFullFiscal = "SI" Then
    Call FISCAL_.SS_FiscalPrinterStatus(bFiscalPrinterStatus)
    If Not bFiscalPrinterStatus Then
        ShowMsg "EL ESTADO DE LA IMPRESORA FISCAL ES INCORRECTO" & vbCrLf & "SALGA DEL PROGRAMA, APAGUE LA IMPRESORA y PRENDALA" & vbCrLf & "VUELVA A ENTRAR AL PROGRAMA DE VENTAS", vbRed, vbYellow
    End If
End If
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~

'----DESPIERTA LA IMPRESORA----'
'PLU.BlockJournal

'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Rem If Not Check_FE_Website("", 250) Then
    Rem If Not Check_FE_Website("", 250) Then
If cLogMode = "SI" Then
    'MsgBox "FAIL: " & GetTickCount - nCounterTime
    EscribeLog "WEB SERVICE DE FACTURA ELECTRONICA NO ESTA DISPONIBLE. PASANDO A MODO (OFF-LINE)"
    Label2(2).FontBold = True
    Label2(2).ForeColor = vbYellow
    Label2(2).BackColor = vbRed
    Label2(2).FontSize = 12
    Label2(2) = "MODO OFF LINE"
Else
    'MsgBox "OK: " & GetTickCount - nCounterTime
    Rem Call LoadPagoElectronico
    Rem Call GetFacilRegistryValues
    Label2(2).ForeColor = &HFF0000
    Label2(2).BackColor = &HEA6E67
    Label2(2).FontSize = 10
    Label2(2).FontBold = False
    Label2(2) = "FACTURA ELECTRONICA OK"
    EscribeLog Label2(2)
End If


Rem If App.LogMode = 0 Then
'If cLogMode = "SI" Then
'    FE_Facil_ONLINE = False
'Else
'    FE_Facil_ONLINE = True
'End If

'INFO: UPDATE MODE OFF LINE SIN PREGUNTAR
'||||||| 31DIC2025
'|||||||If cLogMode = "SI" Then
'|||||||    If FE_Facil_ONLINE = True Then
'|||||||        vResp = ShowMsg("PASAR A OFF LINE. Y/N", , , vbYesNo)
'|||||||        If vResp = vbYes Then
'|||||||            FE_Facil_ONLINE = False
'|||||||        End If
'|||||||    End If
'|||||||Else
    Rem MsgBox "You are not in debug mode."
    Rem FE_Facil_ONLINE = True
'|||||||End If

'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

On Error Resume Next
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "BEGIN")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    Print2_OPOS_Dev "Login:" & Now
    For i = 1 To 10
        Print2_OPOS_Dev " "
    Next
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "END")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Sleep MUNBYNSeconds
    Coptr1.CutPaper 100
    If Err.Number = 482 Then
        ShowMsg "POR FAVOR ENCIENDA LA IMPRESORA" & Err.description, vbYellow, vbRed
        Err.Clear
    End If
On Error GoTo 0
'------------------------------'

On Error Resume Next
Text1.SetFocus
On Error GoTo 0
'INFO: ABRIENDO A PANTALLA CERRADA 19DIC2014
'PLU.Show 1

Rem FE_Facil_ONLINE = False

PLU.Show
Exit Sub

ErrorADO:
  Dim ADOError As Error
  For Each ADOError In msConn.Errors
     sError = sError & ADOError.Number & " - " & ADOError.description + vbCrLf
  Next ADOError
  MsgBox sError, vbCritical, "Error Grave. ANOTE EL NUMERO"
  Resume Next
End Sub

Private Sub Command2_Click()
Dim hwnd As Integer
Dim Mifrm As Form
Dim hnd_window As Long
Dim hnd_destroy As Long
Dim eeError As String

On Error Resume Next

Me.MousePointer = vbHourglass

'INFO: 12SEP2013
rc = Coptr1.State
If rc = OposSBusy Then
    nClosingTries = nClosingTries + 1
    If nClosingTries < 4 Then
        ShowMsg "Intento # (" & nClosingTries & ") LA IMPRESORA ESTA OCUPADA" & vbCrLf & "INTENTE OTRA VEZ CUANDO TERMINE DE IMPRIMIR o INTENTE SALIR OTRA VEZ." & vbCrLf & " Error # " & rc, vbRed, vbYellow
        Command2.Enabled = True
        Me.MousePointer = vbDefault
        Exit Sub
    Else
        EscribeLog ("Sys_Pos.IMPRESORA DE PRE-CUENTA / PEDIDOS ESTA OCUPADA, PERO LE VA PERMITIR SALIR DEL PROGRAMA DE VENTAS")
        ShowMsg "ATENCION: EL SISTEMA LE VA A PERMITIR SALIR, " & "AUNQUE LA IMPRESORA LE ESTA INFORMANDO QUE ESTA OCUPADA." & vbCrLf & "CUANDO SALGA DEL PROGRAMA DE VENTAS, DEBE APAGAR LA IMPRESORA " & vbCrLf & " Y ENCENDERLA ANTES DE VOLVER A ENTRAR AL PROGRAMA", vbYellow
    End If
End If

Label2(2).ForeColor = &HFF&
Label2(2) = "EL SISTEMA SE ESTA CERRANDO... ESPERE UNOS SEGUNDOS": Label2(2).Refresh
nVeriSalida = 2

Open App.Path & "\INOUTLOG.TXT" For Output As #1
    Print #1, "OK"
Close #1

'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
'''29SEP2011
If cFullFiscal = "SI" Then
    If FISCAL_.SS_FiscalClose Then
    Else
        EscribeLog "Ventas. FALLO AL CERRAR LA IMPRESORA (FISCAL) "
        ShowMsg "FALLO AL CERRAR LA IMPRESORA (FISCAL) ", vbYellow, vbRed
    End If
    Set FISCAL_ = Nothing
End If
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~

'DoEvents : 19SEP2013. REMOVIDO POR SUGERENCIA DE EPSON EN OPOS
Coptr1.DeviceEnabled = False        'INFO: HKA80. CIERRE CORRECTO DE SISTEMA. INFO: 6ABR2018
Coptr1.ReleaseDevice: Label2(2).Caption = "Impresora.ReleaseDevice": Label2(2).Refresh
Coptr1.Close: Label2(2).Caption = "Impresora.Close": Label2(2).Refresh

'DoEvents : 19SEP2013. REMOVIDO POR SUGERENCIA DE EPSON EN OPOS
'INFO: 28NOV2014, SI ES LA FISCAL NO HACER NADA CON LA GAVETA DE DINERO.
If cFullFiscal = "SI" Then
Else
    Cocash1.DeviceEnabled = False
    Cocash1.ReleaseDevice: Label2(2).Caption = "Cash Drawer.ReleaseDevice": Label2(2).Refresh
    Cocash1.Close: Label2(2).Caption = "Cash Drawer.Close": Label2(2).Refresh
End If

'INFO: FEB2010
On Error Resume Next
hnd_window = FindWindow(vbNullString, "FastPrintServer")
If hnd_window > 0 Then
    'APPLICACION FASTPRINTSERVER ESTA EN EJECUCION
    'CLOSE IT
    Call SendMessage(hnd_window, WM_CLOSE, 0&, ByVal 0&)
    'hnd_destroy = DestroyWindow(hnd_window)
End If
On Error GoTo 0

'INFO: FEB2010 INDICA QUE EL PRINTSERVER SE ESTA APAGANDO
If nMesa = 1 Then
    'UNICAMENTE CUANDO ES LA CAJA PRINCIPAL
    'Call WriteToINI("Meseros", "PrintServerStatus", "OFF", DATA_PATH & "soloini.ini")
    Call WriteToINI("Meseros", "PrintServerStatus", "ON", DATA_PATH & "soloini.ini")
End If

EscribeLog ("Salida de FASTPOS (" & App.Major & "." & App.Minor & "." & App.Revision & ")")

Me.MousePointer = vbDefault

For Each Mifrm In Forms
    DoEvents
    Mifrm.Hide          ' hide the form
    Unload Mifrm        ' deactivate the form
    Set frm = Nothing   ' remove from memory
Next

On Error GoTo 0
On Error GoTo ErrAdm:

Set rs = Nothing
msConn.Close
Set msConn = Nothing
On Error GoTo 0
Unload Me
End

ErrAdm:
    eeError = Err.description
    EscribeLog "ERROR AL SALIR. " & eeError
    ShowMsg eeError, vbYellow, vbRed
End Sub

Private Sub Form_Load()
Dim rs As Recordset
Dim nHandle As Integer
Dim iLoc As Integer
Dim DomicilioTemp As String

'INSTALADO EN EL ASADOR. LOS PUEBLOS DE ALBROOK
'Verifica si App esta abierta, para solamente cargarla una vez
If App.PrevInstance Then ActivatePrevInstance

'PARA DEMO
'Call VerificaFecha

'abrir coneccion
BoxTit = "MENSAJE DEL SISTEMA DE VENTAS"
Set msConn = New Connection

GENERICO_NO_JOURNAL = Chr$(27) + Chr$(99) + Chr$(48) + Chr$(0)
SOLOFAST_CUT = Chr$(27) & Chr$(105) & vbFormFeed

'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~
'28ASEP011
cFullFiscal = GetFromINI("Fiscal", "FullFiscal", App.Path & "\soloini.ini")

'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
cLogMode = GetFromINI("Fiscal", "Logmode", App.Path & "\soloini.ini")
'REVISAR SI MODO OFF LINE ESTA ACTIVADO
Rem On Error GoTo ErrFEON_OFF_LINE:
Rem FE_Facil_ONLINE = GetFromINI("Fiscal", "FEMode", App.Path & "\soloini.ini")
On Error Resume Next
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

'22JUN2012
If cFullFiscal = "SI" Then
    'no trabaja en debug.mode del lado de la libreria fiscal
    'FUNCION CLOSE LA LIBRERIA
    ''' Call VerificaApp("SoloFiscal", "Close")
End If

nMarca = -1
nMarca = CInt(GetFromINI("Fiscal", "Marca", App.Path & "\soloini.ini"))
'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~
App.OleRequestPendingTimeout = 8000        '8 SEGS HASTA QUE APAREZCA EL MSG DE ABAJO
App.OleRequestPendingMsgTitle = "LA IMPRESORA FISCAL ESTA OCUPADA"
App.OleRequestPendingMsgText = "La Impresora Fiscal aún está procesando Información." & vbCrLf & "POR FAVOR ESPERE"
Set FISCAL_ = CreateObject("SoloFiscal.ClaseFiscal")
'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~

Call VerificaCierre

ON_LINE = True
nVeriSalida = 1
Show

'-------------------------------------
Label2(2) = "Verificando Impresora y Gaveta de Dinero. Favor Esperar ...": Label2(2).Refresh
lbVersion.Caption = "Ver. " & App.Major & "." & App.Minor & "." & App.Revision
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
NOM_PRN_COCINA = GetFromINI("SoloPosDisp", "Cocina", App.Path & "\Soloini.ini")
NOM_PRN_BAR = GetFromINI("SoloPosDisp", "Bar", App.Path & "\Soloini.ini")

If GetFromINI("Facturacion", "PrintTicket", App.Path & "\Soloini.ini") = "Pereza" Then
    IMPRIME_TICKET = True
Else
    IMPRIME_TICKET = False
End If
If GetFromINI("Facturacion", "PrintPedidoEnCocina", App.Path & "\Soloini.ini") = "Pereza" Then
    IMPRIME_PEDIDO = True
Else
    IMPRIME_PEDIDO = False
End If
If GetFromINI("Facturacion", "DuplicarEnFactura", App.Path & "\Soloini.ini") = "Pereza" Then
    DUPLICA_PEDIDO = True
Else
    DUPLICA_PEDIDO = False
End If

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'INFO: 13ENE2017. VERIFICACION DE DOMICILIO
DomicilioTemp = GetFromINI("Facturacion", "DeptoDomicilio", App.Path & "\Soloini.ini")
If DomicilioTemp = "" Then
    nHayDomicilio = 0
Else
    nHayDomicilio = Int(DomicilioTemp)
    Call OpenDBDomicilio
End If
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

On Error GoTo ErrAdm:

'CAPTURA # DE CAJA PARA MULTI-CAJA
cCaja = GetFromINI("Facturacion", "CajaNumero", App.Path & "\Soloini.ini")
nMesa = cCaja - 1

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'INFO: INTRODUCCION DE DECIMALES
'INFO: 4OCT2017
If GetFromINI("Facturacion", "Decimales", App.Path & "\Soloini.ini") = "Pereza" Then
    HayDecimales = True
Else
    HayDecimales = False
End If
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'DoEvents : 19SEP2013. REMOVIDO POR SUGERENCIA DE EPSON EN OPOS
Label2(2) = "Impresora Precuenta. Apertura": Label2(2).Refresh
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~ FUNCIONANDO SIN IMPRESORA DE PRECUENTA ~~~~~~~
'~~~~~~~~ 17ENERO 2021 ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
NOM_PRN_FACTURA = GetFromINI("SoloPosDisp", "Facturacion", App.Path & "\Soloini.ini")
If NOM_PRN_FACTURA = "" Then
    Label2(2).BackColor = vbBlue
    Label2(2).Refresh
    bHayPrecuenta = False
Else
    bHayPrecuenta = True
    'Label2(2).BackColor = vbYellow
    'Label2(2).Refresh
    'ShowMsg "before open-"
    'rc = Coptr1.Open(NOM_PRN_FACTURA)
    'Coptr1.Open NOM_PRN_FACTURA
    'ShowMsg "before rc-"
    'rc = 0
    'ShowMsg "lolo0-" & rc
    'If rc <> OposSuccess Then
    
    If Coptr1.Open(NOM_PRN_FACTURA) <> 0 Then
        ShowMsg "LA IMPRESORA DE PRE-CUENTA/PEDIDOS NO ESTA ENCENDIDA. OPEN.Error  ENCIENDA LA IMPRESORA", vbYellow, vbRed
    Else
    '    Label2(2).BackColor = vbCyan
    '    Label2(2).Refresh
    '    ShowMsg "lolo"
        Coptr1.PowerNotify = 1
    End If
End If

'Label2(2).BackColor = vbRed
'Label2(2).Refresh
'If Coptr1.Open(NOM_PRN_FACTURA) > 0 Then
'    ShowMsg "LA IMPRESORA DE PRE-CUENTA/PEDIDOS NO ESTA ENCENDIDA. OPEN.Error  ENCIENDA LA IMPRESORA", vbYellow, vbRed
'End If
'INFO: 20AGO2014

'DoEvents : 19SEP2013. REMOVIDO POR SUGERENCIA DE EPSON EN OPOS

If bHayPrecuenta = False Then
Else
    Label2(2) = "Impresora Precuenta. CLAIM": Label2(2).Refresh
    
    If Coptr1.ClaimDevice(1000) > 0 Then
        'COMANDO CLAIM ES LO QUE MAS TARDA AL ABRIR IMPRESORA
        ShowMsg "ERROR AL CONECTARSE A LA IMPRESORA DE PRE-CUENTA/PEDIDOS", vbYellow, vbRed
    End If
    Label2(2) = "Impresora Precuenta. DeviceEnabled ": Label2(2).Refresh
    Coptr1.DeviceEnabled = True
End If


'Coptr1.AsyncMode = True (VER MAS ABAJO)

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'INFO: 22NOV2017
'INFO: ACTIVA LOGO PARA NEGOCIOS NO FISCALES
If bHayPrecuenta = False Then
Else
    If UCase(GetFromINI("Facturacion", "Logo", App.Path & "\Soloini.ini")) = "PEREZA" Then
        bLogo = True
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        'INFO: <<NO>> IMPRIME TITULO, UNICAMENTE LO MONTA EN LA MEMORIA DE LA IMPRESORA
        Call PrintImage(App.Path + "\Logo3.bmp")
        'Sys_Pos.Coptr1.SetBitmap 1, PtrSReceipt, App.Path + "\Logo.bmp", Sys_Pos.Coptr1.RecLineWidth / 2, PtrBmCenter
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    End If
End If
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

Me.MousePointer = vbHourglass

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
If bHayPrecuenta = False Then
Else
    If Coptr1.RecEmpty Or Coptr1.RecNearEnd Then
            ShowMsg "¡¡ A T E N C I O N !!" & vbCrLf & vbCrLf & "SE ESTA ACABANDO EL PAPEL DE PRE-CUENTA / PEDIDOS, SE LE RECOMIENDA SALIR DEL PROGRAMA PARA CAMBIAR EL PAPEL Y LUEGO REGRESAR AL PROGRAMA DE VENTAS", vbRed, vbYellow
    End If
End If
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'INFO: 20AGO2014
If bHayPrecuenta = False Then
Else
    Select Case Coptr1.PowerState
        Case 2000       'DESCONOCIDO, NO TIENE LA CAPACIDAD, NO HACER NADA
        Case 2001       'EN LINEA
        Case 2002, 2003, 2004    'APAGADA 'FUERA DE LINEA 'APAGADA
            ShowMsg "ERROR AL CONECTARSE A LA IMPRESORA DE PRE-CUENTA/PEDIDOS" & vbCrLf & "IMPRESORA.PowerState = 'APAGADA'." & vbCrLf & "Error: " & Coptr1.PowerState, vbYellow, vbRed
    End Select
End If
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

'Cocash1.Close
'DoEvents
'
'RC = Cocash1.Open(GetFromINI("SoloPosDisp", "Gaveta", App.Path & "\Soloini.ini"))
'DoEvents
'Cocash1.DeviceEnabled = True

If IMPRIME_PEDIDO = True Then
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    'INFO: QUITANDO LA IMPRESORA DE COCINA DEL SISTEMA FAST
    'YA QUE LA DATA A LAS IMPRESORAS REMOTAS LA REALIZARA EL FASTPRINTSERVER
    'DIC2009
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    
    If NOM_PRN_BAR = "" Then
        'DO NOTHING
    Else
        rc = Sys_Pos.ImprDevice.Open(NOM_PRN_BAR)
        If rc <> OposSuccess Then
            ShowMsg "FAVOR RECUERDE ENCENDER LA IMPRESORA DE BAR. Error # " & rc & " - ENCIENDA LA IMPRESORA DE BAR", vbYellow, vbRed
        End If
        
        rc = 0
        rc = Sys_Pos.ImprDevice.ClaimDevice(1000) 'COMANDO CLAIM ES LO QUE MAS TARDA AL ABRIR IMPRESORA
        If rc <> OposSuccess Then ShowMsg "LA IMPRESORA DE BAR NO ESTA CONECTADA o NO ESTA ENCENDIDA. ERROR NUMERO : " & rc, vbYellow, vbRed
        
        Sys_Pos.ImprDevice.DeviceEnabled = True
    End If
    
    'INFO: FEB2010
    If NOM_PRN_COCINA = "" Then
        'INFO: DO NOTHING. NO HAY IMPRESORA DE COCINA
    Else
        On Error Resume Next
        If FindWindow(vbNullString, "FastPrintServer") > 0 Then
            'APPLICACION FASTPRINTSERVER ESTA EN EJECUCION
        Else
            retVal = Shell(App.Path & "\FastPrintServer.exe", vbNormalFocus)
        End If
        On Error GoTo 0
    End If
    
End If

msConn.Mode = adModeShareDenyNone
Label2(2) = "Abrir Base de Datos"

On Error GoTo 0

Call AbrirFile   'Verifica Conección

'Label2(2) = App.CompanyName & " Listo"
'INFO: 21MAR2017
Label2(2) = "LISTO"

EscribeLog ("Inicio de FASTPOS")

'INFO: FEB2010
'INFO: VERIFICA EL PRINT SERVER, SOLAMENTE SI ES LA CAJA SECUNDARIA
If nMesa = 1 Then
    'INFO: DO NOTHING. PRINT SERVER SHOULD BE RUNNING
    '(1) ES LA CAJA PRIMARIA.
Else
    If GetFromINI("Meseros", "PrintServerStatus", DATA_PATH & "soloini.ini") = "ON" Then
        'INFO: DO NOTHING, PRINTSERVER IS UP
    Else
        ShowMsg "LA CAJA PRINCIPAL NO HA INICIADO CORRECTAMENTE." & vbCrLf & "SALGA DEL PROGRAMA Y REVISE TODAS LAS IMPRESORAS DE LA CAJA PRINCIPAL.", vbYellow, vbRed
        Call EscribeLog("Ventas. LA CAJA PRINCIPAL NO HA INICIADO CORRECTAMENTE. SALGA DEL PROGRAMA Y REVISE TODAS LAS IMPRESORAS DE LA CAJA PRINCIPAL.")
        Label2(2).ForeColor = vbYellow
        Label2(2).BackColor = vbRed
        Label2(2) = "Revise Caja Principal": Label2(2).Refresh
    End If
    
End If

'INFO: 2OCT2012
Call Verifica_TMP_TRANS

'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
'28SEP2011
'~~~~||||   APERTURA IMPRESORA FISCAL    ~~~~||||
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
If cFullFiscal = "SI" Then
    If FISCAL_.SS_Fiscal_IsOn = 1 Then
        If FISCAL_.SS_FiscalOpen Then
            'MsgBox "IMPRESORA ESTA PRENDIDA"
        Else
            ShowMsg "LA IMPRESORA  (FISCAL) NO FUE DETECTADA, FAVOR VERIFIQUE." & vbCrLf & "SALGA DEL PROGRAMA y VUELVA A ENTRAR", vbYellow, vbRed
            'INFO: DESABILITANDO PARA QUE SALGA UNICAMENTE
            Text1.Enabled = False
            Text2.Enabled = False
            Label2(2).ForeColor = vbYellow
            Label2(2).BackColor = vbRed
            Label2(2) = "FISCAL APAGADA. SALIR": Label2(2).Refresh
        End If
    Else
        ShowMsg "LA IMPRESORA (FISCAL) ESTA APAGADA, SALGA DEL PROGRAMA y VUELVA A ENTRAR", vbYellow, vbRed
        'INFO: DESABILITANDO PARA QUE SALGA UNICAMENTE
        Text1.Enabled = False
        Text2.Enabled = False
        Label2(2).ForeColor = vbYellow
        Label2(2).BackColor = vbRed
        Label2(2) = "FISCAL APAGADA. SALIR": Label2(2).Refresh
    End If
End If
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~'~~~~||||'~~~~||||


'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
'28SEP2011
If cFullFiscal = "SI" Then
    Call FISCAL_.SS_FiscalPrinterStatus(bFiscalPrinterStatus)
    If Not bFiscalPrinterStatus Then
        ShowMsg "EL ESTADO DE LA IMPRESORA FISCAL ES INCORRECTO" & vbCrLf & "SALGA DEL PROGRAMA, APAGUE LA IMPRESORA y PRENDALA." & vbCrLf & "VUELVA A ENTRAR AL PROGRAMA DE VENTAS", vbRed, vbYellow
    End If
End If
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~

'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
'28SEP2011
If cFullFiscal = "SI" Then
    If FISCAL_.SS_CashDrawerCheck(1) = 1 Then
    Else
        ShowMsg "Ventas. GAVETA DE DINERO NO ESTA CONECTADA CORRECTAMENTE", vbRed, vbYellow
    End If
Else
    Cocash1.Close
    'DoEvents : 19SEP2013. REMOVIDO POR SUGERENCIA DE EPSON EN OPOS
    rc = Cocash1.Open(GetFromINI("SoloPosDisp", "Gaveta", App.Path & "\Soloini.ini"))
    'DoEvents : 19SEP2013. REMOVIDO POR SUGERENCIA DE EPSON EN OPOS
    'INFO: ADDED 26/AGO/2023
    Cocash1.ClaimDevice 1000
    If Cocash1.Claimed Then
        Cocash1.DeviceEnabled = True
    Else
        ShowMsg "Ventas. Claimed GAVETA DE DINERO NO ESTA CONECTADA CORRECTAMENTE", vbRed, vbYellow
    End If
End If
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~


'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~|||
'INFO: HISTORICO DE TRANSACCIONES FISCALES y CORRELACIONAL DE SOLO SOFTWARE
'28SEP2011
If cFullFiscal = "SI" Then
    Call VerificaTabla_II("TRANSAC_FISCAL")
    'INFO: 24NOV2011
    Call VerificaTabla_II("Z_FISCAL")
End If
'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||

If nHayDomicilio > 0 Then Call VerificaTabla_II("DOMICILIO")

'INFO: 02OCT2012
OPOS_DevName = Sys_Pos.Coptr1.DeviceName
Call PUT_AsyncMode

'INFO: 20SEP2013. ELIMINADO 2MAYO2014
'If OPOS_DevName = "TM-U950P" Then
'    Sys_Pos.Coptr1.AsyncMode = False
'Else
'    Sys_Pos.Coptr1.AsyncMode = True
'End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO: VALIDA SI DESEAN RESTRINGIR LAS ANULACIONES y DESCUENTOS
If UCase(GetFromINI("Facturacion", "Seguridad", App.Path & "\soloini.ini")) = "PEREZA" Then
    HaySeguridad = True
Else
    HaySeguridad = False
End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'NO LE VEO FUNCION. DESACTIVANDO
'INFO: 23ENE2016
'Timer1.Enabled = True

Set rsAcomPLU = New ADODB.Recordset
rsAcomPLU.Open "SELECT PLU_ID FROM PLU_ACOMP", msConn, adOpenStatic, adLockOptimistic


Me.MousePointer = vbDefault


'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

lineWidth = Sys_Pos.Coptr1.RecLineChars

If Not LoadPagoElectronico Then
    Text1.BackColor = vbBlack
    Text2.BackColor = vbBlack
    Text1.Enabled = False
    Text2.Enabled = False
End If

'frmSplash.lblWarning(0).Caption = "CHECK FACTURA ELECTRONICA": frmSplash.lblWarning(0).Refresh
Label2(2) = "REVISA FACTURA ELECTRONICA"

Call GetFacilRegistryValues

'nCounterTime = GetTickCount
Rem If Not Check_FE_Website Then
If cLogMode = "SI" Then
    FE_Facil_ONLINE = False
    'ShowMsg "SERVICIO DE FACTURA ELECTRONICA ESTA CAIDO" & vbCrLf & "WEB SERVICE IS DOWN", vbYellow, vbRed
    ShowMsg "WEB SERVICE DE FACTURA ELECTRONICA" & vbCrLf & "NO ESTA DISPONIBLE" & vbCrLf & "PASANDO A MODO" & vbCrLf & vbCrLf & " (OFF-LINE)", vbBlue, vbCyan
    EscribeLog "WEB SERVICE DE FACTURA ELECTRONICA NO ESTA DISPONIBLE. PASANDO A MODO (OFF-LINE)"
    Label2(2).FontBold = True
    Label2(2).ForeColor = vbYellow
    Label2(2).BackColor = vbRed
    Label2(2).FontSize = 12
    Label2(2) = "MODO OFF LINE"
Else
    FE_Facil_ONLINE = True
    Call LoadPagoElectronico
    Call GetFacilRegistryValues
    Label2(2) = "FACTURA ELECTRONICA OK"
End If

If mi_Ambiente = 2 Then Me.BackColor = vbCyan

mi_App_Name = App.EXEName

'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||


On Error Resume Next
Text1.SetFocus
On Error GoTo 0

Exit Sub

ErrFEON_OFF_LINE:
ShowMsg "ON LINE // OFF LINE" & vbCrLf & vbCrLf & "NO ESTA ACTIVADO" & vbCrLf & "REVISAR VALOR DE (FEMode) CONFIGURACION EN SOLOINI.INI", vbRed, vbYellow
Resume Next

ErrAdm:
    If Err.Number = 13 Then
        ShowMsg "ERROR EN EL ARCHIVO DE CONFIGURACION" & vbCrLf & "SoloIni.ini", vbRed, vbYellow
    Else
        ShowMsg Err.Number & " - " & Err.description, vbRed, vbYellow
    End If
    Label2(2).FontBold = True
    Label2(2).ForeColor = vbYellow
    Label2(2).BackColor = vbRed
    Label2(2) = "Error En Sistema": Label2(2).Refresh
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If nVeriSalida = 1 Then
    MsgBox "¡ Favor utilize el boton Salir !", vbExclamation, BoxTit
    Cancel = True
End If
End Sub
Private Sub Image1_DblClick()
'ShowMsg "Empresa : " & App.LegalCopyright & Chr(13) & "Nombre  : " & App.EXEName & " (" & App.Major & "." & App.Minor & "." & App.Revision & ")" & vbCrLf & cTAXText & ": Calculado por Producto", vbGreen, vbBlue
'INFO: 12SEP2013
'ShowMsg "Empresa : " & App.LegalCopyright & Chr(13) & '       "Nombre  : " & App.EXEName & ", Versión : " & App.Major & "." & App.Minor & "." & App.Revision & Chr(13) & '       "Impuesto: Calculado por Producto" & Chr(13), vbGreen, vbBlue
'If cFullFiscal = "SI" Then Call FISCAL_.SS_PrintProgramacion


Rem ShowMsg GetLicenciaFE, vbBlue, vbYellow
'https://api.facturafacil.com.pa/api/core/subscriptions/status/

Rem ShowMsg "Ambiente: " & EvalAmbiente(mi_Ambiente) & vbCrLf & "Empresa : " & App.LegalCopyright & Chr(13) & "Nombre  : " & App.EXEName & " (" & App.Major & "." & App.Minor & "." & App.Revision & ")" & vbCrLf & cTAXText & ": Calculado por Producto", vbGreen, vbBlue

If cFullFiscal = "SI" Then Call FISCAL_.SS_PrintProgramacion

OFF_ON.Show 1

If cLogMode = "SI" Then
    'ShowMsg "SERVICIO DE FACTURA ELECTRONICA ESTA CAIDO" & vbCrLf & "WEB SERVICE IS DOWN", vbYellow, vbRed
    Rem ShowMsg "WEB SERVICE DE FACTURA ELECTRONICA" & vbCrLf & "NO ESTA DISPONIBLE" & vbCrLf & "PASANDO A MODO" & vbCrLf & vbCrLf & " (OFF-LINE)", vbBlue, vbCyan
    EscribeLog "F.E. PASANDO A MODO (OFF-LINE)"
    Label2(2).FontBold = True
    Label2(2).ForeColor = vbYellow
    Label2(2).BackColor = vbRed
    Label2(2).FontSize = 12
    Label2(2) = "MODO OFF LINE"
Else
    Call LoadPagoElectronico
    Call GetFacilRegistryValues
    Label2(2).BackColor = &HEA6E67
    Label2(2).FontSize = 10
    Label2(2) = "FACTURA ELECTRONICA OK"
End If


End Sub
Private Function EvalAmbiente(nAmb As Integer) As String
If nAmb = 2 Then EvalAmbiente = "PRUEBAS" Else EvalAmbiente = "<<PRODUCCION>>"

End Function
Private Sub Label2_DblClick(Index As Integer)
On Error Resume Next
Select Case Index
    Case 2
        If ShowMsg("¿ LA IMPRESORA FISCAL ESTA APAGADA o NO ESTA CONECTADA, DESEA ENTRAR DE TODAS MANERAS ?", , , vbYesNo) = vbYes Then
            Text1.Enabled = True
            Text2.Enabled = True
            Text1.SetFocus
        End If
    Case Else
End Select
On Error GoTo 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text1 = "" Then
        Text1.SetFocus
    ElseIf Not IsNumeric(Text1) Then
        Text1.SetFocus
    Else
        Text2.SetFocus
        Text2.SelStart = 0
        Text2.SelLength = Len(Text2.Text)
    End If
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text2 = "" Then
        Text2.SetFocus
    ElseIf Not IsNumeric(Text2) Then
        Text2.SetFocus
    Else
        Command1.SetFocus
    End If
End If
End Sub

Private Sub Timer1_Timer()
Dim retVal
Dim nFile As Integer
Dim nLin As Integer
nFile = FreeFile()

On Error GoTo ErrAdm:
'Debug.Print Format(Time, "hh:mm:sss")
nLin = 10
'If IsFormLoaded(PLU) = True Then
'    If PLU.Hora.Caption <> CStr(Time) Then
'        ' It's now a different second than the one displayed.
'        PLU.Hora.Caption = Time
'        'PLU.Hora.Caption = nNumTrans
'    End If
'End If
nLin = 20

If ON_LINE = True Then
nLin = 21
    Open DATA_PATH & "SOLOLINE.TXT" For Input As #nFile
    Do Until EOF(nFile)
        Line Input #nFile, a$
    Loop
    Close #nFile
nLin = 22
    If a$ = "OFF_LINE" Then
        ''''''''''''ES NECESARIO SALIR DEL PROGRAMA UN MOMENTO
        ''''''''''''YA QUE HASTA AHORA HABIAMOS TRABAJADO ON_LINE
        ''''''''''''EL PROGRAMA MSGUSER BORRA DB-LOCAL
        ''''''''''''BorraLocal
nLin = 23
        nVeriSalida = 2
        On Error Resume Next
        retVal = Shell(App.Path & "\MsgUser.exe", vbNormalFocus)
        On Error GoTo 0
        Unload Me
        End
    End If
nLin = 24
Else
nLin = 25
    'NO SE PUEDE VERIFICAR DE ESTA MANERA, PONE AL SISTEMA MUY LENTO
End If
nLin = 26
'Debug.Print Format(Time, "hh:mm:sss")
'Debug.Print "=========================="
On Error GoTo 0
Exit Sub

ErrAdm:
MsgBox Err.Number & " - " & Err.description, , nLin
MsgBox DATA_PATH & "SOLOLINE.TXT", , nLin
MsgBox a$, , nLin

End Sub
Private Sub VerificaFecha()
Dim cMaxFecha As Date
Dim cMaxDia As String
Dim cMaxMes As String
Dim cMaxYear As String
Dim cLocalFecha As String

cMaxFecha = Date
cMaxMes = Mid(Format(cMaxFecha, "short date"), 4, 2)
cMaxDia = Mid(Format(cMaxFecha, "short date"), 1, 2)
cMaxYear = Mid(Format(cMaxFecha, "short date"), 7, 4)

cLocalFecha = cMaxYear & cMaxMes & cMaxDia
If Val(cLocalFecha) > Val("20010430") Then
    MsgBox "***** SU PERIODO DE EVALUACION A TERMINADO *****" & vbCrLf & "- GRACIAS POR PROBAR PRODUCTOS DE SOLO SOFTWARE DEVELOPMENT" & vbCrLf & "- CONTACTE A SU PROVEEDOR" & vbCrLf & "" & vbCrLf & "El programa terminara AHORA", vbCritical, "CONTACTE A SU PROVEEDOR"
    Unload Me
    End
End If
End Sub

Private Function VerificaTabla_II(cTableName As String)
Dim rsTempTable As ADODB.Recordset
Dim cSQL As String

Set rsTempTable = New ADODB.Recordset

On Error Resume Next
rsTempTable.Open "SELECT TOP 2 * FROM " & cTableName, msConn, adOpenStatic, adLockOptimistic
If Err.Number = -2147217865 Then
        'TABLA NO EXISTE. ENTRA A RUTINA
        
        On Error GoTo 0
        
        Dim CAT As ADOX.Catalog
        Dim IDX As ADOX.Index
        
        Set CAT = New ADOX.Catalog
        ' ===[Create Database]===
        CAT.ActiveConnection = cDataPath
                
        Select Case cTableName
            Case "Z_FISCAL"
                ' ===[Create Table 'Z_FISCAL']===
                Set TBL = New ADOX.Table
                Set TBL.ParentCatalog = CAT
                TBL.name = "Z_FISCAL"
                TBL.Columns.Append "FISCAL", adInteger, 0
                TBL.Columns("FISCAL").Properties("NullAble") = True
                TBL.Columns("FISCAL").Properties("Description") = "CONTADOR_Z FISCAL"
                TBL.Columns.Append "FECHA", adVarWChar, 6
                TBL.Columns("FECHA").Properties("NullAble") = True
                TBL.Columns("FECHA").Properties("Jet OLEDB:Allow Zero Length") = True
                TBL.Columns("FECHA").Properties("Description") = "FECHA FISCAL"
                TBL.Columns.Append "HORA", adVarWChar, 4
                TBL.Columns("HORA").Properties("NullAble") = True
                TBL.Columns("HORA").Properties("Jet OLEDB:Allow Zero Length") = True
                TBL.Columns("HORA").Properties("Description") = "HORA FISCAL"
                TBL.Columns.Append "VENTAS", adSingle, 0
                TBL.Columns("VENTAS").Properties("NullAble") = True
                TBL.Columns("VENTAS").Properties("Description") = "VENTAS NETAS DE LA FISCAL"
                TBL.Columns.Append "ITBMS", adSingle, 0
                TBL.Columns("ITBMS").Properties("NullAble") = True
                TBL.Columns("ITBMS").Properties("Description") = "ITBMS DE LA FISCA"
                TBL.Columns.Append "Z_REF", adVarWChar, 10
                TBL.Columns("Z_REF").Properties("NullAble") = True
                TBL.Columns("Z_REF").Properties("Jet OLEDB:Allow Zero Length") = True
                TBL.Columns("Z_REF").Properties("Description") = "REFERENCIA Z_ COUNTER DE SOLO SOFTWARE"
                
                CAT.Tables.Append TBL
                
                Set IDX = New ADOX.Index
                IDX.name = "PrimaryKey"
                IDX.Columns.Append "FISCAL"
                IDX.PrimaryKey = True
                IDX.Unique = True
                IDX.Clustered = False
                IDX.IndexNulls = adIndexNullsDisallow
                
                CAT.Tables("Z_FISCAL").Indexes.Append IDX
                
                Set IDX = Nothing
                Set TBL = Nothing
                Set CAT = Nothing

            Case "TRANSAC_FISCAL"
                ' ===[Create Table 'TRANSAC_FISCAL']===
                Set TBL = New ADOX.Table
                Set TBL.ParentCatalog = CAT
                TBL.name = "TRANSAC_FISCAL"
                TBL.Columns.Append "FISCAL", adVarWChar, 15
                TBL.Columns("FISCAL").Properties("NullAble") = True
                TBL.Columns("FISCAL").Properties("Jet OLEDB:Allow Zero Length") = True
                TBL.Columns("FISCAL").Properties("Description") = "NUMERO DE FACTURA/DOCUMENTO DE LA IMPRESORA FISCAL"
                TBL.Columns.Append "DOC_SOLO", adInteger, 0
                TBL.Columns("DOC_SOLO").Properties("NullAble") = True
                TBL.Columns("DOC_SOLO").Properties("Description") = "CONTADOR INTERNO DE SOLOSOFTWARE"
                TBL.Columns.Append "FECHA_FULL", adDate, 0
                TBL.Columns("FECHA_FULL").Properties("NullAble") = True
                TBL.Columns("FECHA_FULL").Properties("Description") = "FECHA HORA DE LA FACTURA"
                TBL.Columns("FECHA_FULL").Properties("Default") = "Now()"
                TBL.Columns.Append "FECHA", adVarWChar, 8
                TBL.Columns("FECHA").Properties("NullAble") = True
                TBL.Columns("FECHA").Properties("Jet OLEDB:Allow Zero Length") = True
                TBL.Columns.Append "HORA", adVarWChar, 6
                TBL.Columns("HORA").Properties("NullAble") = True
                TBL.Columns("HORA").Properties("Jet OLEDB:Allow Zero Length") = True
                TBL.Columns.Append "SUB_TOTAL", adSingle, 0
                TBL.Columns("SUB_TOTAL").Properties("NullAble") = True
                TBL.Columns("SUB_TOTAL").Properties("Description") = "SUB TOTAL"
                TBL.Columns.Append "DESCUENTO", adSingle, 0
                TBL.Columns("DESCUENTO").Properties("NullAble") = True
                TBL.Columns("DESCUENTO").Properties("Description") = "DESCUENTO GLOBAL"
                TBL.Columns.Append "ITBM", adSingle, 0
                TBL.Columns("ITBM").Properties("NullAble") = True
                TBL.Columns("ITBM").Properties("Description") = "IMPUESTO"
                
                CAT.Tables.Append TBL
                
                Set TBL = Nothing
                Set CAT = Nothing
            
            Case "DOMICILIO"
                ' ===[Create Table 'DOMICILIO']===
                Set TBL = New ADOX.Table
                Set TBL.ParentCatalog = CAT
                TBL.name = "DOMICILIO"
                
                TBL.Columns.Append "FECHA", adVarWChar, 8
                TBL.Columns("FECHA").Properties("NullAble") = True
                TBL.Columns("FECHA").Properties("Jet OLEDB:Allow Zero Length") = True
                TBL.Columns("FECHA").Properties("Description") = "FECHA"
                TBL.Columns.Append "HORA", adVarWChar, 6
                TBL.Columns("HORA").Properties("NullAble") = True
                TBL.Columns("HORA").Properties("Jet OLEDB:Allow Zero Length") = True
                TBL.Columns("HORA").Properties("Description") = "HORA"
                TBL.Columns.Append "REPORTEZ", adVarWChar, 1
                TBL.Columns("REPORTEZ").Properties("NullAble") = True
                TBL.Columns("REPORTEZ").Properties("Description") = "REPORTEZ"
                TBL.Columns.Append "CANT", adInteger, 0
                TBL.Columns("CANT").Properties("NullAble") = True
                TBL.Columns("CANT").Properties("Description") = "CANT"
                TBL.Columns.Append "VALOR_DOMICILIO", adSingle, 0
                TBL.Columns("VALOR_DOMICILIO").Properties("NullAble") = True
                TBL.Columns("VALOR_DOMICILIO").Properties("Description") = "VALOR DOMI TRANS"
                TBL.Columns.Append "VALOR_FACTURA", adSingle, 0
                TBL.Columns("VALOR_FACTURA").Properties("NullAble") = True
                TBL.Columns("VALOR_FACTURA").Properties("Description") = "VALOR_FACTURA"
                TBL.Columns.Append "DOC_SOLO", adInteger, 0
                TBL.Columns("DOC_SOLO").Properties("NullAble") = True
                TBL.Columns("DOC_SOLO").Properties("Jet OLEDB:Allow Zero Length") = True
                TBL.Columns("DOC_SOLO").Properties("Description") = "CONTADOR INTERNO DE SOLOSOFTWARE"
                
                CAT.Tables.Append TBL
                
'                CAT.Tables.Refresh
'
'                Set IDX = New ADOX.Index
'                IDX.Name = "PrimaryKey"
'                IDX.Columns.Append "REPORTEZ"
'                IDX.PrimaryKey = True
'                IDX.Unique = False
'                IDX.Clustered = False
'                IDX.IndexNulls = adIndexNullsDisallow
'
'                Call CAT.Tables("DOMICILIO").Indexes.Append(IDX)
'
'                Set IDX = Nothing
                Set TBL = Nothing
                Set CAT = Nothing
        End Select
End If

On Error Resume Next
rsTempTable.Close
Set rsTempTable = Nothing
On Error GoTo 0
End Function
'---------------------------------------------------------------------------------------
' Procedimiento : VerificaTabla_III
' Autor       : hsequeira
' Fecha       : 06/02/2015
' Proposito   : VERIFICA LA TABLA DESIGNADA PARA UN SISTEMA MULTICAJA
'---------------------------------------------------------------------------------------
'
Private Function VerificaTabla_III(cTableName As String, cFieldName As String, nNewTransNumber As Long)
Dim rsTempTable As ADODB.Recordset
Dim cSQL As String

Set rsTempTable = New ADODB.Recordset

On Error Resume Next
rsTempTable.Open "SELECT " & cFieldName & " FROM " & cTableName, msConn, adOpenStatic, adLockOptimistic
If Err.Number = -2147217904 Then
        'CAMPO NO EXISTE. ENTRA A RUTINA
        'rsTempTable.Close
        On Error GoTo 0
        msConn.Execute "ALTER TABLE " & cTableName & " ADD COLUMN " & cFieldName & " INTEGER NULL "
        msConn.Execute "UPDATE " & cTableName & " SET " & cFieldName & " = " & nNewTransNumber
End If
Set rsTempTable = Nothing
End Function
'---------------------------------------------------------------------------------------
' Procedure : Verifica_TMP_TRANS
' Author    : hsequeira
' Date      : 02/10/2012
' Purpose   : AVISA AL USUARIO QUE EL CAJERO ANTERIOR CERRO EL PROGRAMA
' DEJANDO PRODUCTOS MARCADOS
'---------------------------------------------------------------------------------------
'
Private Function Verifica_TMP_TRANS() As Boolean
Dim rsT_TMP_TRANS As ADODB.Recordset
Dim bError As Boolean

Set rsT_TMP_TRANS = New ADODB.Recordset

'If nMesa = 0 Then
If nMesa = 120 Then
    rsT_TMP_TRANS.Open "SELECT CAJERO, DESCRIP, PRECIO FROM TMP_TRANS WHERE MESA = " & nMesa, msConn, adOpenStatic, adLockOptimistic
Else
    rsT_TMP_TRANS.Open "SELECT CAJERO, DESCRIP, PRECIO FROM TMP_TRANS", msConn, adOpenStatic, adLockOptimistic
End If
If rsT_TMP_TRANS.EOF Then
    'TODO ESTA BIEN. PUEDE PROCEDER A ENTRAR AL SISTEMA
    bError = False
Else
    bError = True
    Do While Not rsT_TMP_TRANS.EOF
        ShowMsg "EL CAJERO # " & rsT_TMP_TRANS!CAJERO & ", DEJO PRODUCTOS MARCADOS: " & rsT_TMP_TRANS!DESCRIP & " (" & Format(rsT_TMP_TRANS!precio, "CURRENCY") & ")", vbRed, vbYellow
        Me.MousePointer = vbHourglass
        EscribeLog "EL CAJERO # " & rsT_TMP_TRANS!CAJERO & ", DEJO PRODUCTOS MARCADOS: " & rsT_TMP_TRANS!DESCRIP
        rsT_TMP_TRANS.MoveNext
        Me.MousePointer = vbDefault
    Loop
End If
If bError Then
    ShowMsg "CONTACTE INMEDIATAMENTE A SU SUPERVISOR PARA CORREGIR EL ERROR" & vbCrLf & vbCrLf & "Si el ERROR no se corrige INMEDIATAMENTE, USTED tendra FALTANTE en su arqueo de CAJA", vbYellow, vbRed
    Me.MousePointer = vbHourglass
    EscribeLog "CONTACTE INMEDIATAMENTE A SU SUPERVISOR PARA CORREGIR EL ERROR"
    Me.MousePointer = vbDefault
End If
rsT_TMP_TRANS.Close
Set rsT_TMP_TRANS = Nothing
End Function
