VERSION 5.00
Begin VB.Form Sys_Pos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ENTRADA DE CAJEROS"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5040
      Top             =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2640
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   855
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      MaxLength       =   4
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   3840
      Picture         =   "Sys_pos.frx":0000
      Top             =   120
      Width           =   2070
   End
   Begin VB.Label Label2 
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Número de Cajero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Bienvenido, Introduzca su número de cajero y contraseña"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Sys_Pos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsUsr As Recordset
Private Sub VerificaCierre()
Open "C:\SYS_POS\INOUTLOG.TXT" For Input As #1
Do Until EOF(1)
   Line Input #1, a$
Loop
Close #1
If a$ = "OK" Then
    'CERRO BIEN
Else
    Open "C:\SYS_POS\VENTALOG.TXT" For Append As #1
        Print #1, "-- SOLO POS NO CERRO BIEN --" & Date & " " & Time
    Close #1
End If
Open "C:\SYS_POS\INOUTLOG.TXT" For Output As #1
    Print #1, "NUNCA BORRAR ESTE ARCHIVO"
Close #1
End Sub

Private Sub AbrirFile()
'VERIFICA SI ES NECESARIO BORRAR TRANS LOCAL
Dim FecHost As Variant
Dim FecLoc As Variant
Dim RSLOC01 As Recordset
Dim nUpdateFlag As Integer
Dim iInt As Integer

Set rsLoc00 = New Recordset
Set RSLOC01 = New Recordset
nUpdateFlag = 0 'CERO, NO HAY QUE ACTUALIZAR
iInt = 0

On Error GoTo ErrorAdm:
   ' Open DATA_PATH + "SOLOLINE.TXT" For Input As #1
   ' Close #1
On Error GoTo 0

If ON_LINE = True Then
    
    '\\SOLO11\ACCESS\SOLO.mdb;"
    On Error GoTo ErrDBMSOpen:
    Sys_Pos.Caption = Sys_Pos.Caption + ".ON LINE"
    msConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
            + "Data Source=\\SOLO11\ACCESS\SOLO.mdb;" _
            + "Jet OLEDB:Database Password=master24"

            '"Provider=Microsoft.Jet.OLEDB.4.0;" _
            + "Data Source=\\SOLO11\ACCESS\SOLO.mdb;" _
            + "Persist Security Info=False"
    On Error GoTo 0
Else
    
    Sys_Pos.Caption = Sys_Pos.Caption + ".OFF LINE"
    On Error GoTo ErrDBMSOpen:
    msConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
            + "Data Source=C:\SYS_POS\LOCAL\SOLO.mdb;" _
            + "Persist Security Info=False"
    On Error GoTo 0
    'msConn.Open "Provider=Microsoft.Jet.OLEDB.3.51" _
        + ";Persist Security Info=False;Data Source=" _
        + "C:\SYS_POS\LOCAL\SOLO.mdb"
    MsgBox "TRABAJANDO OFF-LINE (Fuera de Linea). Puede Continuar. Presione Enter", vbInformation, BoxTit
End If

RSLOC01.Open "SELECT * FROM OPCIONES", msConn, adOpenStatic, adLockOptimistic

SLIP_OK = False
If Not RSLOC01.EOF Then
    SLIP_OK = RSLOC01!SLIP_PRINTER
    REPCAJAX_OK = RSLOC01!REPORTX_OK
    RSLOC01.Close
End If
Exit Sub

ErrorAdm:
ON_LINE = False
Resume Next

ErrorCopiaON:
    ' La BD no se pudo copiar alguien lo esta usando en la oficina
    MsgBox "ON LINE ¡ ERROR AL COPIAR BASES DE DATOS ! POSIBLEMENTE " & _
           "LA ESTEN USANDO EN LA OFICINA. EL PROGRAMA TERMINARA AHORA.", vbCritical, BoxTit
    Unload Me
    End

ErrDBMSOpen:
'Error grave NO SE ABRE DBMS
Dim OBJERR As Error
MsgBox Err.Description
For Each OBJERR In msConn.Errors
     MsgBox OBJERR.Number & " <-> " & OBJERR.Description, vbCritical, "Error Grave. ANOTE EL NUMERO"
Next
Unload Me
End
End Sub
Private Sub Command1_Click()

If Len(Text1) < 1 Or Len(Text2) < 1 Then Exit Sub
If Not IsNumeric(Text1) Then Exit Sub

On Error GoTo ErrorADO:
Set rs = New Recordset
Set rsUsr = New Recordset

rs.Open "SELECT numero,nombre FROM cajeros WHERE numero = " & _
    Text1 & " and clave = " & "'" & Text2 & "'", msConn, adOpenForwardOnly, adLockReadOnly

rsUsr.Open "SELECT numero,nombre FROM USUARIOS WHERE numero = " & _
    Text1 & " and clave = " & "'" & Text2 & "'", msConn, adOpenForwardOnly, adLockReadOnly

Set rs00 = New Recordset
rs00.Open "SELECT * FROM ORGANIZACION ", msConn, adOpenForwardOnly, adLockReadOnly

If rsUsr.EOF Then   'SI NO ES ADMINISTRADOR BUSCA EN CAJEROS
    If rs.EOF Then
        MsgBox "Informacion es INCORRECTA, Intente de Nuevo", vbInformation, BoxTit
        Exit Sub
    End If
Else
    RptCajas.Show 1
    Text1 = "": Text2 = ""
    Text1.SetFocus
    Exit Sub
End If

npNumCaj = rs!numero
cNomCaj = rs!nombre
nDesc01 = rs00!desc_01
nDesc02 = rs00!desc_02
nMesaBarra = rs00!mesa_barra
Text1 = "": Text2 = ""
On Error GoTo 0

'----DESPIERTA LA IMPRESORA----'
On Error Resume Next
    Printer.FontSize = 9
    Printer.FontName = "12.5 CPI"
    'Printer.FontName = "14 CPI"
    Printer.Print "Login : " & Now
    Printer.EndDoc
    If Err.Number = 482 Then
        MsgBox "POR FAVOR ENCIENDA LA IMPRESORA", vbInformation, Err.Description
        Err.Clear
    End If
On Error GoTo 0
'------------------------------'
Text1.SetFocus
PLU.Show
Exit Sub

ErrorADO:
  Dim ADOError As Error
  For Each ADOError In msConn.Errors
     sError = sError & ADOError.Number & " - " & ADOError.Description + vbCrLf
  Next ADOError
  MsgBox sError, vbCritical, "Error Grave. ANOTE EL NUMERO"
  Resume Next
End Sub

Private Sub Command2_Click()
Dim hwnd As Integer
nVeriSalida = 2
Open "C:\SYS_POS\INOUTLOG.TXT" For Output As #1
    Print #1, "OK"
Close #1
'hwnd = GetWindow(nHandle, GW_OWNER)
'hwnd = FindWindow("ThunderRT6Main", "Sys_pos")
'MsgBox Str(hwnd)
'''''''''''''''''''If ON_LINE = True Then BorraLocal
Set rs = Nothing
msConn.Close
Unload Me
End
End Sub

Private Sub Form_Load()
Dim rs As Recordset
Dim nHandle As Integer

'Verifica si App esta abierta, para solamente cargarla una vez
If App.PrevInstance Then ActivatePrevInstance

'abrir coneccion
BoxTit = "MENSAJE DEL SISTEMA DE VENTAS"
Set msConn = New Connection

DEFAULT_PRINTER = GetCurrPrinter()

VerificaCierre
ON_LINE = True
nVeriSalida = 1

msConn.Mode = adModeShareDenyNone
''''''''''''''''''AbrirDBTransLocal
AbrirFile   'Verifica Conección

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If nVeriSalida = 1 Then
    MsgBox "¡ Favor utilize el boton Salir !", vbExclamation, BoxTit
    Cancel = True
End If
End Sub

Private Sub Image1_DblClick()
MsgBox "Empresa : " & App.CompanyName & Chr(13) & _
       "Derechos Reservados : " & App.LegalCopyright & Chr(13) & _
       "Nombre  : " & App.Title & ".EXE" & Chr(13) & _
       "Versión : " & App.Major & "." & App.Minor & "." & App.Revision, vbInformation, "Informacion de la Aplicación"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text1 = "" Then
        Text1.SetFocus
    ElseIf Not IsNumeric(Text1) Then
        Text1.SetFocus
    Else
        Text2.SetFocus
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
Dim RetVal
If IsFormLoaded(PLU) = True Then
    If PLU.Hora.Caption <> CStr(Time) Then
        ' It's now a different second than the one displayed.
            PLU.Hora.Caption = Time
    End If
End If

If ON_LINE = True Then
    
    Open DATA_PATH + "SOLOLINE.TXT" For Input As #1
    Do Until EOF(1)
        Line Input #1, a$
    Loop
    Close #1
    
    If a$ = "OFF_LINE" Then
        ''''''''''''ES NECESARIO SALIR DEL PROGRAMA UN MOMENTO
        ''''''''''''YA QUE HASTA AHORA HABIAMOS TRABAJADO ON_LINE
        ''''''''''''EL PROGRAMA MSGUSER BORRA DB-LOCAL
        ''''''''''''BorraLocal
        nVeriSalida = 2
        RetVal = Shell("c:\sys_pos\MsgUser.exe", vbNormalFocus)
        Unload Me
        End
    End If
Else
    'NO SE PUEDE VERIFICAR DE ESTA MANERA, PONE AL SISTEMA MUY LENTO
End If

End Sub
