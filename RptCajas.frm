VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form RptCajas 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REPORTES X y Z"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   7965
   Icon            =   "RptCajas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRZ 
      Caption         =   "REPORTE TERMINAL DE CAJEROS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   4440
      TabIndex        =   10
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton cmdRZ 
      Caption         =   "REPORTE DE MESAS"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   4440
      TabIndex        =   9
      Top             =   1800
      Width           =   2895
   End
   Begin VB.CommandButton cmdRZ 
      Caption         =   "REPORTE DE MESEROS"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   4440
      TabIndex        =   8
      Top             =   2760
      Width           =   2895
   End
   Begin VB.CommandButton cmdRX 
      Caption         =   "REPORTE DE CAJEROS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   720
      TabIndex        =   6
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton cmdRX 
      Caption         =   "REPORTE DE MESAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   720
      TabIndex        =   5
      Top             =   1800
      Width           =   2895
   End
   Begin VB.CommandButton cmdRX 
      Caption         =   "REPORTE DE MESEROS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   720
      TabIndex        =   4
      Top             =   2760
      Width           =   2895
   End
   Begin VB.CommandButton cmdRX 
      Caption         =   "REPORTE DEPARTAMENTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   720
      TabIndex        =   3
      Top             =   3720
      Width           =   2895
   End
   Begin VB.CommandButton cmdRX 
      Caption         =   "NO DISPONIBLE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   720
      TabIndex        =   2
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CommandButton cmdOpenGaveta 
      Height          =   615
      Left            =   360
      Picture         =   "RptCajas.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   0
      Top             =   6240
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   285
      Left            =   480
      TabIndex        =   14
      Top             =   5760
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label LbFiscal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOTA DE CREDITO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1920
      TabIndex        =   13
      ToolTipText     =   "Opciones de Facturacion Electronica"
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label lbFactRapida 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FACTURA RAPIDA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   3960
      TabIndex        =   12
      ToolTipText     =   "Factura Rapida Consumidor Final"
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "REPORTES Z"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   11
      Top             =   120
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   5055
      Index           =   1
      Left            =   4200
      Top             =   600
      Width           =   3345
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "REPORTES X"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   7
      Top             =   120
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   5055
      Index           =   0
      Left            =   480
      Top             =   600
      Width           =   3345
   End
   Begin VB.Menu mnuRepZ 
      Caption         =   "Reporte Z de Archivo"
   End
End
Attribute VB_Name = "RptCajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsPluInvent As New ADODB.Recordset
Private rsPLURecetas As New ADODB.Recordset         'INFO: PLU vs. RECETAS (12DIC2010)
Private bPrintOk As Boolean '20030829
Private rsLOCALPLUInvent As New ADODB.Recordset
Private rsLOCALPLURecetas As New ADODB.Recordset        'RELACION PLU vs. RECETAS (12DIC2010)
'INFO: EXPORT REPORT Z
Private nFileNumber As Integer
Private cFileName As String
Private ErrorReporte_Email As Boolean
'INFO: 12JUN2012
'Private OPOS_DevName As String ' YA ES UNA VARIABLE PUBLICA


'---------------------------------------------------------------------------------------
' Procedure : Print2_OPOS_Dev
' Author    : hsequeira
' Date      : 27/05/2012
' Purpose   : IMPRIME DE FORMA CORRECTA EN LOS ROLLOS DE LA IMPRESORA QUE ESTA CONECTADA.
'                    YA SEA UNA GRANDE (950) o UNA CHICA
' PARAMETROS: LOS DATOS QUE SE DESEAN IMPRIMIR
'---------------------------------------------------------------------------------------
'
Private Function Print2_OPOS_Dev(cParams As String, Optional cParam2 As String) As Boolean
'Public Function Print2_OPOS_Dev(cParams As String, Optional UnRollo As Boolean) As Boolean

If cParams = Space(1) Or cParams = Space(2) Then
    'VIENE UNA LINEA EN BLANCO
    cParams = Chr(&HD) & Chr(&HA)
End If
Rem Debug.Print cParams
Select Case OPOS_DevName
    Case "SRP-350plus"  'INFO: 20SEP2013
        If cParams = Chr(&HD) & Chr(&HA) Then
            Sys_Pos.Coptr1.PrintNormal PtrSReceipt, Chr(&HD)
        Else
            Sys_Pos.Coptr1.PrintNormal PtrSReceipt, cParams & Chr(&HD)
        End If
    'Case "LR3000", "TM-U200B", "SRP270", "MP4200TH", "TM-T20E", "TM-T20U", "TM-U220B"
    Case "LR3000", "TM-U200B", "SRP270", "MP4200TH", "SRP270P", "SEMOPOS.SO.SERIAL.POSPrinter", "SRP-275", "SRP-275P", "SRP-270", "SRP-270P", "TM-T20-42CU", "TM-T20-42CE", "TM-T20II-42CU", "TM-T20II-42CE", "TM-U200BP", "TM-U220BP", "TM-U200B"
        If cParams = Chr(&HD) & Chr(&HA) Then
            Sys_Pos.Coptr1.PrintNormal PtrSReceipt, cParams
        Else
            Sys_Pos.Coptr1.PrintNormal PtrSReceipt, cParams & Chr(&HD) & Chr(&HA)
        End If
    Case "TM-U950P", "TM-U950"
        UnRollo = True
        If cParams = Chr(&HD) & Chr(&HA) Then
            If UnRollo Then
                Sys_Pos.Coptr1.PrintNormal PtrSReceipt, cParams
            Else
                Sys_Pos.Coptr1.PrintTwoNormal PtrSJournalReceipt, Space(1), Space(1)
            End If
            'Sys_Pos.Coptr1.PrintTwoNormal PTR_S_JOURNAL_RECEIPT, Space(1), Space(1)
        Else
            'Sys_Pos.Coptr1.PrintTwoNormal PTR_S_JOURNAL_RECEIPT, cParams, cParams
            If UnRollo Then
                Sys_Pos.Coptr1.PrintNormal PtrSReceipt, cParams & Chr(&HD) & Chr(&HA)
            Else
                Sys_Pos.Coptr1.PrintTwoNormal PtrSJournalReceipt, cParams, cParams
            End If
        End If
    Case Else
        If cParams = Chr(&HD) & Chr(&HA) Then
            Sys_Pos.Coptr1.PrintNormal PtrSReceipt, cParams
        Else
            Sys_Pos.Coptr1.PrintNormal PtrSReceipt, cParams & Chr(&HD) & Chr(&HA)
        End If
End Select

Eval_OPOS_Dev (Sys_Pos.Coptr1.State)
End Function

'---------------------------------------------------------------------------------------
' Procedure : Eval_OPOS_Dev
' Author    : hsequeira
' Date      : 27/06/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function Eval_OPOS_Dev(rc As Long) As Boolean
Select Case rc
    Case OposSClosed
        'Debug.Print "OPOS_S_CLOSED - " & Sys_Pos.Coptr1.ResultCode
    Case OposSIdle
        'Debug.Print "OPOS_S_IDLE - " & LoginMesas.ImpresoraCuentas.ResultCode
        'Debug.Print "OPOS_S_IDLE - " & Sys_Pos.Coptr1.ResultCode
    Case OposSBusy
        'Debug.Print "OPOS_S_BUSY - " & LoginMesas.ImpresoraCuentas.ResultCode
        'Debug.Print "OPOS_S_BUSY - " & Sys_Pos.Coptr1.ResultCode
    Case OposSError
        'Debug.Print "OPOS_S_ERROR - " & LoginMesas.ImpresoraCuentas.ResultCode
        'Debug.Print "OPOS_S_ERROR - " & Sys_Pos.Coptr1.ResultCode
    Case Else
End Select
End Function
Private Function PutMail(cAttachment As String) As Boolean
'  The mailman object is used for sending and receiving email.
'INFO: 17MAY2011
Dim mailman As New ChilkatMailMan2
Dim cBody As String
Dim hasRecipient As Boolean
Dim cRecipient_A As String, cRecipient_B As String

If Dir(cAttachment) = "" Then
    EscribeLog "Error #2 en Folders de Correo. PutMail Function"
    Exit Function
End If

'  Any string argument automatically begins the 30-day trial.
Dim success As Long
'success = mailman.UnlockComponent("30-day trial")
success = mailman.UnlockComponent("HUMBERMAILQ_p4xn0rEO7R3G")

If (success <> 1) Then
    MsgBox "Component unlock failed", vbCritical
    Exit Function
End If

'  Set the SMTP server. (GOOGLE)
mailman.SmtpHost = "smtp.gmail.com"
mailman.SmtpPort = 465
'HOTMAIL
'mailman.SmtpHost = "smtp.live.com"
'mailman.SmtpPort = 587
mailman.SmtpSsl = 1
'  Set the SMTP login/password (if required)
mailman.SmtpUsername = "email.solosoftware"
'mailman.SmtpUsername = "humberto_sequeira@hotmail.com"
'mailman.SmtpUsername = "humberto_sequeira"
mailman.SmtpPassword = "master24master2424"

'  Create a new email object
Dim email As New ChilkatEmail2

email.subject = "Reporte Z Terminal " & rs00!DESCRIP
'email.AddFileAttachment "D:\SOLOSOFTWARE\SoloMix\habibis\RZ.txt"
'email.AddFileAttachment2 "D:\SOLOSOFTWARE\SoloMix\habibis\RZ.txt", "text"   'lo pone como in-line y lo muestra en el correo
email.AddFileAttachment2 cAttachment, "text"

cBody = "Envio de Correo del Reporte de Caja" & vbCrLf
cBody = cBody & vbCrLf

'email.SetHtmlBody cBody

email.HasHtmlBody       'CON ESTO LA REPRESENTACION DE LOS vbCrLf se ve +/- correcta.
email.Body = cBody
'email.SetTextBody cBody, "text/html"
'email.Body = cBody
email.From = "email.solosoftware@gmail.com"
'email.From = "humberto_sequeira@ hotmail.com"
'email.From = "servicios@laposta.com"

If Left(GetFromINI("Correos", "email1", App.Path & "\email.ini"), 19) <> "email1@sitioweb.com" Then
    hasRecipient = True
    cRecipient_A = Chr(34) & GetFromINI("Correos", "email1", App.Path & "\email.ini") & Chr(34)
    cRecipient_B = Chr(34) & "<" & GetFromINI("Correos", "email1", App.Path & "\email.ini") & ">" & Chr(34)
    email.AddTo cRecipient_A, cRecipient_B
End If

If Left(GetFromINI("Correos", "email2", App.Path & "\email.ini"), 19) <> "email2@sitioweb.com" Then
    hasRecipient = True
    cRecipient_A = Chr(34) & GetFromINI("Correos", "email2", App.Path & "\email.ini") & Chr(34)
    cRecipient_B = Chr(34) & "<" & GetFromINI("Correos", "email2", App.Path & "\email.ini") & ">" & Chr(34)
    email.AddTo cRecipient_A, cRecipient_B
End If

If Left(GetFromINI("Correos", "email3", App.Path & "\email.ini"), 19) <> "email3@sitioweb.com" Then
    hasRecipient = True
    cRecipient_A = Chr(34) & GetFromINI("Correos", "email3", App.Path & "\email.ini") & Chr(34)
    cRecipient_B = Chr(34) & "<" & GetFromINI("Correos", "email3", App.Path & "\email.ini") & ">" & Chr(34)
    email.AddTo cRecipient_A, cRecipient_B
End If

If Left(GetFromINI("Correos", "email4", App.Path & "\email.ini"), 19) <> "email4@sitioweb.com" Then
    hasRecipient = True
    cRecipient_A = Chr(34) & GetFromINI("Correos", "email4", App.Path & "\email.ini") & Chr(34)
    cRecipient_B = Chr(34) & "<" & GetFromINI("Correos", "email4", App.Path & "\email.ini") & ">" & Chr(34)
    email.AddTo cRecipient_A, cRecipient_B
End If

If Not hasRecipient Then
    'NO HAY A QUIEN ENVIAR CORREO. SALIR DE LA RUTINA
    'Exit Function
    email.AddTo "Walter Sequeira", "<sequeira@cableonda.net>"
    email.AddTo "Humberto Yahoo", "<solo_software@yahoo.com>"
End If

'email.AddTo "Humberto Yahoo", "<solo_software@yahoo.com>"
'email.AddTo "Walter Sequeira", "<sequeira@cableonda.net>"
'email.AddTo "Walter Sequeira", "<sequeira@cableonda.net>"

'  Call SendEmail to connect to the SMTP server and send.
'  The connection (i.e. session) to the SMTP server remains
'  open so that subsequent SendEmail calls may use the
'  same connection.
'///////success = mailman.SendEmail(email)

'success = mailman.SendQ(email)          'SI EL SERVICIO ESTA INSTALADO EN EL MISMO FOLDER
'success = mailman.SendQ2(email, DATA_PATH & "\queue")
'INFO: (7NO2011) CORREGIDO, SE LE QUITA EL SLASH
success = mailman.SendQ2(email, DATA_PATH & "queue")

If (success <> 1) Then
    EscribeLog "Error en la Creacion de Correo: " & mailman.LastErrorText
    Exit Function
End If

'  Some SMTP servers do not actually send the email until
'  the connection is closed.  In these cases, it is necessary to
'  call CloseSmtpConnection for the mail to be  sent.
'  Most SMTP servers send the email immediately, and it is
'  not required to close the connection.  We'll close it here
'  for the example:
'///////success = mailman.CloseSmtpConnection()
'///////If (success <> 1) Then
    '///////MsgBox "Connection to SMTP server not closed cleanly.", vbInformation
'///////End If

'///////MsgBox "Mail Sent!", vbInformation

End Function


Private Function GetFileName(cTipo As String) As String
'INFO: 17MAY2011
Dim cFile As String
cFile = cTipo & "_" & Format(Date, "dd_mmm_yyyy") & "_" & Format(Time, "HH_MM") & ".txt"

If Dir(DATA_PATH & cFile) = "" Then
    'NO FILE. GO AHEAD AND CREATE IT.
Else
    'HAY UN FILE DEL MISMO NOMBRE
    Kill DATA_PATH & "queue\old_" & cFile
    FileCopy DATA_PATH & "queue\" & cFile, DATA_PATH & "queue\old_" & cFile
    Kill DATA_PATH & "queue\" & cFile
End If
GetFileName = DATA_PATH & "queue\" & cFile
End Function
Private Function TEXTFileZ(cMode As String, Optional cTexto As String, Optional cCant As Variant, Optional cMonto As Variant, Optional cInfoAdicional As String) As Boolean
'INFO: ARCHIVO DE TEXTO CON EL REPORTE TERMINAL. 17MAY2011
Dim nCant As Long
Dim nMonto As Single

On Error GoTo ErrAdm:

Select Case cMode
    
    Case "CREATE_XG"     'REPORTE X DE GRUPOS
        
        nFileNumber = FreeFile()
        cFileName = GetFileName("RG")
        Open cFileName For Output As #nFileNumber
        Print #nFileNumber, cTexto
        Print #nFileNumber, Format(Date, "LONG DATE") & " - " & Format(Time, "LONG TIME")
        Print #nFileNumber, String(43, "=")
    
    Case "CREATE_X"     'REPORTE X
        
        nFileNumber = FreeFile()
        cFileName = GetFileName("RX")
        Open cFileName For Output As #nFileNumber
        Print #nFileNumber, cTexto
        Print #nFileNumber, Format(Date, "LONG DATE") & " - " & Format(Time, "LONG TIME")
        Print #nFileNumber, String(43, "=")
    
    Case "CREATE"   'REPORTE TERMINAL Z

        nFileNumber = FreeFile()
        cFileName = GetFileName("RZ")
        
        Open GetFileName("RZ") For Output As #nFileNumber
        Print #nFileNumber, cTexto
        Print #nFileNumber, Format(Date, "LONG DATE") & " - " & Format(Time, "LONG TIME")
        Print #nFileNumber, String(43, "=")
    
    Case "APPEND"
        '//////////////////////////////////////////////////////////////////////////////
        Select Case cInfoAdicional
            Case "Y"
                nCant = CLng(cCant)
                nMonto = CSng(cMonto)
                If nCant = -1 Then
                        Print #nFileNumber, Format(cTexto, "!@@@@@@@@@@@@@@@@@@@@@@@@@") & Space(4) & Format(Format(nMonto, "STANDARD"), "@@@@@@@@@@@@@@")
                Else
                        Print #nFileNumber, Format(cTexto, "!@@@@@@@@@@@@@@@@@@@@@@@@@") & Format(nCant, "@@@@") & Format(Format(nMonto, "STANDARD"), "@@@@@@@@@@@@@@")
                End If
            Case "N"

                Print #nFileNumber, Format(cTexto, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
                'Write #nFileNumber, String(Len(cTexto), "-")
        End Select
        '//////////////////////////////////////////////////////////////////////////////
    Case "CLOSE"

        Print #nFileNumber, String(43, "=")
        Close #nFileNumber
        nFileNumber = 0
        
        'Call PutMail(cFileName)
        'Call Phyton_Send_Mail(cFileName, rs00!DESCRIP)
        
        cFileName = ""
        
End Select
On Error GoTo 0
Exit Function

ErrAdm:

If Not ErrorReporte_Email Then
    Select Case Err.Number
        Case 76
            ShowMsg "Error (" & cMode & ") Folders de Correo. " & vbCrLf & Err.Number & " - " & Err.description, vbYellow, vbRed
            EscribeLog "Error (" & cMode & ") Folders de Correo. " & Err.Number & " - " & Err.description
        Case 52
            ShowMsg "Error (" & cMode & ") Folders de Correo. " & vbCrLf & Err.Number & " - " & Err.description, vbYellow, vbRed
            EscribeLog "Error (" & cMode & ") Folders de Correo. " & Err.Number & " - " & Err.description
        Case Else
            EscribeLog "Error (" & cMode & ") Folders de Correo. " & Err.Number & " - " & Err.description
    End Select
End If
ErrorReporte_Email = True
Resume Next
End Function

Private Function GetPagoParcial(nMesa As Long, rsTABLA As ADODB.Recordset) As Single
On Error Resume Next
rsTABLA.MoveFirst
On Error GoTo 0
rsTABLA.Find "MESA = " & nMesa
GetPagoParcial = 0#
If rsTABLA.EOF Then
Else
    Do While rsTABLA!MESA = nMesa
        GetPagoParcial = GetPagoParcial + rsTABLA!MONTO
        rsTABLA.MoveNext
        If rsTABLA.EOF Then Exit Function
    Loop
End If
End Function
Private Function GetPagosParcialesFromCAJERO(nCajero As Long, ByRef nContador As Integer, rsTABLA As ADODB.Recordset) As Single
On Error Resume Next
rsTABLA.MoveFirst
On Error GoTo 0
If nCajero > 0 Then
    rsTABLA.Find "CAJERO = " & nCajero
    GetPagosParcialesFromCAJERO = 0#
    If rsTABLA.EOF Then
    Else
        Do While rsTABLA!CAJERO = nCajero
            GetPagosParcialesFromCAJERO = GetPagosParcialesFromCAJERO + rsTABLA!MONTO
            nContador = nContador + 1
            rsTABLA.MoveNext
            If rsTABLA.EOF Then Exit Function
        Loop
    End If
Else
    Do While Not rsTABLA.EOF
        GetPagosParcialesFromCAJERO = GetPagosParcialesFromCAJERO + rsTABLA!MONTO
        nContador = nContador + 1
        rsTABLA.MoveNext
    Loop
End If
End Function
Private Sub MesasAbiertas()
Dim rsAbiertas As Recordset
Dim cSQL As String
Dim nOpenCounter As Integer
Dim nPagoParcial As Single
Dim rsLocal As Recordset

Set rsAbiertas = New Recordset
Set rsLocal = New Recordset

nOpenCounter = 0

rsLocal.Open "SELECT MESA,MESERO,MONTO FROM TMP_PAR_PAGO ORDER BY MESA", msConn, adOpenStatic, adLockOptimistic

cSQL = "SELECT MESA, MAX(B.NOMBRE) AS NOMBRE, " & "MAX(B.APELLIDO) AS APELLIDO, SUM(PRECIO) AS PEND " & "FROM TMP_TRANS AS A, MESEROS AS B " & "WHERE A.MESERO = B.NUMERO " & "GROUP BY A.MESA " & "ORDER BY 2,3,1"

rsAbiertas.Open cSQL, msConn, adOpenStatic, adLockOptimistic

If rsAbiertas.EOF Then
    rsAbiertas.Close
    Set rsAbiertas = Nothing
    Exit Sub
End If
 
Print2_OPOS_Dev rs00!DESCRIP
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev Date & Space(2) & Time
Print2_OPOS_Dev "REPORTE DE MESAS ABIERTAS"
Print2_OPOS_Dev "AL CAMBIO DE TURNO"
Print2_OPOS_Dev Space(2)
'''''''Printer.FontUnderline = True
Print2_OPOS_Dev "Mesa #   Pendiente"
Print2_OPOS_Dev "=================="
'''''''Printer.FontUnderline = False
cMesero = Trim(rsAbiertas!nombre) & " " & Trim(rsAbiertas!APELLIDO)
'Print2_OPOS_Dev "Mesero: " & cMesero
Do Until rsAbiertas.EOF
    Print2_OPOS_Dev Format(rsAbiertas!MESA, "@@") & Space(8) & Format(Format(rsAbiertas!PEND, "#,##0.00"), "@@@@@@@@")
    nPagoParcial = 0
    nPagoParcial = GetPagoParcial(rsAbiertas!MESA, rsLocal)
    If nPagoParcial > 0 Then
        Print2_OPOS_Dev "PAGO PARCIAL : " & Format(Format(nPagoParcial, "#,##0.00"), "@@@@@@@@") & " **"
    End If
    nOpenCounter = nOpenCounter + 1
    rsAbiertas.MoveNext
    On Error Resume Next
    If rsAbiertas.EOF = True Then
        Exit Do
    End If
    If cMesero <> Trim(rsAbiertas!nombre) & " " & Trim(rsAbiertas!APELLIDO) Then
        Print2_OPOS_Dev "Mesero: " & cMesero & " (" & nOpenCounter & ")"
        nOpenCounter = 0
        cMesero = Trim(rsAbiertas!nombre) & " " & Trim(rsAbiertas!APELLIDO)
    End If
    
    On Error GoTo 0
Loop
Print2_OPOS_Dev "Mesero: " & cMesero & " (" & nOpenCounter & ")"

For i = 1 To 12
    Print2_OPOS_Dev Space(2)
Next
Sleep MUNBYNSeconds
Sys_Pos.Coptr1.CutPaper 100

rsLocal.Close
rsAbiertas.Close
Set rsLocal = Nothing
End Sub

'''''''Private Sub MesasAbiertas()
'''''''Dim rsAbiertas As Recordset
'''''''
'''''''Set rsAbiertas = New Recordset
'''''''
'''''''rsAbiertas.Open "SELECT MESA,SUM(PRECIO) AS PEND FROM TMP_TRANS" & '''''''        " GROUP BY MESA", msConn, adOpenStatic, adLockOptimistic
'''''''
'''''''If rsAbiertas.EOF Then
'''''''    rsAbiertas.Close
'''''''    Set rsAbiertas = Nothing
'''''''    Exit Sub
'''''''End If
'''''''
'''''''Sys_Pos.Coptr1.PrintNormal PtrSReceipt, rs00!DESCRIP & Chr(&HD) & Chr(&HA)
'''''''Sys_Pos.Coptr1.PrintNormal PtrSReceipt, Space(2) & Chr(&HD) & Chr(&HA)
'''''''Sys_Pos.Coptr1.PrintNormal PtrSReceipt, Date & Chr(&HD) & Chr(&HA)
'''''''Sys_Pos.Coptr1.PrintNormal PtrSReceipt, "REPORTE DE MESAS ABIERTAS" & Chr(&HD) & Chr(&HA)
'''''''Sys_Pos.Coptr1.PrintNormal PtrSReceipt, "AL CAMBIO DE TURNO" & Chr(&HD) & Chr(&HA)
'''''''Sys_Pos.Coptr1.PrintNormal PtrSReceipt, Space(2) & Chr(&HD) & Chr(&HA)
''''''''''''''Printer.FontUnderline = True
'''''''Sys_Pos.Coptr1.PrintNormal PtrSReceipt, "Mesa    Pendiente" & Chr(&HD) & Chr(&HA)
''''''''''''''Printer.FontUnderline = False
'''''''Do Until rsAbiertas.EOF
'''''''    Sys_Pos.Coptr1.PrintNormal PtrSReceipt, rsAbiertas!mesa & Space(8) & Format(rsAbiertas!PEND, "STANDARD") & Chr(&HD) & Chr(&HA)
'''''''    rsAbiertas.MoveNext
'''''''Loop
'''''''
'''''''For i = 1 To 10
'''''''    Sys_Pos.Coptr1.PrintNormal PtrSReceipt, Space(2) & Chr(&HD) & Chr(&HA)
'''''''Next
'''''''Sys_Pos.Coptr1.CutPaper 100
'''''''
'''''''rsAbiertas.Close
'''''''End Sub
Private Sub RepDepX()
Dim rsTran As Recordset
Dim rsLocalDepto As Recordset
Dim nTotDepto As Double
Dim MiLen1 As Integer
Dim Milen2 As Integer

nTotDepto = 0#
Set rsTran = New Recordset
Set rsLocalDepto = New Recordset

'ABRE DEPARTAMENTOS
rsLocalDepto.Open "SELECT CODIGO,DESCRIP,CORTO FROM DEPTO", msConn, adOpenDynamic, adLockOptimistic

'"SELECT DEPTO,COUNT(DEPTO) AS X_COUNT,SUM(PRECIO) AS VALOR"
rsTran.Open "SELECT DEPTO,SUM(CANT) AS X_COUNT,SUM(PRECIO) AS VALOR " & " FROM TRANSAC " & "GROUP BY DEPTO ORDER BY DEPTO", msConn, adOpenStatic, adLockOptimistic


Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev Date & Space(2) & Time
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev "REPORTE DEPARTAMENTAL (X)"
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev rs00!DESCRIP
Print2_OPOS_Dev rs00!RAZ_SOC
Print2_OPOS_Dev "RUC:" & rs00!ruc
Print2_OPOS_Dev "SERIAL:" & rs00!SERIAL
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev "CONTADOR TRANS : " & (nNUMTrans)
Print2_OPOS_Dev "CONTADOR X : " & (rs00!X_CDEP + 1)
Print2_OPOS_Dev "REPORTE DEPART."
Print2_OPOS_Dev Space(2)

Do Until rsTran.EOF
    rsLocalDepto.MoveFirst
    rsLocalDepto.Find "CODIGO = " & rsTran!depto
    If Not rsLocalDepto.EOF Then
        MiLen1 = Len(rsTran!X_COUNT)
        Milen2 = Len(Format(rsTran!Valor, "STANDARD"))
        Print2_OPOS_Dev FormatTexto(rsLocalDepto!corto, 13) & Space(4 - MiLen1) & rsTran!X_COUNT & Space(9 - Milen2) & Format(rsTran!Valor, "STANDARD")
    Else
        MsgBox "Hay un Error con los datos Departamentales", vbCritical, BoxTit
    End If
    nTotDepto = nTotDepto + rsTran!Valor
    rsTran.MoveNext
Loop

'***************************
'INFO: 15NOV2018
'***************************
Call InfoDomicilio(nTotDepto, "X")
'***************************

Print2_OPOS_Dev Space(2)
MiLen1 = Len(Format(nTotDepto, "CURRENCY"))
If MiLen1 > 11 Then MiLen1 = 11
Print2_OPOS_Dev "TOTAL DEPTOS : " & Space(11 - MiLen1) & Format(nTotDepto, "CURRENCY")
Print2_OPOS_Dev Space(2)
For i = 1 To 10
    Print2_OPOS_Dev Space(2)
Next
Sleep MUNBYNSeconds
Sys_Pos.Coptr1.CutPaper 100
rsTran.Close

msConn.BeginTrans
'Actualiza el conteo de X's y de TRANSAC
msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1,X_CDEP = X_CDEP + 1"
msConn.CommitTrans

End Sub
Private Sub RepMeseroZ()
Dim rsMeseros As Recordset
Dim nValorMesas As Double
Dim MiLen1 As Integer, Milen2 As Integer

Set rsMeseros = New Recordset
nValorMesas = 0#

rsMeseros.Open "SELECT * FROM MESEROS WHERE VALOR > 0 " & " AND NUMERO <> 999 ORDER BY NUMERO", msConn, adOpenStatic, adLockOptimistic
ProgBar.value = 30
'Open "c:\mifile.txt" For Output As #1   ' Open file.
Print2_OPOS_Dev Date & Space(2) & Time
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev "REPORTE MESEROS (Z)"
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev rs00!DESCRIP
Print2_OPOS_Dev rs00!RAZ_SOC
Print2_OPOS_Dev "RUC:" & rs00!ruc
Print2_OPOS_Dev "SERIAL:" & rs00!SERIAL
Print2_OPOS_Dev Space(2)
'''- SEGUN HACIENDA Y TESORO -'''Printer.Print "CONTADOR TRANS : " & (rs00!TRANS + 1)
Print2_OPOS_Dev "CONTADOR Z : " & (rs00!Z_CMESEROS + 1)
Print2_OPOS_Dev Space(2)
    
Do Until rsMeseros.EOF
    MiLen1 = Len(Format(rsMeseros!Valor, "STANDARD"))
    Milen2 = Len(rsMeseros!numero)
    Print2_OPOS_Dev "Mesero : " & Space(5 - Milen2) & rsMeseros!numero & Space(15 - MiLen1) & Format(rsMeseros!Valor, "STANDARD")
    nValorMesas = nValorMesas + Format(rsMeseros!Valor, "STANDARD")
    rsMeseros.MoveNext
Loop
ProgBar.value = 60
Print2_OPOS_Dev Space(2)
Milen2 = Len(Format(nValorMesas, "CURRENCY"))
Print2_OPOS_Dev "TOTAL : " & Space(21 - Milen2) & Format(nValorMesas, "CURRENCY")
Print2_OPOS_Dev Space(2)
'Close #1
For i = 1 To 10
    Print2_OPOS_Dev Space(2)
Next
Sleep MUNBYNSeconds
Sys_Pos.Coptr1.CutPaper 100
ProgBar.value = 100

msConn.BeginTrans
'''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1,X_CMESEROS = 0,Z_CMESEROS = Z_CMESEROS + 1"
msConn.Execute "UPDATE ORGANIZACION SET X_CMESEROS = 0,Z_CMESEROS = Z_CMESEROS + 1"
msConn.Execute "UPDATE MESEROS SET VALOR = 0.00, X_COUNT = 0, Z_COUNT = 0 " & "WHERE NUMERO <> 999"
msConn.CommitTrans
ProgBar.value = 0
End Sub
Private Sub RepMeseroX()
Dim rsMeseros As Recordset
Dim nValorMesas As Double
Dim MiLen1 As Integer
Dim Milen2 As Integer

Set rsMeseros = New Recordset
nValorMesas = 0#

rsMeseros.Open "SELECT * FROM MESEROS WHERE VALOR > 0 " & " AND NUMERO <> 999 ORDER BY NUMERO", msConn, adOpenStatic, adLockOptimistic

'Open "c:\mifile.txt" For Output As #1   ' Open file.
Print2_OPOS_Dev Date & Space(2) & Time
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev "REPORTE MESEROS (X)"
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev rs00!DESCRIP
Print2_OPOS_Dev rs00!RAZ_SOC
Print2_OPOS_Dev "RUC:" & rs00!ruc
Print2_OPOS_Dev "SERIAL:" & rs00!SERIAL
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev "CONTADOR TRANS : " & (nNUMTrans)
Print2_OPOS_Dev "CONTADOR X : " & (rs00!X_CMESEROS + 1)
Print2_OPOS_Dev Space(2)
    
Do Until rsMeseros.EOF
    MiLen1 = Len(Format(rsMeseros!Valor, "STANDARD"))
    Milen2 = Len(rsMeseros!numero)
    Print2_OPOS_Dev "Mesero : " & Space(5 - Milen2) & rsMeseros!numero & Space(11 - MiLen1) & Format(rsMeseros!Valor, "STANDARD")
    nValorMesas = nValorMesas + Format(rsMeseros!Valor, "STANDARD")
    rsMeseros.MoveNext
Loop
Print2_OPOS_Dev Space(2)
Milen2 = Len(Format(nValorMesas, "CURRENCY"))
Print2_OPOS_Dev "TOTAL : " & Space(17 - Milen2) & Format(nValorMesas, "CURRENCY")
Print2_OPOS_Dev Space(2)
'Close #1
For i = 1 To 10
    Print2_OPOS_Dev Space(2)
Next
Sleep MUNBYNSeconds
Sys_Pos.Coptr1.CutPaper 100

msConn.BeginTrans
msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1, X_CMESEROS = X_CMESEROS + 1"
msConn.CommitTrans
End Sub
Private Sub RepMesZ()
Dim rsMesas As Recordset
Dim nValorMesas As Double
Dim MiLen1 As Integer
Dim Milen2 As Integer

Set rsMesas = New Recordset
nValorMesas = 0#

rsMesas.Open "SELECT * FROM MESAS WHERE VALOR > 0 AND NUMERO > 0" & " ORDER BY NUMERO", msConn, adOpenStatic, adLockOptimistic
ProgBar.value = 40
'Open "c:\mifile.txt" For Output As #1   ' Open file.
Print2_OPOS_Dev Date & Space(2) & Time
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev "REPORTE DE MESAS (Z)"
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev rs00!DESCRIP
Print2_OPOS_Dev rs00!RAZ_SOC
Print2_OPOS_Dev "RUC:" & rs00!ruc
Print2_OPOS_Dev "SERIAL:" & rs00!SERIAL
Print2_OPOS_Dev Space(2)
'''- SEGUN HACIENDA Y TESORO -'''Printer.Print "CONTADOR TRANS : " & (rs00!TRANS + 1)
Print2_OPOS_Dev "CONTADOR Z : " & (rs00!Z_CMESAS + 1)
Print2_OPOS_Dev Space(2)
    
Do Until rsMesas.EOF
    MiLen1 = Len(Format(rsMesas!Valor, "STANDARD"))
    Milen2 = Len(rsMesas!numero)
    Print2_OPOS_Dev "Mesa : " & Space(5 - Milen2) & rsMesas!numero & Space(15 - MiLen1) & Format(rsMesas!Valor, "STANDARD")
    nValorMesas = nValorMesas + Format(rsMesas!Valor, "STANDARD")
    rsMesas.MoveNext
Loop
ProgBar.value = 80
Print2_OPOS_Dev Space(2)
Milen2 = Len(Format(nValorMesas, "CURRENCY"))
Print2_OPOS_Dev "TOTAL : " & Space(19 - Milen2) & Format(nValorMesas, "CURRENCY")
Print2_OPOS_Dev Space(2)
'Close #1
For i = 1 To 10
    Print2_OPOS_Dev Space(2)
Next
Sleep MUNBYNSeconds
Sys_Pos.Coptr1.CutPaper 100

ProgBar.value = 100
msConn.BeginTrans
'''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1, X_CMESAS = 0, Z_CMESAS = Z_CMESAS + 1"
msConn.Execute "UPDATE ORGANIZACION SET X_CMESAS = 0, Z_CMESAS = Z_CMESAS + 1"
msConn.Execute "UPDATE MESAS SET VALOR = 0.00 WHERE NUMERO > 0"
msConn.CommitTrans
ProgBar.value = 0
End Sub
Private Sub RepMesX()
Dim rsMesas As Recordset
Dim nValorMesas As Double
Dim MiLen1 As Integer
Dim Milen2 As Integer

Set rsMesas = New Recordset
nValorMesas = 0#

rsMesas.Open "SELECT * FROM MESAS WHERE VALOR > 0 AND NUMERO > 0" & " ORDER BY NUMERO", msConn, adOpenStatic, adLockOptimistic

'Open "c:\mifile.txt" For Output As #1   ' Open file.
Print2_OPOS_Dev Date & Space(2) & Time
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev "REPORTE DE MESAS (X)"
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev rs00!DESCRIP
Print2_OPOS_Dev rs00!RAZ_SOC
Print2_OPOS_Dev "RUC:" & rs00!ruc
Print2_OPOS_Dev "SERIAL:" & rs00!SERIAL
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev "CONTADOR TRANS : " & (nNUMTrans)
Print2_OPOS_Dev "CONTADOR X : " & (rs00!X_CMESAS + 1)
Print2_OPOS_Dev Space(2)
    
Do Until rsMesas.EOF
    MiLen1 = Len(Format(rsMesas!Valor, "STANDARD"))
    Milen2 = Len(rsMesas!numero)
    Print2_OPOS_Dev "Mesa : " & Space(5 - Milen2) & rsMesas!numero & Space(15 - MiLen1) & Format(rsMesas!Valor, "STANDARD")
    nValorMesas = nValorMesas + Format(rsMesas!Valor, "STANDARD")
    rsMesas.MoveNext
Loop
Print2_OPOS_Dev Space(2)
MiLen1 = Len(Format(nValorMesas, "CURRENCY"))
Print2_OPOS_Dev "TOTAL : " & Space(19 - MiLen1) & Format(nValorMesas, "CURRENCY")
Print2_OPOS_Dev Space(2)
'Close #1
For i = 1 To 10
    Print2_OPOS_Dev Space(2)
Next
Sleep MUNBYNSeconds
Sys_Pos.Coptr1.CutPaper 100

msConn.BeginTrans
msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1, X_CMESAS = X_CMESAS + 1"
msConn.CommitTrans

End Sub
Private Sub RepCajZ()
'REPORTE DE CAJEROS - TERMINAL - DEPARTAMENTOS
Dim rsVta_Z As Recordset
Dim rsPgo_Z As Recordset
Dim rsTran As Recordset
Dim RSPAGOS As Recordset    'Pagos
Dim rsAjustes As Recordset
Dim rsCajeros As Recordset
Dim rsProp As Recordset
Dim rsLocalDepto As Recordset
Dim rsLocTerminal As Recordset
Dim rsSuperGrp As Recordset
Dim rsHash As Recordset
Dim sqltext As String
Dim nSumVta As Double
Dim nSumTotal As Double
'INFO: nTrans Cambiando de Integer a LONG
'12ABR2013
Dim nTrans As Long
Dim MiLen1 As Integer
Dim Milen2 As Integer
Dim nErrInd As Integer
Dim errCounter As Integer
'AGREGANDO PAGO PARCIAL
Dim rsPAGOParcial As Recordset
Dim nMontoPagoParcial As Single
Dim nCounterPagoParcial As Integer

Rem NOTAS DE CREDITO DE LA FACTURACION ELECTRONICA
Dim n1 As Long
Dim n2 As Single

nSumVta = 0
nSumTotal = 0
nErrInd = 0

Set rsHash = New Recordset
Set rsVta_Z = New Recordset
Set rsPgo_Z = New Recordset
Set rsTran = New Recordset
Set RSPAGOS = New Recordset
Set rsAjustes = New Recordset
Set rsCajeros = New Recordset
Set rsProp = New Recordset
Set rsLocalDepto = New Recordset

'AGREGANDO PAGO PARCIAL
Set rsPAGOParcial = New Recordset
rsPAGOParcial.Open "SELECT MESA,CAJERO,MESERO,MONTO FROM TMP_PAR_PAGO ORDER BY CAJERO", msConn, adOpenStatic, adLockOptimistic

'cSQL = "SELECT B.ID_PROD_INV, MAX(B.DESCRIP) AS ART_INVENT,"
'cSQL = cSQL & "MAX(C.DESCRIP) AS UNID_CONSUMO,SUM(A.CANT) * MAX(B.CANT) AS CONSUMIDO "
'cSQL = cSQL & "FROM TRANSAC AS A, PLU_INVENT AS B,UNID_CONSUMO AS C "
'cSQL = cSQL & "WHERE A.CANT <> 0 AND A.PLU = B.ID_PLU AND A.ENVASE = B.ID_ENV "
'cSQL = cSQL & "AND B.ID_UNID_CONSUMO = C.ID GROUP BY B.ID_PROD_INV"
'
'INFO: SETUP DE DATOS PARA ACTUALIZAR INVENTARIO
cSQL = "SELECT B.ID_PROD_INV, "
cSQL = cSQL & "A.CANT * B.CANT AS CONSUMIDO "
cSQL = cSQL & "FROM TRANSAC AS A, PLU_INVENT AS B "
cSQL = cSQL & "WHERE A.PLU = B.ID_PLU "
cSQL = cSQL & "AND A.ENVASE = B.ID_ENV "
cSQL = cSQL & "AND A.DESCRIP NOT LIKE '%DESCUENTO%' "
cSQL = cSQL & "ORDER BY B.ID_PROD_INV"

'INFO: (12DIC2010)
With rsLOCALPLUInvent
    .Fields.Append "ID_PROD_INV", adInteger, , adFldUpdatable
    .Fields.Append "CONSUMIDO", adSingle, , adFldUpdatable
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open
End With

rsPluInvent.Open cSQL, msConn, adOpenDynamic, adLockOptimistic

On Error Resume Next
rsPluInvent.MoveFirst
On Error GoTo 0
Do While Not rsPluInvent.EOF
    rsLOCALPLUInvent.AddNew
    rsLOCALPLUInvent!ID_PROD_INV = rsPluInvent!ID_PROD_INV
    rsLOCALPLUInvent!CONSUMIDO = rsPluInvent!CONSUMIDO
    rsLOCALPLUInvent.Update
    rsPluInvent.MoveNext
Loop

rsPluInvent.Close

'INFO: RECETAS (16AAGO2009)
'INFO: SETUP DE DATOS PARA ACTUALIZAR RECETAS

cSQL = "SELECT B.ID_RECETA, A.CANT * B.CANT AS CONSUMIDO "
cSQL = cSQL & "FROM TRANSAC AS A, PLU_RECETAS AS B "
cSQL = cSQL & "WHERE A.PLU = B.ID_PLU "
cSQL = cSQL & "AND A.ENVASE = B.ID_ENV "
cSQL = cSQL & "AND A.DESCRIP NOT LIKE '%DESCUENTO%' "
cSQL = cSQL & "ORDER BY B.ID_RECETA"

With rsLOCALPLURecetas
    .Fields.Append "ID_RECETA", adInteger, , adFldUpdatable
    .Fields.Append "CONSUMIDO", adSingle, , adFldUpdatable
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open
End With

rsPLURecetas.Open cSQL, msConn, adOpenDynamic, adLockOptimistic

On Error Resume Next
rsPLURecetas.MoveFirst
On Error GoTo 0
Do While Not rsPLURecetas.EOF
    rsLOCALPLURecetas.AddNew
    rsLOCALPLURecetas!ID_RECETA = rsPLURecetas!ID_RECETA
    rsLOCALPLURecetas!CONSUMIDO = rsPLURecetas!CONSUMIDO
    rsLOCALPLURecetas.Update
    rsPLURecetas.MoveNext
Loop

rsPLURecetas.Close
'ABRE DEPARTAMENTOS
rsLocalDepto.Open "SELECT CODIGO,DESCRIP,CORTO FROM DEPTO", msConn, adOpenDynamic, adLockOptimistic

sqltext = "SELECT DISTINCT CAJERO FROM TRANSAC ORDER BY CAJERO"
rsCajeros.Open sqltext, msConn, adOpenStatic, adLockOptimistic

If rsCajeros.EOF Then
    'MsgBox "IMPRIMIENDO EL REPORTE, AUNQUE NO EXISTEN TRANSACCIONES", vbInformation, BoxTit
    ShowMsg "IMPRIMIENDO EL REPORTE, AUNQUE NO EXISTEN TRANSACCIONES", vbYellow, vbRed
'    Exit Sub
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO: 28NOV2014
Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "BEGIN")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

MesasAbiertas

On Error Resume Next    'NUEVO PARA CORRER REPORTE EN ZZZZ
rsCajeros.MoveFirst
rsCajeros.MoveLast

If rsCajeros.RecordCount = 0 Or rsCajeros!CAJERO = "" Then
    'MsgBox "IMPRIMIENDO EL REPORTE, AUNQUE NO EXISTEN TRANSACCIONES", vbInformation, BoxTit
    ShowMsg "IMPRIMIENDO EL REPORTE, AUNQUE NO EXISTEN TRANSACCIONES", vbYellow, vbRed
    'Exit Sub
End If
On Error GoTo 0         'NUEVO PARA CORRER REPORTE EN ZZZZ

'Open "c:\mifile.txt" For Output As #1   ' Open file.

On Error GoTo AjustaMilen:

ProgBar.value = 5
sqltext = "SELECT * FROM pagos WHERE CODIGO <> 999 ORDER BY CODIGO"
RSPAGOS.Open sqltext, msConn, adOpenStatic, adLockOptimistic

ProgBar.value = 10

rsCajeros.MoveFirst
Do Until rsCajeros.EOF

    ProgBar.value = 20
    
    rsTran.Open "SELECT distinct NUM_TRANS FROM TRANSAC " & " WHERE CAJERO = " & rsCajeros!CAJERO, msConn, adOpenStatic, adLockOptimistic
    
    rsTran.MoveFirst
    rsTran.MoveLast
    nTrans = rsTran.RecordCount
    rsTran.MoveFirst
    
''''sqltext = "SELECT b.nombre,b.apellido,a.cajero, sum(a.precio) as Ventas " & " FROM TRANSAC as a, cajeros as b WHERE a.cajero = " & rsCajeros!CAJERO & " AND b.numero = " & rsCajeros!CAJERO & " GROUP BY a.cajero,b.nombre,b.apellido "

    '------------------------------------------------------------------------
    'INFO: 17AGO2014. CORRIGIENDO ERROR EN LA IMPRESION DEL IMPUESTO
    sqltext = "SELECT b.nombre,b.apellido,a.cajero,b.z_c, sum(a.precio) as Ventas, "
    sqltext = sqltext & " sum(a.precio * (A.CON_TAX/100)) AS TAX_CAJERO"
    '------------------------------------------------------------------------
    sqltext = sqltext & " FROM TRANSAC as a, cajeros as b "
    sqltext = sqltext & " WHERE a.cajero = " & rsCajeros!CAJERO
    sqltext = sqltext & " AND b.numero = " & rsCajeros!CAJERO
    sqltext = sqltext & " GROUP BY a.cajero,b.nombre,b.apellido,b.z_c "

    'VALOR EN VENTAS DEL CAJERO CON rsVta_Z
    rsVta_Z.Open sqltext, msConn, adOpenStatic, adLockOptimistic
    
    If rsVta_Z.RecordCount = 0 Then
        'MsgBox "EL CAJERO NO TIENE VENTAS, NO SE IMPRIMIRA REPORTE EN Z"
        rsCajeros.MoveNext
    End If
    
    'TODOS LOS PAGOS RECIBIDOS
    'ESTOY SACANDO EL DESCUENTO GLOBAL (99) DE AQUI PARA PONERLO EN
    'LOS AJUSTES
    'INFO: 26MAR2019. CARGOS DELIVERY
    sqltext = "SELECT a.cajero,a.tipo_pago,SUM(a.monto) AS Valor, "
    sqltext = sqltext & " COUNT(a.tipo_pago) as Z_COUNT, MAX(B.TIPO) AS T_PAGO "
    sqltext = sqltext & " FROM transac_pago as a, PAGOS AS B "
    sqltext = sqltext & " WHERE a.cajero = " & rsCajeros!CAJERO
    sqltext = sqltext & " AND A.TIPO_PAGO = B.CODIGO "
    sqltext = sqltext & " AND a.tipo_pago <> 99 "
    sqltext = sqltext & " AND B.TIPO <> 'CA' "
    sqltext = sqltext & " GROUP BY a.cajero,a.TIPO_PAGO"
    
    rsPgo_Z.Open sqltext, msConn, adOpenStatic, adLockOptimistic
    
    'sqltext = "SELECT MID(a.TIPO,1,2) AS CORREC, COUNT(a.lin) as Z_COUNT, SUM(a.precio) as valor FROM TRANSAC as a WHERE A.CAJERO = " & rsCajeros!CAJERO & " AND MID(A.TIPO,1,1) <> ' ' GROUP BY MID(a.TIPO,1,2)"
    'sqltext = "SELECT MID(a.TIPO,1,2) AS CORREC,DESCUENTO, "
    '    " COUNT(a.lin) as Z_COUNT, SUM(a.precio) as valor "
    '2 de Nov 1999
    'COUNT(a.lin) as Z_COUNT, SUM(a.precio_unit) as valor
    '5 de Nov
    sqltext = "SELECT MID(a.TIPO,1,2) AS CORREC,a.DESCUENTO, "
    sqltext = sqltext & " COUNT(a.lin) as Z_COUNT, SUM(a.precio) as valor "
    sqltext = sqltext & " FROM TRANSAC as a "
    sqltext = sqltext & " WHERE A.CAJERO = " & rsCajeros!CAJERO
    sqltext = sqltext & " AND MID(A.TIPO,1,1) <> ' ' "
    sqltext = sqltext & " GROUP BY MID(a.TIPO,1,2),a.DESCUENTO"
    
    rsAjustes.Open sqltext, msConn, adOpenStatic, adLockOptimistic

    Print2_OPOS_Dev Date & Space(2) & Time
    Print2_OPOS_Dev Space(2)
    Print2_OPOS_Dev "REPORTE DE CAJEROS (Z)"
    Print2_OPOS_Dev Space(2)
    Print2_OPOS_Dev rs00!DESCRIP
    'INFO: AGREGANDO RAZON SOCIAL (AGO 2009)
    Print2_OPOS_Dev rs00!RAZ_SOC
    Print2_OPOS_Dev "RUC:" & rs00!ruc
    Print2_OPOS_Dev "SERIAL:" & rs00!SERIAL
    Print2_OPOS_Dev Space(2)
    '''- SEGUN HACIENDA Y TESORO -'''Printer.Print "CONTADOR TRANS : " & (rs00!TRANS + 1)
    Print2_OPOS_Dev "CONTADOR Z : " & (rsVta_Z!z_c + 1)
    
    'INFO: AGO2009, FIX EN EL LARGO DEL NOMBRE DEL CAJERO
    'Print2_OPOS_Dev "CAJERO : " & rsVta_Z!nombre & ", " & rsVta_Z!apellido, "CAJERO : " & rsVta_Z!nombre & ", " & rsVta_Z!apellido
    Print2_OPOS_Dev "CAJERO : " & rsVta_Z!nombre & ", " & Mid(rsVta_Z!APELLIDO, 1, 1) & "."
    Print2_OPOS_Dev Space(2)
    
    Do Until rsVta_Z.EOF
        MiLen1 = Len(nTrans)
        Milen2 = Len(Format(rsVta_Z!VENTAS, "STANDARD"))
        'INFO: nTrans SI LAS CANTIDADES SON MUY GRANDES, ESTOY CAMBIANDO LA LONGITUD DEL CAMPO A 4
        '12ABR2013
        If MiLen1 > 4 Then MiLen1 = 4
        Print2_OPOS_Dev "VENTA DEL DIA:" & Space(4 - MiLen1) & nTrans & Space(11 - Milen2) & Format(rsVta_Z!VENTAS, "STANDARD")
        nSumTotal = nSumTotal + rsVta_Z!VENTAS
        rsVta_Z.MoveNext
    Loop
    
    Print2_OPOS_Dev Space(2)
    ''''''''''Printer.FontUnderline = True
    'Printer.Print "DESGLOSE DE INGRESOS"
    Print2_OPOS_Dev "TOTALES DE CAJA"
    ''''''''Printer.FontUnderline = False
    Print2_OPOS_Dev Space(2)
    
    'AGREGANDO DETALLE DE PAGO PARCIAL
    'GRACIAS A BAR GRILL 50 (25MAR2004)
    nMontoPagoParcial = 0#
    nCounterPagoParcial = 0
    nMontoPagoParcial = GetPagosParcialesFromCAJERO(rsCajeros!CAJERO, nCounterPagoParcial, rsPAGOParcial)
    If nMontoPagoParcial > 0 Then
        MiLen1 = Len(nCounterPagoParcial)
        Milen2 = Len(Format(nMontoPagoParcial, "STANDARD"))
        'INFO: nTrans SI LAS CANTIDADES SON MUY GRANDES, ESTOY CAMBIANDO LA LONGITUD DEL CAMPO A 4
        '12ABR2013
        If MiLen1 > 4 Then MiLen1 = 4
        Print2_OPOS_Dev "PAGO PARCIAL  " & Space(4 - MiLen1) & nCounterPagoParcial & Space(13 - Milen2) & Format(nMontoPagoParcial, "STANDARD")
        nSumVta = nSumVta + nMontoPagoParcial
    End If

    Do Until rsPgo_Z.EOF
        RSPAGOS.MoveFirst
        RSPAGOS.Find "CODIGO = " & rsPgo_Z!TIPO_PAGO
        If Not RSPAGOS.EOF Then
            MiLen1 = Len(rsPgo_Z!Z_COUNT)
            Milen2 = Len(Format(rsPgo_Z!Valor, "STANDARD"))
            'INFO: nTrans SI LAS CANTIDADES SON MUY GRANDES, ESTOY CAMBIANDO LA LONGITUD DEL CAMPO A 4
            '12ABR2013
            If MiLen1 > 4 Then MiLen1 = 4
            Print2_OPOS_Dev FormatTexto(RSPAGOS!DESCRIP, 13) & Space(4 - MiLen1) & rsPgo_Z!Z_COUNT & Space(13 - Milen2) & Format(rsPgo_Z!Valor, "STANDARD")
        Else
            MiLen1 = 1
            Milen2 = Len(Format(0#, "STANDARD"))
            'INFO: nTrans SI LAS CANTIDADES SON MUY GRANDES, ESTOY CAMBIANDO LA LONGITUD DEL CAMPO A 4
            '12ABR2013
            If MiLen1 > 4 Then MiLen1 = 4
            Print2_OPOS_Dev "OTRO PAGO    " & Space(4 - MiLen1) & 0 & Space(13 - MiLen1) & Format(0#, "STANDARD")
        End If
        nSumVta = nSumVta + rsPgo_Z!Valor
        rsPgo_Z.MoveNext
    Loop

    Print2_OPOS_Dev Space(2)
    MiLen1 = Len(Format(nSumVta, "CURRENCY"))
    Print2_OPOS_Dev "SUBTOTAL: " & Space(20 - MiLen1) & Format(nSumVta, "CURRENCY")
    Print2_OPOS_Dev "------------------------------"
    
    sqltext = "SELECT TIPO_PAGO,COUNT(TIPO_PAGO) AS Z_COUNT, "
    sqltext = sqltext & " SUM(MONTO) AS VALOR FROM TRANSAC_PROP "
    sqltext = sqltext & " WHERE CAJERO = " & rsCajeros!CAJERO
    sqltext = sqltext & " GROUP BY TIPO_PAGO "
    rsProp.Open sqltext, msConn, adOpenStatic, adLockOptimistic
    
    Do Until rsProp.EOF
        RSPAGOS.MoveFirst
        RSPAGOS.Find "CODIGO = " & rsProp!TIPO_PAGO
        If Not RSPAGOS.EOF Then
            MiLen1 = Len(rsProp!Z_COUNT)
            Milen2 = Len(Format(rsProp!Valor, "STANDARD"))
            'INFO: nTrans SI LAS CANTIDADES SON MUY GRANDES, ESTOY CAMBIANDO LA LONGITUD DEL CAMPO A 4
            '12ABR2013
            If MiLen1 > 4 Then MiLen1 = 4
            'INFO: 26MAR2019. CARGOS DELIVERY
            'Print2_OPOS_Dev "Propina " & FormatTexto(RSPAGOS!DESCRIP, 5) & Space(4 - MiLen1) & rsProp!Z_COUNT & Space(13 - Milen2) & Format(rsProp!Valor, "STANDARD")
            Print2_OPOS_Dev FormatTexto(RSPAGOS!DESCRIP, 15) & Space(4 - MiLen1) & rsProp!Z_COUNT & Space(13 - Milen2) & Format(rsProp!Valor, "STANDARD")
        End If
        rsProp.MoveNext
    Loop
    
    rsProp.Close
    
    Print2_OPOS_Dev Space(2)
    '''''Printer.FontUnderline = True
    Print2_OPOS_Dev "AJUSTES"
    '''''Printer.FontUnderline = False
    Print2_OPOS_Dev Space(2)
    
'    If CHECK_NOTAS_DE_CREDITO_FE(True, n1, n2) Then
'        Print2_OPOS_Dev Space(2)
'        Print2_OPOS_Dev Space(2)
'        Print2_OPOS_Dev "NOTAS DE CREDITO    (" & n1 & ")    " & Format(n2, "STANDARD")
'        Print2_OPOS_Dev Space(2)
'        Print2_OPOS_Dev Space(2)
'    End If
    
    Do Until rsAjustes.EOF
        MiLen1 = Len(rsAjustes!Z_COUNT)
        Milen2 = Len(Format(rsAjustes!Valor, "STANDARD"))
        If rsAjustes!CORREC = "EC" Then
            Print2_OPOS_Dev "CORRECCION " & Space(8 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!Valor, "STANDARD")
        ElseIf rsAjustes!CORREC = "VO" Then
            Print2_OPOS_Dev "ANULACION " & Space(8 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!Valor, "STANDARD")
        ElseIf rsAjustes!CORREC = "DC" Then
            'INFO: nTrans SI LAS CANTIDADES SON MUY GRANDES, ESTOY CAMBIANDO LA LONGITUD DEL CAMPO A 4
            '12ABR2013
            If MiLen1 > 4 Then MiLen1 = 4
            Print2_OPOS_Dev "DESCUENTO " & Format(rsAjustes!DESCUENTO, "0.00") & Space(4 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!Valor, "STANDARD")
        End If
        rsAjustes.MoveNext
    Loop

    'PREPARA INFO PARA DESCUENTO GLOBAL
    rsAjustes.Close
    sqltext = "SELECT a.cajero,a.tipo_pago,SUM(a.monto) AS Valor, "
    sqltext = sqltext & " COUNT(a.tipo_pago) as Z_COUNT "
    sqltext = sqltext & " FROM transac_pago as a "
    sqltext = sqltext & " WHERE a.cajero = " & rsCajeros!CAJERO
    sqltext = sqltext & " AND a.tipo_pago = 99 "
    sqltext = sqltext & " GROUP BY a.cajero,a.TIPO_PAGO"
    
    rsAjustes.Open sqltext, msConn, adOpenStatic, adLockOptimistic
    
    If Not rsAjustes.EOF Then
        RSPAGOS.MoveFirst
        RSPAGOS.Find "CODIGO = " & rsAjustes!TIPO_PAGO
        If Not RSPAGOS.EOF Then
            MiLen1 = Len(rsAjustes!Z_COUNT)
            Milen2 = Len(Format(rsAjustes!Valor, "STANDARD"))
            Print2_OPOS_Dev FormatTexto(RSPAGOS!DESCRIP, 13) & Space(5 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!Valor, "STANDARD")
        End If
    End If
    
    msConn.BeginTrans
    '''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
    'Incrementa el contador en Z y Resetea el de X a 0
    msConn.Execute "UPDATE CAJEROS SET Z_C = Z_C + 1, X_C = 0 " & " WHERE NUMERO = " & rsCajeros!CAJERO
    msConn.CommitTrans
    
    rsCajeros.MoveNext
    If Not rsCajeros.EOF = True Then
        For i = 1 To 10
            Print2_OPOS_Dev Space(2)
        Next
        Sleep MUNBYNSeconds
        Sys_Pos.Coptr1.CutPaper 100
    End If
    
    rsVta_Z.Close
    rsPgo_Z.Close
    rsTran.Close
    rsAjustes.Close
    nSumVta = 0
Loop

nSumVta = 0
nSumTotal = 0
ProgBar.value = 40


'INFO: LA IMPRESORA EPSON, O LA QUE SE PONGA, SE LE LLENA EL BUFFER
'Y EN UN MOMENTO DADO, DEJA DE IMPRIMIR Y EL PROGRAMA REVIENTA Y EN OCACIONES NO CIERRA LA Z
'14MAY2012 / 27MAY2012 / 14JUN2012

Me.MousePointer = vbHourglass
Select Case OPOS_DevName
    Case "LR3000", "MP4200TH", "MP4200", "TM-T20E", "TM-T20U"
        'INFO: SI ES UNA IMPRESORA TERMICA, ESPERAR 3 SEGUNDOS
        Sleep 3000
    Case "TM-U950P", "TM-U950", "TM-U200B", "SRP270", "MP200", "TM-U220B"
        'INFO: SI ES UNA IMPRESORA DE MATRIZ, ESPERAR 8 SEGUNDOS
        '11 SEGUNDOS EN SOLOMIX, 8 SEGUNDOS EN FAST
        Sleep 8000
    Case Else
        Sleep 8000
End Select
Me.MousePointer = vbDefault

'-------------------   EL CAJERO TERMINAL  -------------------
'-------------------   EL CAJERO TERMINAL  -------------------
'-------------------   EL CAJERO TERMINAL  -------------------
'-------------------   EL CAJERO TERMINAL  -------------------
'-------------------   EL CAJERO TERMINAL  -------------------

For i = 1 To 10
    Print2_OPOS_Dev Space(2)
Next
Sleep MUNBYNSeconds
Sys_Pos.Coptr1.CutPaper 100

Set rsLocTerminal = New Recordset
rsLocTerminal.Open "SELECT Z_C FROM CAJEROS WHERE NUMERO = 999", msConn, adOpenStatic, adLockOptimistic

rsTran.Open "SELECT distinct NUM_TRANS FROM TRANSAC", msConn, adOpenStatic, adLockOptimistic

nTrans = rsTran.RecordCount

sqltext = "SELECT sum(a.precio) as Ventas FROM TRANSAC as a "
'VALOR EN VENTAS DEL CAJERO CON rsVta_Z
rsVta_Z.Open sqltext, msConn, adOpenStatic, adLockOptimistic

'''- SEGUN HACIENDA Y TESORO -'''rsHash.Open "SELECT SUM(ABS(A.PRECIO)) AS HASH_DIA FROM TRANSAC AS A", msConn, adOpenStatic, adLockOptimistic
rsHash.Open "SELECT SUM(A.PRECIO) AS HASH_DIA " & " FROM TRANSAC AS A " & " WHERE A.PRECIO > 0 ", msConn, adOpenStatic, adLockOptimistic

'TODOS LOS PAGOS RECIBIDOS
'MENOS EL DESCUENTO GLOBAL
'sqltext = "SELECT a.tipo_pago,SUM(a.monto) AS Valor, " & " COUNT(a.tipo_pago) as Z_COUNT " & " FROM transac_pago as a " & " GROUP BY a.TIPO_PAGO"

'INFO: 26MAR2019. CARGOS DELIVERY
sqltext = "SELECT a.tipo_pago,SUM(a.monto) AS Valor, "
sqltext = sqltext & " COUNT(a.tipo_pago) as Z_COUNT, MAX(B.TIPO) AS T_PAGO "
sqltext = sqltext & " FROM transac_pago as a, PAGOS AS B "
sqltext = sqltext & " WHERE A.TIPO_PAGO = B.CODIGO "
sqltext = sqltext & " AND a.tipo_pago <> 99 "
sqltext = sqltext & " AND B.TIPO <> 'CA' "
sqltext = sqltext & " GROUP BY a.TIPO_PAGO"

rsPgo_Z.Open sqltext, msConn, adOpenStatic, adLockOptimistic

'sqltext = "SELECT MID(a.TIPO,1,2) AS CORREC,DESCUENTO, "
        '" COUNT(a.lin) as Z_COUNT, SUM(a.precio) as valor "
'2 de Nov 1999
'COUNT(a.lin) as Z_COUNT, SUM(a.precio_unit) as valor
'5 de Nov
sqltext = "SELECT MID(a.TIPO,1,2) AS CORREC,DESCUENTO, "
sqltext = sqltext & " COUNT(a.lin) as Z_COUNT, SUM(a.precio) as valor "
sqltext = sqltext & " FROM TRANSAC as a "
sqltext = sqltext & " WHERE MID(A.TIPO,1,1) <> ' ' "
sqltext = sqltext & " GROUP BY MID(a.TIPO,1,2),DESCUENTO"

rsAjustes.Open sqltext, msConn, adOpenStatic, adLockOptimistic

Print2_OPOS_Dev Date & Space(2) & Time
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev "REPORTE DE TERMINAL (Z)"
        Call TEXTFileZ("CREATE", "REPORTE DE TERMINAL (Z)", , , "N")
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev rs00!DESCRIP
        Call TEXTFileZ("APPEND", rs00!DESCRIP, , , "N")
'info: AGO2009
Print2_OPOS_Dev rs00!RAZ_SOC
        Call TEXTFileZ("APPEND", rs00!RAZ_SOC, , , "N")
Print2_OPOS_Dev "RUC:" & rs00!ruc
        Call TEXTFileZ("APPEND", "RUC:" & rs00!ruc, , , "N")
Print2_OPOS_Dev "SERIAL:" & rs00!SERIAL
        Call TEXTFileZ("APPEND", "SERIAL:" & rs00!SERIAL, , , "N")
Print2_OPOS_Dev Space(2)
'''- SEGUN HACIENDA Y TESORO -'''Printer.Print "CONTADOR TRANS : " & (rs00!TRANS + 1)
Print2_OPOS_Dev "CONTADOR Z : " & (rsLocTerminal!z_c + 1)
        Call TEXTFileZ("APPEND", "CONTADOR Z : " & (rsLocTerminal!z_c + 1), , , "N")
        Call TEXTFileZ("APPEND", Space(10), , , "N")
Print2_OPOS_Dev "CAJERO : REPORTE/TERMINAL"
Print2_OPOS_Dev Space(2)

iISCTransaccion = 0
Do Until rsVta_Z.EOF
    MiLen1 = Len(nTrans)
    Milen2 = Len(Format(rsVta_Z!VENTAS, "STANDARD"))
    Print2_OPOS_Dev "VENTA DEL DIA:" & Space(5 - MiLen1) & nTrans & Space(10 - Milen2) & Format(rsVta_Z!VENTAS, "STANDARD")
            Call TEXTFileZ("APPEND", "VENTA DEL DIA:", nTrans, rsVta_Z!VENTAS, "Y")
            Call TEXTFileZ("APPEND", Space(10), , , "N")
    nSumTotal = nSumTotal + IIf(IsNull(rsVta_Z!VENTAS), 0, rsVta_Z!VENTAS)
    rsVta_Z.MoveNext
Loop

ProgBar.value = 45
Print2_OPOS_Dev Space(2)
'Printer.Print "DESGLOSE DE INGRESOS"
Print2_OPOS_Dev "TOTALES DE CAJA"
        Call TEXTFileZ("APPEND", "TOTALES DE CAJA", , , "N")
        Call TEXTFileZ("APPEND", Space(10), , , "N")
Print2_OPOS_Dev Space(2)

'AGREGANDO DETALLE DE PAGO PARCIAL
'GRACIAS A BAR GRILL 50 (25MAR2004)
nMontoPagoParcial = 0#
nCounterPagoParcial = 0
nMontoPagoParcial = GetPagosParcialesFromCAJERO(-1, nCounterPagoParcial, rsPAGOParcial)
If nMontoPagoParcial > 0 Then
    MiLen1 = Len(nCounterPagoParcial)
    Milen2 = Len(Format(nMontoPagoParcial, "STANDARD"))
    'INFO: nTrans SI LAS CANTIDADES SON MUY GRANDES, ESTOY CAMBIANDO LA LONGITUD DEL CAMPO A 4
    '12ABR2013
    If MiLen1 > 4 Then MiLen1 = 4
    Print2_OPOS_Dev "PAGO PARCIAL  " & Space(4 - MiLen1) & nCounterPagoParcial & Space(13 - Milen2) & Format(nMontoPagoParcial, "STANDARD")
            Call TEXTFileZ("APPEND", "PAGO PARCIAL", nCounterPagoParcial, nMontoPagoParcial, "Y")
    nSumVta = nSumVta + nMontoPagoParcial
End If


'*********************************************
'EMPIEZAN LAS FORMAS DE PAGO REPORTE TERMINAL Z
'EMPIEZAN LAS FORMAS DE PAGO REPORTE TERMINAL Z
'EMPIEZAN LAS FORMAS DE PAGO REPORTE TERMINAL Z
'*********************************************

Do Until rsPgo_Z.EOF
    RSPAGOS.MoveFirst
    RSPAGOS.Find "CODIGO = " & rsPgo_Z!TIPO_PAGO
    'Debug.Print "----------->" & rsPgo_Z!TIPO_PAGO
    If Not RSPAGOS.EOF Then
        MiLen1 = Len(rsPgo_Z!Z_COUNT)
        Milen2 = Len(Format(rsPgo_Z!Valor, "STANDARD"))
        'INFO: nTrans SI LAS CANTIDADES SON MUY GRANDES, ESTOY CAMBIANDO LA LONGITUD DEL CAMPO A 4
        '12ABR2013
        If MiLen1 > 4 Then MiLen1 = 4
        Print2_OPOS_Dev FormatTexto(RSPAGOS!DESCRIP, 13) & Space(4 - MiLen1) & rsPgo_Z!Z_COUNT & Space(13 - Milen2) & Format(rsPgo_Z!Valor, "STANDARD")
                Call TEXTFileZ("APPEND", RSPAGOS!DESCRIP, rsPgo_Z!Z_COUNT, rsPgo_Z!Valor, "Y")
        'PRINT#1, rsPagos!descrip & Chr(9) & rsPgo_Z!Z_COUNT & Chr(9) & Format(rsPgo_Z!VALOR, "STANDARD")
    Else
        MiLen1 = 1
        Milen2 = Len(Format(0#, "STANDARD"))
        'INFO: nTrans SI LAS CANTIDADES SON MUY GRANDES, ESTOY CAMBIANDO LA LONGITUD DEL CAMPO A 4
        '12ABR2013
        If MiLen1 > 4 Then MiLen1 = 4
        Print2_OPOS_Dev "OTRO PAGO    " & Space(4 - MiLen1) & 0 & Space(13 - MiLen1) & Format(0#, "STANDARD")
                Call TEXTFileZ("APPEND", "OTRO PAGO    ", 0, 0#, "Y")
        'PRINT#1, rsPagos!descrip & Chr(9) & 0 & Chr(9) & Format(0#, "STANDARD")
    End If
    'Debug.Print "----------->" & RSPAGOS!descrip
    nSumVta = nSumVta + rsPgo_Z!Valor
    rsPgo_Z.MoveNext
Loop

ProgBar.value = 50
Print2_OPOS_Dev Space(2)
MiLen1 = Len(Format(nSumVta, "CURRENCY"))
Print2_OPOS_Dev "SUBTOTAL: " & Space(20 - MiLen1) & Format(nSumVta, "CURRENCY")
        Call TEXTFileZ("APPEND", "SUBTOTAL: ", -1, nSumVta, "Y")
Print2_OPOS_Dev "------------------------------"


'*********************************************
'TERMINAN LAS FORMAS DE PAGO REPORTE TERMINAL Z
'TERMINAN LAS FORMAS DE PAGO REPORTE TERMINAL Z
'TERMINAN LAS FORMAS DE PAGO REPORTE TERMINAL Z
'*********************************************
On Error Resume Next
Dim nITBMSToHash As Single
'rsISC.Open "SELECT * FROM ISC", msConn, adOpenDynamic, adLockOptimistic

'INFO: 24NOV2011. FIX DE ISC PARA QUE LOGUARDE X AÑO
rsISC.Open "SELECT * FROM ISC WHERE ISC_YEAR = " & Year(Date), msConn, adOpenDynamic, adLockOptimistic

'INFO: JULIO
'INFO: JULIO 2010 NO SE PUEDE PONER EL 5%, EL ITBMS ES CALCULADO x PRODUCTO
'Print2_OPOS_Dev "ITBMS (5%):" & Space(18 - MiLen1) & Format(rsISC!DIARIO, "CURRENCY"), "ITBMS (5%):" & Space(18 - MiLen1) & Format(rsISC!DIARIO, "CURRENCY")
'Print2_OPOS_Dev "ITBMS        :" & Space(18 - MiLen1) & Format(rsISC!DIARIO, "CURRENCY"), "ITBMS        :" & Space(18 - MiLen1) & Format(rsISC!DIARIO, "CURRENCY")
'UPDATE IMPRESION 26OCT2011
'INFO: SISTEMA FAST NO GUARDA EL CAMBO ISC
'20AGO2014
If rsISC.EOF Then
    MiLen1 = Len(Format(0, "CURRENCY"))
    Print2_OPOS_Dev "ITBMS    :" & Space(18 - MiLen1) & Format(0, "CURRENCY")
    Call TEXTFileZ("APPEND", "ITBMS        :", -1, 0, "Y")
    nITBMSToHash = 0#
    iISCTransaccion = 0#
Else
    MiLen1 = Len(Format(rsISC!DIARIO, "CURRENCY"))
    Print2_OPOS_Dev "ITBMS    :" & Space(18 - MiLen1) & Format(rsISC!DIARIO, "CURRENCY")
    Call TEXTFileZ("APPEND", "ITBMS        :", -1, rsISC!DIARIO, "Y")
    rsISC.Fields(Month(Date)).value = rsISC.Fields(Month(Date)).value + rsISC!DIARIO
    nITBMSToHash = rsISC!DIARIO
    rsISC!DIARIO = 0#
    rsISC.Update
    iISCTransaccion = rsISC.Fields(Month(Date)).value
End If
Call TEXTFileZ("APPEND", Space(10), , , "N")


rsISC.Close

On Error GoTo 0
'*********************************************
'VIENEN LAS PROPINAS REPORTE TERMINAL Z
'VIENEN LAS PROPINAS REPORTE TERMINAL Z
'VIENEN LAS PROPINAS REPORTE TERMINAL Z
'*********************************************

sqltext = "SELECT TIPO_PAGO,COUNT(TIPO_PAGO) AS Z_COUNT, " & " SUM(MONTO) AS VALOR FROM TRANSAC_PROP " & " GROUP BY TIPO_PAGO "
rsProp.Open sqltext, msConn, adOpenStatic, adLockOptimistic

Do Until rsProp.EOF
    RSPAGOS.MoveFirst
    RSPAGOS.Find "CODIGO = " & rsProp!TIPO_PAGO
    If Not RSPAGOS.EOF Then
        MiLen1 = Len(rsProp!Z_COUNT)
        Milen2 = Len(Format(rsProp!Valor, "STANDARD"))
        'INFO: nTrans SI LAS CANTIDADES SON MUY GRANDES, ESTOY CAMBIANDO LA LONGITUD DEL CAMPO A 4
        '12ABR2013
        If MiLen1 > 4 Then MiLen1 = 4
        'INFO: 26MAR2019. CARGOS DELIVERY
        'Print2_OPOS_Dev "Propina " & FormatTexto(RSPAGOS!DESCRIP, 5) & Space(4 - MiLen1) & rsProp!Z_COUNT & Space(13 - Milen2) & Format(rsProp!Valor, "STANDARD")
        '        Call TEXTFileZ("APPEND", "Propina " & RSPAGOS!DESCRIP, rsProp!Z_COUNT, rsProp!Valor, "Y")
        Print2_OPOS_Dev FormatTexto(RSPAGOS!DESCRIP, 15) & Space(4 - MiLen1) & rsProp!Z_COUNT & Space(13 - Milen2) & Format(rsProp!Valor, "STANDARD")
                Call TEXTFileZ("APPEND", RSPAGOS!DESCRIP, rsProp!Z_COUNT, rsProp!Valor, "Y")
        
        'PRINT#1, "Propina " & Mid(rsPagos!descrip, 1, 5) & Chr(9) & rsProp!X_COUNT & Chr(9) & Format(rsProp!VALOR, "STANDARD")
    End If
    rsProp.MoveNext
Loop
rsProp.Close

'*********************************************
'TERMINAN LAS PROPINAS REPORTE TERMINAL Z
'TERMINAN LAS PROPINAS REPORTE TERMINAL Z
'TERMINAN LAS PROPINAS REPORTE TERMINAL Z
'*********************************************

ProgBar.value = 60
Print2_OPOS_Dev Space(2)
'''''''''Printer.FontUnderline = True
Print2_OPOS_Dev "AJUSTES"
        Call TEXTFileZ("APPEND", Space(10), , , "N")
        Call TEXTFileZ("APPEND", "AJUSTES", 0, 0, "N")
        Call TEXTFileZ("APPEND", Space(10), , , "N")
''''''''Printer.FontUnderline = False
Print2_OPOS_Dev Space(2)

'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Rem NOTAS DE CREDITO ELECTRONICA y DIFERENCIAS CON LA FACTURACION NORMAL
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

    If CHECK_NOTAS_DE_CREDITO_FE(True, n1, n2) Then
        Print2_OPOS_Dev Space(2): Call TEXTFileZ("APPEND", Space(10), , , "N")
        Print2_OPOS_Dev Space(2)
        Print2_OPOS_Dev "NOTAS DE CREDITO    (" & n1 & ")    " & Format(n2, "STANDARD")
        
            Call TEXTFileZ("APPEND", "NOTAS DE CREDITO", n1, n2, "Y")
            Call TEXTFileZ("APPEND", Space(10), , , "N")
        Print2_OPOS_Dev Space(2)
    End If
    
    On Error Resume Next
    cTemp = RegRead("HKCU\Software\SoloSoftware\SoloMix\FE_Sales")
    cTemp2 = RegRead("HKCU\Software\SoloSoftware\SoloMix\FE_Counter")
    On Error GoTo 0

    If Format(CSng(cTemp), "STANDARD") <> Format(nSumVta, "STANDARD") Then
        Print2_OPOS_Dev Space(2)
        Print2_OPOS_Dev "|||||||||||||||||||||||||||||||||||||||"
        Print2_OPOS_Dev "Venta FE " & Format(CSng(cTemp), "STANDARD") & " || Venta POS " & Format(nSumVta, "STANDARD")
            Call TEXTFileZ("APPEND", "||||||||||||||||||||||||||||||||||||||||||||||||||", , , "N")
            Call TEXTFileZ("APPEND", "Venta FE: ", -1, CSng(cTemp), "Y")
            Call TEXTFileZ("APPEND", "Venta POS: ", -1, nSumVta, "Y")
            Call TEXTFileZ("APPEND", "Ajuste en FE: ", -1, CSng(cTemp) - nSumVta, "Y")
            Call TEXTFileZ("APPEND", "||||||||||||||||||||||||||||||||||||||||||||||||||", , , "N")
            Call TEXTFileZ("APPEND", Space(10), , , "N")
        'QUE DIGA AJUSTE EN FE // 26ENE2025
        Rem Print2_OPOS_Dev "Diferencia: " & Format(CSng(cTemp) - nSumVta, "STANDARD")
        Print2_OPOS_Dev "Ajuste en FE: " & Format(CSng(cTemp) - nSumVta, "STANDARD")
        Print2_OPOS_Dev "|||||||||||||||||||||||||||||||||||||||"
        Print2_OPOS_Dev Space(2)
        EscribeLog "Ajuste en FE (" & Format(CSng(cTemp) - nSumVta, "STANDARD") & ")"
    End If
    
    On Error Resume Next
    RegWrite "HKCU\Software\SoloSoftware\SoloMix\FE_Sales", 0
    RegWrite "HKCU\Software\SoloSoftware\SoloMix\FE_Counter", 0
    On Error GoTo 0
    
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

Do Until rsAjustes.EOF
    MiLen1 = Len(rsAjustes!Z_COUNT)
    Milen2 = Len(Format(rsAjustes!Valor, "STANDARD"))
    If rsAjustes!CORREC = "EC" Then
        Print2_OPOS_Dev "CORRECCION " & Space(8 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!Valor, "STANDARD")
                Call TEXTFileZ("APPEND", "CORRECCION ", rsAjustes!Z_COUNT, rsAjustes!Valor * -1, "Y")
    ElseIf rsAjustes!CORREC = "VO" Then
        Print2_OPOS_Dev "ANULACION " & Space(8 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!Valor, "STANDARD")
                Call TEXTFileZ("APPEND", "ANULACION ", rsAjustes!Z_COUNT, rsAjustes!Valor, "Y")
    ElseIf rsAjustes!CORREC = "DC" Then
        'INFO: nTrans SI LAS CANTIDADES SON MUY GRANDES, ESTOY CAMBIANDO LA LONGITUD DEL CAMPO A 4
        '12ABR2013
        If MiLen1 > 4 Then MiLen1 = 4
        Print2_OPOS_Dev "DESCUENTO " & Format(rsAjustes!DESCUENTO, "0.00") & Space(4 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!Valor, "STANDARD")
                Call TEXTFileZ("APPEND", "DESCUENTO " & Format(rsAjustes!DESCUENTO, "0.00"), rsAjustes!Z_COUNT, rsAjustes!Valor, "Y")
    End If
    rsAjustes.MoveNext
Loop

'PREPARA INFO PARA DESCUENTO GLOBAL
rsAjustes.Close
sqltext = "SELECT a.cajero,a.tipo_pago,SUM(a.monto) AS Valor, " & " COUNT(a.tipo_pago) as Z_COUNT " & " FROM transac_pago as a " & " WHERE a.tipo_pago = 99 " & " GROUP BY a.cajero,a.TIPO_PAGO"

rsAjustes.Open sqltext, msConn, adOpenStatic, adLockOptimistic

If Not rsAjustes.EOF Then
    RSPAGOS.MoveFirst
    RSPAGOS.Find "CODIGO = " & rsAjustes!TIPO_PAGO
    If Not RSPAGOS.EOF Then
        MiLen1 = Len(rsAjustes!Z_COUNT)
        Milen2 = Len(Format(rsAjustes!Valor, "STANDARD"))
        Print2_OPOS_Dev FormatTexto(RSPAGOS!DESCRIP, 13) & Space(5 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!Valor, "STANDARD")
                Call TEXTFileZ("APPEND", RSPAGOS!DESCRIP, rsAjustes!Z_COUNT, rsAjustes!Valor * -1, "Y")
    End If
End If

ProgBar.value = 70

ssVtatot = rs00!VTA_TOT
ssHashTot = rs00!tot_hash

msConn.BeginTrans
'''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1, " & " VTA_TOT = VTA_TOT + " & nSumTotal & ", TOT_HASH = TOT_HASH + " & IIf(IsNull(rsHash!HASH_DIA), 0, rsHash!HASH_DIA)
'1DEC2003 . AGREGAR EL ITBMS AL TOTAL HASH
msConn.Execute "UPDATE ORGANIZACION SET VTA_TOT = VTA_TOT + " & nSumTotal & ", TOT_HASH = TOT_HASH + " & IIf(IsNull(rsHash!HASH_DIA), 0, rsHash!HASH_DIA + nITBMSToHash)
msConn.Execute "UPDATE CAJEROS SET Z_C = Z_C + 1, X_C = 0 " & " WHERE NUMERO = 999"
msConn.CommitTrans

rsVta_Z.Close
rsPgo_Z.Close
rsTran.Close
rsAjustes.Close

rs00.Requery    'DESPUES DEL COMMIT, LOS DATOS DEBEN DE ESTAR EN EL SERVIDOR

Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev "VENTAS ACUMULADAS"
        Call TEXTFileZ("APPEND", Space(10), , , "N")
        Call TEXTFileZ("APPEND", "VENTAS ACUMULADAS", , , "N")
MiLen1 = Len(Format(nSumTotal, "CURRENCY"))
Print2_OPOS_Dev "HOY        : " & Format(nSumTotal, "CURRENCY")
        Call TEXTFileZ("APPEND", "HOY        : ", -1, Format((nSumTotal), "CURRENCY"), "Y")
Print2_OPOS_Dev "TOTAL ANT. : " & Format(ssVtatot, "CURRENCY")
        Call TEXTFileZ("APPEND", "TOTAL ANT. :", -1, Format(ssVtatot, "CURRENCY"), "Y")
Print2_OPOS_Dev "TOTAL NUEVO: " & Format(rs00!VTA_TOT, "CURRENCY")
        Call TEXTFileZ("APPEND", "TOTAL NUEVO: ", -1, Format(rs00!VTA_TOT, "CURRENCY"), "Y")
'*********************************************
On Error Resume Next
Print2_OPOS_Dev "   ITBMS  : " & Format(iISCTransaccion, "CURRENCY")
        Call TEXTFileZ("APPEND", "   ITBMS  : ", -1, Format(iISCTransaccion, "CURRENCY"), "Y")
iISCTransaccion = 0#
On Error GoTo 0
'*********************************************
Print2_OPOS_Dev Space(2)
'Printer.Print "HASH ANT.  : " & Format(ssHashTot, "CURRENCY")
'--- SEGUN HACIENDA Y TESORO -''Printer.Print "HASH NUEVO : " & Format(rs00!tot_hash, "CURRENCY")
Print2_OPOS_Dev "GRAN TOTAL  : " & Format(rs00!tot_hash, "CURRENCY")
        Call TEXTFileZ("APPEND", "GRAN TOTAL : ", -1, Format(rs00!tot_hash, "CURRENCY"), "Y")
'Print2_OPOS_Dev "GRAN TOTAL  : " & Format(4.73, "CURRENCY"), "GRAN TOTAL  : " & Format(rs00!tot_hash, "CURRENCY")

'rs00!tot_hash = rs00!tot_hash + nITBMSToHash

For i = 1 To 10
    Print2_OPOS_Dev Space(2)
Next

Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "END")

Sleep MUNBYNSeconds
Sys_Pos.Coptr1.CutPaper 100

Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "BEGIN")

'------------- DEPARTAMENTAL --------------

ProgBar.value = 80
'rsTran.Open "SELECT DEPTO,COUNT(DEPTO) AS X_COUNT,SUM(PRECIO) AS VALOR "
rsTran.Open "SELECT DEPTO,SUM(CANT) AS X_COUNT,SUM(PRECIO) AS VALOR " & " FROM TRANSAC " & " GROUP BY DEPTO " & " ORDER BY DEPTO", msConn, adOpenStatic, adLockOptimistic

'''- SEGUN HACIENDA Y TESORO -'''msConn.BeginTrans
'''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
'''- SEGUN HACIENDA Y TESORO -'''msConn.CommitTrans

Print2_OPOS_Dev Date & Space(2) & Time
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev "REPORTE DEPARTAMENTAL (Z)"
        Call TEXTFileZ("APPEND", "///////////////////////////////////////////////////", , , "N")
        Call TEXTFileZ("APPEND", "///////////////////////////////////////////////////", , , "N")
        Call TEXTFileZ("APPEND", "///////////////////////////////////////////////////", , , "N")
        Call TEXTFileZ("APPEND", "REPORTE DEPARTAMENTAL (Z)", , , "N")
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev rs00!DESCRIP
'INFO: AGO(2009)
Print2_OPOS_Dev rs00!RAZ_SOC
Print2_OPOS_Dev "RUC:" & rs00!ruc
Print2_OPOS_Dev "SERIAL:" & rs00!SERIAL
Print2_OPOS_Dev Space(2)
'''- SEGUN HACIENDA Y TESORO -'''Printer.Print "CONTADOR TRANS : " & (rs00!TRANS + 1)
Print2_OPOS_Dev "CONTADOR Z : " & (rs00!Z_CDEP + 1)
Print2_OPOS_Dev "DEPART. REPORTE/TERMINAL"
        Call TEXTFileZ("APPEND", "DEPART. REPORTE/TERMINAL", , , "N")
Print2_OPOS_Dev Space(2)

Dim nTotDepto As Double

nTotDepto = 0#

Do Until rsTran.EOF
    On Error Resume Next
    rsLocalDepto.MoveFirst
    On Error GoTo 0
    rsLocalDepto.Find "CODIGO = " & rsTran!depto
    If Not rsLocalDepto.EOF Then
        MiLen1 = Len(rsTran!X_COUNT)
        Milen2 = Len(Format(rsTran!Valor, "STANDARD"))
        'INFO: nTrans SI LAS CANTIDADES SON MUY GRANDES, ESTOY CAMBIANDO LA LONGITUD DEL CAMPO A 4
        '12ABR2013
        If MiLen1 > 4 Then MiLen1 = 4
        Print2_OPOS_Dev FormatTexto(rsLocalDepto!corto, 13) & Space(4 - MiLen1) & rsTran!X_COUNT & Space(9 - Milen2) & Format(rsTran!Valor, "STANDARD")

                Call TEXTFileZ("APPEND", rsLocalDepto!corto, rsTran!X_COUNT, Format(rsTran!Valor, "STANDARD"), "Y")
    Else
        MsgBox "Error en Depto. ", vbCritical, BoxTit
    End If
    nTotDepto = nTotDepto + rsTran!Valor
    rsTran.MoveNext
Loop

'/***********************************/
'INFO: 23OCT2018. APPETITO24.
Call InfoDomicilio(nTotDepto, "Z")
'/***********************************/

Print2_OPOS_Dev Space(2)
MiLen1 = Len(Format(nTotDepto, "CURRENCY"))
If MiLen1 > 11 Then MiLen1 = 11
Print2_OPOS_Dev "TOTAL DEPTOS :" & Space(11 - MiLen1) & Format(nTotDepto, "CURRENCY")
        Call TEXTFileZ("APPEND", "TOTAL DEPTOS :", -1, Format(nTotDepto, "CURRENCY"), "Y")
Print2_OPOS_Dev Space(2)

'INFO: IMPRIME LAS VENTAS DE DOMICILIO
'08MAR2017
Call GetVentasDomicilio(nTotDepto, "Z")


For i = 1 To 10
    Print2_OPOS_Dev Space(2)
Next

Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "END")

Sleep MUNBYNSeconds
Sys_Pos.Coptr1.CutPaper 100

rsTran.Close
ProgBar.value = 90

msConn.BeginTrans
'Tambien actualiza el conteo de Z's y X's
'''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1,X_CDEP = 0,Z_CDEP = Z_CDEP + 1"
msConn.Execute "UPDATE ORGANIZACION SET X_CDEP = 0,Z_CDEP = Z_CDEP + 1"
msConn.CommitTrans



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO: REPORTE DE GRUPOS ESPECIALES
'INFO: PEDIDO POR JAP JAP
'INFO: '17AGO2015
'INFO: UPDATE 2FEB2017 - 08MAR2017
'INFO: SE ADICIONAN AQUI LOS REPORTES DE DOMICILIO
'INFO: '20JUL2017
'SE SOLICITA ARCHIVO CON FORMATO (CSV) PARA QUE SE IMPORTE A PEACHTREE
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

If UCase(GetFromINI("Facturacion", "ReporteEspecial", App.Path & "\soloini.ini")) = "PEREZA" Then
    'Call ReporteEspecial
    Call ReporteEspecial_v2
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'--------------- SUPER GRUPOS ------------------------

Dim ccc As String

Set rsSuperGrp = New Recordset

ccc = "SELECT A.GRUPO,A.DESCRIP,SUM(C.PRECIO) AS VENTAS"
ccc = ccc & " FROM SUPER_GRP AS A,SUPER_DET AS B, TRANSAC AS C"
ccc = ccc & " Where A.GRUPO = B.GRUPO And B.DEPTO = C.DEPTO "
ccc = ccc & " GROUP BY A.GRUPO,A.DESCRIP"
ccc = ccc & " ORDER BY A.DESCRIP"

rsSuperGrp.Open ccc, msConn, adOpenStatic, adLockOptimistic

'''- SEGUN HACIENDA Y TESORO -'''msConn.BeginTrans
'''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
'''- SEGUN HACIENDA Y TESORO -'''msConn.CommitTrans

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "BEGIN")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Print2_OPOS_Dev Date & Space(2) & Time
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev "REPORTE DE GRUPOS (Z)"
        Call TEXTFileZ("APPEND", "///////////////////////////////////////////////////", , , "N")
        Call TEXTFileZ("APPEND", "///////////////////////////////////////////////////", , , "N")
        Call TEXTFileZ("APPEND", "///////////////////////////////////////////////////", , , "N")
        Call TEXTFileZ("APPEND", "REPORTE DE GRUPOS (Z)", , , "N")
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev rs00!DESCRIP
'INFO: AGO 2009
Print2_OPOS_Dev rs00!RAZ_SOC
Print2_OPOS_Dev "RUC:" & rs00!ruc
Print2_OPOS_Dev "SERIAL:" & rs00!SERIAL
Print2_OPOS_Dev Space(2)
'''- SEGUN HACIENDA Y TESORO -'''Printer.Print "CONTADOR TRANS : " & (rs00!TRANS + 1)
Print2_OPOS_Dev "CONTADOR Z : " & (rs00!Z_CDEP + 1)
Print2_OPOS_Dev "DEPART. GRUPO/TERMINAL"
        Call TEXTFileZ("APPEND", "DEPART. GRUPO/TERMINAL", , , "N")
Print2_OPOS_Dev Space(2)

nTotDepto = 0#

Do Until rsSuperGrp.EOF
    Milen2 = Len(Format(rsSuperGrp!VENTAS, "STANDARD"))
    Print2_OPOS_Dev FormatTexto(rsSuperGrp!DESCRIP, 13) & Space(13 - Milen2) & Format(rsSuperGrp!VENTAS, "STANDARD")
            Call TEXTFileZ("APPEND", rsSuperGrp!DESCRIP, -1, Format(rsSuperGrp!VENTAS, "STANDARD"), "Y")
    nTotDepto = nTotDepto + rsSuperGrp!VENTAS
    rsSuperGrp.MoveNext
Loop

Print2_OPOS_Dev Space(2)
MiLen1 = Len(Format(nTotDepto, "CURRENCY"))
'INFO: nTrans SI LAS CANTIDADES SON MUY GRANDES, ESTOY CAMBIANDO LA LONGITUD DEL CAMPO A 13
'12ABR2013
If MiLen1 > 4 Then MiLen1 = 4
Print2_OPOS_Dev "TOTAL GRUPOS:" & Space(13 - MiLen1) & Format(nTotDepto, "CURRENCY")
        Call TEXTFileZ("APPEND", "TOTAL GRUPOS:", -1, Format(nTotDepto, "CURRENCY"), "Y")
Print2_OPOS_Dev Space(2)

For i = 1 To 10
    Print2_OPOS_Dev Space(2)
Next

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO: 28NOV2014
Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "END")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sleep MUNBYNSeconds
Sys_Pos.Coptr1.CutPaper 100


'''- SEGUN HACIENDA Y TESORO -'''msConn.BeginTrans
'''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
'''- SEGUN HACIENDA Y TESORO -'''msConn.CommitTrans

'--------------- FIN DE SUPER GRUPOS ------------------------

        Call TEXTFileZ("APPEND", "///////////////////////////////////////////////////", , , "N")
        Call TEXTFileZ("CLOSE", "NADA", , , "N")

'----------------------------------------
rsTran.Open "SELECT * FROM TRANSAC", msConn, adOpenStatic, adLockOptimistic

On Error GoTo ErrorZZZ:

msConn.BeginTrans

Do Until rsTran.EOF
    CadenaSql = "INSERT INTO HIST_TR "
    CadenaSql = CadenaSql & "(NUM_TRANS,CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE,"
    CadenaSql = CadenaSql & "PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,FECHA_TRANS,HORA_TRANS,"
    CadenaSql = CadenaSql & "Z_COUNTER,CON_TAX) VALUES ("
    CadenaSql = CadenaSql & "" & rsTran!NUM_TRANS & "," & rsTran!CAJA & "," & rsTran!CAJERO & "," & rsTran!MESA & ","
    CadenaSql = CadenaSql & rsTran!MESERO & "," & GetBoolean(rsTran!VALID) & "," & rsTran!LIN & "," & "'"
    CadenaSql = CadenaSql & rsTran!DESCRIP & "'" & "," & rsTran!CANT & "," & rsTran!depto & "," & rsTran!PLU & ","
    CadenaSql = CadenaSql & rsTran!envase & "," & rsTran!precio_unit & "," & rsTran!precio & "," & "'"
    CadenaSql = CadenaSql & rsTran!FECHA & "'" & "," & "'" & rsTran!Hora & "'"
    CadenaSql = CadenaSql & ",'" & rsTran!Tipo & "'," & rsTran!DESCUENTO & ",'"
    CadenaSql = CadenaSql & rsTran!FECHA_TRANS & "','" & IIf(IsNull(rsTran!HORA_TRANS), "00:00", rsTran!HORA_TRANS)
    CadenaSql = CadenaSql & "'," & rs00!Z_CDEP & ","
    CadenaSql = CadenaSql & rsTran!CON_TAX & ")"
    'Debug.Print CadenaSql
    msConn.Execute CadenaSql
    rsTran.MoveNext
Loop

msConn.Execute "DELETE FROM TRANSAC"
msConn.CommitTrans

'ACTUALIZA Z_COUNTER
'INFO: 8AGO2018. SISTEMA FAST AHORA SE USA CON EMPRESAS QUE COBRAN EL 7%
msConn.BeginTrans
'cSQL = "INSERT INTO Z_COUNTER (CONTADOR, FECHA, HORA, PRINT_OK, TOTAL_ANTERIOR, TOTAL_NUEVO, GRAN_TOTAL) "
cSQL = "INSERT INTO Z_COUNTER (CONTADOR, FECHA, HORA, PRINT_OK, "
cSQL = cSQL & " TOTAL_ANTERIOR, TOTAL_NUEVO, GRAN_TOTAL, ITBMS) "
cSQL = cSQL & " VALUES ('"
cSQL = cSQL & rs00!Z_CDEP & "','" & Format(Date, "YYYYMMDD") & "','"
'INFO: 22AGO2017
'VUELVE Y APARECE EL FANTASMA DE LA CONFIGURACION REGIONAL, SE PONE LA FUNCION GetBoolean
'cSQL = cSQL & Format(Time, "HHMM") & "'," & bPrintOk & "," & ssVtatot & ","
cSQL = cSQL & Format(Time, "HHMM") & "'," & GetBoolean(bPrintOk) & "," & ssVtatot & ","
cSQL = cSQL & rs00!VTA_TOT & "," & rs00!tot_hash & "," & nITBMSToHash & ")"

msConn.Execute cSQL
msConn.CommitTrans

'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~
'INFO: ACTUALIZA INFORMACION EN TABLA DE DATOS FISCALES
'21NOV2011
If cFullFiscal = "SI" Then
    If UpdateFISCALInfoOnZ_Counter Then
        'HAY INFORMACION NUEVA, ES UNA Z REAL
        EscribeLog "Ventas.~~~~||||~~~~ NUEVO REPORTE Z.FISCAL ~~~~||||~~~~ "
    Else
        'DO NOTHING.
        EscribeLog "Ventas.~~~~||||~~~~ REPETICION REPORTE Z.FISCAL ~~~~||||~~~~ "
    End If
End If
'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~


rsTran.Close

'----------------------------------------
rsTran.Open "SELECT * FROM TRANSAC_PAGO", msConn, adOpenStatic, adLockOptimistic

msConn.BeginTrans

Do Until rsTran.EOF
    '20030829
    sqltext = "INSERT INTO HIST_TR_PAGO (NUM_TRANS,TIPO_PAGO,CAJERO,LIN,MONTO,MESA,Z_COUNTER) "
    sqltext = sqltext & " VALUES (" & rsTran!NUM_TRANS & ","
    sqltext = sqltext & rsTran!TIPO_PAGO & ","
    sqltext = sqltext & rsTran!CAJERO & "," & rsTran!LIN & ","
    sqltext = sqltext & Format(rsTran!MONTO, "#0.00") & ","
    sqltext = sqltext & rsTran!MESA & "," & rs00!Z_CDEP & ")"
    msConn.Execute sqltext
    rsTran.MoveNext
Loop

msConn.Execute "DELETE FROM TRANSAC_PAGO"
msConn.CommitTrans
rsTran.Close

'-------------LIMPIA PROPINAS------------------------
rsTran.Open "SELECT * FROM TRANSAC_PROP", msConn, adOpenStatic, adLockOptimistic

msConn.BeginTrans
Do Until rsTran.EOF
    'sqltext = "INSERT INTO HIST_TR_PROP " & " (NUM_TRANS,MESERO,CAJERO,LIN,TIPO_PAGO,MONTO,MESA,Z_COUNTER) " & " VALUES (" & rsTran!NUM_TRANS & "," & rsTran!mesero & "," & rsTran!CAJERO & "," & rsTran!LIN & "," & rsTran!TIPO_PAGO & "," & Format(rsTran!MONTO, "#0.00") & "," & rsTran!mesa & "," & rs00!Z_CDEP & ")"

        sqltext = "INSERT INTO HIST_TR_PROP "
        sqltext = sqltext & " (NUM_TRANS,MESERO,CAJERO,LIN,TIPO_PAGO,MONTO,Z_COUNTER) "
        sqltext = sqltext & " VALUES ("
        sqltext = sqltext & rsTran!NUM_TRANS & "," & rsTran!MESERO & ","
        sqltext = sqltext & rsTran!CAJERO & ","
        sqltext = sqltext & rsTran!LIN & "," & rsTran!TIPO_PAGO & ","
        sqltext = sqltext & Format(rsTran!MONTO, "#0.00") & ","
        sqltext = sqltext & rs00!Z_CDEP & ")"

    msConn.Execute sqltext

    rsTran.MoveNext
Loop

msConn.Execute "DELETE FROM TRANSAC_PROP"
msConn.CommitTrans
rsTran.Close

'--------PASA CLIENTES AL HISTORICO-----------------
rsTran.Open "SELECT * FROM TRANSAC_CLI", msConn, adOpenStatic, adLockOptimistic
    
msConn.BeginTrans
Dim iStat As Integer
Do Until rsTran.EOF
    If rsTran!RECIBIDO > 0# Then
        If rsTran!RECIBIDO = rsTran!MONTO Then
            iStat = 2
        ElseIf rsTran!RECIBIDO < rsTran!MONTO Then
            iStat = 1
        End If
    Else
        iStat = 0
    End If
'    sqltext = "INSERT INTO HIST_TR_CLI " & " (CODIGO_TP,CODIGO_CLI,NUM_TRANS,MONTO,FECHA,STATUS,RECIBIDO,SALDO,TIPO_TRANS,COMMENT,Z_COUNTER) VALUES (" & rsTran!CODIGO_TP & "," & rsTran!CODIGO_CLI & "," & rsTran!NUM_TRANS & "," & Format(rsTran!MONTO - rsTran!RECIBIDO, "#0.00") & ",'FA','Factura de Ventas', " & rs00!Z_CDEP & ")"
    sqltext = "INSERT INTO HIST_TR_CLI "
    sqltext = sqltext & " (CODIGO_TP,CODIGO_CLI,NUM_TRANS,MONTO,FECHA,STATUS,RECIBIDO,SALDO,TIPO_TRANS,COMMENT) VALUES ("
    sqltext = sqltext & rsTran!CODIGO_TP & "," & rsTran!CODIGO_CLI & "," & rsTran!NUM_TRANS & ","
    sqltext = sqltext & Format(rsTran!MONTO, "#0.00") & "," & rsTran!FECHA & ","
    sqltext = sqltext & iStat & "," & Format(rsTran!RECIBIDO, "#0.00") & ","
    sqltext = sqltext & Format(rsTran!MONTO - rsTran!RECIBIDO, "#0.00")
    sqltext = sqltext & ",'FA','Factura de Ventas')"

    msConn.Execute sqltext

    rsTran.MoveNext
Loop

msConn.Execute "DELETE FROM TRANSAC_CLI"
msConn.CommitTrans
rsTran.Close
ProgBar.value = 100
'Close #1

On Error GoTo 0
If nErrInd = 0 Then
    ''''''''''''''''''''If ON_LINE = True Then BorraLocal
    'MsgBox "REPORTE EN 'Z' DE TERMINAL ESTA LISTO", vbInformation, BoxTit
    ShowMsg "REPORTE EN 'Z' DE TERMINAL ESTA LISTO" & vbCrLf & vbCrLf & "ESPERE A QUE TERMINE DE IMPRIMIR EL REPORTE PARA PODER SALIR DEL PROGRAMA", vbYellow, vbBlue
Else
  
    'MsgBox "HA OCURRIDO MAS DE UN ERROR EN EL REPORTE (Z). CONTACTE A SOLO SOFTWARE", vbCritical, BoxTit
    ShowMsg "HA OCURRIDO MAS DE UN ERROR EN EL REPORTE (Z). CONTACTE A SOLO POS MENU", vbRed, vbYellow
End If
ProgBar.value = 0

'REVISION DE PAGOS PARCIALES
rsPAGOParcial.Close
Set rsPAGOParcial = Nothing

Call ActualizaInvent ' ACTUALIZACION DE INVENTARIO

Exit Sub

AjustaMilen:
errCounter = errCounter + 1
Milen2 = 11
If errCounter < 4 Then
    Resume
Else
    '3021
    If Err.Number <> 3021 Then
        'EL ERROR 3021, ES EL EOF DEL ARCHIVO DE CAJEROS,
        'ESTO OCURRE CUANDO NO HAY TRANSACCIONES, SI HAY OTRO TIPO
        'DE ERROR EL SUPERVISOR LO VERA EN LA PANTALLA
        MsgBox "# " & Err.Number & " ----> " & Err.description, vbCritical, "ANOTE LOS DATOS EN PANTALLA: " & Err.Source
        'MsgBox "EXISTE UN PROBLEMA DE IMPRESION. UNA VEZ TERMINADO REVISE EL LISTADO Y VERIFIQUE LOS DATOS", vbCritical, "LA IMPRESION DEL REPORTE TIENE PROBLEMAS"
    End If
    Resume Next
End If

ErrorZZZ:
Dim ADOError As Error
For Each ADOError In msConn.Errors
    sError = sError & ADOError.Number & " - " & ADOError.description & vbCrLf
Next ADOError
ShowMsg "ERROR EN EL REPORTE Z.ANOTE EL NUMERO/DESCRIPCION Y CONTACTE A SOLO POS MENU", vbYellow, vbRed
'ShowMsg sError, vbYellow, vbRed
EscribeLog "ErrorZZZ: " & sError
nErrInd = 1
Resume Next
End Sub
Public Sub RepCajX()
Dim rsVta_X As Recordset
Dim rsPgo_X As Recordset
Dim rsTran As Recordset
Dim RSPAGOS As Recordset    'Pagos
Dim rsAjustes As Recordset
Dim rsCajeros As Recordset
Dim rsProp As Recordset
Dim sqltext As String
Dim nSumVta As Double
Dim nSumTotal As Double
'INFO: nTrans Cambiando de Integer a LONG
'12ABR2013
Dim nTrans As Long
Dim MiLen1 As Integer
Dim Milen2 As Integer
Dim errCounter As Integer
'AGREGANDO PAGO PARCIAL (25MAR2004)
Dim rsPAGOParcial As Recordset
Dim nMontoPagoParcial As Single
Dim nCounterPagoParcial As Integer

Rem NOTAS DE CREDITO DE LA FACTURACION ELECTRONICA
Dim n1 As Long
Dim n2 As Single

nSumVta = 0
nSumTotal = 0

On Error GoTo AjustaMilen:

Set rsVta_X = New Recordset
Set rsPgo_X = New Recordset
Set rsTran = New Recordset
Set RSPAGOS = New Recordset
Set rsAjustes = New Recordset
Set rsCajeros = New Recordset
Set rsProp = New Recordset

Set rsPAGOParcial = New Recordset
rsPAGOParcial.Open "SELECT MESA,CAJERO,MESERO,MONTO FROM TMP_PAR_PAGO ORDER BY CAJERO", msConn, adOpenStatic, adLockOptimistic

sqltext = "SELECT DISTINCT CAJERO FROM TRANSAC ORDER BY CAJERO"
rsCajeros.Open sqltext, msConn, adOpenStatic, adLockOptimistic

MesasAbiertas

If rsCajeros.EOF Then
    MsgBox "NO EXISTEN TRANSACCIONES, NO SE IMPRIMIRA REPORTE", vbInformation, BoxTit
    Exit Sub
End If

If rsCajeros.RecordCount = 0 Or rsCajeros!CAJERO = "" Then
    MsgBox "NO EXISTEN TRANSACCIONES, NO SE IMPRIMIRA REPORTE", vbInformation, BoxTit
    Exit Sub
End If

'Open "c:\mifile.txt" For Output As #1   ' Open file.

sqltext = "SELECT * FROM pagos WHERE CODIGO <> 999 ORDER BY CODIGO"
RSPAGOS.Open sqltext, msConn, adOpenStatic, adLockOptimistic

ProgBar.value = 10
Do Until rsCajeros.EOF

    rsTran.Open "SELECT distinct NUM_TRANS FROM TRANSAC WHERE CAJERO = " & rsCajeros!CAJERO, msConn, adOpenStatic, adLockOptimistic

    nTrans = rsTran.RecordCount

    '------------------------------------------------------------------------
    'INFO: 17AGO2014. CORRIGIENDO ERROR EN LA IMPRESION DEL IMPUESTO
    sqltext = "SELECT b.nombre,b.apellido,a.cajero,b.x_c, sum(a.precio) as Ventas, "
    sqltext = sqltext & " sum(a.precio * (A.CON_TAX/100)) AS TAX_CAJERO"
    '------------------------------------------------------------------------
    sqltext = sqltext & " FROM TRANSAC as a, cajeros as b "
    sqltext = sqltext & " WHERE a.cajero = " & rsCajeros!CAJERO
    sqltext = sqltext & " AND b.numero = " & rsCajeros!CAJERO
    sqltext = sqltext & " GROUP BY a.cajero,b.nombre,b.apellido,b.x_c "
            
    'VALOR EN VENTAS DEL CAJERO CON rsVta_X
    rsVta_X.Open sqltext, msConn, adOpenStatic, adLockOptimistic
    
    If rsVta_X.RecordCount = 0 Then
        'MsgBox "EL CAJERO NO TIENE VENTAS, NO SE IMPRIMIRA REPORTE EN X"
        rsCajeros.MoveNext
    End If
    ProgBar.value = 20
    'TODOS LOS PAGOS RECIBIDOS
    ' SACANDO LOS DESCUENTOS GLOBALES PARA PONERLOS
    ' EN LOS AJUSTES
    'INFO: 26MAR2019. CARGOS DELIVERY
    sqltext = "SELECT a.cajero,a.tipo_pago,SUM(a.monto) AS Valor, "
    sqltext = sqltext & " COUNT(a.tipo_pago) as X_COUNT, MAX(B.TIPO) AS T_PAGO "
    sqltext = sqltext & " FROM transac_pago as a, PAGOS AS B "
    sqltext = sqltext & " WHERE a.cajero = " & rsCajeros!CAJERO
    sqltext = sqltext & " AND A.TIPO_PAGO = B.CODIGO "
    sqltext = sqltext & " AND a.tipo_pago <> 99 "
    sqltext = sqltext & " AND B.TIPO <> 'CA' "
    sqltext = sqltext & " GROUP BY a.cajero,a.TIPO_PAGO"

    rsPgo_X.Open sqltext, msConn, adOpenStatic, adLockOptimistic

    'sqltext = "SELECT MID(a.TIPO,1,2) AS CORREC,DESCUENTO, "
        '" COUNT(a.lin) as X_COUNT, SUM(a.precio) as valor "
    ' 2 de Nov 1999
    sqltext = "SELECT MID(a.TIPO,1,2) AS CORREC,DESCUENTO, "
    sqltext = sqltext & " COUNT(a.lin) as X_COUNT, SUM(a.PRECIO) as valor "
    sqltext = sqltext & " FROM TRANSAC as a WHERE A.CAJERO = " & rsCajeros!CAJERO
    sqltext = sqltext & " AND MID(A.TIPO,1,1) <> ' ' "
    sqltext = sqltext & " GROUP BY MID(a.TIPO,1,2),DESCUENTO"

    rsAjustes.Open sqltext, msConn, adOpenStatic, adLockOptimistic
    
    Print2_OPOS_Dev Date & Space(2) & Time, Date & Space(2) & Time
    Print2_OPOS_Dev Space(2)
    Print2_OPOS_Dev "REPORTE DE CAJEROS (X)", "REPORTE DE CAJEROS (X)"
    Print2_OPOS_Dev Space(2)
    Print2_OPOS_Dev rs00!DESCRIP, rs00!DESCRIP
    Print2_OPOS_Dev rs00!RAZ_SOC, rs00!RAZ_SOC
    Print2_OPOS_Dev "RUC:" & rs00!ruc, "RUC:" & rs00!ruc
    Print2_OPOS_Dev "SERIAL:" & rs00!SERIAL, "SERIAL:" & rs00!SERIAL
    Print2_OPOS_Dev Space(2)
    Print2_OPOS_Dev "CONTADOR TRANS : " & (nNUMTrans), "CONTADOR TRANS : " & (nNUMTrans)
    Print2_OPOS_Dev "CONTADOR X : " & (rsVta_X!x_c + 1), "CONTADOR X : " & (rsVta_X!x_c + 1)
    
    Print2_OPOS_Dev "CAJERO : " & rsVta_X!nombre & ", " & Mid(rsVta_X!APELLIDO, 1, 1) & ".", "CAJERO : " & rsVta_X!nombre & ", " & Mid(rsVta_X!APELLIDO, 1, 1) & "."
    Print2_OPOS_Dev Space(2)
    
    Do Until rsVta_X.EOF
        MiLen1 = Len(nTrans)
        Milen2 = Len(Format(rsVta_X!VENTAS, "STANDARD"))
        Print2_OPOS_Dev "VENTA DEL DIA:" & Space(4 - MiLen1) & nTrans & Space(11 - Milen2) & Format(rsVta_X!VENTAS, "STANDARD"), "VENTA DEL DIA:" & Space(4 - MiLen1) & nTrans & Space(11 - Milen2) & Format(rsVta_X!VENTAS, "STANDARD")
        nSumTotal = nSumTotal + rsVta_X!VENTAS
        rsVta_X.MoveNext
    Loop
    
    ProgBar.value = 30
    Print2_OPOS_Dev Space(2)
    '''''''Printer.FontUnderline = True
    'Printer.Print "DESGLOSE DE INGRESOS"
    Print2_OPOS_Dev "TOTALES DE CAJA", "TOTALES DE CAJA"
    ''''''''Printer.FontUnderline = False
    Print2_OPOS_Dev Space(2)
    
    'AGREGANDO DETALLE DE PAGO PARCIAL
    'GRACIAS A BAR GRILL 50 (25MAR2004)
    nMontoPagoParcial = 0#
    nCounterPagoParcial = 0
    nMontoPagoParcial = GetPagosParcialesFromCAJERO(rsCajeros!CAJERO, nCounterPagoParcial, rsPAGOParcial)
    If nMontoPagoParcial > 0 Then
        MiLen1 = Len(nCounterPagoParcial)
        Milen2 = Len(Format(nMontoPagoParcial, "STANDARD"))
        'Coptr1.PrintTwoNormal FptrSJournalReceipt , "PAGO PARCIAL " & Space(17 - MiLen1) & Format(nMontoPagoParcial, "STANDARD"), "PAGO PARCIAL " & Space(17 - MiLen1) & Format(nMontoPagoParcial, "STANDARD")
        Print2_OPOS_Dev "PAGO PARCIAL  " & Space(4 - MiLen1) & nCounterPagoParcial & Space(13 - Milen2) & Format(nMontoPagoParcial, "STANDARD"), "PAGO PARCIAL " & Space(4 - MiLen1) & nCounterPagoParcial & Space(13 - Milen2) & Format(nMontoPagoParcial, "STANDARD")
        nSumVta = nSumVta + nMontoPagoParcial
    End If
    
    Do Until rsPgo_X.EOF
        RSPAGOS.MoveFirst
        RSPAGOS.Find "CODIGO = " & rsPgo_X!TIPO_PAGO
        If Not RSPAGOS.EOF Then
            MiLen1 = Len(rsPgo_X!X_COUNT)
            Milen2 = Len(Format(rsPgo_X!Valor, "STANDARD"))
            'INFO: SI LOS MONTOS SON MUY GRANDES EL PROGRAMA REVIENTA AL CALCULAR EL LARGO DEL CAMPO
            '12ABR2013
            If MiLen1 > 4 Then MiLen1 = 4
            Print2_OPOS_Dev FormatTexto(RSPAGOS!DESCRIP, 13) & Space(4 - MiLen1) & rsPgo_X!X_COUNT & Space(13 - Milen2) & Format(rsPgo_X!Valor, "STANDARD"), FormatTexto(RSPAGOS!DESCRIP, 13) & Space(4 - MiLen1) & rsPgo_X!X_COUNT & Space(13 - Milen2) & Format(rsPgo_X!Valor, "STANDARD")
        Else
            MiLen1 = 1
            Milen2 = Len(Format(0#, "STANDARD"))
            Print2_OPOS_Dev "OTRO PAGO    " & Space(4 - MiLen1) & 0 & Space(13 - MiLen1) & Format(0#, "STANDARD"), "OTRO PAGO    " & Space(4 - MiLen1) & 0 & Space(13 - MiLen1) & Format(0#, "STANDARD")
            'Printer.Print rsPagos!descrip & Chr(9) & 0 & Chr(9) & Format(0#, "STANDARD")
        End If
        nSumVta = nSumVta + rsPgo_X!Valor
        rsPgo_X.MoveNext
    Loop
    Print2_OPOS_Dev Space(2)
    MiLen1 = Len(Format(nSumVta, "CURRENCY"))
    Print2_OPOS_Dev "SUBTOTAL: " & Space(20 - MiLen1) & Format(nSumVta, "CURRENCY"), "SUBTOTAL: " & Space(20 - MiLen1) & Format(nSumVta, "CURRENCY")
    Print2_OPOS_Dev "------------------------------", "------------------------------"
    
'''    '30 NOV 2003 - AGREGANDO ITBMS AL REPORTE DE CAJEROS (X)
    On Error Resume Next
    rsVta_X.MoveFirst
    
    'INFO: JULIO 2010 NO SE PUEDE PONER EL 5%, EL ITBMS ES CALCULADO x PRODUCTO
    'Print2_OPOS_Dev "ITBMS (5%):" & Space(18 - MiLen1) & Format(nSumVta - (rsVta_X!VENTAS + nMontoPagoParcial), "CURRENCY"), "ITBMS (5%):" & Space(18 - MiLen1) & Format(nSumVta - (rsVta_X!VENTAS + nMontoPagoParcial), "CURRENCY")
    'Print2_OPOS_Dev "ITBMS         :" & Space(18 - MiLen1) & Format(nSumVta - (rsVta_X!VENTAS + nMontoPagoParcial), "CURRENCY"), "ITBMS         :" & Space(18 - MiLen1) & Format(nSumVta - (rsVta_X!VENTAS + nMontoPagoParcial), "CURRENCY")
    '------------------------------------------------------------------------
    'INFO: 17AGO2014. CORRIGIENDO ERROR EN LA IMPRESION DEL IMPUESTO
    Print2_OPOS_Dev "ITBMS         :" & Space(18 - MiLen1) & Format(rsVta_X!TAX_CAJERO, "CURRENCY"), "ITBMS         :" & Space(18 - MiLen1) & Format(rsVta_X!TAX_CAJERO, "CURRENCY")
    '------------------------------------------------------------------------
    'Print2_OPOS_Dev "ITBMS (5%):" & Space(18 - MiLen1) & Format(nSumVta - rsVta_X!VENTAS, "CURRENCY"), "ITBMS (5%):" & Space(18 - MiLen1) & Format(nSumVta - rsVta_X!VENTAS, "CURRENCY")
    Print2_OPOS_Dev Space(2)
    On Error GoTo 0

    sqltext = "SELECT TIPO_PAGO,COUNT(TIPO_PAGO) AS X_COUNT, "
    sqltext = sqltext & " SUM(MONTO) AS VALOR "
    sqltext = sqltext & " FROM TRANSAC_PROP "
    sqltext = sqltext & " WHERE CAJERO = " & rsCajeros!CAJERO
    sqltext = sqltext & " GROUP BY TIPO_PAGO "

    rsProp.Open sqltext, msConn, adOpenStatic, adLockOptimistic
    
    Do Until rsProp.EOF
        RSPAGOS.MoveFirst
        RSPAGOS.Find "CODIGO = " & rsProp!TIPO_PAGO
        If Not RSPAGOS.EOF Then
            MiLen1 = Len(rsProp!X_COUNT)
            Milen2 = Len(Format(rsProp!Valor, "STANDARD"))
            'INFO: 26MAR2019. CARGOS DELIVERY
            'Print2_OPOS_Dev "Propina " & FormatTexto(RSPAGOS!DESCRIP, 5) & Space(4 - MiLen1) & rsProp!X_COUNT & Space(13 - Milen2) & Format(rsProp!Valor, "STANDARD"), "Propina " & FormatTexto(RSPAGOS!DESCRIP, 5) & Space(4 - MiLen1) & rsProp!X_COUNT & Space(13 - Milen2) & Format(rsProp!Valor, "STANDARD")
            Print2_OPOS_Dev FormatTexto(RSPAGOS!DESCRIP, 15) & Space(4 - MiLen1) & rsProp!X_COUNT & Space(13 - Milen2) & Format(rsProp!Valor, "STANDARD"), FormatTexto(RSPAGOS!DESCRIP, 15) & Space(4 - MiLen1) & rsProp!X_COUNT & Space(13 - Milen2) & Format(rsProp!Valor, "STANDARD")
        End If
        rsProp.MoveNext
    Loop
    rsProp.Close
    ProgBar.value = 40
    Print2_OPOS_Dev Space(2)
    ''''''Printer.FontUnderline = True
    Print2_OPOS_Dev "AJUSTES", "AJUSTES"
    ''''''Printer.FontUnderline = False
    Print2_OPOS_Dev Space(2)
    
    If CHECK_NOTAS_DE_CREDITO_FE(False, n1, n2) Then
        Print2_OPOS_Dev Space(2)
        Print2_OPOS_Dev Space(2)
        Print2_OPOS_Dev "NOTAS DE CREDITO    (" & n1 & ")    " & Format(n2, "STANDARD")
        Print2_OPOS_Dev Space(2)
        Print2_OPOS_Dev Space(2)
    End If
    
    Do Until rsAjustes.EOF
        MiLen1 = Len(rsAjustes!X_COUNT)
        Milen2 = Len(Format(rsAjustes!Valor, "STANDARD"))
        If rsAjustes!CORREC = "EC" Then
            Print2_OPOS_Dev "CORRECCION " & Space(8 - MiLen1) & rsAjustes!X_COUNT & Space(12 - Milen2) & Format(rsAjustes!Valor, "STANDARD"), "CORRECCION " & Space(8 - MiLen1) & rsAjustes!X_COUNT & Space(12 - Milen2) & Format(rsAjustes!Valor, "STANDARD")
        ElseIf rsAjustes!CORREC = "VO" Then
            Print2_OPOS_Dev "ANULACION " & Space(8 - MiLen1) & rsAjustes!X_COUNT & Space(12 - Milen2) & Format(rsAjustes!Valor, "STANDARD"), "ANULACION " & Space(8 - MiLen1) & rsAjustes!X_COUNT & Space(12 - Milen2) & Format(rsAjustes!Valor, "STANDARD")
        ElseIf rsAjustes!CORREC = "DC" Then
            Print2_OPOS_Dev "DESCUENTO " & Format(rsAjustes!DESCUENTO, "0.00") & Space(4 - MiLen1) & rsAjustes!X_COUNT & Space(12 - Milen2) & Format(rsAjustes!Valor, "STANDARD"), "DESCUENTO " & Format(rsAjustes!DESCUENTO, "0.00") & Space(4 - MiLen1) & rsAjustes!X_COUNT & Space(12 - Milen2) & Format(rsAjustes!Valor, "STANDARD")
        End If
        rsAjustes.MoveNext
    Loop

    'PREPARA INFO PARA DESCUENTO GLOBAL
    rsAjustes.Close
    sqltext = "SELECT a.cajero,a.tipo_pago,SUM(a.monto) AS Valor, " & " COUNT(a.tipo_pago) as X_COUNT " & " FROM transac_pago as a " & " WHERE a.cajero = " & rsCajeros!CAJERO & " AND a.tipo_pago = 99 " & " GROUP BY a.cajero,a.TIPO_PAGO"
    
    rsAjustes.Open sqltext, msConn, adOpenStatic, adLockOptimistic
    
    If Not rsAjustes.EOF Then
        RSPAGOS.MoveFirst
        RSPAGOS.Find "CODIGO = " & rsAjustes!TIPO_PAGO
        If Not RSPAGOS.EOF Then
            MiLen1 = Len(rsAjustes!X_COUNT)
            Milen2 = Len(Format(rsAjustes!Valor, "STANDARD"))
            Print2_OPOS_Dev FormatTexto(RSPAGOS!DESCRIP, 13) & Space(5 - MiLen1) & rsAjustes!X_COUNT & Space(12 - Milen2) & Format(rsAjustes!Valor, "STANDARD"), FormatTexto(RSPAGOS!DESCRIP, 13) & Space(5 - MiLen1) & rsAjustes!X_COUNT & Space(12 - Milen2) & Format(rsAjustes!Valor, "STANDARD")
        End If
    End If
    ProgBar.value = 50
    msConn.BeginTrans
    msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
    msConn.Execute "UPDATE CAJEROS SET X_C = X_C + 1 " & " WHERE NUMERO = " & rsCajeros!CAJERO
    msConn.CommitTrans
    
    rsCajeros.MoveNext
    If Not rsCajeros.EOF = True Then
        For i = 1 To 10
            Print2_OPOS_Dev Space(2)
        Next
        Sleep MUNBYNSeconds
        Sys_Pos.Coptr1.CutPaper 100
    End If
    
    rsVta_X.Close
    rsPgo_X.Close
    rsTran.Close
    rsAjustes.Close

Loop
ProgBar.value = 60
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev Space(2)
MiLen1 = Len(Format(nSumTotal, "CURRENCY"))
Print2_OPOS_Dev "TOTAL: " & Space(23 - MiLen1) & Format(nSumTotal, "CURRENCY"), "TOTAL: " & Space(23 - MiLen1) & Format(nSumTotal, "CURRENCY")
For i = 1 To 10
    Print2_OPOS_Dev Space(2)
Next
Sleep MUNBYNSeconds
Sys_Pos.Coptr1.CutPaper 100
ProgBar.value = 0
rsPAGOParcial.Close
Set rsPAGOParcial = Nothing
'Close #1
On Error GoTo 0
Exit Sub

AjustaMilen:
errCounter = errCounter + 1
Milen2 = 11
If errCounter < 4 Then
    Resume
Else
    'MsgBox "Error. Reporte X. EXISTE UN PROBLEMA DE IMPRESION. UNA VEZ TERMINADO REVISE EL LISTADO Y VERIFIQUE LOS DATOS" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical, "LA IMPRESION DEL REPORTE TIENE PROBLEMAS"
    ShowMsg "Error. Reporte X. EXISTE UN PROBLEMA DE IMPRESION. UNA VEZ TERMINADO REVISE EL LISTADO Y VERIFIQUE LOS DATOS", vbYellow, vbRed
    Resume Next
End If
End Sub

'''''Private Sub oActualizaInvent()
''''''ACTUALIZACION DE INVENTARIO
'''''Dim iIntento As Integer
'''''iIntento = 0
'''''
'''''On Error GoTo AdmError:
'''''Do Until rsPluInvent.EOF
'''''    msConn.BeginTrans
'''''    msConn.Execute "UPDATE INVENT SET EXIST2 = EXIST2 - " & rsPluInvent!CONSUMIDO & " WHERE ID = " & rsPluInvent!ID_PROD_INV
'''''    msConn.CommitTrans
'''''    rsPluInvent.MoveNext
'''''Loop
'''''rsPluInvent.Close
'''''Set rsPluInvent = Nothing
'''''On Error GoTo 0
'''''Exit Sub
'''''
'''''AdmError:
'''''If iIntento < 4 Then
'''''    MsgBox Err.Description, vbCritical, "Número del Error: " & Err.Number
'''''    iIntento = iIntento + 1
'''''    Resume
'''''Else
'''''    MsgBox "ERROR GRAVE EN ACTUALIZACION DE INVENTARIO : " & Err.Description & vbCrLf & '''''            "LUGAR DEL ERROR : " & Err.Source, vbCritical, "Número del Error: " & Err.Number
'''''    Resume Next
'''''End If
'''''End Sub

Private Sub ActualizaInvent()
'ACTUALIZACION DE INVENTARIO
'INFO (AGO2009) ACTUALIZA RECETAS
Dim iIntento As Integer
Dim cSQL2 As String
iIntento = 0

On Error GoTo AdmError:
On Error Resume Next
rsLOCALPLUInvent.MoveFirst
On Error GoTo 0
Do Until rsLOCALPLUInvent.EOF
    
    cSQL2 = "UPDATE INVENT SET EXIST2 = EXIST2 - " & rsLOCALPLUInvent!CONSUMIDO
    cSQL2 = cSQL2 & " WHERE ID = " & rsLOCALPLUInvent!ID_PROD_INV
    
    msConn.BeginTrans
    msConn.Execute cSQL2
    msConn.CommitTrans
    
    rsLOCALPLUInvent.MoveNext
Loop
rsLOCALPLUInvent.Close
Set rsLOCALPLUInvent = Nothing
On Error GoTo 0

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'ACTUALIZACION DE RECETAS
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
iIntento = 0

On Error GoTo AdmError:
On Error Resume Next
rsLOCALPLURecetas.MoveFirst
On Error GoTo 0
Do Until rsLOCALPLURecetas.EOF
    
    cSQL2 = "UPDATE RECETAS SET EXIST2 = EXIST2 - " & rsLOCALPLURecetas!CONSUMIDO
    cSQL2 = cSQL2 & " WHERE ID = " & rsLOCALPLURecetas!ID_RECETA
    
    msConn.BeginTrans
    msConn.Execute cSQL2
    msConn.CommitTrans
    
    rsLOCALPLURecetas.MoveNext
Loop
rsLOCALPLURecetas.Close
Set rsLOCALPLURecetas = Nothing
On Error GoTo 0

Exit Sub

AdmError:
If iIntento < 4 Then
    EscribeLog "Error en ActualizaInvent(). " & Err.Number & " - " & Err.description
    EscribeLog "COMANDO: " & cSQL2
    MsgBox Err.Number & " - " & Err.description, vbCritical, "Error en ActualizaInvent()"
    iIntento = iIntento + 1
    Resume
Else
    EscribeLog "ERROR GRAVE EN ACTUALIZACION DE INVENTARIO ActualizaInvent(): " & Err.description
    EscribeLog "COMANDO: " & cSQL2
    MsgBox "ERROR GRAVE EN ACTUALIZACION DE INVENTARIO (ActualizaInvent()): " & Err.description & vbCrLf & "LUGAR DEL ERROR : " & Err.Source, vbCritical, "Número del Error: " & Err.Number
    Resume Next
End If
End Sub
Private Sub cmdOpenGaveta_Click()
'abrir faveta
''Sys_Pos.Cocash1.ClaimDevice 5000
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
If cFullFiscal = "SI" Then
    Call FISCAL_.SS_CashDrawerOpen
Else
    'Sys_Pos.Cocash1.ReleaseDevice
    'Sys_Pos.Cocash1.ClaimDevice 5000   'INFO: REMOVIDO 26/AGO/2023
    'rc = Sys_Pos.CoCash1.OpenDrawer        'INFO: 14FEB2017. USANDO DIRECTIO QUE NO ESPERA RESPUESTA DE LA GAVETA
    'Call Sys_Pos.Cocash1.DirectIO(DRW_DI_OPEN_DRAWER, 0, "")
    Call OpenGavetaDinero
End If
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~

End Sub

Private Sub cmdRX_Click(Index As Integer)
Dim cReportesExtra As String

'REVISA LA IMPRESORA ANTES DE IMPRIMIR.
'17AGO2014
If OPOS_Eval("Coptr1.RecEmpty") Then
    ShowMsg "NO HAY PAPEL (RECIBO/Izquierda) EN LA IMPRESORA, CIERRE EL PROGRAMA, CAMBIELO AHORA e INTENTE IMPRIMIR NUEVAMENTE", vbRed, vbYellow
    'ReportesEscribeLog ("Reporte X - NO HAY PAPEL (RECIBO/Izquierda) EN LA IMPRESORA, CAMBIELO AHORA e INTENTE IMPRIMIR NUEVAMENTE")
    EscribeLog ("Reporte X - NO HAY PAPEL (RECIBO/Izquierda) EN LA IMPRESORA, CAMBIELO AHORA e INTENTE IMPRIMIR NUEVAMENTE")
    Me.MousePointer = vbDefault
'            Exit Sub
End If

If OPOS_Eval("Coptr1.RecNearEnd") Then
    ShowMsg "POR FAVOR DESPUES DE IMPRIMIR, REVISE EL PAPEL (RECIBO/Izquierda) EN LA IMPRESORA, PUEDE QUE SE ESTE ACABANDO", vbRed, vbYellow
    'ReportesEscribeLog ("Reporte X - POR FAVOR REVISE EL PAPEL (RECIBO/Izquierda) EN LA IMPRESORA, PUEDE QUE SE ESTE ACABANDO")
    EscribeLog ("Reporte X - POR FAVOR REVISE EL PAPEL (RECIBO/Izquierda) EN LA IMPRESORA, PUEDE QUE SE ESTE ACABANDO")
End If
    
Select Case Index
Case 0  'REPORTE DE CAJEROS X

    '~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~
    '08SEP2011
    '13DIC2012 = HABILITANDO LA PREGUNTA PARA CLIENTES QUE NO TIENEN LA FISCAL
    If cFullFiscal = "SI" Then
        On Error Resume Next
        Call FISCAL_.SS_FiscalReporteCierre("X")
        On Error GoTo 0
    End If
    '~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~
    cmdRX(Index).Enabled = False

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "BEGIN")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    'INFO: 29OCT2014
    'PARA LOS QUE DESEAN UN CONTEO DE PRODUCTOS EN EL REPORTE X
    cReportesExtra = UCase(GetFromINI("Facturacion", "ReportesExtra", App.Path & "\soloini.ini"))
    If cReportesExtra = "PEREZAX" Then
        If ReporteFacturas_y_Productos(cReportesExtra) Then
            Call PrintReportesExtra
        End If
    End If
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "END")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "BEGIN")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    RepCajX
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "END")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    cmdRX(Index).Enabled = True
Case 1
    cmdRX(Index).Enabled = False
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "BEGIN")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    RepMesX
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "END")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    cmdRX(Index).Enabled = True
Case 2
    cmdRX(Index).Enabled = False
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "BEGIN")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    RepMeseroX
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "END")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    cmdRX(Index).Enabled = True
Case 3
    cmdRX(Index).Enabled = False
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "BEGIN")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    RepDepX
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "END")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    cmdRX(Index).Enabled = True
Case Else
    MsgBox "AUN NO ESTA DISPONIBLE", vbInformation, BoxTit
End Select
End Sub

Private Sub cmdRZ_Click(Index As Integer)
Dim cReportesExtra As String

'REVISA LA IMPRESORA ANTES DE IMPRIMIR.
'17AGO2014
If OPOS_Eval("Coptr1.RecEmpty") Then
    ShowMsg "NO HAY PAPEL (RECIBO/Izquierda) EN LA IMPRESORA, CIERRE EL PROGRAMA, CAMBIELO AHORA e INTENTE IMPRIMIR NUEVAMENTE", vbYellow, vbBlue
    'ReportesEscribeLog ("Reporte X - NO HAY PAPEL (RECIBO/Izquierda) EN LA IMPRESORA, CAMBIELO AHORA e INTENTE IMPRIMIR NUEVAMENTE")
    EscribeLog ("Reporte Z - NO HAY PAPEL (RECIBO/Izquierda) EN LA IMPRESORA, CAMBIELO AHORA e INTENTE IMPRIMIR NUEVAMENTE")
    Me.MousePointer = vbDefault
'            Exit Sub
End If

If OPOS_Eval("Coptr1.RecNearEnd") Then
    ShowMsg "POR FAVOR DESPUES DE IMPRIMIR, REVISE EL PAPEL (RECIBO/Izquierda) EN LA IMPRESORA, PUEDE QUE SE ESTE ACABANDO", vbYellow, vbBlue
    'ReportesEscribeLog ("Reporte X - POR FAVOR REVISE EL PAPEL (RECIBO/Izquierda) EN LA IMPRESORA, PUEDE QUE SE ESTE ACABANDO")
    EscribeLog ("Reporte Z - POR FAVOR REVISE EL PAPEL (RECIBO/Izquierda) EN LA IMPRESORA, PUEDE QUE SE ESTE ACABANDO")
End If

Select Case Index
Case 0
'    BoxResp = MsgBox("¿ Desea realizar REPORTE Z ?", vbQuestion + vbYesNo, BoxTit)
'    If BoxResp = vbYes Then
    If ShowMsg(vbCrLf & vbCrLf & "¿ Desea Realizar REPORTE Z ?", vbBlue, vbYellow, vbYesNo) = vbYes Then
    
            '~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~
        '06OCT2011
        If cFullFiscal = "SI" Then
            On Error Resume Next
             'ELIMINADO (21NOV2011). EL IMPUESTO SE CALCULA x TRANSACCION
            'Call PutISC(FISCAL_.GetITBMS())
            cmdRZ(0).ToolTipText = FISCAL_.GetLastReporteZ
            If FISCAL_.SS_FiscalReporteCierre("Z") Then
                'REPORTE OK
            Else
                'REPORTE FAILED
            End If
            Call FISCAL_.SS_CashDrawerOpen
            On Error GoTo 0
        End If
        '~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~
    
        cmdRZ(0).Enabled = False
        
        EscribeLog ("INICIO DE REPORTE TERMINAL (Z) DE CAJEROS")
        
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "BEGIN")
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        
        'INFO: 29OCT2014
        'PARA LOS QUE DESEAN UN CONTEO DE PRODUCTOS EN EL REPORTE X
        cReportesExtra = UCase(GetFromINI("Facturacion", "ReportesExtra", App.Path & "\soloini.ini"))
        If Left(cReportesExtra, 6) = "PEREZA" Then
            If ReporteFacturas_y_Productos(cReportesExtra) Then
                Call PrintReportesExtra
            End If
        End If
        
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "END")
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        
        RepCajZ
        
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'        Call FE_MoveFacturas
'        If FE_Online Then
'            Call FE_JAVA_ProcesarFacturas
'        End If
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        
        cmdRZ(0).Enabled = True
        cmdRZ(1).Enabled = True
        cmdRZ(2).Enabled = True
    End If
Case 1
    cmdRZ(Index).Enabled = False
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "BEGIN")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    RepMesZ
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "END")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    cmdRZ(Index).Enabled = True
Case 2
    cmdRZ(Index).Enabled = False
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "BEGIN")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    RepMeseroZ
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "END")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    cmdRZ(Index).Enabled = True
Case Else
    MsgBox "AUN NO ESTA DISPONIBLE", vbInformation, BoxTit
End Select
End Sub

Private Sub Command1_Click()
rc = Sys_Pos.Coptr1.State
If rc = OposSBusy Then
    EscribeLog ("LA IMPRESORA DE REPORTES ESTA OCUPADA INTENTE OTRA VEZ CUANDO TERMINE DE IMPRIMIR")
    ShowMsg "LA IMPRESORA DE REPORTES ESTA OCUPADA " & vbCrLf & "INTENTE OTRA VEZ CUANDO TERMINE DE IMPRIMIR." & vbCrLf & " Error # " & rc, vbBlue, vbRed
    Exit Sub
End If

EscribeLog "Ventas. Reportes. SALIDA DE REPORTES X y Z"

Unload Me
End Sub

Private Sub Form_Load()
bPrintOk = True
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
'~~~~|||| 25AGO2011
'~~~~|||| APERTURA IMPRESORA FISCAL    ~~~~||||
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
If cFullFiscal = "SI" Then
    If FISCAL_.SS_Fiscal_IsOn = 1 Then
        'IMPRESORA DEBE ESTA OK
    Else
        ShowMsg "LA IMPRESORA  (FISCAL) NO FUE DETECTADA, SALGA DEL PROGRAMA y VUELVA A ENTRAR", vbYellow, vbRed
        LbFiscal.Enabled = False
    End If
Else
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    Rem Call Check_FE_Website("SILENCIO")
    
    If cLogMode = "SI" Then
        FE_Facil_ONLINE = False
    Else
        FE_Facil_ONLINE = True
    End If
    
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
End If

'~~~~ 17ENERO2021 ~~~~~~~~~~~~~~~~~~~~
If NOM_PRN_FACTURA = "" Then
    RptCajas.cmdRX(1).Enabled = False
    RptCajas.cmdRX(2).Enabled = False
    RptCajas.cmdRX(3).Enabled = False
    RptCajas.cmdRX(4).Enabled = False
    
    RptCajas.cmdRZ(1).Enabled = False
    RptCajas.cmdRZ(2).Enabled = False
End If
'~~~~ 17ENERO2021 ~~~~~~~~~~~~~~~~~~~~


End Sub
Private Function PrintReportesExtra() As Boolean
Dim nFreefileTR As Integer
Dim a$

nFreefileTR = FreeFile()
Open App.Path & "\TR_FISCAL.txt" For Input As #nFreefileTR
Do Until EOF(nFreefileTR)
    Line Input #nFreefileTR, a$
    Print2_OPOS_Dev a$
Loop
Sleep MUNBYNSeconds
Sys_Pos.Coptr1.CutPaper 100
Close #nFreefileTR

End Function
'---------------------------------------------------------------------------------------
' Procedure : ReporteFacturas_y_Productos
' Author    : hsequeira
' Date      : 24/06/2012  ~~ 28/06/2012
' Purpose   : CREA ARCHIVOS CON LAS FACTURAS DEL DIA y UN REPORTE DE VENTAS
'                     DE PRODUCTOS
' 8MAR2015
' SE ACTUALIZA PARA QUE SOLAMENTE IMPRIMA LOS PRODUCTOS Y NO LAS TRANSACCIONES
' AL PONER EL SOLOINI COMO = PerezaNoTrans
'---------------------------------------------------------------------------------------
'
Private Function ReporteFacturas_y_Productos(cPrintTransac As String) As Boolean
Dim cSQL As String
Dim rsREPORTES As ADODB.Recordset
Dim nFreefileTR As Integer
Dim nFreefilePR As Integer
Dim cTexto As String
Dim nVentas As Single, nDescuentos As Single, nImpuesto As Single
Dim nPVENTAS As Single, nPCANT As Long
Dim nPVENTAS_TOT As Single, nPCANT_TOT As Long
Dim nPlu As Long

'~~~~~~~~~~
'INFO: 6FEB2015
'~~~~~~~~~~
If cPrintTransac = "PEREZANOTRANS" Then
    'INFO: NO DESEAN EL REPORTE DE TRANSACCIONES, NADA MAS EL DE PRODUCTOS
    ' 6FEB2015
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Set rsREPORTES = New ADODB.Recordset
    On Error Resume Next
    Kill App.Path & "\TR_FISCAL.txt"
    Kill App.Path & "\PR_FISCAL.txt"
    On Error GoTo 0
    
    On Error GoTo ErrAdm:
    
    nFreefileTR = FreeFile()
    Open App.Path & "\TR_FISCAL.txt" For Output As #nFreefileTR
    GoTo ProductosOnly:
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Else
End If

cSQL = "SELECT DISTINCT A.NUM_TRANS, B.FISCAL, B.HORA, B.SUB_TOTAL, B.DESCUENTO, B.ITBM "
cSQL = cSQL & " FROM TRANSAC AS A, TRANSAC_FISCAL AS B "
cSQL = cSQL & " WHERE A.NUM_TRANS = B.DOC_SOLO "
cSQL = cSQL & " ORDER BY 1"

Set rsREPORTES = New ADODB.Recordset
rsREPORTES.Open cSQL, msConn, adOpenStatic

If rsREPORTES.EOF Then
    rsREPORTES.Close
    
    Set rsREPORTES = Nothing
    ReporteFacturas_y_Productos = False
    Exit Function
End If

ReporteFacturas_y_Productos = True

On Error Resume Next
Kill App.Path & "\TR_FISCAL.txt"
Kill App.Path & "\PR_FISCAL.txt"
On Error GoTo 0

nFreefileTR = FreeFile()
Open App.Path & "\TR_FISCAL.txt" For Output As #nFreefileTR

'''~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''INFO: REPORTES DE DOMICILIO
'''18ENE2017
'''~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''If nHayDomicilio > 0 Then
''
''    Print #nFreefileTR, "REPORTE DE DOMICILIO"
''    Print #nFreefileTR, Space(2)
''    Print #nFreefileTR, Format(Date, "LONG DATE")
''    Print #nFreefileTR, String(30, "=")
''    Print #nFreefileTR, "TRANS     DOMICILIO    VENTAS"
''    Print #nFreefileTR, String(30, "=")
''
''    Dim rsDOMI As ADODB.Recordset
''    Dim nDCant As Integer, nValorDomi As Single, nValorFactura As Single
''
''    Set rsDOMI = New ADODB.Recordset
''    rsDOMI.Open "SELECT * FROM DOMICILIO WHERE REPORTEZ = 'N' ORDER BY DOC_SOLO", msConn, adOpenStatic, adLockReadOnly
''    If Not rsDOMI.EOF Then
''        Do While Not rsDOMI.EOF
''
''            cTexto = Format(rsDOMI!DOC_SOLO, "@@@@@@@") & Space(6)                            '(9)
''            cTexto = cTexto & Format(Format(rsDOMI!VALOR_DOMICILIO, "##0.00"), "@@@@@@") & Space(2)    '(11 - 26)
''            cTexto = cTexto & Format(Format(rsDOMI!VALOR_FACTURA, "#,##0.00"), "@@@@@@@@")
''
''            'Print #nFreefileTR, rsDomi!DOC_SOLO & Space(5) & rsDomi!VALOR_DOMICILIO & Space(5) & rsDomi!VALOR_FACTURA
''            Print #nFreefileTR, cTexto
''            nDCant = nDCant + rsDOMI!CANT
''            nValorDomi = nValorDomi + rsDOMI!VALOR_DOMICILIO
''            nValorFactura = nValorFactura + rsDOMI!VALOR_FACTURA
''            rsDOMI.MoveNext
''        Loop
''        Print #nFreefileTR, String(30, "=")
''        cTexto = Format("TOTAL", "@@@@@") & "(" & nDCant & ")" & Space(3)
''        cTexto = cTexto & Format(Format(nValorDomi, "##0.00"), "@@@@@@") & Space(2)    '(11 - 26)
''        cTexto = cTexto & Format(Format(nValorFactura, "#,##0.00"), "@@@@@@@@")
''        Print #nFreefileTR, cTexto
''        'Print #nFreefileTR, "TOTALES " & nDCant & Space(5) & nValorDomi & Space(5) & nValorFactura
''    Else
''        Print #nFreefileTR, String(30, "=")
''        Print #nFreefileTR, "NO HAY VENTAS DE DOMICILIO"
''    End If
''
''    Print #nFreefileTR, Space(2)
''    Print #nFreefileTR, Space(2)
''    Print #nFreefileTR, Space(2)
''    Print #nFreefileTR, Space(2)
''    Print #nFreefileTR, Space(2)
''
''    rsDOMI.Close
''    Set rsDOMI = Nothing
''
''    msConn.Execute "UPDATE DOMICILIO SET REPORTEZ = 'S' WHERE REPORTEZ='N'"
''
''End If
'''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On Error Resume Next
rsREPORTES.MoveFirst
On Error GoTo 0

On Error GoTo ErrAdm:


Print #nFreefileTR, "REPORTE DE TRANSACCIONES"
Print #nFreefileTR, Space(2)
Print #nFreefileTR, Format(Date, "LONG DATE")
Print #nFreefileTR, String(30, "=")
Print #nFreefileTR, "  TRANS  HORA    VENTAS   IMP"
Print #nFreefileTR, String(30, "=")

Do While Not rsREPORTES.EOF

    cTexto = Format(CLng(LTrim(RTrim(rsREPORTES!FISCAL))), "@@@@@@") & Space(1) '(9)
    cTexto = cTexto & rsREPORTES!Hora & Space(1) '(6 - 15)
    cTexto = cTexto & Format(Format(rsREPORTES!SUB_TOTAL, "#,##0.00"), "@@@@@@@@") & Space(2)    '(11 - 26)
    cTexto = cTexto & Format(Format(rsREPORTES!ITBM, "##0.00"), "@@@@@@") '(6 - 42)
    
    nVentas = nVentas + rsREPORTES!SUB_TOTAL
    nDescuento = nDescuento + rsREPORTES!DESCUENTO
    nImpuesto = nImpuesto + rsREPORTES!ITBM
    
    Print #nFreefileTR, cTexto
    cTexto = ""
    If rsREPORTES!DESCUENTO <> 0 Then
        'SI HAY DESCUENTO GLOBAL LO IMPRIME EN LA SIGUIENTE LINEA
        Print #nFreefileTR, " > DESCT. GLOBAL: " & Format(rsREPORTES!DESCUENTO, "#,##0.00")
    End If

    rsREPORTES.MoveNext
Loop

Print #nFreefileTR, String(30, "=")
Print #nFreefileTR, Space(2)
Print #nFreefileTR, String(30, "=")
Print #nFreefileTR, Format("VENTAS = " & vbTab & Format(nVentas, "CURRENCY"), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
Print #nFreefileTR, Format("DESCUENTOS = " & vbTab & Format(nDescuentos, "CURRENCY"), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
Print #nFreefileTR, Format("IMPUESTO = " & vbTab & Format(nImpuesto, "CURRENCY"), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
Print #nFreefileTR, Format("TOTAL = " & vbTab & Format(nVentas + nImpuesto - nDescuento, "CURRENCY"), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
Print #nFreefileTR, String(30, "=")

Print #nFreefileTR, Space(2)
Print #nFreefileTR, Space(2)
Print #nFreefileTR, Space(2)
Print #nFreefileTR, Space(2)
Print #nFreefileTR, Space(2)

rsREPORTES.Close
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO: TAG PARA IMPRESION DE UNICAMENTE REPORTE DE PRODUCTOS
' 6FEB2015
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ProductosOnly:
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Print #nFreefileTR, "REPORTE DE PRODUCTOS"
Print #nFreefileTR, Space(2)
Print #nFreefileTR, Format(Date, "LONG DATE")
Print #nFreefileTR, String(30, "=")
Print #nFreefileTR, "PRODUCTO          CANT  VENTAS"
Print #nFreefileTR, String(30, "=")

cSQL = "SELECT PLU, DESCRIP AS PRODUCTO, SUM(CANT) AS CANT, "
cSQL = cSQL & " MAX(PRECIO_UNIT) AS P_UNIT, SUM(PRECIO) AS VENTAS"
cSQL = cSQL & " INTO LOLO8 "
cSQL = cSQL & " FROM TRANSAC WHERE LEFT(DESCRIP,2) <> '@@' "
cSQL = cSQL & " GROUP BY PLU, DESCRIP "
cSQL = cSQL & " ORDER BY 1"

msConn.Execute cSQL

cSQL = "SELECT * FROM LOLO8 ORDER BY PLU, P_UNIT DESC "
rsREPORTES.Open cSQL, msConn, adOpenStatic

If rsREPORTES.EOF Then
    rsREPORTES.Close
    
    Set rsREPORTES = Nothing
    Close #nFreefileTR
    
    On Error Resume Next
    msConn.Execute "DROP TABLE LOLO8"
    On Error GoTo 0
    
    ReporteFacturas_y_Productos = False
    Exit Function
Else
    ReporteFacturas_y_Productos = True
End If

Do While Not rsREPORTES.EOF
    nPCANT = nPCANT + rsREPORTES!CANT
    nPCANT_TOT = nPCANT_TOT + rsREPORTES!CANT
    nPVENTAS = nPVENTAS + rsREPORTES!VENTAS
    nPVENTAS_TOT = nPVENTAS_TOT + rsREPORTES!VENTAS
    cPRODUCTO = rsREPORTES!PRODUCTO
    nPlu = rsREPORTES!PLU
    
    cTexto = Format(Left(cPRODUCTO, 15), "!@@@@@@@@@@@@@@@") & Space(1)        '16
    cTexto = cTexto & Format(Format(nPCANT, "####"), "@@@@") & Space(1) '21
    cTexto = cTexto & Format(Format(nPVENTAS, "##,##0.00"), "@@@@@@@@@")        '30
    Print #nFreefileTR, cTexto
    
    On Error Resume Next
    rsREPORTES.MoveNext
    Do While nPlu = rsREPORTES!PLU
        'VIENEN LOS AJUSTES (EC, VO, DESCUENTOS)
        Select Case rsREPORTES!CANT
            Case Is <= 0
                If Left(rsREPORTES!PRODUCTO, 3) = "EC-" Then
                    cTexto = Format("<<CORRECCION>>", "!@@@@@@@@@@@@@@@") & Space(1)
                    cTexto = cTexto & Format(Format(Abs(rsREPORTES!CANT), "####"), "@@@@") & Space(1)
                    cTexto = cTexto & Format(Format(rsREPORTES!VENTAS, "##,##0.00"), "@@@@@@@@@")
                    Print #nFreefileTR, cTexto
                End If
                If Left(rsREPORTES!PRODUCTO, 3) = "VO-" Then
                    cTexto = Format("<<ANULACION>>", "!@@@@@@@@@@@@@@@") & Space(1)
                    cTexto = cTexto & Format(Format(Abs(rsREPORTES!CANT), "####"), "@@@@") & Space(1)
                    cTexto = cTexto & Format(Format(rsREPORTES!VENTAS, "##,##0.00"), "@@@@@@@@@")
                    Print #nFreefileTR, cTexto
                End If
            Case Else
                'ES DESCUENTO
                cTexto = Format("<<" & rsREPORTES!PRODUCTO & ">>", "!@@@@@@@@@@@@@@@") & Space(1)
                cTexto = cTexto & Format(Format(rsREPORTES!CANT, "####"), "@@@@") & Space(1)
                cTexto = cTexto & Format(Format(rsREPORTES!VENTAS, "##,##0.00"), "@@@@@@@@@")
                Print #nFreefileTR, cTexto
        End Select
        nPVENTAS = nPVENTAS + rsREPORTES!VENTAS
        nPVENTAS_TOT = nPVENTAS_TOT + rsREPORTES!VENTAS
        rsREPORTES.MoveNext
        If rsREPORTES.EOF Then Exit Do
    Loop
    On Error GoTo 0

    nPCANT = 0
    nPVENTAS = 0
Loop

Print #nFreefileTR, String(30, "=")
Print #nFreefileTR, Space(2)
Print #nFreefileTR, String(30, "=")
Print #nFreefileTR, Format("VENTAS = " & vbTab & Format(nPVENTAS_TOT, "CURRENCY"), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
Print #nFreefileTR, String(30, "=")

Print #nFreefileTR, Space(2)
Print #nFreefileTR, Space(2)
Print #nFreefileTR, Space(2)
Print #nFreefileTR, Space(2)
Print #nFreefileTR, Space(2)
Print #nFreefileTR, Space(2)
Print #nFreefileTR, Space(2)
Print #nFreefileTR, Space(2)
Print #nFreefileTR, Space(2)
Print #nFreefileTR, Space(2)
Print #nFreefileTR, Space(2)
Print #nFreefileTR, Space(2)
Print #nFreefileTR, Space(2)

Close #nFreefileTR
rsREPORTES.Close
Set rsREPORTES = Nothing

msConn.Execute "DROP TABLE LOLO8"

On Error GoTo 0

Exit Function

ErrAdm:
If Err.Number = 13 Then
    'FALTA UN CAMPO EN LA DESCRIPCION
    ShowMsg "Falta un dato en El Reporte de Facturas y Productos, en la impresion de una linea, el reporte va a continuar, por favor reviselo", vbYellow, vbBlue
Else
    If Err.Number = 2147217900 Then
        'INFO: LA TABLA LOLO8 YA EXISTES, ASI QUE NO LA VA A CREAR DE NUEVO. NO ENVIAR MENSAJE.
    Else
        ShowMsg "Error en El Reporte de Facturas y Productos. El Reporte va a Continuar" & vbCrLf & Err.Number & " - " & Err.description, vbYellow, vbBlue
    End If
End If
Resume Next
End Function

Private Sub lbFactRapida_Click()
If Not Check_FE_Website Then
    Rem ShowMsg "WEB SERVICE DE FACTURA ELECTRONICA" & vbCrLf & "NO ESTA DISPONIBLE" & vbCrLf & "PASANDO A MODO" & vbCrLf & vbCrLf & " (OFF-LINE)", vbBlue, vbCyan
    Rem EscribeLog "WEB SERVICE DE FACTURA ELECTRONICA NO ESTA DISPONIBLE"
    Rem Exit Sub
End If

FERapida.Show 1
End Sub

Private Sub LbFiscal_Click()
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
If cFullFiscal = "SI" Then
    FiscalReports.Show 1
Else
'    If VerificaURL Then
'        'WEB SERVICE DE FACTURA ELECTRONICA ESTA ARRIBA
'        FE_Online = True
'    Else
'        FE_Online = False
'        ShowMsg "WEB SERVICE DE FACTURA ELECTRONICA NO ESTA DISPONIBLE" & vbCrLf & "PASANDO A MODO (OFF-LINE)", vbBlue, vbCyan
'    End If
    
    'OpcFactElectonica.Show 1
    
    'ShowMsg "IMPRESORA FISCAL NO ESTA HABILITADA", vbYellow, vbRed
    Rem Me.MousePointer = vbHourglass
    Rem If Not Check_FE_Website Then
    Rem    Me.MousePointer = vbNormal
    Rem    ShowMsg "WEB SERVICE DE FACTURA ELECTRONICA" & vbCrLf & "NO ESTA DISPONIBLE" & vbCrLf & "PASANDO A MODO" & vbCrLf & vbCrLf & " (OFF-LINE)", vbBlue, vbCyan
    Rem    EscribeLog "WEB SERVICE DE FACTURA ELECTRONICA NO ESTA DISPONIBLE"
    Rem    Exit Sub
    Rem End If
    Rem Me.MousePointer = vbNormal
    'FEReports.Show 1
    fast_FEReports.Show 1
End If
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
End Sub

Private Sub mnuRepZ_Click()
Dim nFreefile As Integer
Dim a$
Dim FACTURA_FILE As String
Dim DATA_PATH  As String
Dim vResp

'vResp = MsgBox("¿ Desea Imprimir Copia del Reporte Z ?", vbQuestion + vbYesNo, BoxTit)

'If vResp = vbYes Then
If ShowMsg("¿ Desea Imprimir Copia del Reporte Z ?", vbYellow, vbBlue, vbYesNo) = vbYes Then

    FACTURA_FILE = "REPITEZ.TXT"
    
    DATA_PATH = GetFromINI("General", "DirectorioDatos", App.Path & "\soloini.ini")
    cFactFile = DATA_PATH + FACTURA_FILE
    
    On Error Resume Next
    If Dir(cFactFile) = "" Then
        EscribeLog ("Se intento imprimir Reporte Z que no existe")
        ShowMsg "No se ha generado una copia del Reporte Z para imprimir", vbYellow, vbRed
        Exit Sub
    Else
        EscribeLog ("Impresión de Copia de Reporte Z")
        nFreefile = FreeFile
        Open cFactFile For Input As nFreefile
        Do Until EOF(nFreefile)
            Line Input #nFreefile, a$
            Print2_OPOS_Dev a$
        Loop
        Close nFreefile
        Sleep MUNBYNSeconds
        Coptr1.CutPaper (100)
        Kill cFactFile
    End If
End If
On Error GoTo 0
End Sub



'---------------------------------------------------------------------------------------
' Procedure : ReporteEspecial
' Author    : hsequeira
' Date      : 28/08/2015
' Purpose   : REPORTE ESPECIAL PARA JAP JAP
' 1SEP2015: SE AJUSTAN LAS LINEAS PARA QUE LA TM-U950 PUEDA IMPRIMIR TODO EL TEXTO
' ESTA IMPRESORA USA UN MAXIMO DE 30 LINEAS, MIENTRAS QUE LAS NUEVAS IMPRIMEN 40
'---------------------------------------------------------------------------------------
'
Private Sub ReporteEspecial()

Dim cSQL As String, cSQL2 As String
Dim rsGGRUPOS As ADODB.Recordset
Dim rsGGRUPOS_2 As ADODB.Recordset
Dim iFila As Long
Dim g_GRUPO As String
Dim cCad As String
Dim nTotaRepEspcial As Single
Dim i As Integer

'ES EL DIA DE HOY
g_TABLE = "TRANSAC AS B"

'OBTIENE LOS VALORES SIN INCLUIR DESCUENTO
cSQL = "SELECT A.G_DESCRIP, SUM(A.G_VALOR * B.CANT)  AS CANTIDAD, "
'cSQL = cSQL & " SUM(B.CANT * B.PRECIO_UNIT) AS VENTAS, 0 AS VTA_NETA "
'INFO: VALOR ABSOLUTO DE CANTIDAD. 20SEP2015
cSQL = cSQL & " SUM(ABS(B.CANT) * B.PRECIO_UNIT) AS VENTAS, 0 AS VTA_NETA "
cSQL = cSQL & " FROM G_GRUPOS AS A, " & g_TABLE
cSQL = cSQL & " WHERE A.G_PLU = B.PLU"
'cSQL = cSQL & " AND B.FECHA BETWEEN '" & dF1 & "'"
'cSQL = cSQL & " AND '" & dF2 & "'"
cSQL = cSQL & " AND '%' NOT IN (B.DESCRIP) "
cSQL = cSQL & " AND B.DESCRIP NOT LIKE '%DESCUENTO%' "
cSQL = cSQL & " AND B.DESCRIP NOT LIKE  '%@@%' "
cSQL = cSQL & " GROUP BY A.G_DESCRIP"

cSQL2 = "SELECT A.G_DESCRIP, SUM(A.G_VALOR * B.CANT)  AS CANTIDAD, "
'cSQL2 = cSQL2 & " SUM(B.CANT * B.PRECIO_UNIT) AS VTA_NETA "
'INFO: VALOR ABSOLUTO DE CANTIDAD. 20SEP2015
cSQL2 = cSQL2 & " SUM(ABS(B.CANT) * B.PRECIO_UNIT) AS VTA_NETA "
cSQL2 = cSQL2 & " FROM G_GRUPOS AS A, " & g_TABLE
cSQL2 = cSQL2 & " WHERE A.G_PLU = B.PLU"
'cSQL2 = cSQL2 & " AND B.FECHA BETWEEN '" & dF1 & "'"
'cSQL2 = cSQL2 & " AND '" & dF2 & "'"
cSQL2 = cSQL2 & " AND '%' NOT IN (B.DESCRIP) "
'cSQL2 = cSQL2 & " AND B.DESCRIP NOT LIKE '%DESCUENTO%' "
cSQL2 = cSQL2 & " AND B.DESCRIP NOT LIKE  '%@@%' "
cSQL2 = cSQL2 & " GROUP BY A.G_DESCRIP"

Set rsGGRUPOS = New ADODB.Recordset
Set rsGGRUPOS_2 = New ADODB.Recordset
rsGGRUPOS.Open cSQL, msConn, adOpenStatic, adLockOptimistic
rsGGRUPOS_2.Open cSQL2, msConn, adOpenStatic, adLockOptimistic

If rsGGRUPOS.EOF Then
    Me.MousePointer = vbNormal
    'MsgBox "NO EXISTEN DATOS PARA ESTAS FECHAS", vbCritical
    rsGGRUPOS.Close
    rsGGRUPOS_2.Close
    Set rsGGRUPOS = Nothing
    Set rsGGRUPOS_2 = Nothing
    Exit Sub
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "BEGIN")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Print2_OPOS_Dev Date & Space(2) & Time
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev "REPORTE ESPECIAL (Z)"
Print2_OPOS_Dev Chr(&HD) & Chr(&HA)
Print2_OPOS_Dev rs00!DESCRIP
Print2_OPOS_Dev rs00!RAZ_SOC
Print2_OPOS_Dev "RUC:" & rs00!ruc
Print2_OPOS_Dev "SERIAL:" & rs00!SERIAL
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev "CONTADOR Z : " & (rs00!Z_CDEP)
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev "=============================="

Do While Not rsGGRUPOS.EOF
    cCad = Format(rsGGRUPOS!G_DESCRIP, "!@@@@@@@@@@@@@@@") & Space(1)
    cCad = cCad & Format(Format(rsGGRUPOS!Cantidad, "STANDARD"), "@@@@@@") & Space(2)
    rsGGRUPOS_2.Find "G_DESCRIP = '" & rsGGRUPOS!G_DESCRIP & "'"
    cCad = cCad & Format(Format(rsGGRUPOS_2!VTA_NETA, "0.00"), "@@@@@@")
    nTotaRepEspcial = nTotaRepEspcial + rsGGRUPOS_2!VTA_NETA
    Print2_OPOS_Dev cCad
    rsGGRUPOS.MoveNext
Loop
Print2_OPOS_Dev "=============================="
Print2_OPOS_Dev "TOTAL VTA NETA: " & Format(nTotaRepEspcial, "CURRENCY")

For i = 1 To 10
    Print2_OPOS_Dev Space(2)
Next

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "END")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sleep MUNBYNSeconds
Sys_Pos.Coptr1.CutPaper 100

rsGGRUPOS.Close
rsGGRUPOS_2.Close
Set rsGGRUPOS = Nothing
Set rsGGRUPOS_2 = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetVentasDomicilio
' Author    : hsequeira
' Date      : 08/03/2017
' Purpose   : OBTIENE LA VENTAS DE DOMICILIO
'---------------------------------------------------------------------------------------
'
Private Function GetVentasDomicilio(nnDVentasDepartamentales As Double, cTipoReporte As String)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO: REPORTES DE DOMICILIO
'18ENE2017 / 02FEB2017
'08MAR2017
'~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If nHayDomicilio > 0 Then
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Call OPOSTransactionPrint(Sys_Pos.Coptr1.Name, "BEGIN")
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    'Print2_OPOS_Dev Space(2)
    'Print2_OPOS_Dev Format(Date, "LONG DATE")
    'Print2_OPOS_Dev String(30, "=")
    'Print2_OPOS_Dev "TRANS     DOMICILIO    VENTAS"
    'Print2_OPOS_Dev String(30, "=")
    
    Dim rsDOMI As ADODB.Recordset
    Dim nDCant As Integer, nValorDomi As Single, nValorFactura As Single
    Dim cTexto As String
    
    Set rsDOMI = New ADODB.Recordset
    rsDOMI.Open "SELECT * FROM DOMICILIO WHERE REPORTEZ = 'N' ORDER BY DOC_SOLO", msConn, adOpenStatic, adLockReadOnly
    If Not rsDOMI.EOF Then
    
        Print2_OPOS_Dev String(30, "=")
        Print2_OPOS_Dev "REPORTE DE DOMICILIO"
    
        Do While Not rsDOMI.EOF
        
            'cTexto = Format(rsDOMI!DOC_SOLO, "@@@@@@@") & Space(6)                            '(9)
            'cTexto = cTexto & Format(Format(rsDOMI!VALOR_DOMICILIO, "##0.00"), "@@@@@@") & Space(2)    '(11 - 26)
            'cTexto = cTexto & Format(Format(rsDOMI!VALOR_FACTURA, "#,##0.00"), "@@@@@@@@")
        
            'Print2_OPOS_Dev  rsDomi!DOC_SOLO & Space(5) & rsDomi!VALOR_DOMICILIO & Space(5) & rsDomi!VALOR_FACTURA
            'Print2_OPOS_Dev cTexto
            nDCant = nDCant + rsDOMI!CANT
            nValorDomi = nValorDomi + rsDOMI!VALOR_DOMICILIO
            nValorFactura = nValorFactura + rsDOMI!VALOR_FACTURA
            rsDOMI.MoveNext
        Loop
        Print2_OPOS_Dev String(30, "=")
        cTexto = Format("TOTAL", "@@@@@") & "(" & nDCant & ")" & Space(4)
        cTexto = cTexto & Format(Format(nValorDomi, "##0.00"), "@@@@@@") & Space(2)    '(11 - 26)
        cTexto = cTexto & Format(Format(nValorFactura, "#,##0.00"), "@@@@@@@@")
        Print2_OPOS_Dev cTexto
            'Call TEXTFileZ("APPEND", "DOMICILIO:", -1, Format(nValorDomi, "CURRENCY"), "Y")
            Call TEXTFileZ("APPEND", "DOMICILIO:", nDCant, Format(nValorFactura, "STANDARD"), "Y")
        Print2_OPOS_Dev String(30, "-")
        'Print2_OPOS_Dev "NUEVO TOTAL: " & Format(nnDVentasDepartamentales - nValorFactura, "CURRENCY")
        Print2_OPOS_Dev "NUEVO TOTAL: " & Format(Format(nnDVentasDepartamentales - nValorFactura, "CURRENCY"), "@@@@@@@@@@@@@")
            Call TEXTFileZ("APPEND", "NUEVO TOTAL:", -1, Format(nnDVentasDepartamentales - nValorFactura, "CURRENCY"), "Y")
        'Print2_OPOS_Dev  "TOTALES " & nDCant & Space(5) & nValorDomi & Space(5) & nValorFactura
        Print2_OPOS_Dev String(30, "=")
    Else
        Print2_OPOS_Dev String(30, "=")
        Print2_OPOS_Dev "NO HAY VENTAS DE DOMICILIO"
        Print2_OPOS_Dev String(30, "=")
    End If

    rsDOMI.Close
    Set rsDOMI = Nothing
    
    If cTipoReporte = "Z" Then
        msConn.Execute "UPDATE DOMICILIO SET REPORTEZ = 'S' WHERE REPORTEZ='N'"
    End If

End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

End Function


'---------------------------------------------------------------------------------------
' Procedure : ReporteEspecial
' Author    : hsequeira
' Date      : 28/08/2015
' Purpose   : REPORTE ESPECIAL PARA JAP JAP
' 1SEP2015: SE AJUSTAN LAS LINEAS PARA QUE LA TM-U950 PUEDA IMPRIMIR TODO EL TEXTO
' ESTA IMPRESORA USA UN MAXIMO DE 30 LINEAS, MIENTRAS QUE LAS NUEVAS IMPRIMEN 40
' Date      : 26/07/2017
' Se crea tabla GRUPOS_FAST (GrupoID, GrupoDescripcion) Ò ITEM_ID (10), GL_ACCOUNT (15)
' En la Tabla Opciones, se crea campos: FAST_CUST_ID (10), FAST_ACC_REC (15), FAST_DISTRIBUTION (2)
' Se crea ReporteEspecial_v2 en ReptCajas.frm con la función GetFromTable
'---------------------------------------------------------------------------------------
'
Private Sub ReporteEspecial_v2()

Dim cSQL As String, cSQL2 As String
Dim rsGGRUPOS As ADODB.Recordset
Dim rsGGRUPOS_2 As ADODB.Recordset
Dim iFila As Long
Dim g_GRUPO As String
Dim cCad As String
Dim nTotaRepEspcial As Single
Dim i As Integer

'@@@ === @@@ === @@@ === @@@
Dim txtRepo1 As String, txtRepo2 As String
Dim txtFecha As String, dFecha As Date
Dim nFileNumber As Integer

On Error GoTo ErrAmdRE:

nFileNumber = FreeFile()
'@@@ === @@@ === @@@ === @@@

'ES EL DIA DE HOY
g_TABLE = "TRANSAC AS B"

'OBTIENE LOS VALORES SIN INCLUIR DESCUENTO
cSQL = "SELECT A.G_DESCRIP, SUM(A.G_VALOR * B.CANT)  AS CANTIDAD, "
'cSQL = cSQL & " SUM(B.CANT * B.PRECIO_UNIT) AS VENTAS, 0 AS VTA_NETA "
'INFO: VALOR ABSOLUTO DE CANTIDAD. 20SEP2015
cSQL = cSQL & " SUM(ABS(B.CANT) * B.PRECIO_UNIT) AS VENTAS, 0 AS VTA_NETA "
cSQL = cSQL & " FROM G_GRUPOS AS A, " & g_TABLE
cSQL = cSQL & " WHERE A.G_PLU = B.PLU"
'cSQL = cSQL & " AND B.FECHA BETWEEN '" & dF1 & "'"
'cSQL = cSQL & " AND '" & dF2 & "'"
cSQL = cSQL & " AND '%' NOT IN (B.DESCRIP) "
cSQL = cSQL & " AND B.DESCRIP NOT LIKE '%DESCUENTO%' "
cSQL = cSQL & " AND B.DESCRIP NOT LIKE  '%@@%' "
cSQL = cSQL & " GROUP BY A.G_DESCRIP"

cSQL2 = "SELECT A.G_DESCRIP, SUM(A.G_VALOR * B.CANT)  AS CANTIDAD, "
'cSQL2 = cSQL2 & " SUM(B.CANT * B.PRECIO_UNIT) AS VTA_NETA "
'INFO: VALOR ABSOLUTO DE CANTIDAD. 20SEP2015
cSQL2 = cSQL2 & " SUM(ABS(B.CANT) * B.PRECIO_UNIT) AS VTA_NETA "
cSQL2 = cSQL2 & " FROM G_GRUPOS AS A, " & g_TABLE
cSQL2 = cSQL2 & " WHERE A.G_PLU = B.PLU"
'cSQL2 = cSQL2 & " AND B.FECHA BETWEEN '" & dF1 & "'"
'cSQL2 = cSQL2 & " AND '" & dF2 & "'"
cSQL2 = cSQL2 & " AND '%' NOT IN (B.DESCRIP) "
'cSQL2 = cSQL2 & " AND B.DESCRIP NOT LIKE '%DESCUENTO%' "
cSQL2 = cSQL2 & " AND B.DESCRIP NOT LIKE  '%@@%' "
cSQL2 = cSQL2 & " GROUP BY A.G_DESCRIP"

Set rsGGRUPOS = New ADODB.Recordset
Set rsGGRUPOS_2 = New ADODB.Recordset
rsGGRUPOS.Open cSQL, msConn, adOpenStatic, adLockOptimistic
rsGGRUPOS_2.Open cSQL2, msConn, adOpenStatic, adLockOptimistic

If rsGGRUPOS.EOF Then
    Me.MousePointer = vbNormal
    'MsgBox "NO EXISTEN DATOS PARA ESTAS FECHAS", vbCritical
    rsGGRUPOS.Close
    rsGGRUPOS_2.Close
    Set rsGGRUPOS = Nothing
    Set rsGGRUPOS_2 = Nothing
    Exit Sub
End If

'@@@ === @@@ === @@@ === @@@
txtFecha = Format(Date, "YYYYMMDD")
dFecha = Format(Date, "mm/dd/yyyy")
Open DATA_PATH & txtFecha & ".csv" For Output As #nFileNumber
'@@@ === @@@ === @@@ === @@@

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "BEGIN")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Print2_OPOS_Dev Date & Space(2) & Time
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev "REPORTE ESPECIAL (Z)"
Print2_OPOS_Dev Chr(&HD) & Chr(&HA)
Print2_OPOS_Dev rs00!DESCRIP
Print2_OPOS_Dev rs00!RAZ_SOC
Print2_OPOS_Dev "RUC:" & rs00!ruc
Print2_OPOS_Dev "SERIAL:" & rs00!SERIAL
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev "CONTADOR Z : " & (rs00!Z_CDEP)
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev "=============================="


'@@@ === @@@ === @@@ === @@@'@@@ === @@@ === @@@ === @@@
txtRepo1 = GetFromTable(1) & "," & txtFecha & "," & dFecha & "," & GetFromTable(2) & ",0,10,"
'@@@ === @@@ === @@@ === @@@'@@@ === @@@ === @@@ === @@@

Do While Not rsGGRUPOS.EOF
    cCad = Format(rsGGRUPOS!G_DESCRIP, "!@@@@@@@@@@@@@@@") & Space(1)
    cCad = cCad & Format(Format(rsGGRUPOS!Cantidad, "STANDARD"), "@@@@@@") & Space(2)
    
    rsGGRUPOS_2.Find "G_DESCRIP = '" & rsGGRUPOS!G_DESCRIP & "'"
    
    cCad = cCad & Format(Format(rsGGRUPOS_2!VTA_NETA, "0.00"), "@@@@@@")
    nTotaRepEspcial = nTotaRepEspcial + rsGGRUPOS_2!VTA_NETA
    
    '@@@ === @@@ === @@@ === @@@
    If GetFromTable(3, rsGGRUPOS!G_DESCRIP) = "No-Id" Then
        'EXCLUYE LOS QUE NO SE DESEAN CONTABILIZAR
    Else
        txtRepo2 = rsGGRUPOS!Cantidad & "," & GetFromTable(3, rsGGRUPOS!G_DESCRIP) & "," & rsGGRUPOS!G_DESCRIP & "," & GetFromTable(4, rsGGRUPOS!G_DESCRIP) & ",0," & rsGGRUPOS_2!VTA_NETA & ","
        Print #nFileNumber, txtRepo1 & txtRepo2
        txtRepo2 = ""
    End If
    '@@@ === @@@ === @@@ === @@@
    
    Print2_OPOS_Dev cCad
    
    rsGGRUPOS.MoveNext
Loop

'@@@ === @@@ === @@@ === @@@
Close #nFileNumber
'@@@ === @@@ === @@@ === @@@

Print2_OPOS_Dev "=============================="
Print2_OPOS_Dev "TOTAL VTA NETA: " & Format(nTotaRepEspcial, "CURRENCY")

For i = 1 To 10
    Print2_OPOS_Dev Space(2)
Next

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "END")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sleep MUNBYNSeconds
Sys_Pos.Coptr1.CutPaper 100

rsGGRUPOS.Close
rsGGRUPOS_2.Close
Set rsGGRUPOS = Nothing
Set rsGGRUPOS_2 = Nothing

On Error GoTo 0
Exit Sub

ErrAmdRE:
EscribeLog "|||||| NO HAY REPORTE ESPECIAL PARA ESTA BASE DE DATOS ||||||"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetFromTable
' Author    : hsequeira
' Date      : 26/07/2017
' Date      : 25/08/2017: VALIDA EOF y VALORES NULOS
' Purpose   : OBTIENE LOS VALORES SOLICITADOS PARA NUEVO INFORME DE JAP JAP
'---------------------------------------------------------------------------------------
'
Private Function GetFromTable(nTable As Integer, Optional cCampo As String) As String
Dim cRetorno As String
Dim rsGettt As New ADODB.Recordset
Dim nErrorN As Long
Dim cErrDescrip As String
Dim cSQL As String

On Error GoTo GetFromTable_Error

Select Case nTable
    Case 1
        cSQL = "SELECT FAST_CUST_ID FROM OPCIONES"
        rsGettt.Open cSQL, msConn, adOpenStatic, adLockReadOnly
        If rsGettt.EOF Then cRetorno = "No-Id" Else cRetorno = IIf(IsNull(rsGettt!FAST_CUST_ID), "No-Id", rsGettt!FAST_CUST_ID)
    Case 2
        cSQL = "SELECT FAST_ACC_REC FROM OPCIONES"
        rsGettt.Open cSQL, msConn, adOpenStatic, adLockReadOnly
        If rsGettt.EOF Then cRetorno = "No-Id" Else cRetorno = IIf(IsNull(rsGettt!FAST_ACC_REC), "No-Id", rsGettt!FAST_ACC_REC)
    Case 3
        cSQL = "SELECT ITEM_ID FROM GRUPOS_FAST WHERE G_DESCRIP = '" & cCampo & "'"
        rsGettt.Open cSQL, msConn, adOpenStatic, adLockReadOnly
        If rsGettt.EOF Then cRetorno = "No-Id" Else cRetorno = IIf(IsNull(rsGettt!ITEM_ID), "No-Id", rsGettt!ITEM_ID)
    Case 4
        cSQL = "SELECT GL_ACCOUNT FROM GRUPOS_FAST WHERE G_DESCRIP = '" & cCampo & "'"
        rsGettt.Open cSQL, msConn, adOpenStatic, adLockReadOnly
        If rsGettt.EOF Then cRetorno = "No-Id" Else cRetorno = IIf(IsNull(rsGettt!GL_ACCOUNT), "No-Id", rsGettt!GL_ACCOUNT)
End Select
GetFromTable = cRetorno
rsGettt.Close
Set rsGettt = Nothing

On Error GoTo 0
Exit Function

GetFromTable_Error:
    nErrorN = Err.Number
    cErrDescrip = cSQL & " - " & Err.description
    EscribeLog ("Error " & nErrorN & " (" & cErrDescrip & ") EN TABLAS NUEVAS DE CONTABILIDAD")
    ShowMsg "Error " & nErrorN & " (" & cErrDescrip & ")" & vbCrLf & "EN TABLAS NUEVAS DE CONTABILIDAD", vbYellow, vbRed
    Resume
End Function

Private Function InfoDomicilio(nVentasTotales As Double, cXY As String) As Boolean
Dim cSQL As String, cSQL2 As String
Dim rsTMP As ADODB.Recordset
Dim rsTMP2 As ADODB.Recordset
Dim rsDOMI As ADODB.Recordset
Dim iFila As Long
Dim nNum() As Long
Dim cFecha As String
Dim nVentas() As String
Dim i As Integer
Dim j As Integer
Dim nTotalVentas As Single
Dim nRegistros As Integer
Dim nPedidos As Integer
Dim cNumTrans  As String
Dim nFreefile As Integer        'PARA LEER LOS VALORES DE LA STAR
Dim cFile As String, cReturn As String
Dim nDomiPLU As Integer
Dim MiLen1 As Integer
Dim Milen2 As Integer

    On Error GoTo InfoDomicilio_Error
    nDomiPLU = Val(GetFromINI("Facturacion", "DomicilioPLU", App.Path & "\soloini.ini"))
    If nDomiPLU = 0 Then
        On Error GoTo 0
        Exit Function
    End If

        On Error Resume Next
        cFile = App.Path & "\LOG.txt"
        Kill cFile
    On Error GoTo 0

nFreefile = FreeFile()
Open cFile For Output As #nFreefile

Me.MousePointer = vbHourglass

Set rsDOMI = New ADODB.Recordset
Set rsTMP = New ADODB.Recordset
Set rsTMP2 = New ADODB.Recordset

cSQL = "SELECT DISTINCT A.NUM_TRANS, A.FECHA "
cSQL = cSQL & " FROM TRANSAC AS A "
cSQL = cSQL & " WHERE A.PLU = " & nDomiPLU
cSQL = cSQL & " AND A.VALID AND A.PRECIO > 0"
cSQL = cSQL & " ORDER BY A.FECHA"

rsTMP.Open cSQL, msConn, adOpenStatic, adLockOptimistic

'INFO: UPDATE ESTA FALLANDO EL REPORTE, PERO NO SE PORQUE
'INFO: 19MAR2020
If rsTMP.RecordCount = 0 Then
    rsTMP.Close
    Close #nFreefile
    Me.MousePointer = vbNormal
    Exit Function
End If

rsTMP.MoveFirst
txtFecIni = rsTMP!FECHA
rsTMP.MoveLast
txtFecFin = rsTMP!FECHA
rsTMP.MoveFirst

'ReDim nNUM(rsTMP.RecordCount)
'ReDim cFecha(rsTMP.RecordCount)
j = txtFecFin - txtFecIni
'j = 3
ReDim nVentas(2, j) ' FECHA, VENTAS, PEDIDOS

If Not rsTMP.EOF Then cFecha = rsTMP!FECHA Else cFecha = ""

Do While Not rsTMP.EOF
    Do While cFecha = rsTMP!FECHA
        cNumTrans = cNumTrans & rsTMP!NUM_TRANS & ","
        rsTMP.MoveNext
        If rsTMP.EOF Then
            Exit Do
        End If
    Loop
    cNumTrans = Left(cNumTrans, Len(cNumTrans) - 1)
    Print #nFreefile, Now & Space(1) & cNumTrans
    'Debug.Print cNumTrans
    nPedidos = GetRegistros(cNumTrans)
    cSQL2 = "SELECT SUM(PRECIO) AS VENTAS FROM TRANSAC WHERE NUM_TRANS IN (" & cNumTrans & ") "
    rsTMP2.Open cSQL2, msConn, adOpenStatic, adLockOptimistic
    
    If Not rsTMP2.EOF Then
        'nVentas(0, i) = Format(cFecha, "####-##-##")
        nVentas(0, i) = "DELIVERY"
        nVentas(1, i) = Str(rsTMP2!VENTAS)
        'nVentas(2, i) = "(" & nPedidos & ") Pedidos"
        nVentas(2, i) = nPedidos
        nPedidos = 0
        nRegistros = nRegistros + 1
        nTotalVentas = nTotalVentas + rsTMP2!VENTAS
    End If
    i = i + 1
    
    rsTMP2.Close
    cNumTrans = ""
    If rsTMP.EOF Then Exit Do
    cFecha = rsTMP!FECHA
Loop

On Error Resume Next
MiLen1 = Len(nVentas(2, 0))
Milen2 = Len(Format(nVentas(1, 0), "STANDARD"))

Print2_OPOS_Dev "--------------------------"
Print2_OPOS_Dev FormatTexto(nVentas(0, 0), 13) & Space(4 - MiLen1) & nVentas(2, 0) & Space(9 - Milen2) & Format(nVentas(1, 0), "STANDARD")
Print2_OPOS_Dev "PORCENTAJE VENTAS:  " & Format(nVentas(1, 0) / nVentasTotales, "PERCENT")
Print2_OPOS_Dev "--------------------------"

If cXY = "Z" Then
    Call TEXTFileZ("APPEND", "--------------------------", , , "N")
    Call TEXTFileZ("APPEND", FormatTexto(nVentas(0, 0), 13), -1, Format(nVentas(1, 0), "STANDARD"), "Y")
    Call TEXTFileZ("APPEND", "PORCENTAJE VENTAS:  " & Format(nVentas(1, 0) / nVentasTotales, "PERCENT"), , , "N")
    Call TEXTFileZ("APPEND", "--------------------------", , , "N")
End If
On Error GoTo 0
ReDim Preserve nVentas(2, nRegistros)

'cData = "TOTAL PERIODO"
'cData = cData & ";" & Format(nTotalVentas, "STANDARD")
'DD_DOMICILIO.Rows.Add sgFormatCharSeparatedValue, cData, ";"

Me.MousePointer = vbNormal
Close #nFreefile

   On Error GoTo 0
   Exit Function

InfoDomicilio_Error:
    If Err.Number = 9 Then
        ShowMsg "FECHAS INVALIDAS", vbYellow, vbRed
    Else
        ShowMsg "Error " & Err.Number & " (" & Err.description & ") en InfoDomicilio def RptCajas"
    End If
    Close #nFreefile
End Function
Private Function GetRegistros(cCadena As String) As Integer
    Dim asciiToSearchFor As Integer
    Dim Count As Integer
    asciiToSearchFor = Asc(",")
    For i = 1 To Len(cCadena)
        If Asc(Mid$(cCadena, i, 1)) = asciiToSearchFor Then Count = Count + 1
    Next
    GetRegistros = Count + 1
End Function

'---------------------------------------------------------------------------------------
' Procedure : FE_MoveFacturas
' Author    : hsequeira
' Date      : 11/11/2023
' Purpose   : MUEVE LAS FACTURAS PYTHON AL FOLDER HISTORICO
'---------------------------------------------------------------------------------------
'
Private Function FE_MoveFacturas() As Boolean
Dim nFacturasMovidas As Long
Dim iMover As Integer
Dim cmsg As String

'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
   On Error GoTo FE_MoveFacturas_Error

iMover = 1
sFilename = Dir(cFE_Path & "\FE\*_EnviarHKA.PY")

Do While sFilename > ""
    If Left(sFilename, 2) = "vb" Then
    iMover = 2
    Else
        iMover = 3
        Debug.Print sFilename
        FileCopy cFE_Path & "\FE\" & sFilename, FE_Historico_Path & "\HISTORICO\" & sFilename
        iMover = 4
        nFacturasMovidas = nFacturasMovidas + 1
        On Error Resume Next
        Kill cFE_Path & "\FE\" & sFilename
        iMover = 5
        On Error GoTo 0
    End If
    iMover = 6
    sFilename = Dir()
    iMover = 7
Loop
FE_MoveFacturas = True
iMover = 8
EscribeLog "RepCaja. SE RESPALDARON (" & nFacturasMovidas & ") FACTURAS ELECTRONICAS"
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

   On Error GoTo 0
   Exit Function

FE_MoveFacturas_Error:
    cmsg = "Error " & Err.Number & " (" & Err.description & ") in FE_MoveFacturas ( " & iMover & ")"
    ShowMsg cmsg, vbBlue, vbCyan
    EscribeLog cmsg
End Function

'---------------------------------------------------------------------------------------
' Procedure : Phyton_Send_Mail
' Author    : hsequeira
' Date      : 25/11/2023
' Purpose   : LLAMA AL SCRIPT DE PYTHON QUE ENVIA EL EMAIL
' Update    : LLAMA al EXE QUE ENVIA EL EMAIL
'---------------------------------------------------------------------------------------
''
'Private Function Phyton_Send_Mail(cNombreFile As String, cNombreEmpresa As String)
'
'Call WriteToINI("Opciones", "FileName", cNombreFile, App.Path & "\email.ini")
'Call WriteToINI("Opciones", "Empresa", cNombreEmpresa, App.Path & "\email.ini")
'
''retVal = ShellExecute(0&, "open", App.Path & "\SendMail.exe", "", "", vbHide)
'Call ShellExecute(0&, "open", App.Path & "\SendMail.exe", "", "", vbHide)
'
'End Function
