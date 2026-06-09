VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form RptCajas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REPORTE DE CAJEROS"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpenGaveta 
      Height          =   615
      Left            =   360
      Picture         =   "RptCajas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
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
      Height          =   615
      Left            =   5880
      TabIndex        =   9
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Frame frRZ 
      Caption         =   "REPORTES      Z"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5295
      Left            =   4200
      TabIndex        =   1
      Top             =   360
      Width           =   3375
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
         Left            =   240
         TabIndex        =   10
         Top             =   2400
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
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CommandButton cmdRZ 
         Caption         =   "REPORTE TERMINAL DE CAJEROS"
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
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.Frame frRX 
      Caption         =   "REPORTES     X"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   5295
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3375
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
         Left            =   240
         TabIndex        =   6
         Top             =   4320
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
         Left            =   240
         TabIndex        =   5
         Top             =   3360
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
         Left            =   240
         TabIndex        =   4
         Top             =   2400
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
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CommandButton cmdRX 
         Caption         =   "REPORTE DE CAJEROS"
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
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   2895
      End
   End
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   270
      Left            =   360
      TabIndex        =   11
      Top             =   5800
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "RptCajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsPluInvent As New ADODB.Recordset
Private Declare Function GetVolumeInformation Lib "Kernel32" Alias "GetVolumeInformationA" _
        (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, _
        ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, _
        lpMaximumComponentLength As Long, lpFileSystemFlags As Long, _
        ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
'DISCO,NOMBRE_VOLUMEN,255,NUM_SERIE,0,0,FILE_SYSTEM,255
'Private Declare Function BackDataUp Lib "SoloDll" ()
Private Sub BackDataUp()
Dim rsLoc As New ADODB.Recordset
Dim fld As ADODB.Field
Dim cCadena As String
Dim iCampo As Integer
Dim iControl As Integer
Dim nInt As Byte

On Error Resume Next
Open App.Path + "\" + "file" + ".1" For Output As #1
'1,7,14,15,16
iControl = 0

rsLoc.Open "SELECT * FROM TRANSAC", msConn, adOpenKeyset, adLockOptimistic
If rsLoc.RecordCount > 0 Then
    rsLoc.Fields.Refresh
    iCampo = rsLoc.Fields.Count
    Do Until rsLoc.EOF = True
        For Each fld In rsLoc.Fields
            Select Case fld.Name
                Case "CAJA", "DESCRIP", "FECHA", "HORA", "TIPO"
                    cCadena = cCadena & """" & fld.Value & ""","
                Case Else
                    If iControl = (iCampo - 1) Then
                        cCadena = cCadena & fld.Value
                    Else
                        cCadena = cCadena & fld.Value & ","
                    End If
            End Select
            iControl = iControl + 1
        Next
        Print #1, cCadena
        cCadena = ""
        iControl = 0
        rsLoc.MoveNext
    Loop
End If
rsLoc.Close
Close #1
On Error GoTo 0

'------------------------------------------------------
On Error Resume Next
Open App.Path + "\" + "file" + ".2" For Output As #1

rsLoc.Open "SELECT * FROM TRANSAC_CLI", msConn, adOpenKeyset, adLockOptimistic
If rsLoc.RecordCount > 0 Then
    rsLoc.Fields.Refresh
    iCampo = rsLoc.Fields.Count
    Do Until rsLoc.EOF = True
        For Each fld In rsLoc.Fields
            Select Case fld.Name
                Case "FECHA"
                    cCadena = cCadena & """" & fld.Value & ""","
                Case Else
                    If iControl = (iCampo - 1) Then
                        cCadena = cCadena & fld.Value
                    Else
                        cCadena = cCadena & fld.Value & ","
                    End If
            End Select
            iControl = iControl + 1
        Next
        Print #1, cCadena
        cCadena = ""
        iControl = 0
        rsLoc.MoveNext
    Loop
End If
rsLoc.Close
Close #1
On Error GoTo 0

'------------------------------------------------------
On Error Resume Next
Open App.Path + "\" + "file" + ".3" For Output As #1

rsLoc.Open "SELECT * FROM TRANSAC_PAGO", msConn, adOpenKeyset, adLockOptimistic
If rsLoc.RecordCount > 0 Then
    rsLoc.Fields.Refresh
    iCampo = rsLoc.Fields.Count
    Do Until rsLoc.EOF = True
        For Each fld In rsLoc.Fields
            If iControl = (iCampo - 1) Then
                cCadena = cCadena & fld.Value
            Else
                cCadena = cCadena & fld.Value & ","
            End If
            iControl = iControl + 1
        Next
        Print #1, cCadena
        cCadena = ""
        iControl = 0
        rsLoc.MoveNext
    Loop
End If
rsLoc.Close
Close #1
On Error GoTo 0

'------------------------------------------------------
On Error Resume Next
Open App.Path + "\" + "file" + ".4" For Output As #1

rsLoc.Open "SELECT * FROM TRANSAC_PROP", msConn, adOpenKeyset, adLockOptimistic
If rsLoc.RecordCount > 0 Then
    rsLoc.Fields.Refresh
    iCampo = rsLoc.Fields.Count
    Do Until rsLoc.EOF = True
        For Each fld In rsLoc.Fields
            If iControl = (iCampo - 1) Then
                cCadena = cCadena & fld.Value
            Else
                cCadena = cCadena & fld.Value & ","
            End If
            iControl = iControl + 1
        Next
        Print #1, cCadena
        cCadena = ""
        iControl = 0
        rsLoc.MoveNext
    Loop
End If
rsLoc.Close
Close #1
On Error GoTo 0

'---------------------------------------------------
On Error Resume Next
Open App.Path + "\" + "file" + ".5" For Output As #1

rsLoc.Open "SELECT * FROM ORGANIZACION", msConn, adOpenKeyset, adLockOptimistic
If rsLoc.RecordCount > 0 Then
    rsLoc.Fields.Refresh
    iCampo = rsLoc.Fields.Count
    Do Until rsLoc.EOF = True
        For Each fld In rsLoc.Fields
            Select Case fld.Name
                Case "RAZ_SOC", "DESCRIP", "DIRECCION", "RUC", "TELEFONO1", "TELEFONO2", "FAX", "MENSAJE", "CONTACTO", "SERIAL", "CLAVE"
                    cCadena = cCadena & """" & fld.Value & ""","
                Case Else
                    If iControl = (iCampo - 1) Then
                        cCadena = cCadena & fld.Value
                    Else
                        cCadena = cCadena & fld.Value & ","
                    End If
            End Select
            iControl = iControl + 1
        Next
        Print #1, cCadena
        cCadena = ""
        iControl = 0
        rsLoc.MoveNext
    Loop
End If
rsLoc.Close
Close #1
On Error GoTo 0

'---------------------------------------------------
On Error Resume Next
Open App.Path + "\" + "file" + ".6" For Output As #1

rsLoc.Open "SELECT * FROM CAJEROS", msConn, adOpenKeyset, adLockOptimistic
If rsLoc.RecordCount > 0 Then
    rsLoc.Fields.Refresh
    iCampo = rsLoc.Fields.Count
    Do Until rsLoc.EOF = True
        For Each fld In rsLoc.Fields
            Select Case fld.Name
                Case "CLAVE", "NOMBRE", "APELLIDO"
                    cCadena = cCadena & """" & fld.Value & ""","
                Case Else
                    If iControl = (iCampo - 1) Then
                        cCadena = cCadena & fld.Value
                    Else
                        cCadena = cCadena & fld.Value & ","
                    End If
            End Select
            iControl = iControl + 1
        Next
        Print #1, cCadena
        cCadena = ""
        iControl = 0
        rsLoc.MoveNext
    Loop
End If
rsLoc.Close
Close #1
On Error GoTo 0

Set rsLoc = Nothing
If FILE_OK = True Then
    On Error GoTo ErrCopy:
        Call FileCopy(App.Path + "\file.1", "a:\file.1")
        Call FileCopy(App.Path + "\file.2", "a:\file.2")
        Call FileCopy(App.Path + "\file.3", "a:\file.3")
        Call FileCopy(App.Path + "\file.4", "a:\file.4")
        Call FileCopy(App.Path + "\file.5", "a:\file.5")
        Call FileCopy(App.Path + "\file.6", "a:\file.6")
    On Error GoTo 0
End If
On Error Resume Next
    Kill (App.Path + "\file.1")
    Kill (App.Path + "\file.2")
    Kill (App.Path + "\file.3")
    Kill (App.Path + "\file.4")
    Kill (App.Path + "\file.5")
    Kill (App.Path + "\file.6")
On Error GoTo 0
Exit Sub

ErrCopy:
    MsgBox "ERROR AL COPIAR RESPALDOS" & vbCrLf & "SALGA E INTENTE CON OTRO DISQUETE", vbCritical, "VERIFIQUE QUE LA OPERACION"
    nInt = nInt + 1
    If nInt > 4 Then
        MsgBox "NO SE PUEDE HACER RESPALDO, CONSEGUIR OTRO DISQUETE", vbCritical, "OTRO DISQUETE POR FAVOR"
        Exit Sub
    Else
        Resume
    End If
End Sub
Private Sub DameVolumenA()
Dim Serial As Long, VName As String, FSName As String
'Rellena Buffer
VName = String$(255, Chr$(0))  'rellena con caracter ASCII
FSName = String$(255, Chr$(0)) 'rellena con caracter ASCII
On Error Resume Next
    Call GetVolumeInformation("A:\", VName, 255, Serial, 0, 0, FSName, 255)
    'DISCO,NOMBRE_VOLUMEN,255,NUM_SERIE,0,0,FILE_SYSTEM,255
    'Quita los extra chr$(0)'s
    VName = Left$(VName, InStr(1, VName, Chr$(0)) - 1)
    FSName = Left$(FSName, InStr(1, FSName, Chr$(0)) - 1)
    If VName <> "FIL" Then
        FILE_OK = False
    End If
On Error GoTo 0
End Sub

Private Sub MesasAbiertas()
Dim rsAbiertas As Recordset

Set rsAbiertas = New Recordset

rsAbiertas.Open "SELECT MESA,SUM(PRECIO) AS PEND FROM TMP_TRANS" & _
        " GROUP BY MESA", msConn, adOpenStatic, adLockOptimistic

If rsAbiertas.EOF Then
    rsAbiertas.Close
    Set rsAbiertas = Nothing
    Exit Sub
End If

Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, rs00!descrip & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Date & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "REPORTE DE MESAS ABIERTAS" & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "AL CAMBIO DE TURNO" & Chr(&HD) & Chr(&HA)
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
'''''''Printer.FontUnderline = True
Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, "Mesa    Pendiente" & Chr(&HD) & Chr(&HA)
'''''''Printer.FontUnderline = False
Do Until rsAbiertas.EOF
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, rsAbiertas!mesa & Space(8) & Format(rsAbiertas!PEND, "STANDARD") & Chr(&HD) & Chr(&HA)
    rsAbiertas.MoveNext
Loop

For i = 1 To 10
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Next
Sys_Pos.Coptr1.CutPaper 100

rsAbiertas.Close
End Sub
Private Sub RepDepX()
Dim rsTran As Recordset
Dim rsLocalDepto As Recordset
Dim nTotDepto As Single
Dim MiLen1 As Integer
Dim Milen2 As Integer

nTotDepto = 0#
Set rsTran = New Recordset
Set rsLocalDepto = New Recordset

'ABRE DEPARTAMENTOS
rsLocalDepto.Open "SELECT CODIGO,DESCRIP,CORTO FROM DEPTO", msConn, adOpenDynamic, adLockOptimistic

'"SELECT DEPTO,COUNT(DEPTO) AS X_COUNT,SUM(PRECIO) AS VALOR"
rsTran.Open "SELECT DEPTO,SUM(CANT) AS X_COUNT,SUM(PRECIO) AS VALOR " & _
        " FROM TRANSAC " & _
        "GROUP BY DEPTO ORDER BY DEPTO", msConn, adOpenStatic, adLockOptimistic

Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Date & Space(2) & Time, Date & Space(2) & Time
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "REPORTE DEPARTAMENTAL (X)", "REPORTE DEPARTAMENTAL (X)"
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, rs00!descrip, rs00!descrip
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "RUC:" & rs00!RUC, "RUC:" & rs00!RUC
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "SERIAL:" & rs00!Serial, "SERIAL:" & rs00!Serial
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CONTADOR TRANS : " & (rs00!TRANS + 1), "CONTADOR TRANS : " & (rs00!TRANS + 1)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CONTADOR X : " & (rs00!X_CDEP + 1), "CONTADOR X : " & (rs00!X_CDEP + 1)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "REPORTE DEPART.", "REPORTE DEPART."
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)

Do Until rsTran.EOF
    rsLocalDepto.MoveFirst
    rsLocalDepto.Find "CODIGO = " & rsTran!depto
    If Not rsLocalDepto.EOF Then
        MiLen1 = Len(rsTran!X_COUNT)
        Milen2 = Len(Format(rsTran!VALOR, "STANDARD"))
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, FormatTexto(rsLocalDepto!corto, 13) & Space(4 - MiLen1) & rsTran!X_COUNT & Space(9 - Milen2) & Format(rsTran!VALOR, "STANDARD"), FormatTexto(rsLocalDepto!corto, 13) & Space(4 - MiLen1) & rsTran!X_COUNT & Space(9 - Milen2) & Format(rsTran!VALOR, "STANDARD")
    Else
        MsgBox "EXISTE UN PROBLEMA CON LOS DATOS DEPARTAMENTALES", vbCritical, "REVISAR CINTA"
    End If
    nTotDepto = nTotDepto + rsTran!VALOR
    rsTran.MoveNext
Loop

Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
MiLen1 = Len(Format(nTotDepto, "CURRENCY"))
If MiLen1 > 11 Then MiLen1 = 11
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "TOTAL DEPTOS : " & Space(11 - MiLen1) & Format(nTotDepto, "CURRENCY"), "TOTAL DEPTOS : " & Space(11 - MiLen1) & Format(nTotDepto, "CURRENCY")
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
For i = 1 To 10
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Next
Sys_Pos.Coptr1.CutPaper 100
rsTran.Close

'Actualiza el conteo de X's y de TRANSAC
msConn.BeginTrans
msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1,X_CDEP = X_CDEP + 1"
msConn.CommitTrans

End Sub
Private Sub RepMeseroZ()
Dim rsMeseros As Recordset
Dim nValorMesas As Single
Dim MiLen1 As Integer, Milen2 As Integer

Set rsMeseros = New Recordset
nValorMesas = 0#

rsMeseros.Open "SELECT * FROM MESEROS WHERE VALOR > 0 " & _
        " AND NUMERO <> 999 ORDER BY NUMERO", msConn, adOpenStatic, adLockOptimistic
ProgBar.Value = 30
'Open "c:\mifile.txt" For Output As #1   ' Open file.

Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Date & Space(2) & Time, Date & Space(2) & Time
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "REPORTE MESEROS (Z)", "REPORTE MESEROS (Z)"
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, rs00!descrip, rs00!descrip
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "RUC:" & rs00!RUC, "RUC:" & rs00!RUC
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "SERIAL:" & rs00!Serial, "SERIAL:" & rs00!Serial
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
'''- SEGUN HACIENDA Y TESORO -'''Printer.Print "CONTADOR TRANS : " & (rs00!TRANS + 1)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CONTADOR Z : " & (rs00!Z_CMESEROS + 1), "CONTADOR Z : " & (rs00!Z_CMESEROS + 1)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
    
Do Until rsMeseros.EOF
    MiLen1 = Len(Format(rsMeseros!VALOR, "standard"))
    Milen2 = Len(rsMeseros!numero)
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "Mesero : " & Space(5 - Milen2) & rsMeseros!numero & Space(15 - MiLen1) & Format(rsMeseros!VALOR, "standard"), "Mesero : " & Space(5 - Milen2) & rsMeseros!numero & Space(15 - MiLen1) & Format(rsMeseros!VALOR, "standard")
    nValorMesas = nValorMesas + Format(rsMeseros!VALOR, "standard")
    rsMeseros.MoveNext
Loop
ProgBar.Value = 60
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Milen2 = Len(Format(nValorMesas, "CURRENCY"))
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "TOTAL : " & Space(21 - Milen2) & Format(nValorMesas, "CURRENCY"), "TOTAL : " & Space(21 - Milen2) & Format(nValorMesas, "CURRENCY")
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
'Close #1
For i = 1 To 10
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Next
Sys_Pos.Coptr1.CutPaper 100
ProgBar.Value = 100

msConn.BeginTrans
'''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1,X_CMESEROS = 0,Z_CMESEROS = Z_CMESEROS + 1"
msConn.Execute "UPDATE ORGANIZACION SET X_CMESEROS = 0,Z_CMESEROS = Z_CMESEROS + 1"
msConn.Execute "UPDATE MESEROS SET VALOR = 0.00, X_COUNT = 0, Z_COUNT = 0 " & _
               "WHERE NUMERO <> 999"
msConn.CommitTrans

ProgBar.Value = 0
End Sub
Private Sub RepMeseroX()
Dim rsMeseros As Recordset
Dim nValorMesas As Single
Dim MiLen1 As Integer
Dim Milen2 As Integer

Set rsMeseros = New Recordset
nValorMesas = 0#

rsMeseros.Open "SELECT * FROM MESEROS WHERE VALOR > 0 " & _
        " AND NUMERO <> 999 ORDER BY NUMERO", msConn, adOpenStatic, adLockOptimistic

'Open "c:\mifile.txt" For Output As #1   ' Open file.

Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Date & Space(2) & Time, Date & Space(2) & Time
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "REPORTE MESEROS (X)", "REPORTE MESEROS (X)"
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, rs00!descrip, rs00!descrip
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "RUC:" & rs00!RUC, "RUC:" & rs00!RUC
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "SERIAL:" & rs00!Serial, "SERIAL:" & rs00!Serial
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CONTADOR TRANS : " & (rs00!TRANS + 1), "CONTADOR TRANS : " & (rs00!TRANS + 1)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CONTADOR X : " & (rs00!X_CMESEROS + 1), "CONTADOR X : " & (rs00!X_CMESEROS + 1)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)

Do Until rsMeseros.EOF
    MiLen1 = Len(Format(rsMeseros!VALOR, "standard"))
    Milen2 = Len(rsMeseros!numero)
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "Mesero : " & Space(5 - Milen2) & rsMeseros!numero & Space(11 - MiLen1) & Format(rsMeseros!VALOR, "standard"), "Mesero : " & Space(5 - Milen2) & rsMeseros!numero & Space(11 - MiLen1) & Format(rsMeseros!VALOR, "standard")
    nValorMesas = nValorMesas + Format(rsMeseros!VALOR, "standard")
    rsMeseros.MoveNext
Loop
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Milen2 = Len(Format(nValorMesas, "CURRENCY"))
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "TOTAL : " & Space(17 - Milen2) & Format(nValorMesas, "CURRENCY"), "TOTAL : " & Space(17 - Milen2) & Format(nValorMesas, "CURRENCY")
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
'Close #1
For i = 1 To 10
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Next
Sys_Pos.Coptr1.CutPaper 100

msConn.BeginTrans
msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1, X_CMESEROS = X_CMESEROS + 1"
msConn.CommitTrans

End Sub
Private Sub RepMesZ()
Dim rsMesas As Recordset
Dim nValorMesas As Single
Dim MiLen1 As Integer
Dim Milen2 As Integer

Set rsMesas = New Recordset
nValorMesas = 0#

rsMesas.Open "SELECT * FROM MESAS WHERE VALOR > 0 AND NUMERO > 0" & _
        " ORDER BY NUMERO", msConn, adOpenStatic, adLockOptimistic
ProgBar.Value = 40
'Open "c:\mifile.txt" For Output As #1   ' Open file.

Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Date & Space(2) & Time, Date & Space(2) & Time
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "REPORTE DE MESAS (Z)", "REPORTE DE MESAS (Z)"
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, rs00!descrip, rs00!descrip
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "RUC:" & rs00!RUC, "RUC:" & rs00!RUC
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "SERIAL:" & rs00!Serial, "SERIAL:" & rs00!Serial
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
'''- SEGUN HACIENDA Y TESORO -'''Printer.Print "CONTADOR TRANS : " & (rs00!TRANS + 1)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CONTADOR Z : " & (rs00!Z_CMESAS + 1), "CONTADOR Z : " & (rs00!Z_CMESAS + 1)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)

Do Until rsMesas.EOF
    MiLen1 = Len(Format(rsMesas!VALOR, "standard"))
    Milen2 = Len(rsMesas!numero)
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "Mesa : " & Space(5 - Milen2) & rsMesas!numero & Space(15 - MiLen1) & Format(rsMesas!VALOR, "standard"), "Mesa : " & Space(5 - Milen2) & rsMesas!numero & Space(15 - MiLen1) & Format(rsMesas!VALOR, "standard")
    nValorMesas = nValorMesas + Format(rsMesas!VALOR, "standard")
    rsMesas.MoveNext
Loop
ProgBar.Value = 80
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Milen2 = Len(Format(nValorMesas, "CURRENCY"))
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "TOTAL : " & Space(19 - Milen2) & Format(nValorMesas, "CURRENCY"), "TOTAL : " & Space(19 - Milen2) & Format(nValorMesas, "CURRENCY")
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
'Close #1
For i = 1 To 10
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Next
Sys_Pos.Coptr1.CutPaper 100

ProgBar.Value = 100
msConn.BeginTrans
'''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1, X_CMESAS = 0, Z_CMESAS = Z_CMESAS + 1"
msConn.Execute "UPDATE ORGANIZACION SET X_CMESAS = 0, Z_CMESAS = Z_CMESAS + 1"
msConn.Execute "UPDATE MESAS SET VALOR = 0.00 WHERE NUMERO > 0"
msConn.CommitTrans
ProgBar.Value = 0
End Sub
Private Sub RepMesX()
Dim rsMesas As Recordset
Dim nValorMesas As Single
Dim MiLen1 As Integer
Dim Milen2 As Integer

Set rsMesas = New Recordset
nValorMesas = 0#

rsMesas.Open "SELECT * FROM MESAS WHERE VALOR > 0 AND NUMERO > 0" & _
        " ORDER BY NUMERO", msConn, adOpenStatic, adLockOptimistic

'Open "c:\mifile.txt" For Output As #1   ' Open file.

Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Date & Space(2) & Time, Date & Space(2) & Time
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "REPORTE DE MESAS (X)", "REPORTE DE MESAS (X)"
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, rs00!descrip, rs00!descrip
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "RUC:" & rs00!RUC, "RUC:" & rs00!RUC
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "SERIAL:" & rs00!Serial, "SERIAL:" & rs00!Serial
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CONTADOR TRANS : " & (rs00!TRANS + 1), "CONTADOR TRANS : " & (rs00!TRANS + 1)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CONTADOR X : " & (rs00!X_CMESAS + 1), "CONTADOR X : " & (rs00!X_CMESAS + 1)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
    
Do Until rsMesas.EOF
    MiLen1 = Len(Format(rsMesas!VALOR, "standard"))
    Milen2 = Len(rsMesas!numero)
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "Mesa : " & Space(5 - Milen2) & rsMesas!numero & Space(15 - MiLen1) & Format(rsMesas!VALOR, "standard"), "Mesa : " & Space(5 - Milen2) & rsMesas!numero & Space(15 - MiLen1) & Format(rsMesas!VALOR, "standard")
    nValorMesas = nValorMesas + Format(rsMesas!VALOR, "standard")
    rsMesas.MoveNext
Loop
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
MiLen1 = Len(Format(nValorMesas, "CURRENCY"))
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "TOTAL : " & Space(19 - MiLen1) & Format(nValorMesas, "CURRENCY"), "TOTAL : " & Space(19 - MiLen1) & Format(nValorMesas, "CURRENCY")
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
'Close #1
For i = 1 To 10
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Next
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
Dim nSumVta As Single
Dim nSumTotal As Single
Dim nTrans As Integer
Dim MiLen1 As Integer
Dim Milen2 As Integer
Dim nErrInd As Integer
Dim errCounter As Integer

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

'SELECCION DE DATOS DE INVENTARIO
rsPluInvent.Open "SELECT B.ID_PROD_INV, MAX(B.DESCRIP) AS ART_INVENT," & _
        "MAX(C.DESCRIP) AS UNID_CONSUMO,SUM(A.CANT) * MAX(B.CANT) AS CONSUMIDO " & _
        "FROM TRANSAC AS A, PLU_INVENT AS B,UNID_CONSUMO AS C " & _
        "WHERE A.CANT > 0 AND A.PLU = B.ID_PLU AND A.ENVASE = B.ID_ENV " & _
        "AND B.ID_UNID_CONSUMO = C.ID GROUP BY B.ID_PROD_INV", msConn, adOpenDynamic, adLockOptimistic


'ABRE DEPARTAMENTOS
rsLocalDepto.Open "SELECT CODIGO,DESCRIP,CORTO FROM DEPTO", msConn, adOpenDynamic, adLockOptimistic

'SELECCIONA TODOS LOS CAJEROS QUE HAN TRABAJADO EN EL PERIODO A REPORTAR
sqltext = "SELECT DISTINCT CAJERO FROM TRANSAC ORDER BY CAJERO"
rsCajeros.Open sqltext, msConn, adOpenStatic, adLockOptimistic

If rsCajeros.EOF Then
    MsgBox "IMPRIMIENDO EL REPORTE, AUNQUE NO EXISTEN TRANSACCIONES", vbInformation, BoxTit
'    Exit Sub
End If

MesasAbiertas

On Error Resume Next    'NUEVO PARA CORRER REPORTE EN ZZZZ
'SE USA PARA OBTENER EL CORRECTO RECORDCOUNT
rsCajeros.MoveFirst
rsCajeros.MoveLast

If rsCajeros.RecordCount = 0 Or rsCajeros!CAJERO = "" Then
    MsgBox "IMPRIMIENDO EL REPORTE, AUNQUE NO EXISTEN TRANSACCIONES", vbInformation, BoxTit
'    Exit Sub
End If
On Error GoTo 0         'NUEVO PARA CORRER REPORTE EN ZZZZ

'Open "c:\mifile.txt" For Output As #1   ' Open file.

On Error GoTo AjustaMilen:

ProgBar.Value = 5
sqltext = "SELECT * FROM pagos WHERE CODIGO <> 999 ORDER BY CODIGO"
RSPAGOS.Open sqltext, msConn, adOpenStatic, adLockOptimistic

ProgBar.Value = 10

rsCajeros.MoveFirst
Do Until rsCajeros.EOF

    ProgBar.Value = 20
    
    rsTran.Open "SELECT distinct NUM_TRANS FROM TRANSAC " & _
            " WHERE CAJERO = " & rsCajeros!CAJERO, msConn, adOpenStatic, adLockOptimistic
    
    rsTran.MoveFirst
    rsTran.MoveLast
    nTrans = rsTran.RecordCount
    rsTran.MoveFirst

    sqltext = "SELECT b.nombre,b.apellido,a.cajero,b.z_c, sum(a.precio) as Ventas " & _
            " FROM TRANSAC as a, cajeros as b " & _
            " WHERE a.cajero = " & rsCajeros!CAJERO & _
            " AND b.numero = " & rsCajeros!CAJERO & _
            " GROUP BY a.cajero,b.nombre,b.apellido,b.z_c "

    'VALOR EN VENTAS DEL CAJERO CON rsVta_Z
    rsVta_Z.Open sqltext, msConn, adOpenStatic, adLockOptimistic
    
    If rsVta_Z.RecordCount = 0 Then
        'MsgBox "EL CAJERO NO TIENE VENTAS, NO SE IMPRIMIRA REPORTE EN Z"
        rsCajeros.MoveNext
    End If
    
    'TODOS LOS PAGOS RECIBIDOS
    'ESTOY SACANDO EL DESCUENTO GLOBAL (99) DE AQUI PARA PONERLO EN
    'LOS AJUSTES
    sqltext = "SELECT a.cajero,a.tipo_pago,SUM(a.monto) AS Valor, " & _
            " COUNT(a.tipo_pago) as Z_COUNT FROM transac_pago as a " & _
            " WHERE a.cajero = " & rsCajeros!CAJERO & _
            " AND a.tipo_pago <> 99 " & _
            " GROUP BY a.cajero,a.TIPO_PAGO"
    rsPgo_Z.Open sqltext, msConn, adOpenStatic, adLockOptimistic
    
    'sqltext = "SELECT MID(a.TIPO,1,2) AS CORREC, COUNT(a.lin) as Z_COUNT, SUM(a.precio) as valor FROM TRANSAC as a WHERE A.CAJERO = " & rsCajeros!CAJERO & " AND MID(A.TIPO,1,1) <> ' ' GROUP BY MID(a.TIPO,1,2)"
    'sqltext = "SELECT MID(a.TIPO,1,2) AS CORREC,DESCUENTO, "
    '    " COUNT(a.lin) as Z_COUNT, SUM(a.precio) as valor "
    '2 de Nov 1999
    'COUNT(a.lin) as Z_COUNT, SUM(a.precio_unit) as valor
    '5 de Nov
    sqltext = "SELECT MID(a.TIPO,1,2) AS CORREC,a.DESCUENTO, " & _
        " COUNT(a.lin) as Z_COUNT, SUM(a.precio) as valor " & _
        " FROM TRANSAC as a " & _
        " WHERE A.CAJERO = " & rsCajeros!CAJERO & _
        " AND MID(A.TIPO,1,1) <> ' ' " & _
        " GROUP BY MID(a.TIPO,1,2),a.DESCUENTO"
    rsAjustes.Open sqltext, msConn, adOpenStatic, adLockOptimistic

    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Date & Space(2) & Time, Date & Space(2) & Time
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "REPORTE DE CAJEROS (Z)", "REPORTE DE CAJEROS (Z)"
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, rs00!descrip, rs00!descrip
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "RUC:" & rs00!RUC, "RUC:" & rs00!RUC
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "SERIAL:" & rs00!Serial, "SERIAL:" & rs00!Serial
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
    '''- SEGUN HACIENDA Y TESORO -'''Printer.Print "CONTADOR TRANS : " & (rs00!TRANS + 1)
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CONTADOR Z : " & (rsVta_Z!z_c + 1), "CONTADOR Z : " & (rsVta_Z!z_c + 1)
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CAJERO : " & rsVta_Z!nombre & ", " & rsVta_Z!apellido, "CAJERO : " & rsVta_Z!nombre & ", " & rsVta_Z!apellido
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
    
    Do Until rsVta_Z.EOF
        MiLen1 = Len(nTrans)
        Milen2 = Len(Format(rsVta_Z!VENTAS, "STANDARD"))
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "VENTA DEL DIA:" & Space(4 - MiLen1) & nTrans & Space(11 - Milen2) & Format(rsVta_Z!VENTAS, "STANDARD"), "VENTA DEL DIA:" & Space(4 - MiLen1) & nTrans & Space(11 - Milen2) & Format(rsVta_Z!VENTAS, "STANDARD")
        nSumTotal = nSumTotal + rsVta_Z!VENTAS
        rsVta_Z.MoveNext
    Loop

    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
    ''''''''''Printer.FontUnderline = True
    'Printer.Print "DESGLOSE DE INGRESOS"
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "TOTALES DE CAJA", "TOTALES DE CAJA"
    ''''''''Printer.FontUnderline = False
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)

    Do Until rsPgo_Z.EOF
        RSPAGOS.MoveFirst
        RSPAGOS.Find "CODIGO = " & rsPgo_Z!TIPO_PAGO
        If Not RSPAGOS.EOF Then
            MiLen1 = Len(rsPgo_Z!Z_COUNT)
            Milen2 = Len(Format(rsPgo_Z!VALOR, "STANDARD"))
            Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, FormatTexto(RSPAGOS!descrip, 13) & Space(4 - MiLen1) & rsPgo_Z!Z_COUNT & Space(13 - Milen2) & Format(rsPgo_Z!VALOR, "STANDARD"), FormatTexto(RSPAGOS!descrip, 13) & Space(4 - MiLen1) & rsPgo_Z!Z_COUNT & Space(13 - Milen2) & Format(rsPgo_Z!VALOR, "STANDARD")
        Else
            MiLen1 = 1
            Milen2 = Len(Format(0#, "standard"))
            Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "OTRO PAGO    " & Space(4 - MiLen1) & 0 & Space(13 - MiLen1) & Format(0#, "standard"), "OTRO PAGO    " & Space(4 - MiLen1) & 0 & Space(13 - MiLen1) & Format(0#, "standard")
        End If
        nSumVta = nSumVta + rsPgo_Z!VALOR
        rsPgo_Z.MoveNext
    Loop

    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
    MiLen1 = Len(Format(nSumVta, "currency"))
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "SUBTOTAL: " & Space(20 - MiLen1) & Format(nSumVta, "currency"), "SUBTOTAL: " & Space(20 - MiLen1) & Format(nSumVta, "currency")
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "------------------------------", "------------------------------"
    
    sqltext = "SELECT TIPO_PAGO,COUNT(TIPO_PAGO) AS Z_COUNT, " & _
            " SUM(MONTO) AS VALOR FROM TRANSAC_PROP " & _
            " WHERE CAJERO = " & rsCajeros!CAJERO & _
            " GROUP BY TIPO_PAGO "
    rsProp.Open sqltext, msConn, adOpenStatic, adLockOptimistic
    
    Do Until rsProp.EOF
        RSPAGOS.MoveFirst
        RSPAGOS.Find "CODIGO = " & rsProp!TIPO_PAGO
        If Not RSPAGOS.EOF Then
            MiLen1 = Len(rsProp!Z_COUNT)
            Milen2 = Len(Format(rsProp!VALOR, "STANDARD"))
            Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "Propina " & FormatTexto(RSPAGOS!descrip, 5) & Space(4 - MiLen1) & rsProp!Z_COUNT & Space(13 - Milen2) & Format(rsProp!VALOR, "STANDARD"), "Propina " & FormatTexto(RSPAGOS!descrip, 5) & Space(4 - MiLen1) & rsProp!Z_COUNT & Space(13 - Milen2) & Format(rsProp!VALOR, "STANDARD")
        End If
        rsProp.MoveNext
    Loop
    
    rsProp.Close
    
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
    '''''Printer.FontUnderline = True
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "AJUSTES", "AJUSTES"
    '''''Printer.FontUnderline = False
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
    
    Do Until rsAjustes.EOF
        MiLen1 = Len(rsAjustes!Z_COUNT)
        Milen2 = Len(Format(rsAjustes!VALOR, "STANDARD"))
        If rsAjustes!CORREC = "EC" Then
            Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CORECCION " & Space(8 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD"), "CORECCION " & Space(8 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD")
        ElseIf rsAjustes!CORREC = "VO" Then
            Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "ANULACION " & Space(8 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD"), "ANULACION " & Space(8 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD")
        ElseIf rsAjustes!CORREC = "DC" Then
            Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "DESCUENTO " & Format(rsAjustes!DESCUENTO, "0.00") & Space(4 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD"), "DESCUENTO " & Format(rsAjustes!DESCUENTO, "0.00") & Space(4 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD")
        End If
        rsAjustes.MoveNext
    Loop

    'PREPARA INFO PARA DESCUENTO GLOBAL
    rsAjustes.Close
    sqltext = "SELECT a.cajero,a.tipo_pago,SUM(a.monto) AS Valor, " & _
            " COUNT(a.tipo_pago) as Z_COUNT " & _
            " FROM transac_pago as a " & _
            " WHERE a.cajero = " & rsCajeros!CAJERO & _
            " AND a.tipo_pago = 99 " & _
            " GROUP BY a.cajero,a.TIPO_PAGO"
    
    rsAjustes.Open sqltext, msConn, adOpenStatic, adLockOptimistic
    
    If Not rsAjustes.EOF Then
        RSPAGOS.MoveFirst
        RSPAGOS.Find "CODIGO = " & rsAjustes!TIPO_PAGO
        If Not RSPAGOS.EOF Then
            MiLen1 = Len(rsAjustes!Z_COUNT)
            Milen2 = Len(Format(rsAjustes!VALOR, "STANDARD"))
            Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, FormatTexto(RSPAGOS!descrip, 13) & Space(5 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD"), FormatTexto(RSPAGOS!descrip, 13) & Space(5 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD")
        End If
    End If
    
    msConn.BeginTrans
    '''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
    'Incrementa el contador en Z y Resetea el de X a 0
    msConn.Execute "UPDATE CAJEROS SET Z_C = Z_C + 1, X_C = 0 " & _
                   " WHERE NUMERO = " & rsCajeros!CAJERO
    msConn.CommitTrans
    
    rsCajeros.MoveNext
    If Not rsCajeros.EOF = True Then
        For i = 1 To 10
            Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
        Next
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
ProgBar.Value = 40
'-------------------   EL CAJERO TERMINAL  -------------------

For i = 1 To 10
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Next
Sys_Pos.Coptr1.CutPaper 100

Set rsLocTerminal = New Recordset
rsLocTerminal.Open "SELECT Z_C FROM CAJEROS WHERE NUMERO = 999", msConn, adOpenStatic, adLockOptimistic

rsTran.Open "SELECT distinct NUM_TRANS FROM TRANSAC", msConn, adOpenStatic, adLockOptimistic

nTrans = rsTran.RecordCount

sqltext = "SELECT sum(a.precio) as Ventas FROM TRANSAC as a "
'VALOR EN VENTAS DEL CAJERO CON rsVta_Z
rsVta_Z.Open sqltext, msConn, adOpenStatic, adLockOptimistic

'''- SEGUN HACIENDA Y TESORO -'''rsHash.Open "SELECT SUM(ABS(A.PRECIO)) AS HASH_DIA FROM TRANSAC AS A", msConn, adOpenStatic, adLockOptimistic
rsHash.Open "SELECT SUM(A.PRECIO) AS HASH_DIA " & _
    " FROM TRANSAC AS A " & _
    " WHERE A.PRECIO > 0 ", msConn, adOpenStatic, adLockOptimistic

'TODOS LOS PAGOS RECIBIDOS
'MENOS EL DESCUENTO GLOBAL
'sqltext = "SELECT a.tipo_pago,SUM(a.monto) AS Valor, " & _
        " COUNT(a.tipo_pago) as Z_COUNT " & _
        " FROM transac_pago as a " & _
        " GROUP BY a.TIPO_PAGO"
        
sqltext = "SELECT a.tipo_pago,SUM(a.monto) AS Valor, " & _
        " COUNT(a.tipo_pago) as Z_COUNT " & _
        " FROM transac_pago as a " & _
        " WHERE a.tipo_pago <> 99 " & _
        " GROUP BY a.TIPO_PAGO"
rsPgo_Z.Open sqltext, msConn, adOpenStatic, adLockOptimistic

'sqltext = "SELECT MID(a.TIPO,1,2) AS CORREC,DESCUENTO, "
        '" COUNT(a.lin) as Z_COUNT, SUM(a.precio) as valor "
'2 de Nov 1999
'COUNT(a.lin) as Z_COUNT, SUM(a.precio_unit) as valor
'5 de Nov
sqltext = "SELECT MID(a.TIPO,1,2) AS CORREC,DESCUENTO, " & _
        " COUNT(a.lin) as Z_COUNT, SUM(a.precio) as valor " & _
        " FROM TRANSAC as a WHERE MID(A.TIPO,1,1) <> ' ' " & _
        " GROUP BY MID(a.TIPO,1,2),DESCUENTO"

rsAjustes.Open sqltext, msConn, adOpenStatic, adLockOptimistic

Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Date & Space(2) & Time, Date & Space(2) & Time
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "REPORTE DE TERMINAL (Z)", "REPORTE DE TERMINAL (Z)"
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, rs00!descrip, rs00!descrip
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "RUC:" & rs00!RUC, "RUC:" & rs00!RUC
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "SERIAL:" & rs00!Serial, "SERIAL:" & rs00!Serial
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
'''- SEGUN HACIENDA Y TESORO -'''Printer.Print "CONTADOR TRANS : " & (rs00!TRANS + 1)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CONTADOR Z : " & (rsLocTerminal!z_c + 1), "CONTADOR Z : " & (rsLocTerminal!z_c + 1)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CAJERO : REPORTE/TERMINAL", "CAJERO : REPORTE/TERMINAL"
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)

iISCTransaccion = 0
Do Until rsVta_Z.EOF
    MiLen1 = Len(nTrans)
    Milen2 = Len(Format(rsVta_Z!VENTAS, "STANDARD"))
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "VENTA DEL DIA:" & Space(5 - MiLen1) & nTrans & Space(10 - Milen2) & Format(rsVta_Z!VENTAS, "STANDARD"), "VENTA DEL DIA:" & Space(5 - MiLen1) & nTrans & Space(10 - Milen2) & Format(rsVta_Z!VENTAS, "STANDARD")
    nSumTotal = nSumTotal + IIf(IsNull(rsVta_Z!VENTAS), 0, rsVta_Z!VENTAS)
    rsVta_Z.MoveNext
Loop

ProgBar.Value = 45
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
''''''''Printer.FontUnderline = True
'Printer.Print "DESGLOSE DE INGRESOS"
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "TOTALES DE CAJA", "TOTALES DE CAJA"
''''''''Printer.FontUnderline = False
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)

Do Until rsPgo_Z.EOF
    RSPAGOS.MoveFirst
    RSPAGOS.Find "CODIGO = " & rsPgo_Z!TIPO_PAGO
    If Not RSPAGOS.EOF Then
        MiLen1 = Len(rsPgo_Z!Z_COUNT)
        Milen2 = Len(Format(rsPgo_Z!VALOR, "STANDARD"))
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, FormatTexto(RSPAGOS!descrip, 13) & Space(4 - MiLen1) & rsPgo_Z!Z_COUNT & Space(13 - Milen2) & Format(rsPgo_Z!VALOR, "STANDARD"), FormatTexto(RSPAGOS!descrip, 13) & Space(4 - MiLen1) & rsPgo_Z!Z_COUNT & Space(13 - Milen2) & Format(rsPgo_Z!VALOR, "STANDARD")
        'PRINT#1, rsPagos!descrip & Chr(9) & rsPgo_Z!Z_COUNT & Chr(9) & Format(rsPgo_Z!VALOR, "STANDARD")
    Else
        MiLen1 = 1
        Milen2 = Len(Format(0#, "standard"))
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "OTRO PAGO    " & Space(4 - MiLen1) & 0 & Space(13 - MiLen1) & Format(0#, "standard"), "OTRO PAGO    " & Space(4 - MiLen1) & 0 & Space(13 - MiLen1) & Format(0#, "standard")
        'PRINT#1, rsPagos!descrip & Chr(9) & 0 & Chr(9) & Format(0#, "standard")
    End If
    nSumVta = nSumVta + rsPgo_Z!VALOR
    rsPgo_Z.MoveNext
Loop

ProgBar.Value = 50
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
MiLen1 = Len(Format(nSumVta, "currency"))
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "SUBTOTAL: " & Space(20 - MiLen1) & Format(nSumVta, "currency"), "SUBTOTAL: " & Space(20 - MiLen1) & Format(nSumVta, "currency")
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "------------------------------", "------------------------------"
'*********************************************
'*********************************************
On Error Resume Next
rsISC.Open "SELECT * FROM ISC", msConn, adOpenDynamic, adLockOptimistic
MiLen1 = Len(Format(rsISC!DIARIO, "CURRENCY"))
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "I.S.C. (5%):" & Space(18 - MiLen1) & Format(rsISC!DIARIO, "CURRENCY"), "I.S.C. (5%):" & Space(18 - MiLen1) & Format(rsISC!DIARIO, "CURRENCY")
rsISC.Fields(Month(Date)).Value = rsISC.Fields(Month(Date)).Value + rsISC!DIARIO
rsISC!DIARIO = 0#
rsISC.Update
iISCTransaccion = rsISC.Fields(Month(Date)).Value
rsISC.Close
On Error GoTo 0
'*********************************************
'*********************************************
sqltext = "SELECT TIPO_PAGO,COUNT(TIPO_PAGO) AS Z_COUNT, " & _
        " SUM(MONTO) AS VALOR FROM TRANSAC_PROP " & _
        " GROUP BY TIPO_PAGO "
rsProp.Open sqltext, msConn, adOpenStatic, adLockOptimistic

Do Until rsProp.EOF
    RSPAGOS.MoveFirst
    RSPAGOS.Find "CODIGO = " & rsProp!TIPO_PAGO
    If Not RSPAGOS.EOF Then
        MiLen1 = Len(rsProp!Z_COUNT)
        Milen2 = Len(Format(rsProp!VALOR, "STANDARD"))
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "Propina " & FormatTexto(RSPAGOS!descrip, 5) & Space(4 - MiLen1) & rsProp!Z_COUNT & Space(13 - Milen2) & Format(rsProp!VALOR, "STANDARD"), "Propina " & FormatTexto(RSPAGOS!descrip, 5) & Space(4 - MiLen1) & rsProp!Z_COUNT & Space(13 - Milen2) & Format(rsProp!VALOR, "STANDARD")
        'PRINT#1, "Propina " & Mid(rsPagos!descrip, 1, 5) & Chr(9) & rsProp!X_COUNT & Chr(9) & Format(rsProp!VALOR, "STANDARD")
    End If
    rsProp.MoveNext
Loop
rsProp.Close

ProgBar.Value = 60
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
'''''''''Printer.FontUnderline = True
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "AJUSTES", "AJUSTES"
''''''''Printer.FontUnderline = False
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)

Do Until rsAjustes.EOF
    MiLen1 = Len(rsAjustes!Z_COUNT)
    Milen2 = Len(Format(rsAjustes!VALOR, "STANDARD"))
    If rsAjustes!CORREC = "EC" Then
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CORECCION " & Space(8 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD"), "CORECCION " & Space(8 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD")
    ElseIf rsAjustes!CORREC = "VO" Then
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "ANULACION " & Space(8 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD"), "ANULACION " & Space(8 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD")
    ElseIf rsAjustes!CORREC = "DC" Then
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "DESCUENTO " & Format(rsAjustes!DESCUENTO, "0.00") & Space(4 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD"), "DESCUENTO " & Format(rsAjustes!DESCUENTO, "0.00") & Space(4 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD")
    End If
    rsAjustes.MoveNext
Loop

'PREPARA INFO PARA DESCUENTO GLOBAL
rsAjustes.Close
sqltext = "SELECT a.cajero,a.tipo_pago,SUM(a.monto) AS Valor, " & _
        " COUNT(a.tipo_pago) as Z_COUNT " & _
        " FROM transac_pago as a " & _
        " WHERE a.tipo_pago = 99 " & _
        " GROUP BY a.cajero,a.TIPO_PAGO"

rsAjustes.Open sqltext, msConn, adOpenStatic, adLockOptimistic

If Not rsAjustes.EOF Then
    RSPAGOS.MoveFirst
    RSPAGOS.Find "CODIGO = " & rsAjustes!TIPO_PAGO
    If Not RSPAGOS.EOF Then
        MiLen1 = Len(rsAjustes!Z_COUNT)
        Milen2 = Len(Format(rsAjustes!VALOR, "STANDARD"))
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, FormatTexto(RSPAGOS!descrip, 13) & Space(5 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD"), FormatTexto(RSPAGOS!descrip, 13) & Space(5 - MiLen1) & rsAjustes!Z_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD")
    End If
End If

ProgBar.Value = 70

ssVtatot = rs00!VTA_TOT
ssHashTot = rs00!tot_hash

msConn.BeginTrans
'''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1, " & _
        " VTA_TOT = VTA_TOT + " & nSumTotal & _
        ", TOT_HASH = TOT_HASH + " & IIf(IsNull(rsHash!HASH_DIA), 0, rsHash!HASH_DIA)
msConn.Execute "UPDATE ORGANIZACION SET VTA_TOT = VTA_TOT + " & nSumTotal & _
        ", TOT_HASH = TOT_HASH + " & IIf(IsNull(rsHash!HASH_DIA), 0, rsHash!HASH_DIA)
msConn.Execute "UPDATE CAJEROS SET Z_C = Z_C + 1, X_C = 0 " & _
        " WHERE NUMERO = 999"
msConn.CommitTrans

rsVta_Z.Close
rsPgo_Z.Close
rsTran.Close
rsAjustes.Close

rs00.Requery    'DESPUES DEL COMMIT, LOS DATOS DEBEN DE ESTAR EN EL SERVIDOR

Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "VENTAS ACUMULADAS", "VENTAS ACUMULADAS"
MiLen1 = Len(Format(nSumTotal, "currency"))
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "HOY        : " & Format(nSumTotal, "currency"), "HOY        : " & Format(nSumTotal, "currency")
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "TOTAL ANT. : " & Format(ssVtatot, "currency"), "TOTAL ANT. : " & Format(ssVtatot, "currency")
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "TOTAL NUEVO: " & Format(rs00!VTA_TOT, "currency"), "TOTAL NUEVO: " & Format(rs00!VTA_TOT, "currency")
'*********************************************
On Error Resume Next
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "   I.S.C.  : " & Format(iISCTransaccion, "CURRENCY"), "   I.S.C.  : " & Format(iISCTransaccion, "CURRENCY")
iISCTransaccion = 0#
On Error GoTo 0
'*********************************************
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
'Printer.Print "HASH ANT.  : " & Format(ssHashTot, "currency")
'--- SEGUN HACIENDA Y TESORO -''Printer.Print "HASH NUEVO : " & Format(rs00!tot_hash, "currency")
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "GRAN TOTAL  : " & Format(rs00!tot_hash, "currency"), "GRAN TOTAL  : " & Format(rs00!tot_hash, "currency")

For i = 1 To 10
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Next
Sys_Pos.Coptr1.CutPaper 100

'------------- DEPARTAMENTAL --------------

ProgBar.Value = 80
'rsTran.Open "SELECT DEPTO,COUNT(DEPTO) AS X_COUNT,SUM(PRECIO) AS VALOR "
rsTran.Open "SELECT DEPTO,SUM(CANT) AS X_COUNT,SUM(PRECIO) AS VALOR " & _
        " FROM TRANSAC " & _
        " GROUP BY DEPTO " & _
        " ORDER BY DEPTO", msConn, adOpenStatic, adLockOptimistic

'''- SEGUN HACIENDA Y TESORO -'''msConn.BeginTrans
'''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
'''- SEGUN HACIENDA Y TESORO -'''msConn.CommitTrans

Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Date & Space(2) & Time, Date & Space(2) & Time
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "REPORTE DEPARTAMENTAL (Z)", "REPORTE DEPARTAMENTAL (Z)"
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, rs00!descrip, rs00!descrip
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "RUC:" & rs00!RUC, "RUC:" & rs00!RUC
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "SERIAL:" & rs00!Serial, "SERIAL:" & rs00!Serial
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
'''- SEGUN HACIENDA Y TESORO -'''Printer.Print "CONTADOR TRANS : " & (rs00!TRANS + 1)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CONTADOR Z : " & (rs00!Z_CDEP + 1), "CONTADOR Z : " & (rs00!Z_CDEP + 1)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "DEPART. REPORTE/TERMINAL", "DEPART. REPORTE/TERMINAL"
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)

Dim nTotDepto As Single

nTotDepto = 0#

Do Until rsTran.EOF
    On Error Resume Next
    rsLocalDepto.MoveFirst
    On Error GoTo 0
    rsLocalDepto.Find "CODIGO = " & rsTran!depto
    If Not rsLocalDepto.EOF Then
        MiLen1 = Len(rsTran!X_COUNT)
        Milen2 = Len(Format(rsTran!VALOR, "STANDARD"))
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, FormatTexto(rsLocalDepto!corto, 13) & Space(4 - MiLen1) & rsTran!X_COUNT & Space(9 - Milen2) & Format(rsTran!VALOR, "STANDARD"), FormatTexto(rsLocalDepto!corto, 13) & Space(4 - MiLen1) & rsTran!X_COUNT & Space(9 - Milen2) & Format(rsTran!VALOR, "STANDARD")
    Else
        MsgBox "POSIBLE ERROR EN DEPARTAMENTO ", vbCritical, "REVISAR CINTA"
    End If
    nTotDepto = nTotDepto + rsTran!VALOR
    rsTran.MoveNext
Loop

Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
MiLen1 = Len(Format(nTotDepto, "CURRENCY"))
If MiLen1 > 11 Then MiLen1 = 11
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "TOTAL DEPTOS :" & Space(11 - MiLen1) & Format(nTotDepto, "CURRENCY"), "TOTAL DEPTOS :" & Space(11 - MiLen1) & Format(nTotDepto, "CURRENCY")
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)

For i = 1 To 10
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Next
Sys_Pos.Coptr1.CutPaper 100

rsTran.Close
ProgBar.Value = 90

msConn.BeginTrans
'Tambien actualiza el conteo de Z's y X's
'''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1,X_CDEP = 0,Z_CDEP = Z_CDEP + 1"
msConn.Execute "UPDATE ORGANIZACION SET X_CDEP = 0,Z_CDEP = Z_CDEP + 1"
msConn.CommitTrans

'--------------- SUPER GRUPOS ------------------------

Dim ccc As String

Set rsSuperGrp = New Recordset

ccc = "SELECT A.GRUPO,A.DESCRIP,SUM(C.PRECIO) AS VENTAS" & _
    " FROM SUPER_GRP AS A,SUPER_DET AS B, TRANSAC AS C" & _
    " Where A.GRUPO = B.GRUPO And B.DEPTO = C.DEPTO " & _
    " GROUP BY A.GRUPO,A.DESCRIP" & _
    " ORDER BY A.DESCRIP"

rsSuperGrp.Open ccc, msConn, adOpenStatic, adLockOptimistic

'''- SEGUN HACIENDA Y TESORO -'''msConn.BeginTrans
'''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
'''- SEGUN HACIENDA Y TESORO -'''msConn.CommitTrans

Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Date & Space(2) & Time, Date & Space(2) & Time
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "REPORTE DE GRUPOS (Z)", "REPORTE DE GRUPOS (Z)"
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, rs00!descrip, rs00!descrip
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "RUC:" & rs00!RUC, "RUC:" & rs00!RUC
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "SERIAL:" & rs00!Serial, "SERIAL:" & rs00!Serial
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
'''- SEGUN HACIENDA Y TESORO -'''Printer.Print "CONTADOR TRANS : " & (rs00!TRANS + 1)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CONTADOR Z : " & (rs00!Z_CDEP + 1), "CONTADOR Z : " & (rs00!Z_CDEP + 1)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "DEPART. GRUPO/TERMINAL", "DEPART. GRUPO/TERMINAL"
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)

nTotDepto = 0#

Do Until rsSuperGrp.EOF
    Milen2 = Len(Format(rsSuperGrp!VENTAS, "STANDARD"))
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, FormatTexto(rsSuperGrp!descrip, 13) & Space(13 - Milen2) & Format(rsSuperGrp!VENTAS, "STANDARD"), FormatTexto(rsSuperGrp!descrip, 13) & Space(13 - Milen2) & Format(rsSuperGrp!VENTAS, "STANDARD")
    nTotDepto = nTotDepto + rsSuperGrp!VENTAS
    rsSuperGrp.MoveNext
Loop

Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
MiLen1 = Len(Format(nTotDepto, "CURRENCY"))
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "TOTAL GRUPOS:" & Space(13 - MiLen1) & Format(nTotDepto, "CURRENCY"), "TOTAL GRUPOS:" & Space(13 - MiLen1) & Format(nTotDepto, "CURRENCY")
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)

For i = 1 To 10
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Next
Sys_Pos.Coptr1.CutPaper 100

rsSuperGrp.Close

'''- SEGUN HACIENDA Y TESORO -'''msConn.BeginTrans
'''- SEGUN HACIENDA Y TESORO -'''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
'''- SEGUN HACIENDA Y TESORO -'''msConn.CommitTrans

'--------------- FIN DE SUPER GRUPOS ------------------------

'----------------------------------------
rsTran.Open "SELECT * FROM TRANSAC", msConn, adOpenStatic, adLockOptimistic

On Error GoTo ErrorZZZ:

msConn.BeginTrans

Do Until rsTran.EOF
    If FILE_OK = False Then
        'NO HAY DISQUETE ASI QUE HACER HISTORICO
        CadenaSql = "INSERT INTO HIST_TR " & _
            "(NUM_TRANS,CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO) VALUES (" & _
            "" & rsTran!NUM_TRANS & "," & rsTran!caja & "," & rsTran!CAJERO & "," & rsTran!mesa & "," & rsTran!mesero & "," & rsTran!VALID & "," & rsTran!LIN & "," & "'" & _
            rsTran!descrip & "'" & "," & rsTran!cant & "," & rsTran!depto & "," & rsTran!PLU & "," & _
            rsTran!envase & "," & rsTran!precio_unit & "," & rsTran!precio & "," & "'" & rsTran!FECHA & "'" & "," & "'" & rsTran!Hora & "'" & _
            ",'" & rsTran!TIPO & "'," & rsTran!DESCUENTO & ")"
        msConn.Execute CadenaSql
    End If
    rsTran.MoveNext
Loop

msConn.Execute "DELETE * FROM TRANSAC"
msConn.CommitTrans
rsTran.Close

'----------------------------------------
rsTran.Open "SELECT * FROM TRANSAC_PAGO", msConn, adOpenStatic, adLockOptimistic

msConn.BeginTrans

Do Until rsTran.EOF
    If FILE_OK = False Then
        sqltext = "INSERT INTO HIST_TR_PAGO (NUM_TRANS,TIPO_PAGO,CAJERO,LIN,MONTO) " & _
            " VALUES (" & rsTran!NUM_TRANS & "," & rsTran!TIPO_PAGO & "," & _
            rsTran!CAJERO & "," & rsTran!LIN & "," & Format(rsTran!MONTO, "#0.00") & ")"
        msConn.Execute sqltext
    End If
    rsTran.MoveNext
Loop

msConn.Execute "DELETE * FROM TRANSAC_PAGO"
msConn.CommitTrans
rsTran.Close

'-------------LIMPIA PROPINAS------------------------
rsTran.Open "SELECT * FROM TRANSAC_PROP", msConn, adOpenStatic, adLockOptimistic

msConn.BeginTrans
Do Until rsTran.EOF
    If FILE_OK = False Then
        sqltext = "INSERT INTO HIST_TR_PROP " & _
            " (NUM_TRANS,MESERO,CAJERO,LIN,TIPO_PAGO,MONTO) " & _
            " VALUES (" & _
            rsTran!NUM_TRANS & "," & rsTran!mesero & "," & rsTran!CAJERO & "," & _
            rsTran!LIN & "," & rsTran!TIPO_PAGO & "," & Format(rsTran!MONTO, "#0.00") & ")"
    
        msConn.Execute sqltext
    End If
    rsTran.MoveNext
Loop

msConn.Execute "DELETE * FROM TRANSAC_PROP"
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
    If FILE_OK = False Then
        sqltext = "INSERT INTO HIST_TR_CLI " & _
            " (CODIGO_TP,CODIGO_CLI,NUM_TRANS,MONTO,FECHA,STATUS,RECIBIDO,SALDO,TIPO_TRANS,COMMENT) VALUES (" & _
            rsTran!CODIGO_TP & "," & rsTran!CODIGO_CLI & "," & rsTran!NUM_TRANS & "," & _
            Format(rsTran!MONTO, "#0.00") & "," & rsTran!FECHA & "," & _
            iStat & "," & Format(rsTran!RECIBIDO, "#0.00") & "," & _
            Format(rsTran!MONTO - rsTran!RECIBIDO, "#0.00") & _
            ",'FA','Factura de Ventas')"

        msConn.Execute sqltext
    End If
    rsTran.MoveNext
Loop

msConn.Execute "DELETE * FROM TRANSAC_CLI"
msConn.CommitTrans
rsTran.Close
ProgBar.Value = 100
'Close #1

On Error GoTo 0
If nErrInd = 0 Then
    ''''''''''''''''''''If ON_LINE = True Then BorraLocal
    MsgBox "REPORTE EN 'Z' DE TERMINAL ESTA LISTO", vbInformation, BoxTit
Else
    MsgBox "HA OCURRIDO MAS DE UN ERROR EN EL REPORTE (Z). CONTACTE A SOLO SOFTWARE", vbCritical, BoxTit
End If
ProgBar.Value = 0

Call ActualizaInvent ' ACTUALIZACION DE INVENTARIO

Exit Sub

AjustaMilen:
errCounter = errCounter + 1
Milen2 = 11
If errCounter < 3 Then
    Resume
Else
    MsgBox "EXISTE UN PROMLEMA DE IMPRESION. UNA VEZ TERMINADO REVISE EL LISTADO Y VERIFIQUE LOS DATOS", vbCritical, "LA IMPRESION DEL REPORTE TIENE PROBLEMAS"
    Resume Next
End If

ErrorZZZ:
Dim ADOError As Error
For Each ADOError In msConn.Errors
    sError = sError & ADOError.Number & " - " & ADOError.Description + vbCrLf
Next ADOError
MsgBox "ERROR EN EL REPORTE Z.ANOTE EL NUMERO/DESCRIPCION Y CONTACTE A SOLO SOFTWARE", vbCritical
MsgBox sError, vbCritical, "MENSAJE DE ERROR"
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
Dim nSumVta As Single
Dim nSumTotal As Single
Dim nTrans As Integer
Dim MiLen1 As Integer
Dim Milen2 As Integer
Dim errCounter As Integer

On Error GoTo AjustaMilen:
nSumVta = 0
nSumTotal = 0

Set rsVta_X = New Recordset
Set rsPgo_X = New Recordset
Set rsTran = New Recordset
Set RSPAGOS = New Recordset
Set rsAjustes = New Recordset
Set rsCajeros = New Recordset
Set rsProp = New Recordset

sqltext = "SELECT DISTINCT CAJERO FROM TRANSAC ORDER BY CAJERO"
rsCajeros.Open sqltext, msConn, adOpenStatic, adLockOptimistic

Call MesasAbiertas

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

ProgBar.Value = 10
Do Until rsCajeros.EOF

    rsTran.Open "SELECT distinct NUM_TRANS FROM TRANSAC WHERE CAJERO = " & rsCajeros!CAJERO, msConn, adOpenStatic, adLockOptimistic

    nTrans = rsTran.RecordCount

    sqltext = "SELECT b.nombre,b.apellido,a.cajero,b.x_c, sum(a.precio) as Ventas " & _
            " FROM TRANSAC as a, cajeros as b " & _
            " WHERE a.cajero = " & rsCajeros!CAJERO & _
            " AND b.numero = " & rsCajeros!CAJERO & _
            " GROUP BY a.cajero,b.nombre,b.apellido,b.x_c "
            
    'VALOR EN VENTAS DEL CAJERO CON rsVta_X
    rsVta_X.Open sqltext, msConn, adOpenStatic, adLockOptimistic
    
    If rsVta_X.RecordCount = 0 Then
        'MsgBox "EL CAJERO NO TIENE VENTAS, NO SE IMPRIMIRA REPORTE EN X"
        rsCajeros.MoveNext
    End If
    ProgBar.Value = 20
    'TODOS LOS PAGOS RECIBIDOS
    ' SACANDO LOS DESCUENTOS GLOBALES PARA PONERLOS
    ' EN LOS AJUSTES
    sqltext = "SELECT a.cajero,a.tipo_pago,SUM(a.monto) AS Valor, " & _
            " COUNT(a.tipo_pago) as X_COUNT FROM transac_pago as a " & _
            " WHERE a.cajero = " & rsCajeros!CAJERO & _
            " AND a.tipo_pago <> 99 " & _
            " GROUP BY a.cajero,a.TIPO_PAGO"

    rsPgo_X.Open sqltext, msConn, adOpenStatic, adLockOptimistic

    'sqltext = "SELECT MID(a.TIPO,1,2) AS CORREC,DESCUENTO, "
        '" COUNT(a.lin) as X_COUNT, SUM(a.precio) as valor "
    ' 2 de Nov 1999
    sqltext = "SELECT MID(a.TIPO,1,2) AS CORREC,DESCUENTO, " & _
        " COUNT(a.lin) as X_COUNT, SUM(a.PRECIO) as valor " & _
        " FROM TRANSAC as a WHERE A.CAJERO = " & rsCajeros!CAJERO & _
        " AND MID(A.TIPO,1,1) <> ' ' " & _
        " GROUP BY MID(a.TIPO,1,2),DESCUENTO"

    rsAjustes.Open sqltext, msConn, adOpenStatic, adLockOptimistic
    
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Date & Space(2) & Time, Date & Space(2) & Time
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "REPORTE DE CAJEROS (X)", "REPORTE DE CAJEROS (X)"
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, rs00!descrip, rs00!descrip
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "RUC:" & rs00!RUC, "RUC:" & rs00!RUC
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "SERIAL:" & rs00!Serial, "SERIAL:" & rs00!Serial
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CONTADOR TRANS : " & (rs00!TRANS + 1), "CONTADOR TRANS : " & (rs00!TRANS + 1)
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CONTADOR X : " & (rsVta_X!x_c + 1), "CONTADOR X : " & (rsVta_X!x_c + 1)
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CAJERO : " & rsVta_X!nombre & ", " & Mid(rsVta_X!apellido, 1, 1) & ".", "CAJERO : " & rsVta_X!nombre & ", " & Mid(rsVta_X!apellido, 1, 1) & "."
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)

    Do Until rsVta_X.EOF
        MiLen1 = Len(nTrans)
        Milen2 = Len(Format(rsVta_X!VENTAS, "STANDARD"))
        Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "VENTA DEL DIA:" & Space(4 - MiLen1) & nTrans & Space(11 - Milen2) & Format(rsVta_X!VENTAS, "STANDARD"), "VENTA DEL DIA:" & Space(4 - MiLen1) & nTrans & Space(11 - Milen2) & Format(rsVta_X!VENTAS, "STANDARD")
        nSumTotal = nSumTotal + rsVta_X!VENTAS
        rsVta_X.MoveNext
    Loop
    ProgBar.Value = 30
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
    '''''''Printer.FontUnderline = True
    'Printer.Print "DESGLOSE DE INGRESOS"
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "TOTALES DE CAJA", "TOTALES DE CAJA"
    ''''''''Printer.FontUnderline = False
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)

    Do Until rsPgo_X.EOF
        RSPAGOS.MoveFirst
        RSPAGOS.Find "CODIGO = " & rsPgo_X!TIPO_PAGO
        If Not RSPAGOS.EOF Then
            MiLen1 = Len(rsPgo_X!X_COUNT)
            Milen2 = Len(Format(rsPgo_X!VALOR, "STANDARD"))
            Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, FormatTexto(RSPAGOS!descrip, 13) & Space(4 - MiLen1) & rsPgo_X!X_COUNT & Space(13 - Milen2) & Format(rsPgo_X!VALOR, "STANDARD"), FormatTexto(RSPAGOS!descrip, 13) & Space(4 - MiLen1) & rsPgo_X!X_COUNT & Space(13 - Milen2) & Format(rsPgo_X!VALOR, "STANDARD")
        Else
            MiLen1 = 1
            Milen2 = Len(Format(0#, "standard"))
            Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "OTRO PAGO    " & Space(4 - MiLen1) & 0 & Space(13 - MiLen1) & Format(0#, "standard"), "OTRO PAGO    " & Space(4 - MiLen1) & 0 & Space(13 - MiLen1) & Format(0#, "standard")
            'Printer.Print rsPagos!descrip & Chr(9) & 0 & Chr(9) & Format(0#, "standard")
        End If
        nSumVta = nSumVta + rsPgo_X!VALOR
        rsPgo_X.MoveNext
    Loop
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
    MiLen1 = Len(Format(nSumVta, "currency"))
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "SUBTOTAL: " & Space(20 - MiLen1) & Format(nSumVta, "currency"), "SUBTOTAL: " & Space(20 - MiLen1) & Format(nSumVta, "currency")
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "------------------------------", "------------------------------"

    sqltext = "SELECT TIPO_PAGO,COUNT(TIPO_PAGO) AS X_COUNT, " & _
            " SUM(MONTO) AS VALOR FROM TRANSAC_PROP " & _
            " WHERE CAJERO = " & rsCajeros!CAJERO & _
            " GROUP BY TIPO_PAGO "
    rsProp.Open sqltext, msConn, adOpenStatic, adLockOptimistic
    
    Do Until rsProp.EOF
        RSPAGOS.MoveFirst
        RSPAGOS.Find "CODIGO = " & rsProp!TIPO_PAGO
        If Not RSPAGOS.EOF Then
            MiLen1 = Len(rsProp!X_COUNT)
            Milen2 = Len(Format(rsProp!VALOR, "STANDARD"))
            Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "Propina " & FormatTexto(RSPAGOS!descrip, 5) & Space(4 - MiLen1) & rsProp!X_COUNT & Space(13 - Milen2) & Format(rsProp!VALOR, "STANDARD"), "Propina " & FormatTexto(RSPAGOS!descrip, 5) & Space(4 - MiLen1) & rsProp!X_COUNT & Space(13 - Milen2) & Format(rsProp!VALOR, "STANDARD")
        End If
        rsProp.MoveNext
    Loop
    rsProp.Close
    ProgBar.Value = 40
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
    ''''''Printer.FontUnderline = True
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "AJUSTES", "AJUSTES"
    ''''''Printer.FontUnderline = False
    Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)

    Do Until rsAjustes.EOF
        MiLen1 = Len(rsAjustes!X_COUNT)
        Milen2 = Len(Format(rsAjustes!VALOR, "STANDARD"))
        If rsAjustes!CORREC = "EC" Then
            Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "CORECCION " & Space(8 - MiLen1) & rsAjustes!X_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD"), "CORECCION " & Space(8 - MiLen1) & rsAjustes!X_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD")
        ElseIf rsAjustes!CORREC = "VO" Then
            Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "ANULACION " & Space(8 - MiLen1) & rsAjustes!X_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD"), "ANULACION " & Space(8 - MiLen1) & rsAjustes!X_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD")
        ElseIf rsAjustes!CORREC = "DC" Then
            Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "DESCUENTO " & Format(rsAjustes!DESCUENTO, "0.00") & Space(4 - MiLen1) & rsAjustes!X_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD"), "DESCUENTO " & Format(rsAjustes!DESCUENTO, "0.00") & Space(4 - MiLen1) & rsAjustes!X_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD")
        End If
        rsAjustes.MoveNext
    Loop

    'PREPARA INFO PARA DESCUENTO GLOBAL
    rsAjustes.Close
    sqltext = "SELECT a.cajero,a.tipo_pago,SUM(a.monto) AS Valor, " & _
            " COUNT(a.tipo_pago) as X_COUNT " & _
            " FROM transac_pago as a " & _
            " WHERE a.cajero = " & rsCajeros!CAJERO & _
            " AND a.tipo_pago = 99 " & _
            " GROUP BY a.cajero,a.TIPO_PAGO"
    
    rsAjustes.Open sqltext, msConn, adOpenStatic, adLockOptimistic
    
    If Not rsAjustes.EOF Then
        RSPAGOS.MoveFirst
        RSPAGOS.Find "CODIGO = " & rsAjustes!TIPO_PAGO
        If Not RSPAGOS.EOF Then
            MiLen1 = Len(rsAjustes!X_COUNT)
            Milen2 = Len(Format(rsAjustes!VALOR, "STANDARD"))
            Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, FormatTexto(RSPAGOS!descrip, 13) & Space(5 - MiLen1) & rsAjustes!X_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD"), FormatTexto(RSPAGOS!descrip, 13) & Space(5 - MiLen1) & rsAjustes!X_COUNT & Space(12 - Milen2) & Format(rsAjustes!VALOR, "STANDARD")
        End If
    End If
    ProgBar.Value = 50
    
    msConn.BeginTrans
    msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
    msConn.Execute "UPDATE CAJEROS SET X_C = X_C + 1 " & _
                   " WHERE NUMERO = " & rsCajeros!CAJERO
    msConn.CommitTrans
    
    rsCajeros.MoveNext
    If Not rsCajeros.EOF = True Then
        For i = 1 To 10
            Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
        Next
        Sys_Pos.Coptr1.CutPaper 100
    End If
    
    rsVta_X.Close
    rsPgo_X.Close
    rsTran.Close
    rsAjustes.Close

Loop
ProgBar.Value = 60
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, Space(2), Space(2)
MiLen1 = Len(Format(nSumTotal, "currency"))
Sys_Pos.Coptr1.PrintTwoNormal FPTR_S_JOURNAL_RECEIPT, "TOTAL: " & Space(23 - MiLen1) & Format(nSumTotal, "currency"), "TOTAL: " & Space(23 - MiLen1) & Format(nSumTotal, "currency")
For i = 1 To 10
    Sys_Pos.Coptr1.PrintNormal PTR_S_RECEIPT, Space(2) & Chr(&HD) & Chr(&HA)
Next
Sys_Pos.Coptr1.CutPaper 100
ProgBar.Value = 0
'Close #1
On Error GoTo 0
Exit Sub

AjustaMilen:
errCounter = errCounter + 1
Milen2 = 11
If errCounter < 3 Then
    Resume
Else
    MsgBox "EXISTE UN PROMLEMA DE IMPRESION. UNA VEZ TERMINADO REVISE EL LISTADO Y VERIFIQUE LOS DATOS", vbCritical, "LA IMPRESION DEL REPORTE TIENE PROBLEMAS"
    Resume Next
End If

End Sub

Private Sub cmdOpenGaveta_Click()
'abrir gaveta
Sys_Pos.Cocash1.Claim 5000
RC = Sys_Pos.Cocash1.OpenDrawer
End Sub

Private Sub cmdRX_Click(Index As Integer)
Select Case Index
Case 0  'REPORTE DE CAJEROS X
    cmdRX(Index).Enabled = False
    RepCajX
    cmdRX(Index).Enabled = True
Case 1
    cmdRX(Index).Enabled = False
    RepMesX
    cmdRX(Index).Enabled = True
Case 2
    cmdRX(Index).Enabled = False
    RepMeseroX
    cmdRX(Index).Enabled = True
Case 3
    cmdRX(Index).Enabled = False
    RepDepX
    cmdRX(Index).Enabled = True
Case Else
    MsgBox "AUN NO ESTA DISPONIBLE", vbInformation, BoxTit
End Select
End Sub

Private Sub cmdRZ_Click(Index As Integer)
Select Case Index
Case 0
    BoxResp = MsgBox("¿ Desea realizar REPORTE Z ?", vbQuestion + vbYesNo, BoxTit)
    If BoxResp = vbYes Then
        If FILE_OK = True Then
            Call DameVolumenA
            ProgBar.Value = 30
            Call BackDataUp
            ProgBar.Value = 100
        End If
        ProgBar.Value = 0
        cmdRZ(0).Enabled = False
        RepCajZ
        cmdRZ(0).Enabled = True
        cmdRZ(1).Enabled = True
        cmdRZ(2).Enabled = True
    End If
Case 1
    cmdRZ(Index).Enabled = False
    RepMesZ
    cmdRZ(Index).Enabled = True
Case 2
    cmdRZ(Index).Enabled = False
    RepMeseroZ
    cmdRZ(Index).Enabled = True
Case Else
    MsgBox "LA OPCION AUN NO HA SIDO HABILITADA", vbInformation, BoxTit
End Select
End Sub

Private Sub Command1_Click()
Unload Me
End Sub
Private Sub ActualizaInvent()
'ACTUALIZACION DE INVENTARIO
Dim iIntento As Integer
iIntento = 0
On Error GoTo AdmError:
Do Until rsPluInvent.EOF
    msConn.BeginTrans
    msConn.Execute "UPDATE INVENT SET EXIST2 = EXIST2 - " & rsPluInvent!CONSUMIDO & _
            " WHERE ID = " & rsPluInvent!ID_PROD_INV
    msConn.CommitTrans
    rsPluInvent.MoveNext
Loop
rsPluInvent.Close
Set rsPluInvent = Nothing
On Error GoTo 0
Exit Sub

AdmError:
If iIntento < 4 Then
    MsgBox Err.Description, vbCritical, "Error Número : " & Err.Number
    iIntento = iIntento + 1
    Resume
Else
    MsgBox "ERROR GRAVE EN ACTUALIZACION DE INVENTARIO : " & Err.Description & vbcrl & _
                    "LUGAR DEL ERROR : " & Err.Source, vbCritical, "Error Número : " & Err.Number
    Resume Next
End If
End Sub

