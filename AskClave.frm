VERSION 5.00
Begin VB.Form AskClave 
   BackColor       =   &H008080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INTRODUCIR CONTRASEÑA"
   ClientHeight    =   2865
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3975
   ControlBox      =   0   'False
   Icon            =   "AskClave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1692.736
   ScaleMode       =   0  'User
   ScaleWidth      =   3732.31
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Borrar 
      Height          =   615
      Left            =   1680
      Picture         =   "AskClave.frx":0742
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2040
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   3735
      Begin VB.CommandButton Command8 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   60
         TabIndex        =   13
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   780
         TabIndex        =   12
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   1500
         TabIndex        =   11
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   2220
         TabIndex        =   10
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   2940
         TabIndex        =   9
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   60
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   780
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   1500
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   2220
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   2940
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.TextBox txtLin 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "@"
      TabIndex        =   1
      Top             =   120
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   2640
      TabIndex        =   3
      Top             =   2040
      Width           =   1260
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H008080FF&
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
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   2400
   End
End
Attribute VB_Name = "AskClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private nlPase As Integer

Private Sub Borrar_Click()
nlPase = 0
txtLin = ""
End Sub

Private Sub cmdCancel_Click()
OkAnul = 0
Unload Me
End Sub

Private Sub cmdOK_Click()

Dim rsUsr As Recordset

If txtLin = "" Or Not IsNumeric(txtLin) Then
    MsgBox "Clave NO es Valida", vbExclamation, BoxTit
Else
    Set rsUsr = New Recordset
    rsUsr.Open "SELECT numero,nombre FROM USUARIOS " & " WHERE CLAVE = " & "'" & LTrim(txtLin) & "'", msConn, adOpenForwardOnly, adLockReadOnly
    If Not rsUsr.EOF Then
        OkAnul = 1
        'OKGlobal = 1
        rsUsr.Close
    End If
End If
Unload Me

End Sub

Private Sub Command8_Click(Index As Integer)
Dim cCant As String

If nlPase = 0 Then
    txtLin = Command8(Index).Index
Else
    'cCant = Str(txtLin)
    cCant = CStr(txtLin)
    cCant = cCant & Command8(Index).Index
    txtLin = cCant
End If
nlPase = nlPase + 1
End Sub

Private Sub Form_Load()
nlPase = 0
lblLabels = txtInfo
End Sub

Private Sub txtLin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdOK.SetFocus
End Sub


