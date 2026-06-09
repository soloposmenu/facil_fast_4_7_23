VERSION 5.00
Begin VB.Form Vuelto 
   Appearance      =   0  'Flat
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CAMBIO PARA EL CLIENTE"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   ControlBox      =   0   'False
   Icon            =   "Vuelto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Salir 
      Caption         =   "Regresar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Vuelto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Caption         =   "CAMBIO"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Vuelto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
On Error Resume Next
    Vuelto = Format(nCambio, "CURRENCY")
    If SLIP_OK = True Then
        Label2 = "Recuerde Introducir la FACTURA en la boca de la Impresora"
    End If
On Error GoTo 0
End Sub

Private Sub Form_Click()
Salir_Click
End Sub
Private Sub Form_Load()
On Error Resume Next
    Vuelto = Format(nCambio, "CURRENCY")
    If SLIP_OK = True Then
        Label2 = "Recuerde Introducir la FACTURA en la boca de la Impresora"
    End If
On Error GoTo 0
End Sub

Private Sub Label1_Click()
On Error Resume Next
    Vuelto = Format(nCambio, "CURRENCY")
    If SLIP_OK = True Then Label2 = "Recuerde Introducir la FACTURA en la boca de la Impresora"
On Error GoTo 0
Salir_Click
End Sub

Private Sub Label2_Click()
On Error Resume Next
    Vuelto = Format(nCambio, "CURRENCY")
    If SLIP_OK = True Then Label2 = "Recuerde Introducir la FACTURA en la boca de la Impresora"
On Error GoTo 0
Salir_Click
End Sub

Private Sub Salir_Click()
'Me.Visible = False
Me.Hide
'Unload Me
End Sub
Private Sub Vuelto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Salir_Click
End Sub
Private Sub Vuelto_Click()
Salir_Click
End Sub
