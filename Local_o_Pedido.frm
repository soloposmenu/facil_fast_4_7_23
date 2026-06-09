VERSION 5.00
Begin VB.Form Local_o_Pedido 
   BackColor       =   &H00B39665&
   BorderStyle     =   0  'None
   Caption         =   "LLAMADA  o CLIENTE"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6105
   Icon            =   "Local_o_Pedido.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image ImageCliente 
      Height          =   1875
      Left            =   3600
      Picture         =   "Local_o_Pedido.frx":0742
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2010
   End
   Begin VB.Image ImagePhone 
      Height          =   1875
      Left            =   480
      Picture         =   "Local_o_Pedido.frx":850C
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2010
   End
End
Attribute VB_Name = "Local_o_Pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem 13JUN2025. PARA JAP JP

Public nClienteLlamada As Long
Public nClienteCamina As Long

Private Sub ImageCliente_Click()
nClienteCamina = nClienteCamina + 1
Call salida
End Sub

Private Sub ImagePhone_Click()
nClienteLlamada = nClienteLlamada + 1
Call salida
End Sub

Sub salida()


Rem WriteToINI "Administracion", "Llamada", nClienteLlamada, App.Path & "\soloini.ini"
Rem WriteToINI "Administracion", "Persona", nClienteLlamada, App.Path & "\soloini.ini"

Unload Me
End Sub
