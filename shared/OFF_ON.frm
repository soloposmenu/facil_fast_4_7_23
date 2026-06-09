VERSION 5.00
Begin VB.Form OFF_ON 
   BackColor       =   &H00EA6E67&
   BorderStyle     =   0  'None
   Caption         =   "ON LINE // OFF LINE"
   ClientHeight    =   4275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   Icon            =   "OFF_ON.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Regresar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1433
      TabIndex        =   2
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CheckBox chkStatus 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   2640
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   240
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status del Sistema ON LINE u OFF LINE                            Toque el gancho para cambiarlo  "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   420
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "OFF_ON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CheckBox control — 0 is Unchecked (default), 1 is Checked,
Private Sub Command2_Click()
Dim vResp As Variant


Me.MousePointer = vbHourglass
    
If chkStatus.value = 0 Then  ' 1 = Checked (ON LINE)
    Call WriteToINI("Fiscal", "Logmode", "SI", App.Path & "\soloini.ini")
    FE_Facil_ONLINE = False
    Me.MousePointer = vbDefault
    vResp = ShowMsg("EL SISTEMA QUEDA " & vbCrLf & vbCrLf & "<< OFF >>  LINE", vbYellow, vbRed, vbYes)
    EscribeLog "EL SISTEMA QUEDA << OFF >> LINE"
    cLogMode = "SI"
Else
    If Check_FE_Website Then
        Call WriteToINI("Fiscal", "Logmode", "", App.Path & "\soloini.ini")
        FE_Facil_ONLINE = True
        Me.MousePointer = vbDefault
        vResp = ShowMsg("EL SISTEMA QUEDA " & vbCrLf & vbCrLf & "<< ON >>  LINE", vbBlue, vbCyan, vbYes)
        EscribeLog "EL SISTEMA QUEDA << ON >> LINE"
        cLogMode = ""
    Else
        Me.MousePointer = vbDefault
        ShowMsg "MODO << ON LINE >> NO ESTA DISPONIBLE" & vbCrLf & "SE TRABAJARA << OFF LINE >>", vbYellow, vbRed
        Call WriteToINI("Fiscal", "Logmode", "SI", App.Path & "\soloini.ini")
        FE_Facil_ONLINE = False
        EscribeLog "Check_FE_Website NO ESTA DISPONIBLE, SE DEJA << OFF LINE >>"
        cLogMode = "SI"
    End If
End If

Me.MousePointer = vbDefault

Unload OFF_ON
End Sub

Private Sub Form_Load()

cLogMode = GetFromINI("Fiscal", "Logmode", App.Path & "\soloini.ini")

If cLogMode = "" Then
    chkStatus.value = 1
    chkStatus.Caption = "    ON LINE F.E."
    chkStatus.ForeColor = vbCyan
Else
    chkStatus.value = 0
    chkStatus.Caption = "    OFF LINE F.E."
    chkStatus.ForeColor = vbRed
End If

End Sub
Private Sub chkStatus_Click()
    If chkStatus.value = 1 Then  ' 1 = Checked (ON LINE)
        chkStatus.Caption = "    ON LINE F.E." ' Or maybe "SET OFFLINE"? It depends on your UI logic.
        FE_Facil_ONLINE = True
        chkStatus.ForeColor = vbCyan
    Else ' 0 = Unchecked (OFF LINE)
        chkStatus.Caption = "    OFF LINE F.E."
        FE_Facil_ONLINE = False
        chkStatus.ForeColor = vbRed
    End If
End Sub
