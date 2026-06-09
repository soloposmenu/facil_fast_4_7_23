VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form FEReports 
   BackColor       =   &H00B39665&
   Caption         =   "NOTAS DE CREDITO"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11745
   Icon            =   "FEReports.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11745
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox lbValor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   14
      Top             =   1080
      Width           =   2175
   End
   Begin VB.PictureBox Picture 
      Height          =   375
      Left            =   11280
      Picture         =   "FEReports.frx":0442
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   12
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtEmail 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   10
      Top             =   2040
      Width           =   5655
   End
   Begin VB.CommandButton cmdAplicaNC 
      Caption         =   "&APLICAR NOTA DE CREDITO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9480
      TabIndex        =   3
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton cmdFiltroFecha 
      Caption         =   "&FILTRAR x FECHA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin DDSharpGridOLEDB2.SGGrid DD_PEDDETALLE 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   5775
      _cx             =   10186
      _cy             =   11880
      DataMember      =   ""
      DataMode        =   1
      AutoFields      =   -1  'True
      Enabled         =   -1  'True
      GridBorderStyle =   1
      ScrollBars      =   3
      FlatScrollBars  =   0
      ScrollBarTrack  =   0   'False
      DataRowCount    =   2
      BeginProperty HeadingFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataColCount    =   2
      HeadingRowCount =   1
      HeadingColCount =   0
      TextAlignment   =   0
      WordWrap        =   0   'False
      Ellipsis        =   1
      HeadingBackColor=   12632256
      HeadingForeColor=   -2147483630
      HeadingTextAlignment=   0
      HeadingWordWrap =   0   'False
      HeadingEllipsis =   1
      GridLines       =   1
      HeadingGridLines=   2
      GridLinesColor  =   -2147483633
      HeadingGridLinesColor=   -2147483632
      EvenOddStyle    =   1
      ColorEven       =   -2147483628
      ColorOdd        =   -2147483636
      UserResizeAnimate=   1
      UserResizing    =   3
      RowHeightMin    =   450
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      UserDragging    =   2
      UserHiding      =   0
      CellPadding     =   15
      CellBkgStyle    =   1
      CellBackColor   =   -2147483643
      CellForeColor   =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   2
      FocusRectColor  =   0
      FocusRectLineWidth=   1
      TabKeyBehavior  =   0
      EnterKeyBehavior=   1
      NavigationWrapMode=   1
      SkipReadOnly    =   0   'False
      DefaultColWidth =   1219
      DefaultRowHeight=   255
      CellsBorderColor=   0
      CellsBorderVisible=   -1  'True
      RowNumbering    =   0   'False
      EqualRowHeight  =   0   'False
      EqualColWidth   =   0   'False
      HScrollHeight   =   0
      VScrollWidth    =   0
      Format          =   "General"
      Appearance      =   2
      FitLastColumn   =   0   'False
      SelectionMode   =   2
      MultiSelect     =   0
      AllowAddNew     =   0   'False
      AllowDelete     =   0   'False
      AllowEdit       =   0   'False
      ScrollBarTips   =   0
      CellTips        =   0
      CellTipsDelay   =   1000
      SpecialMode     =   0
      OutlineLines    =   1
      CacheAllRecords =   -1  'True
      ColumnClickSort =   0   'False
      PreviewPaneColumn=   ""
      PreviewPaneType =   0
      PreviewPanePosition=   2
      PreviewPaneSize =   2000
      GroupIndentation=   241
      InactiveSelection=   1
      AutoScroll      =   -1  'True
      AutoResize      =   2
      AutoResizeHeadings=   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   0
      Caption         =   ""
      ScrollTipColumn =   ""
      MaxRows         =   4194304
      MaxColumns      =   8192
      NewRowPos       =   1
      CustomBkgDraw   =   0
      AutoGroup       =   0   'False
      GroupByBoxVisible=   0   'False
      GroupByBoxText  =   ""
      AlphaBlendEnabled=   0   'False
      DragAlphaLevel  =   206
      AutoSearch      =   0
      AutoSearchDelay =   2000
      OutlineIcons    =   1
      CellTipsDisplayLength=   3000
      StylesCollection=   $"FEReports.frx":0A18
      ColumnsCollection=   $"FEReports.frx":284B
      ValueItems      =   $"FEReports.frx":31C1
   End
   Begin MSComCtl2.DTPicker txtFecIni 
      Height          =   465
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   145883137
      CurrentDate     =   36418
   End
   Begin DDSharpGridOLEDB2.SGGrid DDCONTENIDO 
      Height          =   4575
      Left            =   6120
      TabIndex        =   6
      Top             =   3600
      Width           =   5535
      _cx             =   9763
      _cy             =   8070
      DataMember      =   ""
      DataMode        =   1
      AutoFields      =   -1  'True
      Enabled         =   -1  'True
      GridBorderStyle =   1
      ScrollBars      =   3
      FlatScrollBars  =   0
      ScrollBarTrack  =   0   'False
      DataRowCount    =   2
      BeginProperty HeadingFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataColCount    =   2
      HeadingRowCount =   1
      HeadingColCount =   0
      TextAlignment   =   0
      WordWrap        =   0   'False
      Ellipsis        =   1
      HeadingBackColor=   12632256
      HeadingForeColor=   -2147483630
      HeadingTextAlignment=   0
      HeadingWordWrap =   0   'False
      HeadingEllipsis =   1
      GridLines       =   1
      HeadingGridLines=   2
      GridLinesColor  =   -2147483633
      HeadingGridLinesColor=   -2147483632
      EvenOddStyle    =   1
      ColorEven       =   -2147483628
      ColorOdd        =   -2147483636
      UserResizeAnimate=   1
      UserResizing    =   3
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      UserDragging    =   2
      UserHiding      =   0
      CellPadding     =   15
      CellBkgStyle    =   1
      CellBackColor   =   -2147483643
      CellForeColor   =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   2
      FocusRectColor  =   0
      FocusRectLineWidth=   1
      TabKeyBehavior  =   0
      EnterKeyBehavior=   1
      NavigationWrapMode=   1
      SkipReadOnly    =   0   'False
      DefaultColWidth =   1219
      DefaultRowHeight=   255
      CellsBorderColor=   0
      CellsBorderVisible=   -1  'True
      RowNumbering    =   0   'False
      EqualRowHeight  =   0   'False
      EqualColWidth   =   0   'False
      HScrollHeight   =   0
      VScrollWidth    =   0
      Format          =   "General"
      Appearance      =   2
      FitLastColumn   =   0   'False
      SelectionMode   =   2
      MultiSelect     =   0
      AllowAddNew     =   0   'False
      AllowDelete     =   0   'False
      AllowEdit       =   -1  'True
      ScrollBarTips   =   0
      CellTips        =   0
      CellTipsDelay   =   1000
      SpecialMode     =   0
      OutlineLines    =   1
      CacheAllRecords =   -1  'True
      ColumnClickSort =   0   'False
      PreviewPaneColumn=   ""
      PreviewPaneType =   0
      PreviewPanePosition=   2
      PreviewPaneSize =   2000
      GroupIndentation=   241
      InactiveSelection=   1
      AutoScroll      =   -1  'True
      AutoResize      =   2
      AutoResizeHeadings=   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   0
      Caption         =   ""
      ScrollTipColumn =   ""
      MaxRows         =   4194304
      MaxColumns      =   8192
      NewRowPos       =   1
      CustomBkgDraw   =   0
      AutoGroup       =   0   'False
      GroupByBoxVisible=   0   'False
      GroupByBoxText  =   ""
      AlphaBlendEnabled=   0   'False
      DragAlphaLevel  =   206
      AutoSearch      =   0
      AutoSearchDelay =   2000
      OutlineIcons    =   1
      CellTipsDisplayLength=   3000
      StylesCollection=   $"FEReports.frx":3261
      ColumnsCollection=   $"FEReports.frx":5093
      ValueItems      =   $"FEReports.frx":5A09
   End
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00B39665&
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   2
      Left            =   7920
      TabIndex        =   13
      Top             =   8180
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00B39665&
      Caption         =   "E-MAIL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   6000
      TabIndex        =   11
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00B39665&
      Caption         =   "FECHA"
      BeginProperty Font 
         Name            =   "Verdana"
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
      TabIndex        =   8
      Top             =   550
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00B39665&
      Caption         =   "VALOR $"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   9720
      TabIndex        =   7
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lb 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00B39665&
      Caption         =   "FACTURA #"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6720
      TabIndex        =   5
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lbFactura 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   1080
      Width           =   3255
   End
End
Attribute VB_Name = "FEReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long


Private Const PBM_SETBARCOLOR As Long = &H409
Private Const PBM_SETBKCOLOR As Long = &H2001
Private Const PROGBAR_DEF_COLOR = &HFF000000 '&H8000000D

Dim dF1 As String
Dim nRowKey As Long
Private cNewSource  As String
Dim cFechaFactura As String
Dim cHoraFactura As String

Private Sub cmdAplicaNC_Click()

Dim FileUbicacion As String
Dim NC_FileUbicacion As String

FileUbicacion = App.Path & "\HIST\" & lbFactura.Caption & ".txt"
NC_FileUbicacion = App.Path & "\HIST\NC" & lbFactura.Caption & ".txt"

If Dir(NC_FileUbicacion) <> "" Then
    ShowMsg "YA SE HIZO NOTA DE CREDITO A ESTA FACTURA", vbRed, vbCyan
    ProgBar.value = 0
    Exit Sub
End If

If Dir(FileUbicacion) = "" Then
    ShowMsg "NO SE PUEDE HACER NOTA DE CREDITO " & vbCrLf & "A ESTA FACTURA # " & vbCrLf & vbCrLf & lbFactura.Caption, vbYellow, vbRed
    EscribeLog "FACTURA # " & lbFactura.Caption & ", NO ENCONTRADA (json)"
    ProgBar.value = 0
    Exit Sub
End If


If cmdAplicaNC.Tag = "OFF-LINE" Then
    'INFO: NO HAY VALIDACION DE CORREO, YA QUE NO SE ENVIA
Else
    If txtEmail.Text = "" Then
        ShowMsg "DEBE ESCRIBIR EL CORREO DEL CLIENTE", vbYellow, vbRed
        txtEmail.BackColor = vbYellow
        txtEmail.SetFocus
        Exit Sub
    End If
    
    If Not InStr(1, txtEmail.Text, "@") > 0 Then
        ShowMsg "DEBE ESCRIBIR UN CORREO VALIDO", vbYellow, vbRed
        txtEmail.BackColor = vbYellow
        txtEmail.SetFocus
        Exit Sub
    End If
    
    If Len(txtEmail.Text) < 8 Then
        ShowMsg "DEBE ESCRIBIR UN CORREO VALIDO", vbYellow, vbRed
        txtEmail.BackColor = vbYellow
        txtEmail.SetFocus
        Exit Sub
    End If
End If

If ShowMsg("¿ Desea Realizar NOTA DE CREDITO ?" & vbCrLf & vbCrLf & "POR VALOR DE: " & lbValor.Text, vbBlue, vbYellow, vbYesNo) = vbYes Then
    If HacerNC Then
        If Send_FE_Mail(txtEmail.Text, FE_CUFE, "") Then
            ProgBar.value = 100
            On Error Resume Next
                txtEmail.BackColor = vbWhite
                txtEmail.Text = ""
                lbValor.Text = ""
                lbFactura.Caption = ""
            On Error GoTo 0
        Else
            ProgBar.value = 20
            If txtEmail.Text <> "" Then
                ShowMsg "FALLO AL ENVIAR FACTURA # " & lbFactura.Caption & vbCrLf & "A ESTE eMAIL" & txtEmail.Text, vbBlue, vbCyan
                EscribeLog "FALLO AL ENVIAR FACTURA # " & lbFactura.Caption & "A ESTE eMAIL " & txtEmail.Text
            End If
            ProgBar.value = 0
        End If
    Else
        If cFE_LastErrorCode = 400 Then
            'INFO: Ajuste de centavo
            lbValor.BackColor = &H8080FF
            lbValor.Text = Format(lbValor.Text, "#.00") + 0.01
            Call HacerNC
            If Send_FE_Mail(txtEmail.Text, FE_CUFE, "") Then
                ProgBar.value = 100
                On Error Resume Next
                    txtEmail.BackColor = vbWhite
                    txtEmail.Text = ""
                    lbValor.Text = ""
                    lbFactura.Caption = ""
                On Error GoTo 0
            Else
                ProgBar.value = 20
                ShowMsg "FALLO AL ENVIAR FACTURA # " & lbFactura.Caption & vbCrLf & "A ESTA DIRECCION " & txtEmail.Text, vbBlue, vbCyan
                EscribeLog "FALLO AL ENVIAR FACTURA # " & lbFactura.Caption & "A ESTA DIRECCION " & txtEmail.Text
                ProgBar.value = 0
            End If
            lbValor.BackColor = vbWhite
        End If
    End If
Else
    ShowMsg "SE CANCELO LA REALIZACION DE" & vbCrLf & "LA NOTA DE CREDITO", vbYellow, vbBlue
End If
End Sub

Private Sub cmdFiltroFecha_Click()
dF1 = Format(txtFecIni, "YYYYMMDD")
Call LoadTransacFE
End Sub

Private Sub DD_PEDDETALLE_Click()
If DD_PEDDETALLE.Rows.Current.Cells(3).value = "0.00" Then
    ShowMsg "FACTURA SIN VALOR COMERCIAL, NO SE PUEDE REALIZAR NOTA DE CREDITO", vbYellow, vbRed
    Exit Sub
End If


If DD_PEDDETALLE.Rows.Current.Cells(5).value = "OFF-LINE" Then
    Rem ShowMsg "FACTURA EMITIDA OFF-LINE", vbYellow, vbRed
    DD_PEDDETALLE.Rows.Current.Style.ForeColor = vbRed
    DD_PEDDETALLE.Rows.Current.Style.Font.Bold = True
    
    cmdAplicaNC.Tag = "OFF-LINE"
    
    Label2(2).FontBold = True
    Label2(2).ForeColor = vbYellow
    Label2(2).BackColor = vbRed
    Label2(2).FontSize = 12
    Label2(2) = "FACTURA OFF LINE"
    
    txtEmail.Text = ""
    txtEmail.BackColor = vbRed
    txtEmail.Enabled = False
Else
    cmdAplicaNC.Tag = ""
    
    Label2(2).ForeColor = vbWhite
    Label2(2).BackColor = &HEA6E67
    Label2(2).FontSize = 10
    Label2(2).FontBold = False
    Label2(2) = "FACTURA NORMAL"
    EscribeLog Label2(2)
    
    txtEmail.BackColor = vbYellow
    txtEmail.Enabled = True
End If

On Error Resume Next
lbFactura.Caption = Val(DD_PEDDETALLE.Rows.Current.Cells(0).value)
lbValor.Text = DD_PEDDETALLE.Rows.Current.Cells(3).value
nRowKey = DD_PEDDETALLE.Rows.Current.Key

cFechaFactura = DD_PEDDETALLE.Rows.Current.Cells(1).value
cHoraFactura = DD_PEDDETALLE.Rows.Current.Cells(2).value

On Error GoTo 0

Call Load_FE_Historico(Val(lbFactura.Caption))



'If DD_PEDDETALLE.Rows.Current.Cells(3).value = "0.00" Then
'    ShowMsg "FACTURA SIN VALOR COMERCIAL, NO SE PUEDE REALIZAR NOTA DE CREDITO", vbYellow, vbRed
'    Exit Sub
'End If
'
'On Error Resume Next
'lbFactura.Caption = Val(DD_PEDDETALLE.Rows.Current.Cells(0).value)
'lbValor.Caption = DD_PEDDETALLE.Rows.Current.Cells(3).value
'nRowKey = DD_PEDDETALLE.Rows.Current.Key
'
'cFechaFactura = DD_PEDDETALLE.Rows.Current.Cells(1).value
'cHoraFactura = DD_PEDDETALLE.Rows.Current.Cells(2).value

'If DD_PEDDETALLE.Rows.Current.Style.ForeColor = 0 Then
'    DD_PEDDETALLE.Rows.Current.Style.ForeColor = vbRed
'Else
'    DD_PEDDETALLE.Rows.Current.Style.ForeColor = vbBlack
'End If

End Sub
Private Sub Form_Load()

If mi_Ambiente = 2 Then Me.BackColor = vbCyan
txtFecIni = Format(Date, "SHORT DATE")

Call GetFacilRegistryValues
Call LoadTransacFE

If FE_Facil_ONLINE Then
Else
    txtEmail.BackColor = vbRed
    txtEmail.Enabled = False
End If

Call ChangeColorOfProgBar

End Sub

Private Sub LoadTransacFE()
Dim cSQL As String
Dim rsTFE As ADODB.Recordset


Set rsTFE = New ADODB.Recordset
If dF1 = "" Then
    cSQL = "SELECT DOC_SOLO AS FACTURA, FORMAT(FECHA,'####-##-##') as FECHA, HORA, "
    cSQL = cSQL & " FORMAT((SUB_TOTAL + ITBM),'STANDARD') AS VALOR, FISCAL, ID_URL"
    cSQL = cSQL & " FROM TRANSAC_FISCAL "
    cSQL = cSQL & " ORDER BY FECHA DESC, HORA DESC"
Else
    cSQL = "SELECT DOC_SOLO AS FACTURA, FORMAT(FECHA,'####-##-##') as FECHA, HORA, "
    cSQL = cSQL & " FORMAT((SUB_TOTAL + ITBM),'STANDARD') AS VALOR, FISCAL, ID_URL"
    cSQL = cSQL & " FROM TRANSAC_FISCAL "
    cSQL = cSQL & " WHERE FECHA = '" & dF1 & "'"
    cSQL = cSQL & " ORDER BY FECHA DESC, HORA DESC"
End If

rsTFE.Open cSQL, msConn, adOpenKeyset, adLockOptimistic
DD_PEDDETALLE.DataMode = sgBound
Set DD_PEDDETALLE.DataSource = rsTFE

DD_PEDDETALLE.Columns(1).Style.TextAlignment = sgAlignLeftCenter
DD_PEDDETALLE.Columns(2).Width = 1500
DD_PEDDETALLE.Columns(3).Style.TextAlignment = sgAlignCenterCenter
DD_PEDDETALLE.Columns(4).Style.TextAlignment = sgAlignRightCenter
DD_PEDDETALLE.Columns(5).Hidden = True
DD_PEDDETALLE.Columns(6).Hidden = True

End Sub

''---------------------------------------------------------------------------------------
'' Procedure : Load_FE_Historico
'' Author    : hsequeira
'' Date      : 26/04/2024
'' Purpose   : nNumFactura PASA A SER LONG EN VEZ DE INTEGER
''---------------------------------------------------------------------------------------
''
''Public Function Load_FE_Historico(nNumFactura As Integer) As Boolean
'Public Function Load_FE_Historico(nNumFactura As Long) As Boolean
'Dim cSQL As String
'Set rsFETrans = New ADODB.Recordset
'
'cSQL = "SELECT LIN, DESCRIP AS description, CANT AS quantity, PRECIO_UNIT  as price, 'und' as mu, PLU as internal_code,"
''cSQL = cSQL & " DESCRIP AS description, '01' as type, '01' as code,  ((PRECIO - FE_DESCUENTO) * (CON_TAX/100)) as amount, "
'cSQL = cSQL & " '01' as type, '01' as code,  "
'cSQL = cSQL & " ID_DESCUENTO as descuento, CON_TAX "
'cSQL = cSQL & " FROM HIST_TR "
'cSQL = cSQL & " WHERE NUM_TRANS = " & nNumFactura
'cSQL = cSQL & " AND LEFT(TIPO,2) NOT IN ('VO','BV','DC','BE','EC') "
'cSQL = cSQL & " ORDER BY LIN"
'
'rsFETrans.Open cSQL, msConn, adOpenStatic, adLockOptimistic
'
'If rsFETrans.RecordCount = 0 Then
'    '
'    'DDCONTENIDO.BackColor = vbRed
'    DDCONTENIDO.BackColor = &HC0&
'Else
'    DDCONTENIDO.BackColor = &H8000000C
'End If
''ShowMsg "AUN NO SE HA EMITIDO EL REPORTE Z", vbYellow, vbRed
'
'DDCONTENIDO.DataMode = sgBound
'Set DDCONTENIDO.DataSource = rsFETrans
'
'DDCONTENIDO.Columns(1).Style.TextAlignment = sgAlignLeftCenter
'DDCONTENIDO.Columns(1).Width = 400
'DDCONTENIDO.Columns(2).Width = 3200
'DDCONTENIDO.Columns(3).Width = 400
'DDCONTENIDO.Columns(4).Width = 1100
'
'DDCONTENIDO.Columns(5).Hidden = True
'DDCONTENIDO.Columns(6).Hidden = True
'DDCONTENIDO.Columns(7).Hidden = True
'DDCONTENIDO.Columns(8).Hidden = True
'DDCONTENIDO.Columns(9).Hidden = True
'DDCONTENIDO.Columns(10).Hidden = True
''DDCONTENIDO.Columns(2).Width = 1500
''DDCONTENIDO.Columns(3).Style.TextAlignment = sgAlignCenterCenter
''DDCONTENIDO.Columns(4).Style.TextAlignment = sgAlignRightCenter
''DDCONTENIDO.Columns(5).Hidden = True
'
'
'End Function


'---------------------------------------------------------------------------------------
' Procedure : Load_FE_Historico
' Author    : hsequeira
' Date      : 26/04/2024
' Purpose   : nNumFactura PASA A SER LONG EN VEZ DE INTEGER
'---------------------------------------------------------------------------------------
'
'Public Function Load_FE_Historico(nNumFactura As Integer) As Boolean
Public Function Load_FE_Historico(nNumFactura As Long) As Boolean
Dim cSQL As String
Set rsFETrans = New ADODB.Recordset

cSQL = "SELECT LIN, DESCRIP AS description, CANT AS quantity, PRECIO_UNIT  as price, 'und' as mu, PLU as internal_code,"
'cSQL = cSQL & " DESCRIP AS description, '01' as type, '01' as code,  ((PRECIO - FE_DESCUENTO) * (CON_TAX/100)) as amount, "
cSQL = cSQL & " '01' as type, '01' as code,  "
cSQL = cSQL & " ID_DESCUENTO as descuento, CON_TAX "
cSQL = cSQL & " FROM HIST_TR "
cSQL = cSQL & " WHERE NUM_TRANS = " & nNumFactura
cSQL = cSQL & " AND LEFT(TIPO,2) NOT IN ('VO','BV','DC','BE','EC') "
cSQL = cSQL & " ORDER BY LIN"

rsFETrans.Open cSQL, msConn, adOpenStatic, adLockOptimistic

If rsFETrans.RecordCount = 0 Then
    '
    'DDCONTENIDO.BackColor = vbRed
    Rem DDCONTENIDO.BackColor = &HC0&
    Call Load_FE_Transac(nNumFactura)
    Exit Function
Else
    Rem DDCONTENIDO.BackColor = &H8000000C
End If
'ShowMsg "AUN NO SE HA EMITIDO EL REPORTE Z", vbYellow, vbRed

DDCONTENIDO.DataMode = sgBound
Set DDCONTENIDO.DataSource = rsFETrans

DDCONTENIDO.Columns(1).Style.TextAlignment = sgAlignLeftCenter
DDCONTENIDO.Columns(1).Width = 400
DDCONTENIDO.Columns(2).Width = 3200
DDCONTENIDO.Columns(3).Width = 400
DDCONTENIDO.Columns(4).Width = 1100

DDCONTENIDO.Columns(5).Hidden = True
DDCONTENIDO.Columns(6).Hidden = True
DDCONTENIDO.Columns(7).Hidden = True
DDCONTENIDO.Columns(8).Hidden = True
DDCONTENIDO.Columns(9).Hidden = True
DDCONTENIDO.Columns(10).Hidden = True
'DDCONTENIDO.Columns(2).Width = 1500
'DDCONTENIDO.Columns(3).Style.TextAlignment = sgAlignCenterCenter
'DDCONTENIDO.Columns(4).Style.TextAlignment = sgAlignRightCenter
'DDCONTENIDO.Columns(5).Hidden = True


End Function

Private Function Load_FE_Transac(nNumFactura As Long) As Boolean
Dim cSQL As String
Set rsFETrans = New ADODB.Recordset

cSQL = "SELECT LIN, DESCRIP AS description, CANT AS quantity, PRECIO_UNIT  as price, 'und' as mu, PLU as internal_code,"
'cSQL = cSQL & " DESCRIP AS description, '01' as type, '01' as code,  ((PRECIO - FE_DESCUENTO) * (CON_TAX/100)) as amount, "
cSQL = cSQL & " '01' as type, '01' as code,  "
cSQL = cSQL & " ID_DESCUENTO as descuento, CON_TAX "
cSQL = cSQL & " FROM TRANSAC "
cSQL = cSQL & " WHERE NUM_TRANS = " & nNumFactura
cSQL = cSQL & " AND LEFT(TIPO,2) NOT IN ('VO','BV','DC','BE','EC') "
cSQL = cSQL & " ORDER BY LIN"

rsFETrans.Open cSQL, msConn, adOpenStatic, adLockOptimistic

If rsFETrans.RecordCount = 0 Then
    '
    'DDCONTENIDO.BackColor = vbRed
    DDCONTENIDO.BackColor = &HC0&
Else
    Rem DDCONTENIDO.BackColor = &H8000000C
    Rem DDCONTENIDO.BackColor = &H8000000C
End If
'ShowMsg "AUN NO SE HA EMITIDO EL REPORTE Z", vbYellow, vbRed

DDCONTENIDO.DataMode = sgBound
Set DDCONTENIDO.DataSource = rsFETrans

DDCONTENIDO.Columns(1).Style.TextAlignment = sgAlignLeftCenter
DDCONTENIDO.Columns(1).Width = 400
DDCONTENIDO.Columns(2).Width = 3200
DDCONTENIDO.Columns(3).Width = 400
DDCONTENIDO.Columns(4).Width = 1100

DDCONTENIDO.Columns(5).Hidden = True
DDCONTENIDO.Columns(6).Hidden = True
DDCONTENIDO.Columns(7).Hidden = True
DDCONTENIDO.Columns(8).Hidden = True
DDCONTENIDO.Columns(9).Hidden = True
DDCONTENIDO.Columns(10).Hidden = True
'DDCONTENIDO.Columns(2).Width = 1500
'DDCONTENIDO.Columns(3).Style.TextAlignment = sgAlignCenterCenter
'DDCONTENIDO.Columns(4).Style.TextAlignment = sgAlignRightCenter
'DDCONTENIDO.Columns(5).Hidden = True


End Function


Private Function HacerNC() As Boolean
Dim nFileNumber As Integer
Dim a$
Dim FileUbicacion As String
Dim NC_FileUbicacion As String
Dim cCadena As String
Dim cMonto As String
Dim docNumber As String

FileUbicacion = App.Path & "\HIST\" & lbFactura.Caption & ".txt"
NC_FileUbicacion = App.Path & "\HIST\NC" & lbFactura.Caption & ".txt"

ProgBar.value = 10
If Dir(NC_FileUbicacion) <> "" Then
    ShowMsg "YA SE HIZO NOTA DE CREDITO A ESTA FACTURA", vbRed, vbCyan
    ProgBar.value = 0
    HacerNC = False
    Exit Function
End If

ProgBar.value = 20
If Dir(FileUbicacion) = "" Then
    ShowMsg "FACTURA # " & lbFactura.Caption & ", NO ENCONTRADA (json)", vbYellow, vbRed
    ProgBar.value = 0
    HacerNC = False
    Exit Function
Else
    nFileNumber = FreeFile()
    Open FileUbicacion For Input As #nFileNumber

    Do Until EOF(nFileNumber)
        Line Input #nFileNumber, a$
    Loop
    Close #nFileNumber

    cMonto = Format(lbValor.Text, "#.00")
    '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    
    cCadena = Chr(34) & "total" & Chr(34) & ":" & Chr(34) & LTrim(RTrim(cMonto)) & Chr(34) & "," & _
                Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "01" & Chr(34)

    ProgBar.value = 50
    Call FindIt_And_ReplaceIt(a$, cCadena, 1)
    
    
    If cmdAplicaNC.Tag = "OFF-LINE" Then
        Rem cCadena = """F"""
        Rem Call FindIt_And_ReplaceIt(a$, cCadena, 1)
        docNumber = ExtractValue(a$, """docNumber""")
        Call FindIt_And_ReplaceIt(a$, docNumber, 1)
    Else
        cCadena = Chr(34) & "total" & Chr(34) & ":" & Chr(34) & LTrim(RTrim(cMonto)) & Chr(34) & "," & Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "01" & Chr(34)
        Call FindIt_And_ReplaceIt(a$, cCadena, 1)
    End If
    ProgBar.value = 50
    
    If Public_TestRequest(cNewSource) Then
        ProgBar.value = 70
        Call StoreJSON_In_OS(lbFactura.Caption, cNewSource, "NC")
        
        ProgBar.value = 80
        Call Add_NC_To_Registry_and_Database(lbFactura.Caption, cMonto, txtEmail.Text)
        ProgBar.value = 85
        Call Print_NC
        Call Print_NC(True)
        '||||||||||||||||||||||||||||
        ShowMsg "NOTA DE CREDITO REALIZADA CON EXITO", vbYellow, vbBlue
        '||||||||||||||||||||||||||||
        '||||||||||||||||||||||||||||
        HacerNC = True
        DDCONTENIDO.DataMode = sgBound
        Set DDCONTENIDO.DataSource = Nothing
    End If
    cNewSource = ""
End If

'ProgBar.value = 0
End Function

'---------------------------------------------------------------------------------------
' Procedure : Add_NC_To_Registry_and_Database
' Author    : hsequeira
' Date      : 03/11/2024
' Purpose   : SE AÑADE PARA QUE EL REPORTE Z LO PUEDA CONTABILIZAR
Rem LUEGO SE AGREGARA A LA BASE DE DATOS
Rem EL REPORTE X LO LEE, EL REPORTE Z LO LEE, LO IMPRIME y LO BORRA
'---------------------------------------------------------------------------------------
'
Private Sub Add_NC_To_Registry_and_Database(cFactura As String, cMonto As String, cEmail As String)

Dim nCurrenCounter As Long
Dim nCurrentValor As Single
Dim cTemp As String
Dim cTempFE As String
Dim nFECurrentCounter As Long
Dim nFECurrentValor As Single
Dim ncurrentcounter As Long
Dim ncurrentcounterFE As Long

cTemp = RegRead("HKCU\Software\SoloSoftware\SoloMix\NCCounter")
cTempFE = RegRead("HKCU\Software\SoloSoftware\SoloMix\FE_Counter")

' SE ENCARGA DEL CONTADOR DE NC
If cTemp = "" Then
    ncurrentcounter = 0
    nCurrentValor = 0
Else
    ncurrentcounter = CLng(cTemp)
    cTemp = RegRead("HKCU\Software\SoloSoftware\SoloMix\NCValor")
    nCurrentValor = CSng(cTemp)
End If

If cTempFE = "" Then
    ncurrentcounterFE = 0
    nFECurrentValor = 0
Else
    ncurrentcounterFE = CLng(cTempFE)
    cTempFE = RegRead("HKCU\Software\SoloSoftware\SoloMix\FE_Sales")
    nFECurrentValor = CSng(cTempFE)
End If

ncurrentcounter = ncurrentcounter + 1
nCurrentValor = nCurrentValor + CSng(cMonto)

ncurrentcounterFE = ncurrentcounterFE - 1
nFECurrentValor = nFECurrentValor - CSng(cMonto)

RegWrite "HKCU\Software\SoloSoftware\SoloMix\NCCounter", ncurrentcounter
RegWrite "HKCU\Software\SoloSoftware\SoloMix\NCValor", nCurrentValor

RegWrite "HKCU\Software\SoloSoftware\SoloMix\FE_Counter", ncurrentcounterFE
RegWrite "HKCU\Software\SoloSoftware\SoloMix\FE_Sales", nFECurrentValor


'Dim nCurrenCounter As Long
'Dim nCurrentValor As Single
'Dim cTemp As String
'
'cTemp = RegRead("HKCU\Software\SoloSoftware\SoloMix\NCCounter")
'If cTemp = "" Then
'    nCurrentCounter = 0
'    nCurrentValor = 0
'Else
'    nCurrentCounter = CLng(cTemp)
'    cTemp = RegRead("HKCU\Software\SoloSoftware\SoloMix\NCValor")
'    nCurrentValor = CSng(cTemp)
'End If
'Rem cTemp = CSng(cMonto)
'nCurrentCounter = nCurrentCounter + 1
'nCurrentValor = nCurrentValor + CSng(cMonto)
'
'RegWrite "HKCU\Software\SoloSoftware\SoloMix\NCCounter", nCurrentCounter
'RegWrite "HKCU\Software\SoloSoftware\SoloMix\NCValor", nCurrentValor

End Sub




'********************************************************************
      'Findit takes three arguments; two required, and one optional. The
      'required arguments, Box and Srch, are a RichTextBox object and a
      'string to search for. The optional argument, Start, is a Long
      'integer that is used in the recursive calls.
      '********************************************************************

Rem Private Function FindIt(Box As String, Srch As String, Optional Start As Long) As Boolean
Private Function FindIt_And_ReplaceIt(Box As String, Srch As String, Optional Start As Long) As Boolean

Dim retVal As Long      'Instr returns a long
Dim Source As String    'variable used in Instr
Dim cNuevoValor As String

   Source = Box   'put the text to search into the variable
   
   If Start = 0 Then Start = 1 'the initial call doesn't pass a value for Start, so it will equal 0

      retVal = InStr(Start, Source, Srch) 'do the first search, starting at the beginning of the text

      If retVal <> 0 Then  'there is at least one more occurrence of the string
'
        Start = retVal + Len(Srch) 'move the starting point past the first occurrence
        Rem cNuevoValor = Mid(Srch, 1, Len(Srch) - 2) & "6" & Chr(34)
        Rem cNewSource = Replace(Source, Srch, cNuevoValor)
        
        If cmdAplicaNC.Tag = "OFF-LINE" Then
            cNewSource = Replace(Source, """F""", """C""")
            bb = Replace(Srch, """", "")
            newNumDoc = "NC-" & bb
            cNewSource = Replace(cNewSource, bb, newNumDoc)
        Else
            cNuevoValor = Mid(Srch, 1, Len(Srch) - 2) & "6" & Chr(34)
            cNewSource = Replace(Source, Srch, cNuevoValor)
        End If
                

        
        FindIt_And_ReplaceIt = True
        

         'FindIt calls itself with new arguments
         'this is what makes it Recursive
         FindIt_And_ReplaceIt = 1 + FindIt_And_ReplaceIt(Box, Srch, Start)
      Else
        'cNewSource = ""
        FindIt_And_ReplaceIt = False
      End If
End Function

Private Sub UpdateControl(frm As Form, ByVal parametros As String)
    ' Update TextBox1 on the passed form
    frm.Coptr1.PrintNormal PtrSReceipt, parametros
    'RptCajas.Coptr1.PrintNormal parametros
End Sub

Private Sub Print_NC(Optional bCopia As Boolean)
Dim nValorMesas As Single
Dim MiLen1 As Integer
Dim Milen2 As Integer
Dim cMeseroName As String
Dim rsLAST As ADODB.Recordset
Dim rsLAST_Pago As ADODB.Recordset
Dim cSQL As String, cSQL2 As String
'Dim oFormulario As Form

' On Error Resume Next

'Select Case App.EXEName     'SoloMix4 o FastPos
'    Case "SoloMix4"
'        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'        Call OPOSTransactionPrint(RptCajas.Coptr1.name, "BEGIN")
'        oFormulario = "RptCajas"
'        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'    Else
'        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'        Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "BEGIN")
'        oFormulario = "Sys_Pos"
'        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'End Select
        
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call OPOSTransactionPrint(RptCajas.Coptr1.name, "BEGIN")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Open "c:\mifile.txt" For Output As #1   ' Open file.
'UpdateControl oFormulario, Date & Space(2) & Time & vbCrLf
'UpdateControl oFormulario, "=================================" & vbCrLf
'UpdateControl oFormulario, "HUMBERTO SEQUEIRA" & vbCrLf

RptCajas.Coptr1.PrintNormal PtrSReceipt, Date & Space(2) & Time & vbCrLf
RptCajas.Coptr1.PrintNormal PtrSReceipt, "=================================" & vbCrLf
RptCajas.Coptr1.PrintNormal PtrSReceipt, Chr(&H1B) & "!" & Chr(32) & Space(3) & "NOTA DE CREDITO" & vbCrLf
RptCajas.Coptr1.PrintNormal PtrSReceipt, Chr(&H1B) & Chr(&H46) & Space(6) & "FECHA/HORA: " & cFechaFactura & " - " & cHoraFactura & vbCrLf
RptCajas.Coptr1.PrintNormal PtrSReceipt, Space(9) & "FACTURA ORIGINAL # " & lbFactura.Caption & vbCrLf
RptCajas.Coptr1.PrintNormal PtrSReceipt, "eMail: " & Me.txtEmail & vbCrLf
If bCopia Then
    RptCajas.Coptr1.PrintNormal PtrSReceipt, Space(14) & "<< COPIA LOCAL >>" & vbCrLf
End If
RptCajas.Coptr1.PrintNormal PtrSReceipt, Chr(&HD) & Chr(&HA)
RptCajas.Coptr1.PrintNormal PtrSReceipt, "=================================" & vbCrLf
RptCajas.Coptr1.PrintNormal PtrSReceipt, rs00!DESCRIP & vbCrLf
'INFO: NO SALIA LA RAZON SOCIAL EN LOS REPORTES
RptCajas.Coptr1.PrintNormal PtrSReceipt, rs00!RAZ_SOC & vbCrLf
RptCajas.Coptr1.PrintNormal PtrSReceipt, "RUC:" & rs00!ruc & vbCrLf
RptCajas.Coptr1.PrintNormal PtrSReceipt, "SERIAL:" & rs00!Serial & vbCrLf
    
RptCajas.Coptr1.PrintNormal PtrSReceipt, Space(3) & vbCrLf

RptCajas.Coptr1.PrintNormal PtrSReceipt, "CUFE: " & FE_CUFE & vbCrLf
RptCajas.Coptr1.PrintNormal PtrSReceipt, Space(3) & vbCrLf

If FE_QRData <> "" Then
    With RptCajas.Coptr1
        .PrintBarCode PtrSReceipt, FE_QRData, PTR_BCS_QRCODE, .RecLineHeight * 0.6, .RecLineWidth * 0.6, PtrBcCenter, PtrBcTextNone&
                    '.PrintBarCode PtrSReceipt, FE_QRData, PTR_BCS_QRCODE, .RecLineHeight, .RecLineWidth, PtrBcCenter, PtrBcTextNone&
    End With
Else
End If

RptCajas.Coptr1.PrintNormal PtrSReceipt, "=================================" & vbCrLf

Set rsLAST = New ADODB.Recordset
cSQL = "SELECT * FROM TRANSAC WHERE NUM_TRANS = " & lbFactura.Caption & " ORDER BY LIN"
rsLAST.Open cSQL, msConn, adOpenStatic, adLockOptimistic

Set rsLAST_Pago = New ADODB.Recordset
cSQL2 = "SELECT  A.TIPO_PAGO, B.DESCRIP, A.MONTO "
cSQL2 = cSQL2 & " FROM TRANSAC_PAGO AS A, PAGOS AS B"
cSQL2 = cSQL2 & " WHERE A.NUM_TRANS = " & lbFactura.Caption
cSQL2 = cSQL2 & " AND A.TIPO_PAGO = B.CODIGO "
cSQL2 = cSQL2 & " ORDER BY A.LIN"
rsLAST_Pago.Open cSQL2, msConn, adOpenStatic, adLockOptimistic

Do While Not rsLAST.EOF
    RptCajas.Coptr1.PrintNormal PtrSReceipt, Format(Left(rsLAST!DESCRIP, 15), "!@@@@@@@@@@@@@@@") & Space(2) & _
                Format(Format(rsLAST!CANT, "###"), "@@@") & Space(2) & Format(Format(rsLAST!precio, "STANDARD"), "@@@@@@@@") & vbCrLf
        rsLAST.MoveNext
Loop

RptCajas.Coptr1.PrintNormal PtrSReceipt, "=================================" & vbCrLf

Do While Not rsLAST_Pago.EOF
    RptCajas.Coptr1.PrintNormal PtrSReceipt, Format(Left(rsLAST_Pago!DESCRIP, 15), "!@@@@@@@@@@@@@@@") & Space(2) & _
                    Format(Format(rsLAST_Pago!MONTO, "STANDARD"), "@@@@@@@@@@") & vbCrLf
    rsLAST_Pago.MoveNext
Loop

RptCajas.Coptr1.PrintNormal PtrSReceipt, Space(3) & vbCrLf
RptCajas.Coptr1.PrintNormal PtrSReceipt, Space(3) & vbCrLf

For i = 1 To 6
    RptCajas.Coptr1.PrintNormal PtrSReceipt, Chr(&HD) & Chr(&HA)
Next

RptCajas.Coptr1.CutPaper 100

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call OPOSTransactionPrint(RptCajas.Coptr1.name, "END")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

rsLAST.Close
rsLAST_Pago.Close

Set rsLAST = Nothing
Set rsLAST_Pago = Nothing
End Sub


Private Sub ChangeColorOfProgBar()
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    ''Call SendMessageLong(ProgBar.hwnd, PBM_SETBARCOLOR, 0&, ByVal 255)      'Red
    ''Call SendMessageLong(ProgBar.hwnd, PBM_SETBKCOLOR, 0&, ByVal 65535)    'yellow
If FE_Facil_ONLINE Then
    Call SendMessageLong(ProgBar.hwnd, PBM_SETBARCOLOR, 0&, ByVal 65280)       'green
    Call SendMessageLong(ProgBar.hwnd, PBM_SETBKCOLOR, 0&, ByVal 0)    'black
Else
    Call SendMessageLong(ProgBar.hwnd, PBM_SETBARCOLOR, 0&, ByVal 255)      'Red
    Call SendMessageLong(ProgBar.hwnd, PBM_SETBKCOLOR, 0&, ByVal 65535)    'yellow
End If
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
End Sub

Private Sub Picture_Click()

If cmdAplicaNC.Tag = "" Then
    If FE_Facil_ONLINE Then
        If txtEmail.Text = "" Then
            txtEmail.Text = RegRead("HKCU\Software\SoloSoftware\SoloMix\LastEmail")
        Else
            RegWrite "HKCU\Software\SoloSoftware\SoloMix\LastEmail", txtEmail.Text
        End If
    Else
    End If
End If

End Sub

Private Function ExtractValue(json As String, Key As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim value As String
    
    startPos = InStr(json, Key) + Len(Key) + 2 ' Position after the ": character
    endPos = InStr(startPos, json, ",")       ' Find coma
    If startPos > 0 And endPos > 0 Then
        value = Mid(json, startPos, endPos - startPos)
    End If
    
    ExtractValue = LTrim(RTrim(value))
End Function
