VERSION 5.00
Begin VB.Form Acompanantes 
   BackColor       =   &H00B39665&
   BorderStyle     =   0  'None
   Caption         =   "cmdMesas"
   ClientHeight    =   2835
   ClientLeft      =   0
   ClientTop       =   9465
   ClientWidth     =   12585
   Icon            =   "Acompanantes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   12585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdMesas 
      BackColor       =   &H00C0FFC0&
      Caption         =   "<<PARA LLEVAR>>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   560
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "1"
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H0000FFFF&
      Caption         =   "APLICAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10800
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H0000FFFF&
      Caption         =   "REGRESAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10800
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Shape ShapeFondo 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   2535
      Left            =   120
      Top             =   120
      Width           =   10575
   End
End
Attribute VB_Name = "Acompanantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SOLO_COLOR_ROJO = &H8080FF
Private Const SOLO_COLOR_VERDE = &HC0FFC0
Private Const SOLO_COLOR_AMBAR = &H80FF
Private aArregAcompa(30) As String
Private ind As Integer
Private SOLO_FECHA As String
Private nCantMesas As Integer
Private nMesaInicial As Integer
Dim rsTmpAco As ADODB.Recordset

Public Function LoadAcompanantes()
Dim cSQL As String
Dim iCnt As Byte
Dim sData As String
Dim i As Integer
i = 1

SOLO_FECHA = Format(Date, "YYYYMMDD")
Set rsTmpAco = New ADODB.Recordset

'INFO: UPDATE 21ABR2016 (cSQL)
cSQL = "SELECT A.PLU_ID,A.ACOMP_ID,B.DESCRIP "
cSQL = cSQL & " FROM PLU_ACOMP AS A, ACOMPA AS B "
cSQL = cSQL & " WHERE A.PLU_ID = " & nPLUPassed
cSQL = cSQL & " AND A.ACOMP_ID = B.CODIGO "
cSQL = cSQL & " ORDER BY B.DESCRIP "

rsTmpAco.Open cSQL, msConn, adOpenStatic, adLockOptimistic

iCnt = 1
'Load cmdMesas(iCnt)

Do While Not rsTmpAco.EOF

   Call ShowMesa(iCnt, (cmdMesas(iCnt - 1).Left + cmdMesas(iCnt - 1).Width + GetAncho(nCantMesas)), _
                        cmdMesas(iCnt - 1).Top, rsTmpAco!DESCRIP)
    iCnt = iCnt + 1

    rsTmpAco.MoveNext
Loop

rsTmpAco.Close
Set rsTmpAco = Nothing

Call MuestraMesas
Me.Show
End Function

Private Function ShowMesa(iNumeroDeMesa As Byte, iLeft As Long, iTop As Long, Optional cTitulo As String) As Long
'INFO: MAY2010
'FUNCION QUE MUESTRA EN PANTALLA LOS BOTONES DE LAS MESAS
Dim nLocalTop As Long

'If IsObject(cmdMesas(iNumeroDeMesa)) Then
Load cmdMesas(iNumeroDeMesa)
cmdMesas(iNumeroDeMesa).Visible = True
cmdMesas(iNumeroDeMesa).Caption = cTitulo
cmdMesas(iNumeroDeMesa).BackColor = SOLO_COLOR_VERDE
cmdMesas(iNumeroDeMesa).Tag = iNumeroDeMesa + 1

Select Case nCantMesas
    Case 70
        Select Case iNumeroDeMesa
            'Case 7, 14, 21, 28, 35, 42, 49, 56, 63
            'Case 6, 13, 20, 27, 34, 41, 48, 55, 62
            Case 5, 10, 15, 20, 25, 40, 47, 54, 61
                Debug.Print cTitulo
                ShowMesa = iTop + 560
                nLocalTop = iTop + 560
                cmdMesas(iNumeroDeMesa).Left = 240
            Case Else
                ShowMesa = iTop
                nLocalTop = iTop
                cmdMesas(iNumeroDeMesa).Left = iLeft
        End Select
End Select

cmdMesas(iNumeroDeMesa).Top = nLocalTop
End Function

Private Sub cmdAceptar_Click()
''nDeptoPassed
''nPLUPassed
Call AddAcompanantes
End Sub
Private Sub AddAcompanantes()
Dim cSQL As String

For i = 0 To cmdMesas.Count - 1
    If cmdMesas(i).BackColor = SOLO_COLOR_ROJO Then
        CajLin = CajLin + 1
        cSQL = "INSERT INTO TMP_TRANS "
        cSQL = cSQL & "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,IMPRESO,IMPRESORA, CON_TAX) "
        cSQL = cSQL & " VALUES ("
        cSQL = cSQL & "" & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & "," & -1 & "," & CajLin & "," & "' @@ "
        'INFO: 3OCT2017
        cSQL = cSQL & cmdMesas(i).Caption & "'" & "," & nCantidad & "," & 0 & "," & 0 & ","
        cSQL = cSQL & 0 & "," & 0# & "," & 0# & "," & "'" & SOLO_FECHA & "'" & "," & "'" & Time & "'"
        cSQL = cSQL & ",'  '," & 0# & "," & 0 & ",0," & 1
        'INFO: AG02009
        cSQL = cSQL & "," & 0 & ")"
        
        msConn.BeginTrans
        msConn.Execute cSQL
        msConn.CommitTrans
    End If
Next

For i = 0 To cmdMesas.Count - 1
    If cmdMesas(i).BackColor = SOLO_COLOR_ROJO Then
        cmdMesas(i).BackColor = SOLO_COLOR_VERDE
    End If
Next

Call MuestraProductosMarcados
Unload Me
End Sub

Private Sub cmdMesas_Click(Index As Integer)
If cmdMesas(Index).BackColor = SOLO_COLOR_ROJO Then
    cmdMesas(Index).BackColor = SOLO_COLOR_VERDE
Else
    cmdMesas(Index).BackColor = SOLO_COLOR_ROJO
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub DD_PEDDETALLE_Click()
If DD_PEDDETALLE.Rows.Current.Style.ForeColor = vbRed Then
    DD_PEDDETALLE.Rows.Current.Style.ForeColor = vbBlack
    aArregAcompa(ind) = ""
    ind = ind - 1
Else
    DD_PEDDETALLE.Rows.Current.Style.ForeColor = vbRed
    aArregAcompa(ind) = " @@ " & DD_PEDDETALLE.Rows.Current.Cells(2).value
    ind = ind + 1
    
End If
DD_PEDDETALLE.Redraw
End Sub

Private Sub Form_Load()
''SOLO_FECHA = Format(Date, "YYYYMMDD")
nMesaInicial = 1
nCantMesas = 70
''Call LoadAcompanantes
Call MuestraMesas
End Sub

Private Sub List_Click()
List.BackColor = vbRed
End Sub

Private Sub MuestraMesas()
'INFO: MAY2010
'PONE LOS BOTONES EN PANTALLA SEGUN LA RESOLUCION DEL BOTON PRINCIPAL
'Y EL NUMERO DE MESAS
Dim iCnt As Byte

On Error GoTo ErrAdm:

cmdMesas(0).Height = 615
Select Case nCantMesas
    Case 70     '70 MESAS
        cmdMesas(0).Width = 1935
    Case 90    '90 MESAS
        cmdMesas(0).Width = 1140
    Case 110    '110 MESAS
        cmdMesas(0).Width = 930
    Case Else   '70 DEFAULT MESAS
        cmdMesas(0).Width = 1455
End Select

'For iCnt = 0 To nCantMesas - 1
    'Call ShowMesa(iCnt, (cmdMesas(iCnt - 1).Left + cmdMesas(iCnt - 1).Width + GetAncho(nCantMesas)), cmdMesas(iCnt - 1).Top)
'    Call ShowMesa(iCnt, (cmdMesas(iCnt - 1).Left + cmdMesas(iCnt - 1).Width + GetAncho(nCantMesas)), cmdMesas(iCnt - 1).Top)
'Next
On Error GoTo 0
Exit Sub

ErrAdm:
Debug.Print "Error.MuestraMesas." & Err.Number & " - " & Err.Description
End Sub


Private Function GetAncho(nNumMesas As Integer) As Byte
Select Case nNumMesas
    Case 70
        GetAncho = 105
    Case 90
        GetAncho = 81
    Case 110
        GetAncho = 65
    Case Else
        GetAncho = 105
End Select
End Function

Private Function GetMesa_AREA(iNumMesa As Integer) As Long
'INFO: MARZO 2010
'REGRESA EL NUMERO DEL COLOR DE LA MESA
On Error Resume Next
rsMESAS_Color.MoveFirst
On Error GoTo 0

On Error GoTo ErrAdm:
rsMESAS_Color.Find "Mesa = " & iNumMesa
If rsMESAS_Color.EOF Then
    GetMesa_AREA = "&HFF00"
    GetMesa_AREA = SOLO_COLOR_VERDE
Else
    GetMesa_AREA = RGB(rsMESAS_Color!Color1, rsMESAS_Color!Color2, rsMESAS_Color!Color3)
End If
On Error GoTo 0
Exit Function

ErrAdm:
'MsgBox "Error.GetMesa_AREA: " & Err.Number & " - " & Err.Description
Debug.Print "Error.GetMesa_AREA: " & Err.Number & " - " & Err.Description
End Function
