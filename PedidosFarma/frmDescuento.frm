VERSION 5.00
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Begin VB.Form frmDescuento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descuento ..."
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6810
   Icon            =   "frmDescuento.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6810
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optDeshacer 
      Caption         =   "Deshacer todos los descuentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   1080
      Width           =   3735
   End
   Begin AIFCmp1.asxPowerButton cmdOk 
      Height          =   615
      Left            =   4920
      TabIndex        =   1
      Top             =   3000
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      BorderStyle     =   4
      Picture         =   "frmDescuento.frx":030A
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtResul 
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
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txtImporte 
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
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtDescuento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   2520
      Width           =   1335
   End
   Begin VB.OptionButton optTotal 
      Caption         =   "Aplicar el Descuento sobre todos los items del facturador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   6375
   End
   Begin VB.OptionButton OptItems 
      Caption         =   "Sobre el Items Seleccionado"
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
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Value           =   -1  'True
      Width           =   3615
   End
   Begin AIFCmp1.asxPowerButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   5760
      TabIndex        =   2
      Top             =   3000
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      BorderStyle     =   4
      Picture         =   "frmDescuento.frx":075C
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Importe - Descuento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Width           =   2160
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Importe:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   7
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descuento %:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Width           =   1440
   End
End
Attribute VB_Name = "frmDescuento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsFacturando As New ADODB.Recordset
Private reg As Integer
Private vPor As Double 'para sacar porcentaje de descuento
Private Sub cmdCancel_Click()
rsFacturando.Close
Unload Me
End Sub

Private Sub cmdOk_Click()
If OptItems.Value = True Then
    rsFacturando!importe = Str(txtResul.Text)
    rsFacturando!Descuento = Str(txtDescuento.Text)
    rsFacturando.Update
End If

rsFacturando.MoveFirst
reg = rsFacturando.RecordCount

If optDeshacer = True Then
   For i = 1 To reg
        rsFacturando!importe = rsFacturando!precio
        rsFacturando!Descuento = 0
        rsFacturando.MoveNext
   Next
End If

If optTotal.Value = True Then
   For i = 1 To reg
        rsFacturando!importe = Round(rsFacturando!importe - ((rsFacturando!importe * txtDescuento.Text) / 100), 2)
        rsFacturando!Descuento = txtDescuento.Text
        rsFacturando.MoveNext
   Next
End If
rsFacturando.Requery
rsFacturando.Close
Unload Me
End Sub

Private Sub Form_Load()
OptItems.Value = True
rsFacturando.Open "select * from facturador", cn, adOpenDynamic, adLockOptimistic, adCmdText
rsFacturando.Find "idfactura = " & vIdFactura, , adSearchForward, 1
txtImporte.Text = rsFacturando!importe
End Sub

Private Sub optTotal_Click()
Dim vTot As Double
reg = rsFacturando.RecordCount
vTot = 0
rsFacturando.MoveFirst
For i = 1 To reg
    vTot = vTot + rsFacturando!precio
    rsFacturando.MoveNext
Next
txtImporte.Text = vTot
txtDescuento.SetFocus

End Sub
Private Sub txtDescuento_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        cmdOk.SetFocus
    Case 44
    Case 45
    Case 46
        KeyAscii = 44
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select

End Sub
Private Sub txtDescuento_LostFocus()
If Len(txtDescuento.Text) > 0 And Val(txtDescuento.Text) <= 100 Then
    vPor = (txtImporte.Text * txtDescuento.Text) / 100
    txtResul.Text = Round(txtImporte.Text - vPor, 2)
End If
End Sub
