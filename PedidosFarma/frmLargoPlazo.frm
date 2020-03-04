VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLargoPlazo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recalcular fecha de vencimiento por largo plazo..."
   ClientHeight    =   2835
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmLargoPlazo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker dtpFechaFac 
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   153157633
      CurrentDate     =   42951
   End
   Begin VB.TextBox txtDias 
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
      Left            =   3360
      MaxLength       =   3
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4680
      Picture         =   "frmLargoPlazo.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   495
      Left            =   3240
      Picture         =   "frmLargoPlazo.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpFechaVto 
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   153092097
      CurrentDate     =   42951
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Vencimiento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha factura:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad de días a vencer factura:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2985
   End
End
Attribute VB_Name = "frmLargoPlazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
rtaLargoPlazo = False
Unload Me
End Sub

Private Sub OKButton_Click()
rtaLargoPlazo = True
vFecVto = dtpFechaVto.Value
Unload Me
End Sub

Private Sub txtDias_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    OKButton.SetFocus
End If
End Sub

Private Sub txtDias_LostFocus()
If Len(txtDias.Text) > 0 Then
    'vFecVto = vFecVto + txtDias.Text
    dtpFechaVto.Value = dtpFechaFac.Value + txtDias.Text
Else
    vFecVto = dtpFechaVto.Value
End If

End Sub
