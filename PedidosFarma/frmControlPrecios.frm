VERSION 5.00
Begin VB.Form frmControlPrecios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Precios ..."
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "frmControlPrecios.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   6210
   Begin VB.TextBox txtLista 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   4440
      TabIndex        =   10
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox txtPublico 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   4440
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtCosto 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtMargen 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtIva 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Control de Precios"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Precio de Lista           ="
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   420
      Left            =   360
      TabIndex        =   12
      Top             =   4200
      Width           =   3630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Precio Publico           ="
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   420
      Left            =   360
      TabIndex        =   5
      Top             =   3480
      Width           =   3600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Precio Costo Factura ="
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   420
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   3645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "% Margen                   ="
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   420
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "IVA %                         ="
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   420
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   3585
   End
End
Attribute VB_Name = "frmControlPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private xPublico As Double
Private xCostoIva As Double
Private xPLista As Double
Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub Form_Load()
Me.Top = 3000
Me.Left = 100
txtIva.Text = 21
txtMargen.Text = 40
End Sub
Private Sub txtCosto_KeyPress(KeyAscii As Integer)
Err.Clear
On Error GoTo Solucion
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        Call CalculaPrecio
        SendKeys "{end}+{home}"
    Case 44
    Case 45
    Case 46
        KeyAscii = 44
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select

Exit Sub
Solucion:
    MsgBox Err.Number & " - " & Err.Description, vbInformation, "Error del Sistema..."

End Sub
Private Sub txtIva_Change()
Call CalculaPrecio
End Sub
Private Sub txtIva_KeyPress(KeyAscii As Integer)

Err.Clear
On Error GoTo Solucion

Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtMargen.SetFocus
        SendKeys "{end}+{home}"
    Case 44
    Case 45
    Case 46
        KeyAscii = 44
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select

Exit Sub
Solucion:
    MsgBox Err.Number & " - " & Err.Description, vbInformation, "Error del Sistema..."

End Sub
Private Sub txtLista_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtMargen_Change()
    Call CalculaPrecio
End Sub
Private Sub txtMargen_KeyPress(KeyAscii As Integer)
Err.Clear
On Error GoTo Solucion

Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtCosto.SetFocus
        SendKeys "{end}+{home}"
    Case 44
    Case 45
    Case 46
        KeyAscii = 44
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select

Exit Sub
Solucion:
    MsgBox Err.Number & " - " & Err.Description, vbInformation, "Error del Sistema..."

End Sub
Private Sub CalculaPrecio()
Err.Clear
On Error GoTo Solucion

If Len(txtCosto.Text) = 0 Then
    Exit Sub
End If
xCostoIva = txtCosto.Text * Val("1." & txtIva.Text)
xPublico = xCostoIva * Val("1." & txtMargen.Text)
txtPublico.Text = Round(CDbl(xPublico), 2)
'calcula el precio lista sumandole un 11.05 %
txtLista.Text = Round(((CDbl(txtPublico.Text) * 11.05 / 100) + CDbl(txtPublico.Text)), 2)

Exit Sub
Solucion:
    MsgBox Err.Number & " - " & Err.Description, vbInformation, "Error del Sistema..."

End Sub
Private Sub txtPublico_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case Else: KeyAscii = 0
End Select
End Sub
