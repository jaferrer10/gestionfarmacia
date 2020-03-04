VERSION 5.00
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Begin VB.Form frmNuevoProductoPrestamo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nuevo Producto ..."
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   Icon            =   "frmNuevoProductoPrestamo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDescripcion 
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
      Left            =   120
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1440
      Width           =   7335
   End
   Begin VB.TextBox txtTroquel 
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
      Left            =   120
      MaxLength       =   20
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin AIFCmp1.asxPowerButton cmdGrabar 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      Picture         =   "frmNuevoProductoPrestamo.frx":058A
      Caption         =   "&Grabar"
      CaptionAlignment=   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignment=   3
   End
   Begin AIFCmp1.asxPowerButton cmdCancelar 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      Picture         =   "frmNuevoProductoPrestamo.frx":09DC
      Caption         =   "&Cancelar"
      CaptionAlignment=   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignment=   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descripcion del Producto:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2700
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nº Troquel:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "frmNuevoProductoPrestamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
Unload Me
End Sub
Private Sub cmdGrabar_Click()

If Len(txtDescripcion.Text) = 0 Then
    MsgBox "DEBE INGRESAR UNA DESCRIPCION DEL NUEVO PRODUCTO ...!", vbExclamation, "ATENCION !"
    txtDescripcion.SetFocus
    Exit Sub
End If
'agrega en la tabla de productos
If Len(txtTroquel.Text) > 0 Then
    cn.Execute "insert into productos (troquel, Descripcion) values (" & txtTroquel.Text & ", '" & LTrim(txtDescripcion.Text) & "')"
End If

Unload Me

End Sub
Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdGrabar.SetFocus
End If
End Sub
Private Sub txtTroquel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtDescripcion.SetFocus
End If
End Sub
