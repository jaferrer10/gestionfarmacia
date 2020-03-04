VERSION 5.00
Begin VB.Form frmPideClave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Permiso Usuario..."
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3825
   Icon            =   "frmPideClave.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtClave 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox txtUsuario 
      Height          =   350
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Usuario:"
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
      TabIndex        =   2
      Top             =   840
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuario:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   885
   End
End
Attribute VB_Name = "frmPideClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsVerUsu As New ADODB.Recordset
Private Sub Form_Load()
vUsu = ""
TempNivel = 0
rsVerUsu.Open "select * from usuarios", cn, adOpenDynamic, adLockReadOnly, adCmdText
If rsVerUsu.RecordCount = 0 Then
    MsgBox "NO HAY USUARIOS REGISTRADOS, LLAME AL ADMINISTRADOR ..!!!", vbCritical, "Atencion !!!"
    Unload Me
End If

End Sub
Private Sub Form_Unload(cancel As Integer)
rsVerUsu.Close
Set rsVerUsu = Nothing
End Sub
Private Sub txtClave_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    rsVerUsu.Find ("nombre = '" & Trim(txtUsuario.Text) & "'"), , adSearchForward, 1
    If rsVerUsu.EOF Then
        MsgBox "EL USUARIO INGRESADO NO EXISTE ...!", vbExclamation, "Usurio Incorrecto ..."
        txtUsuario.SetFocus
        'SendKeys "{home}+{end}"
        Exit Sub
    End If
    'funcion en el modulo que desencrita claves
    If txtClave.Text = (DeCrypt(rsVerUsu!clave)) Then
        TempNivel = rsVerUsu!nivel
        vUsu = rsVerUsu!nombre
    Else
        MsgBox "LA CLAVE ES INCORRECTA...!", vbExclamation, "Seguridad de Usuario..."
        TempNivel = 0
    End If
    Unload Me
End If
If KeyAscii = 27 Then
    Unload Me
End If
End Sub
Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtClave.SetFocus
ElseIf KeyAscii = 27 Then
    Unload Me
End If
End Sub
