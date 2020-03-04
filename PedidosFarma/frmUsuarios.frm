VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Begin VB.Form frmUsuarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuarios ..."
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
   Icon            =   "frmUsuarios.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Archivo de Usuarios"
      Height          =   4095
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   9255
      Begin MSDataGridLib.DataGrid dtgUsuarios 
         Height          =   3735
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Si desea eliminar un usuario presione la Tecla Supr sobre el registro"
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6588
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648384
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "idUsuario"
            Caption         =   "Cod.Usuario"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Nombre"
            Caption         =   "Nombre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "habilitado"
            Caption         =   "Habilitado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Nivel"
            Caption         =   "Nivel"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   1260,284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4334,74
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Usuario"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin AIFCmp1.asxPowerButton cmdGrabar 
         Height          =   375
         Left            =   7680
         TabIndex        =   6
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "&Grabar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextColor       =   32768
      End
      Begin VB.ComboBox cbonivel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmUsuarios.frx":014A
         Left            =   8280
         List            =   "frmUsuarios.frx":0157
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox cboHab 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmUsuarios.frx":0164
         Left            =   6000
         List            =   "frmUsuarios.frx":016E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtConfirma 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   10
         PasswordChar    =   "#"
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtClave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   10
         PasswordChar    =   "#"
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtNombre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   1
         Top             =   240
         Width           =   3495
      End
      Begin AIFCmp1.asxPowerButton cmdCancelar 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   7680
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "&Salir"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextColor       =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Confirmar Clave:"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   1470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nivel Autoriz"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6960
         TabIndex        =   13
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Habilitado ?:"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4680
         TabIndex        =   12
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clave (max.10 caracteres):"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   2370
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsUsuarios As New ADODB.Recordset
Private Sub cboHab_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cbonivel.SetFocus
End If
End Sub
Private Sub cboJefe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
End If
End Sub
Private Sub cbonivel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
End If
End Sub
Private Sub cmdGrabar_Click()
If Len(txtNombre.Text) = 0 Then
    MsgBox "DEBE INGRESAR UN NOMBRE DE USUARIO ...", vbExclamation, "ATENCION !"
    txtNombre.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If
If Len(txtClave.Text) = 0 Or Len(txtConfirma.Text) = 0 Then
    MsgBox "DEBE INGRESAR LAS CLAVES ...!", vbExclamation, "ATENCION !"
    txtClave.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If
If Trim(txtClave.Text) <> Trim(txtConfirma.Text) Then
    MsgBox "LA CONFIRMACION DE CLAVE ES INCORRECTA, REINGRESE CLAVE ...!", vbCritical, "ATENCION !"
    txtClave.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If
'verifica que no haya otro usuario con el mismo nombre
rsUsuarios.Find "nombre = '" & Trim(txtNombre.Text) & "'", , adSearchForward, 1
If rsUsuarios.EOF = False Then
    MsgBox "EL NOMBRE DE USUARIO YA EXISTE, INVENTE OTRO !", vbCritical, "ATENCION !"
    txtNombre.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If

rsUsuarios.AddNew
rsUsuarios!nombre = txtNombre.Text
rsUsuarios!clave = EnCrypt(txtClave.Text)
If cboHab.Text = "Si" Then
    rsUsuarios!habilitado = "S"
Else
    rsUsuarios!habilitado = "N"
End If
rsUsuarios!nivel = cbonivel.Text
rsUsuarios.Update
Me.dtgUsuarios.Refresh
txtNombre.Text = ""
txtClave.Text = ""
txtConfirma.Text = ""
txtNombre.SetFocus
End Sub
Private Sub cmdCancelar_Click()
Unload Me
End Sub
Private Sub dtgUsuarios_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If rsUsuarios.RecordCount > 0 Then
        frmPideClave.Show vbModal
        If TempNivel = 1 Then
            SioNo = MsgBox("ESTA SEGURO DE ELIMINAR ESTE USUARIO DEFINITIVAMENTE ???", vbInformation + vbYesNo, "ATENCION !")
            If SioNo = vbYes Then
                rsUsuarios.Delete
                rsUsuarios.Update
                dtgUsuarios.Refresh
            End If
        Else
            MsgBox "SU NIVEL DE AUTORIZACION NO LE PERMITE ELMINAR REGISTROS ...", vbExclamation, "Seguridad del Sistema...."
            Exit Sub
        End If
    End If
End If
End Sub
Private Sub Form_Load()
If rsUsuarios.State = 1 Then
    rsUsuarios.Close
End If
cbonivel.Text = 1
cboHab.Text = "Si"
rsUsuarios.Open "select * from usuarios", cn, adOpenDynamic, adLockOptimistic, adCmdText
Set dtgUsuarios.DataSource = rsUsuarios
dtgUsuarios.Refresh
End Sub
Private Sub Form_Unload(cancel As Integer)
If rsUsuarios.State = 1 Then
    rsUsuarios.Close
    Set rsUsuarios = Nothing
End If
End Sub
Private Sub txtClave_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtConfirma.SetFocus
End If
End Sub
Private Sub txtConfirma_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cboHab.SetFocus
End If
End Sub
Private Sub txtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtClave.SetFocus
End If
End Sub
Function EnCrypt(strCryptThis)
' para encriptar clave
    g_Key = "O_[:S&&]44AK;;^&*R?ZN^9_7LL),VG;;$=QY,JMM1*2*KW<^@I@T,3YY6V0]$2DA)+T0RZIOC`;>:FA[)6P)#=13N&"
    Dim strChar, iKeyChar, iStringChar, i
    For i = 1 To Len(strCryptThis)
       iKeyChar = Asc(Mid(g_Key, i, 1))
       iStringChar = Asc(Mid(strCryptThis, i, 1))
       iCryptChar = iStringChar + iKeyChar
       strEncrypted = strEncrypted & Chr(iCryptChar)
    Next
    EnCrypt = strEncrypted
End Function
