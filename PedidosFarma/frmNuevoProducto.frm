VERSION 5.00
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Begin VB.Form frmNuevoProducto 
   Caption         =   "Nuevo Producto ..."
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   Icon            =   "frmNuevoProducto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCantidad 
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
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
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
      TabIndex        =   2
      Top             =   1440
      Width           =   2655
   End
   Begin AIFCmp1.asxPowerButton cmdGrabar 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      Picture         =   "frmNuevoProducto.frx":058A
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
      Top             =   480
      Width           =   7335
   End
   Begin AIFCmp1.asxPowerButton cmdCancelar 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   2160
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      Picture         =   "frmNuevoProducto.frx":09DC
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad a Pedir:"
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
      Left            =   3000
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
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
      TabIndex        =   6
      Top             =   1080
      Width           =   1200
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
      TabIndex        =   0
      Top             =   120
      Width           =   2700
   End
End
Attribute VB_Name = "frmNuevoProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsNuevoP As New ADODB.Recordset
Private Sub cmdCancelar_Click()
Unload Me
End Sub
Private Sub cmdGrabar_Click()
If IsNumeric(txtCantidad.Text) = False Then
    MsgBox "SOLO SE ADMITEN DIGITOS ...!", vbCritical, "ATENCION !"
    txtCantidad.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If

If Len(txtDescripcion.Text) = 0 Then
    MsgBox "DEBE INGRESAR UNA DESCRIPCION DEL NUEVO PRODUCTO ...!", vbExclamation, "ATENCION !"
    txtDescripcion.SetFocus
    Exit Sub
End If
'graba en la tabla de pedidos
rsNuevoP.AddNew
rsNuevoP!fecha = frmPedidos.dtpFecha.Value
rsNuevoP!troquel = txtTroquel.Text
rsNuevoP!descripcion = txtDescripcion.Text
rsNuevoP!cantidad = txtCantidad.Text
rsNuevoP!idproveedor = vidPro
rsNuevoP!estado = 3
rsNuevoP.Update
rsNuevoP.Close
Set rsNuevoP = Nothing

'si tiene codigo de barra lo agrega en la tabla de productos
If Len(txtTroquel.Text) > 0 Then
    cn.Execute "insert into productos (troquel, Descripcion, fabricante) values (" & txtTroquel.Text & ", '" & LTrim(txtDescripcion.Text) & "', " & vidPro & ")"
End If


Unload Me
End Sub
Private Sub Form_Load()
rsNuevoP.Open "select * from pedidos", cn, adOpenDynamic, adLockOptimistic, adCmdText
End Sub
Private Sub Form_Unload(cancel As Integer)
If rsNuevoP.State = 1 Then
    rsNuevoP.Close
    Set rsNuevoP = Nothing
End If
End Sub
Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdGrabar.SetFocus
End If
End Sub
Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtTroquel.SetFocus
End If
End Sub
Private Sub txtTroquel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtCantidad.SetFocus
End If
End Sub
