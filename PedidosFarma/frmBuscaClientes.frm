VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBuscaClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Archivo de Clientes ..."
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9930
   Icon            =   "frmBuscaClientes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9930
   Begin VB.Frame Frame1 
      Caption         =   "Lista de Clientes"
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.Frame frameCarga 
         Caption         =   "Datos del Cliente"
         Height          =   5655
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   8415
         Begin MSDataListLib.DataCombo dtcIva 
            Height          =   360
            Left            =   2040
            TabIndex        =   20
            Top             =   3720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin AIFCmp1.asxPowerButton cmdGraba 
            Height          =   615
            Left            =   4680
            TabIndex        =   18
            Top             =   4680
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1085
            Picture         =   "frmBuscaClientes.frx":058A
            Caption         =   "&Grabar"
            CaptionAlignment=   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PictureAlignment=   0
         End
         Begin VB.TextBox txtCuit 
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
            Left            =   5040
            TabIndex        =   17
            Top             =   3720
            Width           =   2415
         End
         Begin VB.TextBox txtObser 
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
            Left            =   1920
            TabIndex        =   16
            Top             =   3000
            Width           =   5535
         End
         Begin MSComCtl2.DTPicker dtpNac 
            Height          =   375
            Left            =   5880
            TabIndex        =   15
            Top             =   2280
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   16711680
            CalendarTitleForeColor=   16711680
            Format          =   68878337
            CurrentDate     =   39365
         End
         Begin VB.TextBox txtTel 
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
            Left            =   1920
            TabIndex        =   14
            Top             =   2280
            Width           =   2295
         End
         Begin VB.TextBox txtDir 
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
            Left            =   1920
            TabIndex        =   13
            Top             =   1560
            Width           =   5535
         End
         Begin VB.TextBox txtApellido 
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
            Left            =   1920
            TabIndex        =   11
            Top             =   840
            Width           =   5535
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
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1920
            TabIndex        =   9
            Top             =   240
            Width           =   5535
         End
         Begin AIFCmp1.asxPowerButton cmdCancelar 
            Height          =   615
            Left            =   6240
            TabIndex        =   19
            Top             =   4680
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1085
            Picture         =   "frmBuscaClientes.frx":09DC
            Caption         =   "&Cancelar"
            CaptionAlignment=   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PictureAlignment=   0
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "CUIT:"
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
            Left            =   4080
            TabIndex        =   12
            Top             =   3840
            Width           =   600
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "F.Nacimiento:"
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
            Left            =   4320
            TabIndex        =   10
            Top             =   2400
            Width           =   1440
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Condicion IVA:"
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
            Left            =   240
            TabIndex        =   8
            Top             =   3840
            Width           =   1530
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
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
            Left            =   240
            TabIndex        =   7
            Top             =   3120
            Width           =   1650
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Telefono:"
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
            Left            =   240
            TabIndex        =   6
            Top             =   2400
            Width           =   1005
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Dirección:"
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
            Left            =   240
            TabIndex        =   5
            Top             =   1680
            Width           =   1065
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Apellidos:"
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
            Left            =   240
            TabIndex        =   4
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nombres:"
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
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   1020
         End
      End
      Begin AIFCmp1.asxToolButton cmdAgregar 
         Height          =   975
         Left            =   8520
         ToolTipText     =   "Agrega cliente nuevo"
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1720
         BorderStyle     =   4
         Picture         =   "frmBuscaClientes.frx":0E2E
         Caption         =   "&Agreg"
         CaptionAlignment=   7
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
      Begin MSDataGridLib.DataGrid dtgLisClie 
         Height          =   5295
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   9340
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   8454143
         HeadLines       =   1
         RowHeight       =   19
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin AIFCmp1.asxToolButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   975
         Left            =   8520
         Top             =   4440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1720
         BorderStyle     =   4
         Picture         =   "frmBuscaClientes.frx":1280
         Caption         =   "&Salir"
         CaptionAlignment=   7
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
      Begin AIFCmp1.asxToolButton cmdModificar 
         Height          =   975
         Left            =   8520
         ToolTipText     =   "Modifica datos del cliente"
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1720
         BorderStyle     =   4
         Picture         =   "frmBuscaClientes.frx":16D2
         Caption         =   "&Modif"
         CaptionAlignment=   7
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
      Begin AIFCmp1.asxToolButton cmdSelec 
         Height          =   975
         Left            =   8520
         ToolTipText     =   "Selecciona cliente para la facturacion"
         Top             =   2760
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1720
         BorderStyle     =   4
         Picture         =   "frmBuscaClientes.frx":1B24
         Caption         =   "&Selecc"
         CaptionAlignment=   7
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
   End
End
Attribute VB_Name = "frmBuscaClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LisCli As New ADODB.Recordset
Private rsCond As New ADODB.Recordset

Private Sub cmdAgregar_Click()
vAgrega = True
frameCarga.Visible = True
txtNombre.Text = ""
txtApellido.Text = ""
txtDir.Text = ""
txtTel.Text = ""
txtObser.Text = ""
txtCuit.Text = ""
txtNombre.SetFocus
End Sub

Private Sub cmdCancelar_Click()
frameCarga.Visible = False
End Sub
Private Sub cmdGraba_Click()
If vAgrega = True Then
    LisCli.AddNew
End If
If Len(txtNombre.Text) = 0 Then
    MsgBox "DEBE INGRESAR EL NOMBRE DEL CLIENTE !", vbExclamation, "ATENCION !"
    txtNombre.SetFocus
    Exit Sub
End If
If Len(txtApellido.Text) = 0 Then
    MsgBox "DEBE INGRESAR EL APELLIDO DEL CLIENTE !", vbExclamation, "ATENCION !"
    txtApellido.SetFocus
    Exit Sub
End If
LisCli!nombre = txtNombre.Text
LisCli!apellido = txtApellido.Text
LisCli!direccion = txtDir.Text
LisCli!telefono = txtTel.Text
LisCli!fechanac = dtpNac.Value
LisCli!observaciones = txtObser.Text
LisCli!condiva = dtcIva.BoundText
LisCli!cuit = txtCuit.Text
LisCli.Update
dtgLisClie.Refresh
frameCarga.Visible = False
dtgLisClie.SetFocus
End Sub

Private Sub cmdModificar_Click()
vAgrega = False
frameCarga.Visible = True
txtNombre.Text = LisCli!nombre
txtApellido.Text = LisCli!apellido
txtDir.Text = LisCli!direccion
txtTel.Text = LisCli!telefono & ""
dtpNac.Value = LisCli!fechanac
dtcIva.BoundText = LisCli!condiva
txtCuit.Text = LisCli!cuit & ""
txtObser.Text = LisCli!observaciones
txtNombre.SetFocus

End Sub

Private Sub cmdSalir_Click()
rsCond.Close
LisCli.Close
Unload Me
End Sub
Private Sub cmdSelec_Click()
vIdCliente = LisCli!idcliente
rsCond.Close
LisCli.Close
Unload Me
End Sub
Private Sub dtgLisClie_DblClick()
vIdCliente = LisCli!idcliente
rsCond.Close
LisCli.Close
Unload Me
End Sub

Private Sub Form_Load()
frameCarga.Visible = False
LisCli.Open "select * from clientes order by apellido", cn, adOpenDynamic, adLockOptimistic, adCmdText
Set dtgLisClie.DataSource = LisCli
dtgLisClie.Refresh

'llena el combo de condicion IVA
rsCond.Open "select * from condicioniva", cn, adOpenDynamic, adLockReadOnly, adCmdText
Set dtcIva.DataSource = rsCond
Set dtcIva.RowSource = rsCond
dtcIva.ListField = "condicion"
dtcIva.BoundColumn = "idcondicion"
End Sub
Private Sub txtApellido_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtDir.SetFocus
End If
End Sub
Private Sub txtCuit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdGraba.SetFocus
End If
End Sub
Private Sub txtDir_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtTel.SetFocus
End If
End Sub
Private Sub txtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtApellido.SetFocus
End If
End Sub
Private Sub txtObser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    dtcIva.SetFocus
End If
End Sub
Private Sub txtTel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    dtpNac.SetFocus
End If
End Sub
