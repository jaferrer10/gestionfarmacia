VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Begin VB.Form frmGestionProveedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestion de Proveedores ..."
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11160
   Icon            =   "frmGestionProveedores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11160
   Begin AIFCmp1.asxPowerButton cmdSalir 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   9240
      TabIndex        =   6
      Top             =   2760
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      Picture         =   "frmGestionProveedores.frx":014A
      Caption         =   "&Salir"
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
      PictureAlignment=   6
   End
   Begin AIFCmp1.asxPowerButton cmdEliminar 
      Height          =   495
      Left            =   9240
      TabIndex        =   5
      Top             =   1920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      Picture         =   "frmGestionProveedores.frx":06D6
      Caption         =   "&Eliminar"
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
   Begin AIFCmp1.asxPowerButton cmdModificar 
      Height          =   495
      Left            =   9240
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      Picture         =   "frmGestionProveedores.frx":0B28
      Caption         =   "&Modificar"
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
   Begin AIFCmp1.asxPowerButton cmdAgregar 
      Height          =   495
      Left            =   9240
      TabIndex        =   3
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      Picture         =   "frmGestionProveedores.frx":0F7A
      Caption         =   "&Agregar"
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
   Begin VB.Frame frameDatos 
      Caption         =   "Datos del Proveedor"
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   9015
      Begin AIFCmp1.asxPowerButton cmdCancelar 
         Height          =   615
         Left            =   4680
         TabIndex        =   18
         Top             =   2400
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         Picture         =   "frmGestionProveedores.frx":13CC
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
         PictureAlignment=   6
      End
      Begin AIFCmp1.asxPowerButton cmdGrabar 
         Height          =   615
         Left            =   2280
         TabIndex        =   17
         Top             =   2400
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         Picture         =   "frmGestionProveedores.frx":181E
         Caption         =   "&Grabar Datos"
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
         PictureAlignment=   6
      End
      Begin VB.TextBox txtObservaciones 
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
         Left            =   2280
         TabIndex        =   16
         Top             =   1680
         Width           =   6615
      End
      Begin VB.TextBox txtMinimo 
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
         Left            =   6600
         TabIndex        =   15
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtMaximo 
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
         Left            =   2280
         TabIndex        =   14
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtTelefono 
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
         Left            =   2280
         TabIndex        =   13
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox txtNombre 
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
         Left            =   2280
         TabIndex        =   12
         Top             =   240
         Width           =   6135
      End
      Begin AIFCmp1.asxLabel asxLabel1 
         Height          =   270
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
         Caption         =   "Nombre Proveedor:"
         AutoSize        =   -1  'True
         UseMnemonic     =   -1  'True
         MouseIcon       =   "frmGestionProveedores.frx":1D25
      End
      Begin AIFCmp1.asxLabel asxLabel2 
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
         Caption         =   "Telefono:"
         AutoSize        =   -1  'True
         UseMnemonic     =   -1  'True
         MouseIcon       =   "frmGestionProveedores.frx":203F
      End
      Begin AIFCmp1.asxLabel asxLabel3 
         Height          =   270
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
         Caption         =   "Monto Maximo:"
         AutoSize        =   -1  'True
         UseMnemonic     =   -1  'True
         MouseIcon       =   "frmGestionProveedores.frx":2359
      End
      Begin AIFCmp1.asxLabel asxLabel4 
         Height          =   270
         Left            =   5040
         TabIndex        =   10
         Top             =   1320
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
         Caption         =   "Monto Minimo:"
         AutoSize        =   -1  'True
         UseMnemonic     =   -1  'True
         MouseIcon       =   "frmGestionProveedores.frx":2673
      End
      Begin AIFCmp1.asxLabel asxLabel5 
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
         Caption         =   "Observaciones:"
         AutoSize        =   -1  'True
         UseMnemonic     =   -1  'True
         MouseIcon       =   "frmGestionProveedores.frx":298D
      End
   End
   Begin VB.Frame frameLista 
      Caption         =   "Archivo de Proveedores "
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin MSDataGridLib.DataGrid dtgLista 
         Height          =   4095
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   7223
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "idproveedor"
            Caption         =   "Codigo"
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
            DataField       =   "nombre"
            Caption         =   "Nombre Proveedor"
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
            DataField       =   "telefono"
            Caption         =   "Telefono"
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
            DataField       =   "observaciones"
            Caption         =   "Observaciones"
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
         BeginProperty Column04 
            DataField       =   "montomaximo"
            Caption         =   "Monto Maximo"
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
         BeginProperty Column05 
            DataField       =   "montominimo"
            Caption         =   "Monto Minimo"
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
            BeginProperty Column00 
               ColumnWidth     =   659,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3644,788
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmGestionProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsListaPro As New ADODB.Recordset
Private Sub cmdAgregar_Click()
frameLista.Enabled = False
frameDatos.Visible = True
vAgrega = True
txtNombre.SetFocus
End Sub
Private Sub cmdCancelar_Click()
Call BlanqueaCampos
End Sub
Private Sub cmdEliminar_Click()
SioNo = MsgBox("ESTA SEGURO DE ELIMINAR ESTE PROVEEDOR ?", vbInformation + vbYesNo, "ATENCION !")
If SioNo = vbYes Then
    rsListaPro.Delete
    rsListaPro.Update
    dtgLista.Refresh
    dtgLista.SetFocus
End If
End Sub
Private Sub cmdGrabar_Click()
If Len(txtNombre.Text) = 0 Then
    MsgBox "DEBE INGRESAR POR LO MENOS EL NOMBRE DEL PROVEEDOR...", vbCritical, "ATENCION !"
    txtNombre.SetFocus
    Exit Sub
End If
SioNo = MsgBox("ESTA SEGURO DE GRABAR LOS DATOS ?", vbInformation + vbYesNo, "ATENCION !")
If SioNo = vbYes Then
    If vAgrega = True Then
        rsListaPro.AddNew
    End If
    rsListaPro!nombre = txtNombre.Text
    rsListaPro!telefono = txtTelefono.Text
    If Len(txtMaximo.Text) > 0 Then
        rsListaPro!montomaximo = Val(txtMaximo.Text)
    End If
    If Len(txtMinimo.Text) > 0 Then
        rsListaPro!montominimo = Val(txtMinimo.Text)
    End If
    rsListaPro!observaciones = txtObservaciones.Text
    rsListaPro.Update
    dtgLista.Refresh
    Call BlanqueaCampos
End If
vAgrega = True
End Sub

Private Sub cmdModificar_Click()
If rsListaPro.RecordCount = 0 Then
    MsgBox "NO HAY INFORAMCION PARA MODIFICAR DATOS...", vbExclamation, "ATENCION !"
    Exit Sub
End If
frameLista.Enabled = False
frameDatos.Visible = True
txtNombre.Text = rsListaPro!nombre
txtTelefono.Text = rsListaPro!telefono & ""
txtMaximo.Text = rsListaPro!montomaximo
txtMinimo.Text = rsListaPro!montominimo
txtObservaciones.Text = rsListaPro!observaciones & ""
txtNombre.SetFocus
vAgrega = False
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub Form_Load()
Me.Top = 1800
Me.Left = 200
vAgrega = False
frameDatos.Visible = False
frameLista.Enabled = True
rsListaPro.Open "Select * from proveedores order by nombre", cn, adOpenDynamic, adLockOptimistic, adCmdText
Set dtgLista.DataSource = rsListaPro
End Sub
Private Sub BlanqueaCampos()
txtNombre.Text = ""
txtTelefono.Text = ""
txtMaximo.Text = ""
txtMinimo.Text = ""
txtObservaciones.Text = ""
frameDatos.Visible = False
frameLista.Enabled = True
dtgLista.SetFocus
End Sub
Private Sub Form_Unload(cancel As Integer)
If rsListaPro.State = 1 Then
    rsListaPro.Close
    Set rsListaPro = Nothing
End If
End Sub
Private Sub txtMaximo_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtMinimo.SetFocus
        SendKeys "{end}+{home}"
    Case 44
    Case 45
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtMinimo_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtObservaciones.SetFocus
        SendKeys "{end}+{home}"
    Case 44
    Case 45
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtTelefono.SetFocus
End If
End Sub
Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdGrabar.SetFocus
End If
End Sub
Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtMaximo.SetFocus
End If
End Sub
