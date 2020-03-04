VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAgenda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda ..."
   ClientHeight    =   8520
   ClientLeft      =   1050
   ClientTop       =   690
   ClientWidth     =   12390
   Icon            =   "frmAgenda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   12390
   Begin VB.Frame Frame2 
      Caption         =   "Datos Registro"
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   12135
      Begin VB.TextBox txtObs 
         Height          =   285
         Left            =   240
         MaxLength       =   50
         TabIndex        =   24
         Top             =   3000
         Width           =   6255
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx8 
         Height          =   240
         Left            =   240
         TabIndex        =   23
         Top             =   2640
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   423
         Caption         =   "Observaciones"
      End
      Begin VB.TextBox txtCuit 
         Height          =   285
         Left            =   3720
         MaxLength       =   50
         TabIndex        =   22
         Top             =   2160
         Width           =   2775
      End
      Begin MSDataListLib.DataCombo dtcIva 
         Height          =   315
         Left            =   240
         TabIndex        =   21
         Top             =   2160
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx7 
         Height          =   240
         Left            =   3720
         TabIndex        =   20
         Top             =   1800
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   423
         Caption         =   "Nº CUIT"
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx6 
         Height          =   240
         Left            =   240
         TabIndex        =   19
         Top             =   1800
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   423
         Caption         =   "Condicion IVA"
      End
      Begin MSComCtl2.DTPicker dtpNac 
         Height          =   405
         Left            =   6720
         TabIndex        =   18
         Top             =   1320
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16646145
         CurrentDate     =   39960
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx5 
         Height          =   240
         Left            =   6720
         TabIndex        =   17
         Top             =   1080
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   423
         Caption         =   "Fecha de Nacimiento"
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
         Height          =   285
         Left            =   240
         MaxLength       =   50
         TabIndex        =   16
         Top             =   1320
         Width           =   6255
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
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   6720
         MaxLength       =   30
         TabIndex        =   15
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   3480
         MaxLength       =   30
         TabIndex        =   14
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtApellido 
         Height          =   285
         Left            =   240
         MaxLength       =   30
         TabIndex        =   13
         Top             =   600
         Width           =   3015
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx4 
         Height          =   240
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   423
         Caption         =   "Direccion"
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx3 
         Height          =   240
         Left            =   6720
         TabIndex        =   11
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   423
         Caption         =   "Telefonos"
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx2 
         Height          =   240
         Left            =   3480
         TabIndex        =   10
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   423
         Caption         =   "Nombres"
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx1 
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   423
         Caption         =   "Apellido"
      End
      Begin AIFCmp1.asxPowerButton cmdCancelar 
         Height          =   615
         Left            =   10200
         TabIndex        =   7
         Top             =   2760
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         Picture         =   "frmAgenda.frx":038A
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
         PictureOffsetX  =   10
         TextColor       =   255
      End
      Begin AIFCmp1.asxPowerButton cmdGrabar 
         Height          =   615
         Left            =   10200
         TabIndex        =   6
         Top             =   2040
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         Picture         =   "frmAgenda.frx":0D9C
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
         PictureOffsetX  =   10
         TextColor       =   32768
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Archivo"
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12135
      Begin MSDataGridLib.DataGrid dtgArchivo 
         Height          =   3495
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   6165
         _Version        =   393216
         AllowUpdate     =   0   'False
         ForeColor       =   8388608
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
               LCID            =   1034
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
               LCID            =   1034
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtBusca 
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
         Height          =   360
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   0
         ToolTipText     =   "Escriba el apellido que desea buscar..."
         Top             =   360
         Width           =   5415
      End
      Begin AIFCmp1.asxPowerButton cmdEliminar 
         Height          =   495
         Left            =   10320
         TabIndex        =   5
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmAgenda.frx":17AE
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
         Left            =   10320
         TabIndex        =   4
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmAgenda.frx":21C0
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
         Left            =   10320
         TabIndex        =   3
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmAgenda.frx":2BD2
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
      Begin AIFCmp1.asxPowerButton cmdSalir 
         Height          =   495
         Left            =   10320
         TabIndex        =   8
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmAgenda.frx":35E4
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
         PictureAlignment=   3
         TextColor       =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busca por Apellido >"
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
         TabIndex        =   25
         Top             =   360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsCtes As New ADODB.Recordset
Private rsIva As New ADODB.Recordset

Private Sub cmdAgregar_Click()
vAgrega = True
Frame1.Enabled = False
Frame2.Enabled = True
txtNombre.Text = ""
txtApellido.Text = ""
txtTel.Text = ""
txtDir.Text = ""
txtObs.Text = ""
txtCuit.Text = ""
dtpNac.Value = Date
txtApellido.SetFocus
End Sub
Private Sub cmdCancelar_Click()
Call TomaDatos
Frame1.Enabled = True
Frame2.Enabled = False
End Sub
Private Sub cmdEliminar_Click()
KeyAscii = 0
SioNo = MsgBox("ESTA SEGURO DE ELMINAR EL REGISTRO SELECCIONADO ?", vbInformation + vbYesNo, "Eliminando Registro...")
If SioNo = vbYes Then
    rsCtes.Delete
    rsCtes.Requery
    dtgArchivo.Refresh
End If
End Sub
Private Sub cmdGrabar_Click()
KeyAscii = 0
If Len(txtNombre.Text) = 0 Or Len(txtApellido.Text) = 0 Or Len(dtcIva.Text) = 0 Then
    MsgBox "FALTAN DATOS PARA GRABAR EL REGISTRO...!", vbCritical, "ATENCION !"
    txtApellido.SetFocus
    Exit Sub
End If
SioNo = MsgBox("ESTA SEGURO DE GRABAR TODOS LOS DATOS ?", vbInformation + vbYesNo, "ATENCION !")
If SioNo = vbNo Then Exit Sub
If vAgrega = True Then
    rsCtes.AddNew
End If
rsCtes!apellido = txtApellido.Text
rsCtes!nombre = txtNombre.Text
rsCtes!direccion = txtDir.Text
rsCtes!telefono = txtTel.Text
rsCtes!fechanac = dtpNac.Value
rsCtes!condiva = dtcIva.BoundText
rsCtes!observaciones = txtObs.Text
rsCtes!cuit = txtCuit.Text
rsCtes.Update
rsCtes.Requery
dtgArchivo.Refresh
Frame1.Enabled = True
Frame2.Enabled = False
dtgArchivo.SetFocus
vAgrega = False
End Sub
Private Sub cmdModificar_Click()
If rsCtes.RecordCount = 0 Then
    MsgBox "NO HAY DATOS PARA MODIFICAR ...!", vbCritical, "ATENCION !!!"
    cmdAgregar.SetFocus
    Exit Sub
End If
vAgrega = False
Frame1.Enabled = False
Frame2.Enabled = True
txtApellido.SetFocus
SendKeys "{home}+{end}"
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub dtcIva_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtCuit.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub
Private Sub dtgArchivo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call TomaDatos
End Sub

Private Sub dtpNac_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyAscii = 0
    dtcIva.SetFocus
End If
End Sub
Private Sub Form_Load()
vAgrega = False
Frame2.Enabled = False
Frame1.Enabled = True
Me.Left = 1400
Me.Top = 300
rsCtes.Open "select * from clientes order by apellido", cn, adOpenDynamic, adLockOptimistic, adCmdText
rsCtes.MoveFirst
Set dtgArchivo.DataSource = rsCtes
dtgArchivo.Refresh
Call TomaDatos
'llena el combo de condicion de iva
rsIva.Open "select * from condicioniva order by condicion", cn, adOpenDynamic, adLockReadOnly, adCmdText
Set dtcIva.DataSource = rsIva
Set dtcIva.RowSource = rsIva
dtcIva.BoundColumn = "idcondicion"
dtcIva.ListField = "condicion"
rsIva.MoveFirst
dtcIva.BoundText = rsIva!idcondicion
End Sub
Private Sub TomaDatos()
If rsCtes.RecordCount = 0 Or rsCtes.EOF = True Then
    txtNombre.Text = ""
    txtApellido.Text = ""
    txtTel.Text = ""
    txtDir.Text = ""
    txtObs.Text = ""
    txtCuit.Text = ""
    Exit Sub
End If

txtNombre.Text = rsCtes!nombre
txtApellido.Text = rsCtes!apellido
txtDir.Text = rsCtes!direccion
txtTel.Text = rsCtes!telefono
txtObs.Text = rsCtes!observaciones
dtcIva.BoundText = rsCtes!condiva
txtCuit.Text = rsCtes!cuit & ""
dtpNac.Value = rsCtes!fechanac

End Sub
Private Sub Form_Unload(cancel As Integer)
If rsCtes.State = 1 Then
    rsCtes.Close
End If
If rsIva.State = 1 Then
    rsIva.Close
End If
End Sub
Private Sub txtApellido_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    KeyAscii = 0
    txtNombre.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub
Private Sub txtBusca_Click()
    SendKeys "{home}+{end}"
End Sub
Private Sub txtBusca_Change()
If Len(txtBusca.Text) = 0 Then Exit Sub
rsCtes.Find "apellido like '" & txtBusca.Text & "%'", , adSearchForward, 1
If rsCtes.EOF = True Then
    rsCtes.MoveFirst
End If
dtgArchivo.Refresh
Call TomaDatos
End Sub
Private Sub txtBusca_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    dtgArchivo.SetFocus
End If
End Sub
Private Sub txtBusca_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtCuit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtObs.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub
Private Sub txtDir_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    dtpNac.SetFocus
End If
End Sub
Private Sub txtNombre_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    KeyAscii = 0
    txtTel.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub
Private Sub txtObs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdGrabar.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub
Private Sub txtTel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtDir.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub
