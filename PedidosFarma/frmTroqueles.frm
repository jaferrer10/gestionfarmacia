VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Begin VB.Form frmTroqueles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Troqueles disponibles ..."
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   Icon            =   "frmTroqueles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   10455
   Begin VB.TextBox txtBusca 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2880
      MaxLength       =   100
      TabIndex        =   12
      Top             =   1920
      Width           =   5895
   End
   Begin VB.Frame Frame2 
      Caption         =   "Archivo de Troqueles disponibles"
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   8655
      Begin MSDataGridLib.DataGrid dtgTroqueles 
         Height          =   3975
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   7011
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   8438015
         HeadLines       =   1
         RowHeight       =   15
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "codtroquel"
            Caption         =   "CodTroquel"
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
            DataField       =   "descripcion"
            Caption         =   "Descripcion Producto"
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
            DataField       =   "cantidad"
            Caption         =   "Cantidad"
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
               ColumnWidth     =   1860,095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   5414,74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   780,095
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Almacenar Troquel"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
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
         Height          =   350
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx2 
         Height          =   240
         Left            =   3120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   423
         Caption         =   "Cantidad"
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
         Height          =   350
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx1 
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   423
         Caption         =   "Codigo Troquel"
      End
      Begin AIFCmp1.asxPowerButton cmdGrabar 
         Height          =   495
         Left            =   8760
         TabIndex        =   6
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Picture         =   "frmTroqueles.frx":0582
         Caption         =   "&Grabar"
         CaptionAlignment=   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureAlignment=   3
      End
      Begin VB.Label lblDescripcion 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   615
      End
   End
   Begin AIFCmp1.asxPowerButton cmdUsar 
      Height          =   495
      Left            =   8880
      TabIndex        =   3
      ToolTipText     =   "Borra el registro seleccionado de Enviados"
      Top             =   5640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Picture         =   "frmTroqueles.frx":09D4
      Caption         =   "&Usar"
      CaptionAlignment=   5
      CaptionOffsetX  =   -10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignment=   3
      PictureOffsetX  =   15
      TextColor       =   255
   End
   Begin AIFCmp1.asxPowerButton cmdSalir 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   8880
      TabIndex        =   7
      Top             =   6360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Picture         =   "frmTroqueles.frx":0B2E
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
      PictureOffsetX  =   10
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Busca por la Descripcion:"
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
      TabIndex        =   11
      Top             =   1920
      Width           =   2685
   End
End
Attribute VB_Name = "frmTroqueles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsTroqueles As New ADODB.Recordset
Private rsDescrip As New ADODB.Recordset

Private Sub cmdGrabar_Click()
If Len(txtTroquel.Text) = 0 Then
    MsgBox "FALTA EL CODIGO DE TROQUEL PARA GRABAR ....!", vbExclamation, "ATENCION !"
    txtTroquel.SetFocus
    Exit Sub
End If
If Len(txtCantidad.Text) = 0 Then
    MsgBox "FALTA LA CANTIDAD DE TROQUELES QUE SE VAN A GRABAR ...!", vbExclamation, "ATENCION !"
    txtCantidad.SetFocus
    Exit Sub
End If
rsTroqueles.Find "codtroquel = " & Trim(txtTroquel.Text), , adSearchForward, 1

'si ya esta registrado suma la cantidad, en caso contrario agrega un registro
If rsTroqueles.EOF = True Then
    rsTroqueles.AddNew
    rsTroqueles!cantidad = txtCantidad.Text
Else
    rsTroqueles!cantidad = rsTroqueles!cantidad + txtCantidad.Text
End If
rsTroqueles!codtroquel = txtTroquel.Text
rsTroqueles!descripcion = lblDescripcion.Caption
rsTroqueles.Update
dtgTroqueles.Refresh
txtTroquel.Text = ""
txtCantidad.Text = 0
lblDescripcion.Visible = False
txtTroquel.SetFocus
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdUsar_Click()
If rsTroqueles.RecordCount = 0 Then
    MsgBox "NO HAY DATOS EN EL ARCHIVO PARA OPERAR !", vbCritical, "ATENCION !"
    txtTroquel.SetFocus
    Exit Sub
End If
If rsTroqueles!cantidad = 1 Then
    rsTroqueles.Delete
Else
    rsTroqueles!cantidad = rsTroqueles!cantidad - 1
End If
rsTroqueles.Update
dtgTroqueles.Refresh
txtTroquel.SetFocus
End Sub

Private Sub Form_Load()
Me.Top = 50
Me.Left = 50
lblDescripcion.Visible = False
rsTroqueles.Open "select codtroquel,descripcion,cantidad from troqueles order by descripcion", cn, adOpenDynamic, adLockOptimistic, adCmdText
Set dtgTroqueles.DataSource = rsTroqueles
dtgTroqueles.Refresh

End Sub
Private Sub Form_Unload(cancel As Integer)
If rsTroqueles.State = 1 Then
    rsTroqueles.Close
    Set rsTroqueles = Nothing
End If
If rsDescrip.State = 1 Then
    rsDescrip.Close
    Set rsDescrip = Nothing
End If
End Sub
Private Sub txtBusca_Change()
If Len(txtBusca.Text) = 0 Then Exit Sub
rsTroqueles.Find "descripcion like '" & txtBusca.Text & "%'", , adSearchForward, 1
If rsTroqueles.EOF = True Then
    rsTroqueles.MoveFirst
End If
End Sub
Private Sub txtBusca_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    dtgTroqueles.SetFocus
End If
End Sub
Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdGrabar.SetFocus
End If
End Sub
Private Sub txtTroquel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If rsDescrip.State = 1 Then
        rsDescrip.Close
        Set rsDescrip = Nothing
    End If
    rsDescrip.Open "select troquel, descripcion from productos where troquel like '" & txtTroquel.Text & "%' order by troquel", cn, adOpenDynamic, adLockReadOnly, adCmdText
    If rsDescrip.RecordCount = 0 Then
        MsgBox "NO EXISTE NINGUN PRODUCTO CON ESTE CODIGO DE TROQUEL, VERIQUE !", vbCritical, "ATENCION !"
        txtTroquel.SetFocus
        SendKeys "{home}+{end}"
        rsDescrip.Close
        Set rsDescrip = Nothing
        Exit Sub
    End If
    lblDescripcion.Visible = True
    lblDescripcion.Caption = rsDescrip!descripcion
    txtCantidad.Text = 1
    txtCantidad.SetFocus
    SendKeys "{home}+{end}"
    rsDescrip.Close
    Set rsDescrip = Nothing
End If
End Sub
