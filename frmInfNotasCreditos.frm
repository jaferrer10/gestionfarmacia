VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Begin VB.Form frmInfNotasCreditos 
   Caption         =   "Informe de Notas de Creditos a Proveedores ..."
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7005
   Icon            =   "frmInfNotasCreditos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7005
   StartUpPosition =   1  'CenterOwner
   Begin AIFCmp1.asxPowerButton cmdVer 
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Picture         =   "frmInfNotasCreditos.frx":0A02
      Caption         =   "Ver el Informe"
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
   Begin VB.Frame Frame1 
      Caption         =   "Opciones de Impresión"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txtano 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         MaxLength       =   5
         TabIndex        =   8
         ToolTipText     =   "Ingrese los digitos del año del periodo de la NC"
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton optAno 
         Caption         =   "Filtrar por año del periodo"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo dtcOs 
         Height          =   315
         Left            =   3360
         TabIndex        =   5
         Tag             =   " "
         ToolTipText     =   "Seleccione una Obra Social"
         Top             =   720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.OptionButton optPro 
         Caption         =   "Filtrar por Proveedor"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton OptOs 
         Caption         =   "Filtrar por Obra Social"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo dtcPro 
         Height          =   315
         Left            =   3360
         TabIndex        =   6
         Tag             =   " "
         ToolTipText     =   "Seleccione un proveedor"
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
   End
   Begin AIFCmp1.asxPowerButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Picture         =   "frmInfNotasCreditos.frx":1414
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
End
Attribute VB_Name = "frmInfNotasCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsPro As New ADODB.Recordset
Private rsOs As New ADODB.Recordset
Private rsInf As New ADODB.Recordset
Private rptNC As New crptInfNCproveedores
Private Sub cmdCancel_Click()
rsPro.Close
rsOs.Close
Unload Me
End Sub

Private Sub cmdVer_Click()
If optPro.Value = True Then
    rsInf.Open "select o.nombre as ObSocial, n.codigo, n.periodo, n.importe, n.resumen, p.nombre as Proveedor, n.observaciones, n.Drogueria, n.idnc, n.osocial " & _
            "from (ncdroguerias n " & _
            "inner join obrasociales o on n.osocial = o.idos) " & _
            "inner join Proveedores p on n.drogueria = p.idproveedor " & _
            "where n.drogueria = " & dtcPro.BoundText & _
            " order by n.periodo", cn, adOpenDynamic, adLockOptimistic, adCmdText
End If

If OptOs.Value = True Then
    rsInf.Open "select o.nombre as ObSocial, n.codigo, n.periodo, n.importe, n.resumen, p.nombre as Proveedor, n.observaciones, n.Drogueria, n.idnc, n.osocial " & _
            "from (ncdroguerias n " & _
            "inner join obrasociales o on n.osocial = o.idos) " & _
            "inner join Proveedores p on n.drogueria = p.idproveedor " & _
            "where n.osocial = " & dtcOs.BoundText & _
            " order by n.periodo", cn, adOpenDynamic, adLockOptimistic, adCmdText
End If

If optAno.Value = True Then
    rsInf.Open "select o.nombre as ObSocial, n.codigo, n.periodo, n.importe, n.resumen, p.nombre as Proveedor, n.observaciones, n.Drogueria, n.idnc, n.osocial " & _
            "from (ncdroguerias n " & _
            "inner join obrasociales o on n.osocial = o.idos) " & _
            "inner join Proveedores p on n.drogueria = p.idproveedor " & _
            "where Right(n.periodo,4) = " & Trim(txtano.Text) & _
            " order by n.periodo", cn, adOpenDynamic, adLockOptimistic, adCmdText
End If

If rsInf.RecordCount = 0 Then
    MsgBox "NO HAY DATOS PARA EL INFORME ...", vbExclamation, "ATENCION !"
    Exit Sub
End If

crptInfNCproveedores.Database.SetDataSource rsInf

Set rptGeneral = crptInfNCproveedores ' Asigna el reporte al objeto reporte general utilizado
                           ' en el Form de la Vista Previa.
frmVistaPrevia.Show vbModal

Set crptInfNCproveedores = Nothing
rsInf.Close
Set rsInf = Nothing

End Sub

Private Sub Form_Load()
optPro.Value = True
dtcPro.Enabled = True

dtcOs.Enabled = False

txtano.Enabled = False

rsPro.Open "select * from proveedores order by nombre", cn, adOpenDynamic, adLockReadOnly, adCmdText
rsOs.Open "select * from obrasociales order by nombre", cn, adOpenDynamic, adLockReadOnly, adCmdText

Dim reg
Set dtcPro.DataSource = rsPro
Set dtcPro.RowSource = rsPro
dtcPro.BoundColumn = "idproveedor"
dtcPro.ListField = "nombre"
'esto se hace para que aparesca de entrada el combo lleno
rsPro.MoveFirst
reg = rsPro!idproveedor
dtcPro.BoundText = reg

Set dtcOs.DataSource = rsOs
Set dtcOs.RowSource = rsOs
dtcOs.BoundColumn = "idos"
dtcOs.ListField = "nombre"
rsOs.MoveFirst
reg = rsOs!idos
dtcOs.BoundText = reg

End Sub
Private Sub optAno_Click()
dtcPro.Enabled = False
dtcOs.Enabled = False
txtano.Enabled = True
txtano.SetFocus
End Sub
Private Sub OptOs_Click()
dtcPro.Enabled = False
dtcOs.Enabled = True
txtano.Enabled = False
End Sub
Private Sub optPro_Click()
dtcPro.Enabled = True
dtcOs.Enabled = False
txtano.Enabled = False
End Sub
Private Sub txtano_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdVer.SetFocus
End If
End Sub
