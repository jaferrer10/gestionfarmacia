VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInfVentasExtras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe de Ventas Extraordinarias ..."
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7440
   Icon            =   "frmInfVentasExtras.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar concepto de Venta"
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   7215
      Begin MSDataListLib.DataCombo dtcConceptos 
         Height          =   360
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
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
   End
   Begin VB.Frame frameFechas 
      Caption         =   "Rango de Fechas para filtrar Ventas"
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7215
      Begin MSComCtl2.DTPicker dtpIni 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   600
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
         Format          =   141099009
         CurrentDate     =   39392
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   375
         Left            =   4680
         TabIndex        =   3
         Top             =   600
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
         Format          =   151912449
         CurrentDate     =   39392
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
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
         Left            =   840
         TabIndex        =   5
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
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
         Left            =   3720
         TabIndex        =   4
         Top             =   720
         Width           =   690
      End
   End
   Begin AIFCmp1.asxPowerButton cmdVisualizar 
      Height          =   735
      Left            =   3600
      TabIndex        =   0
      Top             =   3240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      Picture         =   "frmInfVentasExtras.frx":27A2
      Caption         =   "&Visualizar"
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
   Begin AIFCmp1.asxPowerButton cmdCancelar 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   5520
      TabIndex        =   6
      Top             =   3240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      Picture         =   "frmInfVentasExtras.frx":307C
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
   End
End
Attribute VB_Name = "frmInfVentasExtras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsInforme As New ADODB.Recordset
Private rpt_infExtras As New crptVentasExtras
Private rsConceptos As New ADODB.Recordset

Private Sub cmdCancelar_Click()
Unload Me
End Sub
Private Sub cmdVisualizar_Click()
If rsInforme.State = 1 Then
    rsInforme.Close
Else
    If dtcConceptos.Text = "" Then
        rsInforme.Open "select fecha,ve.descripcion,concepto,importe,ve.observaciones,c.descripcion from ventasExtraordinarias ve" & _
                    " inner join conceptosExtraordinarios c on ve.concepto=c.idconcepto" & _
                    " where fecha between #" & Format(dtpIni.Value, "mm/dd/yyyy") & "# and #" & Format(dtpFin.Value, "mm/dd/yyyy") & "# order by fecha desc", cn, adOpenDynamic, adLockReadOnly, adCmdText
    Else
        rsInforme.Open "select fecha,ve.descripcion,concepto,importe,ve.observaciones,c.descripcion from ventasExtraordinarias ve" & _
                    " inner join conceptosExtraordinarios c on ve.concepto=c.idconcepto" & _
                    " where concepto = " & dtcConceptos.BoundText & " and fecha between #" & Format(dtpIni.Value, "mm/dd/yyyy") & "# and #" & Format(dtpFin.Value, "mm/dd/yyyy") & "# order by fecha desc", cn, adOpenDynamic, adLockReadOnly, adCmdText
    End If
End If
If rsInforme.RecordCount = 0 Then
    MsgBox " NO HAY INFORMACION PARA EL INFORME...", vbInformation, "ATENCION !"
    rsInforme.Close
    Set rsInforme = Nothing
    Exit Sub
End If

rpt_infExtras.Database.SetDataSource rsInforme

Set rptGeneral = rpt_infExtras ' Asigna el reporte al objeto reporte general utilizado
                           ' en el Form de la Vista Previa.
frmVistaPrevia.Show vbModal

Set rpt_InfVentas = Nothing
Set rptGeneral = Nothing
rsInforme.Close
Set rsInforme = Nothing

End Sub
Private Sub dtpFin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdVisualizar.SetFocus
End If
End Sub
Private Sub dtpIni_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    dtpFin.SetFocus
End If
End Sub

Private Sub Form_Load()
Me.Top = 2000
Me.Left = 4000
dtpIni.Value = "01/01/" & Year(Date)
dtpFin.Value = "31/12/" & Year(Date)

'Llena combo de conceptos de ventas extraordinarias
rsConceptos.Open "select idconcepto, descripcion from conceptosextraordinarios order by 2", cn, adOpenDynamic, adLockReadOnly, adCmdText

Set dtcConceptos.DataSource = rsConceptos
Set dtcConceptos.RowSource = rsConceptos
dtcConceptos.ListField = "Descripcion"
dtcConceptos.BoundColumn = "idconcepto"

End Sub

Private Sub Form_Unload(Cancel As Integer)
If rsConceptos.State = 1 Then
    rsConceptos.Close
    Set rsConceptos = Nothing
End If
End Sub
