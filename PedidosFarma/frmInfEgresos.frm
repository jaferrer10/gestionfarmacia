VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInfEgresos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Infome de Egresos de dinero ..."
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7425
   Icon            =   "frmInfEgresos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frConcepto 
      Caption         =   "Seleccione concepto de Egresos"
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   7215
      Begin MSDataListLib.DataCombo dtcConceptos 
         Height          =   360
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ForeColor       =   32768
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
   Begin VB.Frame frameOpc 
      Caption         =   "Opciones de Periodo"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7215
      Begin VB.OptionButton optTodo 
         Caption         =   "Todos los datos del Archivo"
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton optFechas 
         Caption         =   "Establecer Rango de Fecha"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame frameFechas 
      Caption         =   "Rango de Fechas para filtrar Cuenta Corriente"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   7215
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   375
         Left            =   1800
         TabIndex        =   1
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
         Format          =   140902401
         CurrentDate     =   39392
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   375
         Left            =   4680
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
         Format          =   152764417
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   720
         Width           =   690
      End
   End
   Begin AIFCmp1.asxPowerButton cmdImprimir 
      Height          =   735
      Left            =   3480
      TabIndex        =   5
      Top             =   4320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      Picture         =   "frmInfEgresos.frx":0442
      Caption         =   "&Imprimir"
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
      PictureOffsetX  =   5
   End
   Begin AIFCmp1.asxPowerButton cmdCancelar 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   5520
      TabIndex        =   9
      Top             =   4320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      Picture         =   "frmInfEgresos.frx":0D1C
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
      PictureOffsetX  =   5
   End
End
Attribute VB_Name = "frmInfEgresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rptEgresos As New crptInfEgresos
Private rsInfEg As New ADODB.Recordset
Private rsConceptos As New ADODB.Recordset

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdImprimir_Click()
If optTodo.Value = True Then
    If dtcConceptos.BoundText = "" Then
        rsInfEg.Open "select egresos.idegreso,egresos.Fecha,egresos.Descripcion as DE,egresos.concepto,conceptosegresos.Descripcion as DC,egresos.Importe from Egresos inner join ConceptosEgresos on egresos.concepto=conceptosegresos.idconeg " & _
                " order by fecha", cn, adOpenDynamic, adLockOptimistic, adCmdText
    Else
        rsInfEg.Open "select egresos.idegreso,egresos.Fecha,egresos.Descripcion as DE,egresos.concepto,conceptosegresos.Descripcion as DC,egresos.Importe from Egresos inner join ConceptosEgresos on egresos.concepto=conceptosegresos.idconeg " & _
                " where concepto = " & dtcConceptos.BoundText & " order by fecha", cn, adOpenDynamic, adLockOptimistic, adCmdText
    End If
Else
    If dtcConceptos.BoundText = "" Then
        rsInfEg.Open "select egresos.idegreso,egresos.Fecha,egresos.Descripcion as DE,egresos.concepto,conceptosegresos.Descripcion as DC,egresos.Importe from Egresos inner join ConceptosEgresos on egresos.concepto=conceptosegresos.idconeg " & _
                "where fecha between #" & Format(dtpDesde.Value, "mm/dd/yyyy") & "# and #" & Format(dtpHasta.Value, "mm/dd/yyyy") & "# order by fecha", cn, adOpenDynamic, adLockOptimistic, adCmdText
    Else
        rsInfEg.Open "select egresos.idegreso,egresos.Fecha,egresos.Descripcion as DE,egresos.concepto,conceptosegresos.Descripcion as DC,egresos.Importe from Egresos inner join ConceptosEgresos on egresos.concepto=conceptosegresos.idconeg " & _
                "where concepto = " & dtcConceptos.BoundText & " and fecha between #" & Format(dtpDesde.Value, "mm/dd/yyyy") & "# and #" & Format(dtpHasta.Value, "mm/dd/yyyy") & "# order by fecha", cn, adOpenDynamic, adLockOptimistic, adCmdText
    End If
End If
If rsInfEg.RecordCount = 0 Then
    MsgBox "NO HAY INFORMACION PARA EL INFORME ...!", vbInformation, "ATENCION !"
    rsInfEg.Close
    Set rsInfEg = Nothing
    Exit Sub
End If
rptEgresos.Database.SetDataSource rsInfEg

Set rptGeneral = rptEgresos ' Asigna el reporte al objeto reporte general utilizado
                           ' en el Form de la Vista Previa.
frmVistaPrevia.Show vbModal

Set rptEgresos = Nothing
rsInfEg.Close
Set rsInfEg = Nothing
End Sub
Private Sub dtpDesde_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    dtpHasta.SetFocus
End If
End Sub
Private Sub dtpHasta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdImprimir.SetFocus
End If
End Sub

Private Sub Form_Load()
Me.Top = 2000
Me.Left = 4000
'toma el 1º y el ultimo dia de mes para los filtros de fechas
dtpDesde.Value = "01/" & Month(Date) & "/" & Year(Date)
If Month(Date) = 1 Or Month(Date) = 3 Or Month(Date) = 5 Or Month(Date) = 7 Or Month(Date) = 8 Or Month(Date) = 10 Or Month(Date) = 12 Then
    dtpHasta.Value = Format("31/" & Month(Date) & "/" & Year(Date), "dd/mm/yyyy")
End If
If Month(Date) = 4 Or Month(Date) = 6 Or Month(Date) = 9 Or Month(Date) = 11 Then
    dtpHasta.Value = Format("30/" & Month(Date) & "/" & Year(Date), "dd/mm/yyyy")
End If
If Month(Date) = 2 Then
    dtpHasta.Value = Format("28/" & Month(Date) & "/" & Year(Date), "dd/mm/yyyy")
End If
frameFechas.Visible = False

'Lleno combo de conceptos de egresos
rsConceptos.Open "select idConEg,descripcion from conceptosegresos order by 2", cn, adOpenDynamic, adLockReadOnly, adCmdText

Set dtcConceptos.DataSource = rsConceptos
Set dtcConceptos.RowSource = rsConceptos
dtcConceptos.ListField = "Descripcion"
dtcConceptos.BoundColumn = "idConEg"

End Sub
Private Sub optFechas_Click()
optTodo.Value = False
frameFechas.Visible = True
End Sub

Private Sub optFechas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    dtpDesde.SetFocus
End If
End Sub

Private Sub optTodo_Click()
optFechas.Value = False
frameFechas.Visible = False
End Sub
