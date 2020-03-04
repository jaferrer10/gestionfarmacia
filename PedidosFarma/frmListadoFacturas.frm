VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListadoFacturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de Facturas de Compras"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9765
   Icon            =   "frmListadoFacturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione Rubro"
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   9495
      Begin MSDataListLib.DataCombo cbRubro 
         Height          =   420
         Left            =   2280
         TabIndex        =   12
         Top             =   600
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   741
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame frameFechas 
      Caption         =   "Rango de Fechas para filtrar Cuenta Corriente"
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   9495
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
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
         Format          =   141230081
         CurrentDate     =   39392
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   375
         Left            =   5400
         TabIndex        =   11
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
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
         Format          =   141230081
         CurrentDate     =   39392
      End
      Begin VB.Label Label3 
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
         Left            =   4440
         TabIndex        =   8
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label2 
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
         Left            =   1320
         TabIndex        =   7
         Top             =   720
         Width           =   765
      End
   End
   Begin VB.OptionButton optFechas 
      Caption         =   "Facturas por Rango de Fechas"
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   840
      Width           =   3015
   End
   Begin VB.OptionButton optTodas 
      Caption         =   "Todas las Facturas"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Value           =   -1  'True
      Width           =   2655
   End
   Begin MSDataListLib.DataCombo dtcProveedor 
      Height          =   360
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      ForeColor       =   12582912
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
   Begin AIFCmp1.asxPowerButton cmdImprimir 
      Height          =   855
      Left            =   5760
      TabIndex        =   4
      Top             =   4920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1508
      Picture         =   "frmListadoFacturas.frx":27A2
      Caption         =   "&Imprimir"
      CaptionAlignment=   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignment=   3
      PictureOffsetX  =   5
   End
   Begin AIFCmp1.asxPowerButton cmdCancelar 
      Cancel          =   -1  'True
      Height          =   855
      Left            =   7800
      TabIndex        =   5
      Top             =   4920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1508
      Picture         =   "frmListadoFacturas.frx":307C
      Caption         =   "&Cancelar"
      CaptionAlignment=   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignment=   3
      PictureOffsetX  =   5
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Proveedor:"
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
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1170
   End
End
Attribute VB_Name = "frmListadoFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rptLista As New crptListaFacturas
Private rptListaFac As New CrptFacturasCompras
Private rsProv As New ADODB.Recordset
Private rsRubros As New ADODB.Recordset
Private rsFacturas As New ADODB.Recordset

Private Sub cmdCancelar_Click()
Unload Me
End Sub
Private Sub cmdImprimir_Click()
If rsFacturas.State = 1 Then
    rsFacturas.Close
    Set rsFacturas = Nothing
End If
If optTodas.Value = True Then
    rsFacturas.Open "select p.nombre, p.telefono, f.idproveedor, fecha, numero, tipo, importe, depositado, f.observaciones, estado, rubro " & _
                "from facturascompras f inner join Proveedores p " & _
                "on f.idproveedor = p.idproveedor " & _
                "where f.idproveedor = " & dtcProveedor.BoundText & " and tipo <> 'DP'" & _
                " order by fecha, numero asc", cn, adOpenDynamic, adLockReadOnly, adCmdText
Else
    If cbRubro.Text = "" Or cbRubro.Text = "TODOS" Then
        SQL = "select p.nombre, p.telefono, f.idproveedor, fecha, numero, tipo, importe, depositado, f.observaciones, estado, rubro " & _
                "from facturascompras f inner join Proveedores p " & _
                "on f.idproveedor = p.idproveedor " & _
                "where f.idproveedor = " & dtcProveedor.BoundText & " and Fecha between #" & Format(dtpDesde.Value, "mm/dd/yyyy") & _
                "# and #" & Format(dtpHasta.Value, "mm/dd/yyyy") & "# and tipo <> 'DP' order by fecha, numero asc"
    Else
        SQL = "select p.nombre, p.telefono, f.idproveedor, fecha, numero, tipo, importe, depositado, f.observaciones, estado, rubro " & _
                "from facturascompras f inner join Proveedores p " & _
                "on f.idproveedor = p.idproveedor " & _
                "where f.idproveedor = " & dtcProveedor.BoundText & " and (Fecha >= #" & Format(dtpDesde.Value, "mm/dd/yyyy") & _
                "# and fecha <= #" & Format(dtpHasta.Value, "mm/dd/yyyy") & "#) and tipo <> 'DP' and rubro = '" & cbRubro.Text & "' order by fecha, numero asc"
        
    End If
    rsFacturas.Open SQL, cn, adOpenDynamic, adLockReadOnly, adCmdText
    rptListaFac.txtDesde.SetText (dtpDesde.Value)
    rptListaFac.txtHasta.SetText (dtpHasta.Value)
End If
If rsFacturas.RecordCount = 0 Then
    MsgBox "NO HAY IMFORMACION PARA LA IMPRESION...", vbInformation, "ATENCION !"
    rsFacturas.Close
    Set rsFacturas = Nothing
    Exit Sub
End If

'rptLista.Database.SetDataSource rsFacturas
rptListaFac.Database.SetDataSource rsFacturas


Set rptGeneral = rptListaFac 'Asigna el reporte al objeto reporte general utilizado
                           'en el Form de la Vista Previa.
frmVistaPrevia.Show vbModal

Set rptListaFac = Nothing

rsFacturas.Close
Set rsFacturas = Nothing

End Sub
Private Sub dtcProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    optTodas.SetFocus
End If
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
rsProv.Open "select * from proveedores order by nombre", cn, adOpenDynamic, adLockReadOnly, adCmdText

'llena el combo de proveedores
Set dtcProveedor.DataSource = rsProv
Set dtcProveedor.RowSource = rsProv
dtcProveedor.ListField = "Nombre"
dtcProveedor.BoundColumn = "idproveedor"
dtcProveedor.BoundText = 1

'Llena combo de rubros
rsRubros.Open "select idrubro, rubro from Rubros order by rubro", cn, adOpenDynamic, adLockReadOnly, adCmdText
Set cbRubro.DataSource = rsRubros
Set cbRubro.RowSource = rsRubros
cbRubro.ListField = "Rubro"
cbRubro.BoundColumn = "idRubro"


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
End Sub

Private Sub Form_Unload(cancel As Integer)
rsProv.Close
Set rsProv = Nothing

If rsRubros.State = 1 Then
    rsRubros.Close
    Set rsRubros = Nothing
End If
End Sub

Private Sub optFechas_Click()
optTodas.Value = False
frameFechas.Visible = True
End Sub

Private Sub optTodas_Click()
optFechas.Value = False
frameFechas.Visible = False
End Sub

Private Sub optTodas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    dtpDesde.SetFocus
End If
End Sub
