VERSION 5.00
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInfGanancias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Ganancias ..."
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7440
   Icon            =   "frmInfGanancias.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   7440
   Begin VB.Frame frameOpc 
      Caption         =   "Opciones de Periodo"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7215
      Begin VB.OptionButton optTodo 
         Caption         =   "Todos los datos del Archivo"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton optFechas 
         Caption         =   "Establecer Rango de Fecha"
         Height          =   255
         Left            =   3840
         TabIndex        =   6
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame frameFechas 
      Caption         =   "Rango de Fechas para filtrar Cuenta Corriente"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   1200
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
         Format          =   147652609
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
         Format          =   147652609
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
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      Picture         =   "frmInfGanancias.frx":27A2
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
      Height          =   495
      Left            =   5520
      TabIndex        =   9
      Top             =   2880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      Picture         =   "frmInfGanancias.frx":28FC
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
Attribute VB_Name = "frmInfGanancias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsDatosVtas As New ADODB.Recordset
Private rsDatosEgr As New ADODB.Recordset
Private rsDatosVEx As New ADODB.Recordset
Private rsDatosCpras As New ADODB.Recordset
Private rsDatosDeuda As New ADODB.Recordset 'suma facturas impagas
Private rsGanancias As New ADODB.Recordset
Private rptGanancias As New crptInfGanancias
Private Sub cmdCancelar_Click()
Unload Me
End Sub
Private Sub cmdImprimir_Click()
If optTodo.Value = True Then
    rsDatosVtas.Open "select fecha,total from Ventas order by fecha desc", cn, adOpenDynamic, adLockOptimistic, adCmdText
    rsDatosEgr.Open "select fecha,importe from Egresos order by fecha desc", cn, adOpenDynamic, adLockOptimistic, adCmdText
    rsDatosVEx.Open "select fecha,importe from VentasExtraordinarias order by fecha desc", cn, adOpenDynamic, adLockOptimistic, adCmdText
    rsDatosCpras.Open "select fecha,importe from FacturasCompras where estado = 'P' and tipo <> 'Dp' order by fecha desc", cn, adOpenDynamic, adLockOptimistic, adCmdText
    rsDatosDeuda.Open "select fecha,importe from FacturasCompras where estado = 'D' order by fecha desc", cn, adOpenDynamic, adLockOptimistic, adCmdText
Else
    rsDatosVtas.Open "select fecha,total from Ventas where fecha between #" & Format(dtpDesde.Value, "mm/dd/yyyy") & "# and #" & Format(dtpHasta.Value, "mm/dd/yyyy") & "# order by fecha desc", cn, adOpenDynamic, adLockOptimistic, adCmdText
    rsDatosEgr.Open "select fecha,importe from Egresos where fecha between #" & Format(dtpDesde.Value, "mm/dd/yyyy") & "# and #" & Format(dtpHasta.Value, "mm/dd/yyyy") & "# order by fecha desc", cn, adOpenDynamic, adLockOptimistic, adCmdText
    rsDatosVEx.Open "select fecha,importe from VentasExtraordinarias where fecha between #" & Format(dtpDesde.Value, "mm/dd/yyyy") & "# and #" & Format(dtpHasta.Value, "mm/dd/yyyy") & "# order by fecha desc", cn, adOpenDynamic, adLockOptimistic, adCmdText
    rsDatosCpras.Open "select fecha,Importe from FacturasCompras where fecha between #" & Format(dtpDesde.Value, "mm/dd/yyyy") & "# and #" & Format(dtpHasta.Value, "mm/dd/yyyy") & "# and estado='P' and tipo <> 'Dp' order by fecha desc", cn, adOpenDynamic, adLockOptimistic, adCmdText
    rsDatosDeuda.Open "select fecha,Importe from FacturasCompras where fecha between #" & Format(dtpDesde.Value, "mm/dd/yyyy") & "# and #" & Format(dtpHasta.Value, "mm/dd/yyyy") & "# and estado='D' order by fecha desc", cn, adOpenDynamic, adLockOptimistic, adCmdText
End If
If rsDatosVtas.RecordCount = 0 Then
    MsgBox "NO HAY INFORMACION DE VENTAS PARA EL INFORME ...!", vbInformation, "ATENCION !"
    rsDatosVtas.Close
    rsDatosEgr.Close
    rsDatosVEx.Close
    rsDatosCpras.Close
    rsDatosDeuda.Close
    Exit Sub
End If
rsGanancias.Open "select * from infganancias", cn, adOpenDynamic, adLockOptimistic, adCmdText
If rsGanancias.RecordCount > 0 Then
    cn.Execute "delete all * from infganancias"
    rsGanancias.Requery
End If

'Procesamos Ventas
If rsDatosVtas.RecordCount > 0 Then
    rsDatosVtas.MoveFirst
    Dim vmes, vano
    Dim vTotVtas As Double
    Dim vGralVtas As Double
    vGralVtas = 0
    vTotVtas = 0
    vmes = Month(rsDatosVtas!fecha)
    vano = Year(rsDatosVtas!fecha)
    Do While Not rsDatosVtas.EOF
        If vmes <> Month(rsDatosVtas!fecha) Then
            rsGanancias.AddNew
            rsGanancias!mesa�o = Str(vmes) + "/" + Str(vano)
            rsGanancias!total = vTotVtas
            rsGanancias!concepto = "VENTAS"
            rsGanancias.Update
            vmes = Month(rsDatosVtas!fecha)
            vano = Year(rsDatosVtas!fecha)
            vTotVtas = 0
        Else
            vTotVtas = vTotVtas + rsDatosVtas!total
            vGralVtas = vGralVtas + rsDatosVtas!total
            rsDatosVtas.MoveNext
        End If
    Loop
    'Graba el ultimo registro leido
            rsGanancias.AddNew
            rsGanancias!mesa�o = Str(vmes) + "/" + Str(vano)
            rsGanancias!total = vTotVtas
            rsGanancias!concepto = "VENTAS"
            rsGanancias.Update
End If
rsDatosVtas.Close
Set rsDatosVtas = Nothing

'Procesamos Ventas Extraordinarias
If rsDatosVEx.RecordCount > 0 Then
    Dim vTotVE As Double
    Dim vGralVex As Double
    vTotVE = 0
    vGralVex = 0
    rsDatosVEx.MoveFirst
    vmes = Month(rsDatosVEx!fecha)
    vano = Year(rsDatosVEx!fecha)
    Do While Not rsDatosVEx.EOF
        If vmes <> Month(rsDatosVEx!fecha) Then
            rsGanancias.AddNew
            rsGanancias!mesa�o = Str(vmes) + "/" + Str(vano)
            rsGanancias!total = vTotVE
            rsGanancias!concepto = "Vtas.Extras"
            rsGanancias.Update
            vmes = Month(rsDatosVEx!fecha)
            vano = Year(rsDatosVEx!fecha)
            vTotVE = 0
        Else
            vTotVE = vTotVE + rsDatosVEx!Importe
            vGralVex = vGralVex + rsDatosVEx!Importe
            rsDatosVEx.MoveNext
        End If
    Loop
    'Graba el ultimo registro leido
            rsGanancias.AddNew
            rsGanancias!mesa�o = Str(vmes) + "/" + Str(vano)
            rsGanancias!total = vTotVE
            rsGanancias!concepto = "Vtas.Extras"
            rsGanancias.Update
End If
rsDatosVEx.Close
Set rsDatosVEx = Nothing

'Procesamos Egresos
If rsDatosEgr.RecordCount > 0 Then
    Dim vTotEgr As Double
    Dim vGralEgr As Double
    vTotEgr = 0
    vGralEgr = 0
    rsDatosEgr.MoveFirst
    vmes = Month(rsDatosEgr!fecha)
    vano = Year(rsDatosEgr!fecha)
    Do While Not rsDatosEgr.EOF
        If vmes <> Month(rsDatosEgr!fecha) Or i = rsDatosEgr.RecordCount Then
            rsGanancias.AddNew
            rsGanancias!mesa�o = Str(vmes) + "/" + Str(vano)
            rsGanancias!total = vTotEgr
            rsGanancias!concepto = "EGRESOS"
            rsGanancias.Update
            vmes = Month(rsDatosEgr!fecha)
            vano = Year(rsDatosEgr!fecha)
            vTotEgr = 0
        Else
            vTotEgr = vTotEgr + rsDatosEgr!Importe
            vGralEgr = vGralEgr + rsDatosEgr!Importe
            rsDatosEgr.MoveNext
        End If
    Loop
    'Graba el ultimo registro leido
            rsGanancias.AddNew
            rsGanancias!mesa�o = Str(vmes) + "/" + Str(vano)
            rsGanancias!total = vTotEgr
            rsGanancias!concepto = "EGRESOS"
            rsGanancias.Update
End If
rsDatosEgr.Close
Set rsDatosEgr = Nothing

'Procesamos gastos por COMPRAS
Dim vTotCs As Double
Dim vTotGcs As Double
vTotCs = 0
vTotGcs = 0
If rsDatosCpras.RecordCount > 0 Then
    rsDatosCpras.MoveFirst
    vmes = Month(rsDatosCpras!fecha)
    vano = Year(rsDatosCpras!fecha)
    Do While Not rsDatosCpras.EOF
        If vmes <> Month(rsDatosCpras!fecha) Or i = rsDatosCpras.RecordCount Then
            rsGanancias.AddNew
            rsGanancias!mesa�o = Str(vmes) + "/" + Str(vano)
            rsGanancias!total = vTotCs
            rsGanancias!concepto = "Compras"
            rsGanancias.Update
            vmes = Month(rsDatosCpras!fecha)
            vano = Year(rsDatosCpras!fecha)
            vTotCs = 0
        Else
            vTotCs = vTotCs + rsDatosCpras!Importe
            vTotGcs = vTotGcs + rsDatosCpras!Importe
            rsDatosCpras.MoveNext
        End If
    Loop
    'Graba el ultimo registro leido
            rsGanancias.AddNew
            rsGanancias!mesa�o = Str(vmes) + "/" + Str(vano)
            rsGanancias!total = vTotCs
            rsGanancias!concepto = "Compras"
            rsGanancias.Update
End If
rsDatosCpras.Close
Set rsDatosCpras = Nothing

'Procesamos Deudas, osea, facturas impagas
Dim vTotD As Double
Dim vTotDG As Double
vTotD = 0
vTotDG = 0
If rsDatosDeuda.RecordCount > 0 Then
    rsDatosDeuda.MoveFirst
    vmes = Month(rsDatosDeuda!fecha)
    vano = Year(rsDatosDeuda!fecha)
    Do While Not rsDatosDeuda.EOF
        If vmes <> (Month(rsDatosDeuda!fecha)) Then
            rsGanancias.AddNew
            rsGanancias!mesa�o = Str(vmes) + "/" + Str(vano)
            rsGanancias!total = vTotD
            rsGanancias!concepto = "Facturas Impagas"
            rsGanancias.Update
            vmes = Month(rsDatosDeuda!fecha)
            vano = Year(rsDatosDeuda!fecha)
            vTotD = 0
        Else
            vTotD = vTotD + rsDatosDeuda!Importe
            vTotDG = vTotDG + rsDatosDeuda!Importe
            rsDatosDeuda.MoveNext
        End If
    Loop
    'Graba ultimo registro leido
            rsGanancias.AddNew
            rsGanancias!mesa�o = Str(vmes) + "/" + Str(vano)
            rsGanancias!total = vTotD
            rsGanancias!concepto = "Facturas Impagas"
            rsGanancias.Update
End If
rsDatosDeuda.Close
Set rsDatosDeuda = Nothing

'calculo el registro de ganancias apartir de los totales anteriores
Dim vTotGs As Double
vTotGs = 0
vTotGs = vGralVtas + vGralVex
vTotGs = vTotGs - (vGralEgr + vTotGcs)
'grabo el registro de la ganancia calculada
rsGanancias.AddNew
'rsGanancias!mesa�o = Str(vmes) + "/" + Str(vano)
rsGanancias!total = Round(vTotGs, 2)
rsGanancias!concepto = "GANANCIAS"

rsGanancias.Update
rsGanancias.Requery

'------------------------------------------------------------------'

rptGanancias.Database.SetDataSource rsGanancias

Set rptGeneral = rptGanancias ' Asigna el reporte al objeto reporte general utilizado
                           ' en el Form de la Vista Previa.
frmVistaPrevia.Show vbModal

Set rptGanancias = Nothing
Set rptGeneral = Nothing

rsGanancias.Close

Set rsGanancias = Nothing

'blanqueamos la tabla del informe
'cn.Execute "delete from infganancias"
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
'toma el 1� y el ultimo dia de mes para los filtros de fechas
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
Private Sub optFechas_Click()
frameFechas.Visible = True
optTodo.Value = False
End Sub
Private Sub optTodo_Click()
frameFechas.Visible = False
optFechas.Value = False
End Sub
