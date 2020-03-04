VERSION 5.00
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCalculoDeposito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculo de Deposito ..."
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7470
   Icon            =   "frmCalculoDeposito.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameDep 
      Caption         =   "Depósito"
      Height          =   2895
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   7215
      Begin VB.TextBox txtres 
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
         Left            =   5760
         MaxLength       =   10
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin AIFCmp1.asxPowerButton cmdNc 
         Height          =   495
         Left            =   2640
         TabIndex        =   14
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Picture         =   "frmCalculoDeposito.frx":030A
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextColor       =   192
      End
      Begin VB.TextBox txtDeposito 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2640
         TabIndex        =   9
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtFavor 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2640
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtTotal 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin AIFCmp1.asxPowerButton cmdOk 
         Height          =   495
         Left            =   4320
         TabIndex        =   12
         ToolTipText     =   "Queda registrado el Control de Caja"
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Picture         =   "frmCalculoDeposito.frx":268C
         Caption         =   "G&rabar"
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
      Begin AIFCmp1.asxPowerButton cmdCancel 
         Height          =   495
         Left            =   5640
         TabIndex        =   13
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Picture         =   "frmCalculoDeposito.frx":2C26
         Caption         =   "Cancelar"
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
      Begin MSComCtl2.DTPicker dtpFDep 
         Height          =   375
         Left            =   5280
         TabIndex        =   11
         Top             =   1560
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
         Format          =   139329537
         CurrentDate     =   39392
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Resumen Nº:"
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
         Left            =   4200
         TabIndex        =   25
         Top             =   480
         Width           =   1365
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Notas de Cred a desc:"
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
         TabIndex        =   24
         Top             =   2280
         Width           =   2340
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
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
         TabIndex        =   22
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Importe a Depositar:"
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
         TabIndex        =   21
         Top             =   1680
         Width           =   2130
      End
      Begin VB.Label lblSaldo 
         AutoSize        =   -1  'True
         Caption         =   "Saldo a Favor:"
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
         TabIndex        =   20
         Top             =   1080
         Width           =   1545
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Facturas:"
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
         TabIndex        =   19
         Top             =   480
         Width           =   1575
      End
   End
   Begin AIFCmp1.asxPowerButton cmdCalcular 
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Picture         =   "frmCalculoDeposito.frx":31C0
      Caption         =   "C&alcular"
      CaptionAlignment=   5
      CaptionOffsetX  =   -10
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
   Begin VB.Frame frameFechas 
      Caption         =   "Rango de Fechas para filtrar Cuenta Corriente"
      Height          =   1335
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   7215
      Begin VB.ComboBox cbEstado 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frmCalculoDeposito.frx":3612
         Left            =   4680
         List            =   "frmCalculoDeposito.frx":3625
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Indica si la factura se debe o fue pagada"
         Top             =   720
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker dtpIni 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   240
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
         Format          =   139395073
         CurrentDate     =   39392
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   375
         Left            =   4680
         TabIndex        =   3
         Top             =   240
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
         Format          =   139395073
         CurrentDate     =   39392
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Estado de Facturas:"
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
         Left            =   2280
         TabIndex        =   23
         Top             =   840
         Width           =   2100
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
         TabIndex        =   17
         Top             =   360
         Width           =   690
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
         TabIndex        =   16
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.OptionButton optFechas 
      Caption         =   "Calcular segun rango de fechas ..."
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.OptionButton optAnterior 
      Caption         =   "Calcular Todas las Facturas impagas..."
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   3495
   End
   Begin AIFCmp1.asxPowerButton cmdCancelar 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   2400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Picture         =   "frmCalculoDeposito.frx":3657
      Caption         =   "&Cancelar"
      CaptionAlignment=   5
      CaptionOffsetX  =   -10
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
End
Attribute VB_Name = "frmCalculoDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsFacturas As New ADODB.Recordset
Private rsTodasLasFacturas As New ADODB.Recordset
Private vTotal, vDepositos, vDeuda, vCredito, vPagadas, vFavor As Double
Private vPriFact, vUltFact As String

Private Sub cbEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdCalcular.SetFocus
End If
End Sub

Private Sub cmdCalcular_Click()
Err.Clear
On erro GoTo Solucion
If optFechas.Value = True Then
    If Len(cbEstado.Text) = 0 Then
        MsgBox "DEBE SELECCIONAR EL ESTADO DE LAS FACTURAS....!", vbCritical, "ATENCION !"
        cbEstado.SetFocus
        Exit Sub
    End If
    
    Dim testFecha As Integer
    testFecha = DateDiff("d", dtpIni.Value, dtpFin.Value)
    If testFecha < 0 Then
        MsgBox "LA FECHA DE INICIO NO PUEDE SER INFERIOR A LA FECHA FINAL DEL PERIODO...!", vbCritical, "ERROR DE FECHAS ..."
        dtpIni.SetFocus
        Exit Sub
    End If
    If cbEstado.Text = "Todo" Then
        rsFacturas.Open "select * from facturascompras where fechavto between #" & Format(dtpIni.Value, "mm/dd/yyyy") & "# and #" & Format(dtpFin.Value, "mm/dd/yyyy") & _
                    "# and idproveedor = " & vidPro & " order by numero,fechavto", cn, adOpenDynamic, adLockOptimistic, adCmdText
    ElseIf cbEstado.Text = "Debe y Creditos" Then
            rsFacturas.Open "select * from facturascompras where fechavto between #" & Format(dtpIni.Value, "mm/dd/yyyy") & "# and #" & Format(dtpFin.Value, "mm/dd/yyyy") & _
                    "# and idproveedor = " & vidPro & " And estado = 'D' or estado = 'C' order by numero,fechavto", cn, adOpenDynamic, adLockOptimistic, adCmdText
    ElseIf cbEstado.Text = "Debe" Then
            rsFacturas.Open "select * from facturascompras where fechavto between #" & Format(dtpIni.Value, "mm/dd/yyyy") & "# and #" & Format(dtpFin.Value, "mm/dd/yyyy") & _
                    "# and idproveedor = " & vidPro & " And estado = 'D' order by numero,fechavto", cn, adOpenDynamic, adLockOptimistic, adCmdText
    ElseIf cbEstado.Text = "Pagado" Then
            rsFacturas.Open "select * from facturascompras where fechavto between #" & Format(dtpIni.Value, "mm/dd/yyyy") & "# and #" & Format(dtpFin.Value, "mm/dd/yyyy") & _
                    "# and idproveedor = " & vidPro & " And estado = 'P' order by numero,fechavto", cn, adOpenDynamic, adLockOptimistic, adCmdText
    ElseIf cbEstado.Text = "Credito" Then
            rsFacturas.Open "select * from facturascompras where fechavto between #" & Format(dtpIni.Value, "mm/dd/yyyy") & "# and #" & Format(dtpFin.Value, "mm/dd/yyyy") & _
                    "# and idproveedor = " & vidPro & " And estado = 'C' order by numero,fechavto", cn, adOpenDynamic, adLockOptimistic, adCmdText
    End If
Else
    rsFacturas.Open "select * from facturascompras where idproveedor = " & vidPro & " and estado = " & "'D'" & " order by fechavto", cn, adOpenDynamic, adLockOptimistic, adCmdText
End If

If rsFacturas.RecordCount = 0 Then
    MsgBox "NO HAY INFORMACION DE FACTURAS IMPAGAS PARA CALCULAR EL DEPOSITO ...!", vbExclamation, "NO DEBE FACTURAS..."
    rsFacturas.Close
    Set rsFacturas = Nothing
    Exit Sub
End If
'toma el ultimo numero de factura
rsFacturas.MoveLast
vUltFact = rsFacturas!numero
'toma el primer numero de factura
rsFacturas.MoveFirst
vPriFact = rsFacturas!numero

vTotal = 0
Dim vNc As Double
vNc = 0
Dim cr As Integer
'acumula importes de compras
Do While rsFacturas.EOF = False
    If rsFacturas!Importe < 0 Then
        vNc = vNc + rsFacturas!Importe
    Else
       vTotal = vTotal + rsFacturas!Importe
    End If

    cr = cr + 1
    rsFacturas.MoveNext
   
Loop
'al total le resto las notas de creditos acumuladas en vNc
If vNc < 0 Then
    vTotal = vTotal - Abs(vNc)
End If

'muestra el total a depositar
txtTotal.Text = vTotal

'Inhabilita botones
cmdCalcular.Enabled = False
frameFechas.Enabled = False

'calculo de deuda y pagos totales para establecer saldos
frameDep.Visible = True

'rsTodasLasFacturas.Open "select sum(importe) as Total from facturascompras where idproveedor = " & vIdPro & " and estado = 'P' ", cn, adOpenDynamic, adLockReadOnly, adCmdText
rsTodasLasFacturas.Open "SELECT FacturasCompras.idProveedor, FacturasCompras.Estado, Sum(FacturasCompras.Importe) AS SumaDeImporte" & _
                    " From FacturasCompras GROUP BY FacturasCompras.idProveedor, FacturasCompras.Estado" & _
                    " HAVING (((FacturasCompras.idProveedor)= " & vidPro & " ))", cn, adOpenDynamic, adLockReadOnly, adCmdText
'Total facturas pagadas al proveedor
rsTodasLasFacturas.Find "estado = " & "'P'", , adSearchForward, 1

If rsTodasLasFacturas.EOF = True Then
    vPagadas = 0
Else
    vPagadas = rsTodasLasFacturas!sumadeimporte
End If

'total de compras que se deben
rsTodasLasFacturas.Find "estado = " & "'D'", , adSearchForward, 1

If rsTodasLasFacturas.EOF = True Then
    vDeuda = 0
Else
    vDeuda = rsTodasLasFacturas!sumadeimporte
End If

'Total de facturas a Credito
rsTodasLasFacturas.Find "estado = " & "'C'", , adSearchForward, 1

If rsTodasLasFacturas.EOF = True Then
    vCredito = 0
Else
    vCredito = rsTodasLasFacturas!sumadeimporte
End If

'Total DEUDA
'vDeuda = vDeuda + vCredito

rsTodasLasFacturas.Close
Set rsTodasLasFacturas = Nothing

rsTodasLasFacturas.Open "select sum(Depositado) as Depositos from facturascompras where idproveedor = " & vidPro, cn, adOpenDynamic, adLockReadOnly, adCmdText

If IsNull(rsTodasLasFacturas!depositos) = True Then
    vDepositos = 0
Else
    vDepositos = rsTodasLasFacturas!depositos
End If
'establece si hay saldo a favor o en contra
'vFavor = vDepositos - vPagadas

'txtFavor.Text = Round(vFavor, 2)

'al total de deuda le resto el deposito que se haria del rango de fechas
vDeuda = vDeuda - vTotal
txtFavor.Text = vDeuda
lblSaldo.Caption = "Saldo Deudor:"
'establece leyenda
'If vDeuda > vDepositos Then
'    lblSaldo.Caption = "Saldo Deudor:"
'Else
'    lblSaldo.Caption = "Saldo a Favor:"
'End If

txtDeposito.SetFocus

rsTodasLasFacturas.Close
Set rsTodasLasFacturas = Nothing

Exit Sub
Solucion:
   MsgBox Err.Number & "-" & Err.Description, vbInformation, "Error del Sistema ..."
Err.Clear
   On Error Resume Next
   
End Sub
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdCancelar_Click()
Unload Me
End Sub
Private Sub cmdNc_Click()
frmNCdroguerias.Show vbModal
End Sub
Private Sub cmdOk_Click()
Err.Clear
On Error GoTo VerErr

If Len(txtDeposito.Text) = 0 Then
    MsgBox "DEBE INGRESAR EL IMPORTE A DEPOSITAR ...!", vbCritical, "ATENCION !"
    txtDeposito.SetFocus
    Exit Sub
End If

'graba el total a depositar en la tabla
Dim vString As String
vString = "insert into facturascompras (idproveedor,fecha,numero,tipo,importe,depositado,observaciones, estado, Usuario, Rubro, fechavto ) Values ( " & vidPro & ",#" & ((Format(dtpFDep.Value, "mm/dd/yyyy"))) & "#," & "'DEPOSITO'" & "," & "'Dp'" & ",'" & vTotal & "','" & txtDeposito.Text & "'," & "'Para pagar facturas desde la " & vPriFact & " Hasta la " & vUltFact & " /Resumen Nº: " & txtres.Text & "'," & "'P'" & ",'" & vUsu & "','" & frmCompras.cbRubro.Text & "',#" & ((Format(dtpFDep.Value, "mm/dd/yyyy"))) & "#" & ")"

cn.Execute vString

'coloca todas las facturas procesadas en Pagado
rsFacturas.MoveFirst
Do While rsFacturas.EOF = False
    rsFacturas!estado = "P"
    rsFacturas.MoveNext
Loop
rsFacturas.Close
Set rsFacturas = Nothing
Unload Me
Exit Sub
VerErr:
    MsgBox Err.Number & "-" & Err.Description

End Sub
Private Sub dtpFDep_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdOk.SetFocus
End If
End Sub
Private Sub dtpFin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cbEstado.SetFocus
End If
End Sub

Private Sub dtpIni_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    dtpFin.SetFocus
End If
End Sub
Private Sub Form_Load()
dtpFDep.Value = Date + 1
cmdCalcular.Enabled = True
frameFechas.Visible = False
frameDep.Visible = False
End Sub
Private Sub Form_Unload(cancel As Integer)
If rsFacturas.State = 1 Then
    rsFacturas.Close
    Set rsFacturas = Nothing
End If
If rsTodasLasFacturas.State = 1 Then
    rsTodasLasFacturas.Close
    Set rsTodasLasFacturas = Nothing
End If
End Sub
Private Sub num1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    num2.SetFocus
End If
End Sub
Private Sub num2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdCalcular.SetFocus
End If
End Sub
Private Sub optAnterior_Click()
optFechas.Value = False
frameFechas.Visible = False
End Sub
Private Sub optFechas_Click()
optFechas.Value = True
frameFechas.Visible = True
'pone el 1º y utlimo dia del mes actual en los campos fechas para filtros
dtpIni.Value = "01/" & Month(Date) & "/" & Year(Date)
If Month(Date) = 1 Or Month(Date) = 3 Or Month(Date) = 5 Or Month(Date) = 7 Or Month(Date) = 8 Or Month(Date) = 10 Or Month(Date) = 12 Then
    dtpFin.Value = Format("31/" & Month(Date) & "/" & Year(Date), "dd/mm/yyyy")
End If
If Month(Date) = 4 Or Month(Date) = 6 Or Month(Date) = 9 Or Month(Date) = 11 Then
    dtpFin.Value = Format("30/" & Month(Date) & "/" & Year(Date), "dd/mm/yyyy")
End If
If Month(Date) = 2 Then
    dtpFin.Value = Format("28/" & Month(Date) & "/" & Year(Date), "dd/mm/yyyy")
End If
End Sub
Private Sub txtDeposito_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtres.SetFocus
    Case 44 'para que no acepte la entrada de la coma
        KeyAscii = 0
    Case 45
    Case 46
        KeyAscii = 44
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtRes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    dtpFDep.SetFocus
End If
End Sub
