VERSION 5.00
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmImprimeCtaCte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresion de Informe de Cuenta Corriente ..."
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7515
   Icon            =   "frmImprimeCtaCte.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameFechas 
      Caption         =   "Rango de Fechas para filtrar Cuenta Corriente"
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   7215
      Begin MSComCtl2.DTPicker dtpIni 
         Height          =   375
         Left            =   1800
         TabIndex        =   8
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
         Format          =   62062593
         CurrentDate     =   39392
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   375
         Left            =   4680
         TabIndex        =   9
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
         Format          =   62062593
         CurrentDate     =   39392
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
         TabIndex        =   7
         Top             =   720
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
         TabIndex        =   6
         Top             =   720
         Width           =   765
      End
   End
   Begin AIFCmp1.asxPowerButton cmdImprimir 
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   2880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      Picture         =   "frmImprimeCtaCte.frx":0442
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
   Begin VB.Frame frameOpc 
      Caption         =   "Opciones de Periodo"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.OptionButton optFechas 
         Caption         =   "Establecer Rango de Fecha"
         Height          =   255
         Left            =   3840
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
      Begin VB.OptionButton optTodo 
         Caption         =   "Toda la Cuenta ..."
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin AIFCmp1.asxPowerButton cmdCancelar 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   2880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      Picture         =   "frmImprimeCtaCte.frx":059C
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
Attribute VB_Name = "frmImprimeCtaCte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rptCuenta As New crptImpCuenta
Private rsImpCta As New ADODB.Recordset
Private Sub cmdCancelar_Click()
Unload Me
End Sub
Private Sub cmdImprimir_Click()
Err.Clear
On erro GoTo SolucionErr
If rsImpCta.State = 1 Then
    rsImpCta.Close
    Set rsImpCta = Nothing
End If
If optTodo.Value = True Then
    rsImpCta.Open "SELECT Clientes.Nombre, Clientes.Apellido, Clientes.Telefono, Clientes.Direccion, " & _
              "CuentasCorrientes.idcliente, CuentasCorrientes.Fecha, CuentasCorrientes.Concepto, CuentasCorrientes.Cantidad, " & _
              "CuentasCorrientes.Precio, CuentasCorrientes.Descuento, CuentasCorrientes.Importe " & _
              "FROM Clientes INNER JOIN CuentasCorrientes ON Clientes.idCliente = CuentasCorrientes.idCliente " & _
              "WHERE cuentascorrientes.idcliente = " & vIdCliente, cn, adOpenDynamic, adLockReadOnly, adCmdText
Else
    rsImpCta.Open "SELECT Clientes.Nombre, Clientes.Apellido, Clientes.Telefono, Clientes.Direccion, " & _
              "CuentasCorrientes.idcliente, CuentasCorrientes.Fecha, CuentasCorrientes.Concepto, CuentasCorrientes.Cantidad, " & _
              "CuentasCorrientes.Precio, CuentasCorrientes.Descuento, CuentasCorrientes.Importe " & _
              "FROM Clientes INNER JOIN CuentasCorrientes ON Clientes.idCliente = CuentasCorrientes.idCliente " & _
              "WHERE cuentascorrientes.idcliente = " & vIdCliente & " and CuentasCorrientes.Fecha between #" & Format(dtpIni.Value, "mm/dd/yyyy") & _
              "# and #" & Format(dtpFin.Value, "mm/dd/yyyy") & "#", cn, adOpenDynamic, adLockReadOnly, adCmdText
End If
If rsImpCta.RecordCount = 0 Then
    MsgBox "NO HAY IMFORMACION EN LA CUENTA CORRIENTE...", vbInformation, "ATENCION !"
    rsImpCta.Close
    Set rsImpCta = Nothing
    Exit Sub
End If

rptCuenta.Database.SetDataSource rsImpCta
Set rptGeneral = rptCuenta ' Asigna el reporte al objeto reporte general utilizado
                           ' en el Form de la Vista Previa.
frmVistaPrevia.Show vbModal

Set rptCuenta = Nothing

rsImpCta.Close
Set rsImpCta = Nothing
Exit Sub

SolucionErr:
     MsgBox Err.Number & " " & Err.Description, vbInformation, "ATENCION - ERROR ..."

End Sub
Private Sub Form_Load()
Me.Top = 2000
Me.Left = 4000
If Month(Date) = 1 Then 'como es el primer mes del año, debe referirse al dic del año anterior
    dtpIni.Value = Format("01/12/" & (Year(Date)) - 1, "dd/mm/yyyy")
    dtpFin.Value = Format("31/12/" & (Year(Date)) - 1, "dd/mm/yyyy")
Else
    dtpIni.Value = Format("01/" & Str((Month(Date)) - 1) & "/" & Str(Year(Date)), "dd/mm/yyyy")
    If (Month(Date) - 1) = 1 Or (Month(Date) - 1) = 3 Or (Month(Date) - 1) = 5 Or (Month(Date) - 1) = 7 Or (Month(Date) - 1) = 8 Or (Month(Date) - 1) = 10 Or (Month(Date) - 1) = 12 Then
        dtpFin.Value = Format("31/" & (Month(Date) - 1) & "/" & Year(Date), "dd/mm/yyyy")
    End If
    If (Month(Date) - 1) = 4 Or (Month(Date) - 1) = 6 Or (Month(Date) - 1) = 9 Or (Month(Date) - 1) = 11 Then
        dtpFin.Value = Format("30/" & (Month(Date) - 1) & "/" & Year(Date), "dd/mm/yyyy")
    End If
    If (Month(Date) - 1) = 2 Then
        dtpFin.Value = Format("28/" & (Month(Date) - 1) & "/" & Year(Date), "dd/mm/yyyy")
    End If
End If

frameFechas.Visible = False
End Sub
Private Sub Form_Unload(cancel As Integer)
If rsImpCta.State = 1 Then
    rsImpCta.Close
    Set rsImpCta = Nothing
End If
End Sub
Private Sub optFechas_Click()
optTodo.Value = False
frameFechas.Visible = True
End Sub
Private Sub optTodo_Click()
optFechas.Value = False
frameFechas.Visible = False
End Sub
