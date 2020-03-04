VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Begin VB.Form frmEjecutarPuntoControl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ejecucion de Puntos de Control..."
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14370
   Icon            =   "frmEjecutarPuntoControl.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   14370
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Resultados del Punto de Control"
      Height          =   4455
      Left            =   240
      TabIndex        =   2
      Top             =   3840
      Width           =   14055
      Begin VB.TextBox txtCredito 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   21
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtResultado 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   19
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txtExtrac 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   18
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtVentas 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   17
         Top             =   600
         Width           =   1815
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx6 
         Height          =   240
         Left            =   11400
         TabIndex        =   13
         Top             =   2880
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   423
         Caption         =   "Resultado"
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx2 
         Height          =   240
         Left            =   11400
         TabIndex        =   9
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   423
         Caption         =   "Total Extracciones"
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx1 
         Height          =   240
         Left            =   11400
         TabIndex        =   8
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   423
         Caption         =   "Total Ventas"
      End
      Begin MSDataGridLib.DataGrid dtgCajas 
         Height          =   3975
         Left            =   4920
         TabIndex        =   4
         Top             =   360
         Width           =   6375
         _ExtentX        =   11245
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
         Caption         =   "CONTROL CAJAS"
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "Fecha"
            Caption         =   "Fecha"
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
            DataField       =   "Inicio"
            Caption         =   "Inicio"
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
            DataField       =   "Extracciones"
            Caption         =   "Extracciones"
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
            DataField       =   "Credito"
            Caption         =   "Credito"
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
            DataField       =   "Caja"
            Caption         =   "Caja"
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
            DataField       =   "Resultado"
            Caption         =   "Resultado"
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
         BeginProperty Column06 
            DataField       =   "Turno"
            Caption         =   "Turno"
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
         BeginProperty Column07 
            DataField       =   "Observaciones"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   824,882
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   840,189
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1635,024
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dtgVentas 
         Height          =   3975
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
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
         Caption         =   "VENTAS"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "Mañana"
            Caption         =   "Mañana"
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
            DataField       =   "Tarde"
            Caption         =   "Tarde"
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
            DataField       =   "Total"
            Caption         =   "Total Vtas"
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
            DataField       =   "Fecha"
            Caption         =   "Fecha"
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
               ColumnWidth     =   1049,953
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1035,213
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1049,953
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1200,189
            EndProperty
         EndProperty
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx7 
         Height          =   240
         Left            =   11400
         TabIndex        =   20
         Top             =   2040
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   423
         Caption         =   "Total Creditos"
      End
      Begin VB.Label Label1 
         Caption         =   "(Total Entregas - Efectivo de Vtas)"
         Height          =   255
         Left            =   11400
         TabIndex        =   22
         Top             =   3720
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Archivo de Puntos de Control"
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   14055
      Begin VB.TextBox txtExtras 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   9000
         TabIndex        =   23
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtMostrador 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   9000
         TabIndex        =   16
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txtGrande 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   9000
         TabIndex        =   15
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtChica 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   9000
         TabIndex        =   14
         Top             =   480
         Width           =   1815
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx5 
         Height          =   240
         Left            =   8280
         TabIndex        =   12
         Top             =   1680
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   423
         Caption         =   "Entrega Caja Mostrador"
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx4 
         Height          =   240
         Left            =   8280
         TabIndex        =   11
         Top             =   960
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   423
         Caption         =   "Entrega Caja Grande"
      End
      Begin MSDataGridLib.DataGrid dtgPuntos 
         Height          =   3015
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   5318
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   8454016
         HeadLines       =   1
         RowHeight       =   15
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
               LCID            =   11274
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
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin AIFCmp1.asxPowerButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   12000
         TabIndex        =   5
         Top             =   2760
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         BackColor       =   8421631
         Caption         =   "&Salir"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin AIFCmp1.asxPowerButton cmdEliminar 
         Height          =   495
         Left            =   12000
         TabIndex        =   6
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         Caption         =   "&Eliminar Punto"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextColor       =   255
      End
      Begin AIFCmp1.asxPowerButton cmdEjecutar 
         Height          =   495
         Left            =   12000
         TabIndex        =   7
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         Caption         =   "&Ejecutar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextColor       =   32768
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx3 
         Height          =   240
         Left            =   8280
         TabIndex        =   10
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   423
         Caption         =   "Entrega Caja Chica"
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx8 
         Height          =   240
         Left            =   8280
         TabIndex        =   24
         Top             =   2400
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   423
         Caption         =   "Extracciones Extraordinarias"
      End
   End
End
Attribute VB_Name = "frmEjecutarPuntoControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsPuntos As New ADODB.Recordset
Private rsVtas As New ADODB.Recordset
Private rsCajas As New ADODB.Recordset

Private Sub cmdEjecutar_Click()
If rsVtas.State = 1 Then
    rsVtas.Close
    Set rsVtas = Nothing
End If
rsVtas.Open "select mañana,tarde,total,fecha from ventas where fecha >= #" & Format(rsPuntos!fecha, "mm/dd/yyyy") & "# order by fecha", cn, adOpenDynamic, adLockReadOnly, adCmdText
If rsVtas.RecordCount = 0 Then
    MsgBox "NO HAY INFORMACION DE VENTAS APARTIR DE LA FECHA DEL PUNTO DE CONTROL EN ADELANTE !", vbExclamation, "ATENCION !"
    rsVtas.Close
    Set rsVtas = Nothing
    Exit Sub
End If
Set dtgVentas.DataSource = rsVtas
dtgVentas.Refresh

If rsCajas.State = 1 Then
    rsCajas.Close
    Set rsCajas = Nothing
End If
rsCajas.Open "select fecha,inicio,extracciones,credito,caja,resultado,turno,observaciones from controlcajas where fecha >= #" & Format(rsPuntos!fecha, "mm/dd/yyyy") & "# order by idcaja", cn, adOpenDynamic, adLockReadOnly, adCmdText
If rsCajas.RecordCount = 0 Then
    MsgBox "NO HAY INFORMACION DE CAJAS APARTIR DE LA FECHA DEL PUNTO DE CONTROL EN ADELANTE !", vbExclamation, "ATENCION !"
    rsCajas.Close
    Set rsCajas = Nothing
    Exit Sub
End If
Set dtgCajas.DataSource = rsCajas
dtgCajas.Refresh

'Total de ventas del periodo
rsVtas.MoveFirst
Dim vTvtas As Double
vTvtas = 0
Do While rsVtas.EOF = False
    vTvtas = vTvtas + rsVtas!total
    rsVtas.MoveNext
Loop
txtVentas.Text = vTvtas

'Total de extracciones y credito
Dim vExt As Double
Dim vCre As Double
vExt = 0
vCre = 0
rsCajas.MoveFirst
Do While rsCajas.EOF = False
    vExt = vExt + rsCajas!extracciones
    vCre = vCre + rsCajas!credito
    rsCajas.MoveNext
Loop
txtExtrac.Text = vExt
txtCredito.Text = vCre

'Total de efectivo por ventas
Dim TotEfecVtas As Double
TotEfecVtas = vTvtas - vCre

'total de efectivo entregado
Dim vTotEnt As Double
vTotEnt = 0
vTotEnt = (Val(txtChica.Text) + Val(txtGrande.Text) + Val(txtMostrador.Text)) - Val(txtExtras.Text)

'Calculo del resultado del control general
Dim ResulGral As Double
ResulGral = 0
ResulGral = vTotEnt - TotEfecVtas
txtResultado.Text = ResulGral

End Sub
Private Sub cmdEliminar_Click()
frmPideClave.Show vbModal
If TempNivel = 1 Then
    SioNo = MsgBox("ESTA SEGURO DE ELIMINAR ESTE PUNTO DE CONTROL ?", vbInformation + vbYesNo, "ATENCION !")
    If SioNo = vbYes Then
        rsPuntos.Delete
        rsPuntos.Update
        dtgPuntos.Refresh
    End If
Else
    MsgBox " NO TIENE AUTORIZACION PARA LA ELIMINACION DE REGISTROS ...", vbCritical, "ATENCION !"
End If
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub dtgPuntos_DblClick()
txtChica.SetFocus
End Sub
Private Sub Form_Load()
Me.Top = 50
Me.Left = 50
rsPuntos.Open "select * from puntocontrol order by fecha", cn, adOpenDynamic, adLockOptimistic, adCmdText
Set dtgPuntos.DataSource = rsPuntos
dtgPuntos.Refresh
If rsPuntos.RecordCount = 0 Then
    MsgBox "NO HAY PUNTOS DE CONTROL REGISTRADOS....!", vbExclamation, "ATENCION !!!"
    rsPuntos.Close
    Set rsPuntos = Nothing
    Unload Me
End If
End Sub
Private Sub Form_Unload(cancel As Integer)
rsPuntos.Close
Set rsPuntos = Nothing
If rsVtas.State = 1 Then
    rsVtas.Close
    Set rsVtas = Nothing
End If
If rsCajas.State = 1 Then
    rsCajas.Close
    Set rsCajas = Nothing
End If
End Sub
Private Sub txtChica_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtGrande.SetFocus
        SendKeys "{home}+{end}"
    Case 44 'para que no acepte la entrada de la coma
        KeyAscii = 0
    Case 45
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub

Private Sub txtExtras_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        cmdEjecutar.SetFocus
    Case 44 'para que no acepte la entrada de la coma
        KeyAscii = 0
    Case 45
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub

Private Sub txtGrande_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtMostrador.SetFocus
        SendKeys "{home}+{end}"
    Case 44 'para que no acepte la entrada de la coma
        KeyAscii = 0
    Case 45
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtMostrador_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtExtras.SetFocus
        SendKeys "{home}+{end}"
    Case 44 'para que no acepte la entrada de la coma
        KeyAscii = 0
    Case 45
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
