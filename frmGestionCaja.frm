VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmGestionCaja 
   Caption         =   "Movimiento de Caja..."
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   Icon            =   "frmGestionCaja.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8370
   ScaleWidth      =   11070
   Begin VB.TextBox txtInicio 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1034
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin AIFCmp1.asxPowerButton cmdExt 
      Height          =   615
      Left            =   3120
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      Picture         =   "frmGestionCaja.frx":058A
      Caption         =   "Extracción"
      CaptionAlignment=   5
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
      PictureOffsetX  =   10
   End
   Begin AIFCmp1.asxPowerButton cmdSalir 
      Height          =   1335
      Left            =   9000
      TabIndex        =   8
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2355
      Picture         =   "frmGestionCaja.frx":0B24
      Caption         =   "Salir"
      CaptionAlignment=   7
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
   Begin VB.Frame Frame2 
      Caption         =   "Resultados"
      Height          =   6255
      Left            =   6960
      TabIndex        =   12
      Top             =   1920
      Width           =   3975
      Begin VB.TextBox txtVentas 
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
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   4440
         Width           =   1815
      End
      Begin VB.TextBox txtTotCaja 
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
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   5400
         Width           =   1815
      End
      Begin VB.TextBox txtExt 
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
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   3480
         Width           =   1815
      End
      Begin VB.TextBox txtRein 
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
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox txtCredito 
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
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtEfectivo 
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
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   600
         Width           =   1815
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx3 
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   423
         Caption         =   "TOTAL CREDITO"
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx2 
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   423
         Caption         =   "TOTAL EFECTIVO"
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx4 
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   423
         Caption         =   "TOTAL REINTEGROS"
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx5 
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   3120
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   423
         Caption         =   "TOTAL EXTRACCIONES"
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx6 
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   5040
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   423
         Caption         =   "TOTAL CAJA"
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx7 
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   4080
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   423
         Caption         =   "TOTAL VENTAS"
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Archivo"
      Height          =   6255
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   6495
      Begin AIFCmp1.asxPowerButton cmdBorrar 
         Height          =   495
         Left            =   4680
         TabIndex        =   5
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmGestionCaja.frx":112C
         Caption         =   "&Borrar"
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
         PictureOffsetX  =   10
      End
      Begin MSDataGridLib.DataGrid dtgArchivo 
         Height          =   5775
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   10186
         _Version        =   393216
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "fecha"
            Caption         =   "Fecha"
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
            DataField       =   "hora"
            Caption         =   "Hora"
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
         BeginProperty Column02 
            DataField       =   "importe"
            Caption         =   "Importe"
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
         BeginProperty Column03 
            DataField       =   "concepto"
            Caption         =   "Concepto"
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
            BeginProperty Column00 
               ColumnWidth     =   1289,764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   945,071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1365,165
            EndProperty
         EndProperty
      End
      Begin AIFCmp1.asxPowerButton cmdRecal 
         Height          =   495
         Left            =   4680
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmGestionCaja.frx":16C6
         Caption         =   "&Recalcular"
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
      Begin AIFCmp1.asxPowerButton cmdReintegrar 
         Height          =   495
         Left            =   4680
         TabIndex        =   7
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmGestionCaja.frx":1C60
         Caption         =   "Poner Reintegro"
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
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1034
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
   Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx1 
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Importe"
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
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
      Format          =   67960833
      CurrentDate     =   40550
   End
   Begin AIFCmp1.asxPowerButton cmdRein 
      Height          =   615
      Left            =   5040
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      Picture         =   "frmGestionCaja.frx":21FA
      Caption         =   "Reintegro"
      CaptionAlignment=   5
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
      PictureOffsetX  =   10
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Inicio de Caja >"
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
      Height          =   240
      Left            =   3480
      TabIndex        =   24
      Top             =   360
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Caja:"
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
      TabIndex        =   0
      Top             =   360
      Width           =   1260
   End
End
Attribute VB_Name = "frmGestionCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsCaja As New ADODB.Recordset

Private Sub cmdBorrar_Click()
If rsCaja.EOF = True Then
    Exit Sub
End If
vidCaja = rsCaja!idcaja
If rsCaja.RecordCount = 0 Then
    MsgBox "NO HAY REGISTROS PARA BORRAR ...!", vbCritical, "ATENCION !"
    Exit Sub
End If
cn.Execute "delete from movimientoscaja where idcaja = " & vidCaja
Call Recalculo
rsCaja.Requery
dtgArchivo.Refresh

End Sub

Private Sub cmdExt_Click()
If Len(txtImporte.Text) = 0 Then
    MsgBox "DEBE INGRESAR UN IMPORTE PARA LA EXTRACCION..!", vbCritical, "REGISTRANDO REINTEGRO..."
    Exit Sub
End If
rsCaja.AddNew
rsCaja!fecha = dtpFecha.Value
rsCaja!Hora = Time()
rsCaja!Importe = Val(txtImporte.Text) * -1
rsCaja!concepto = "EXTRACCION"

rsCaja.Update
rsCaja.Requery
dtgArchivo.Refresh

'calcular totales para la pantalla
Call Recalculo

txtImporte.SetFocus
SendKeys "{end}+{home}", 2

End Sub

Private Sub cmdRecal_Click()
Call Recalculo
End Sub

Private Sub cmdRein_Click()
If Len(txtImporte.Text) = 0 Then
    MsgBox "DEBE INGRESAR UN IMPORTE PARA EL REINTEGRO..!", vbCritical, "REGISTRANDO REINTEGRO..."
    Exit Sub
End If
rsCaja.AddNew
rsCaja!fecha = dtpFecha.Value
rsCaja!Hora = Time()
rsCaja!Importe = Val(txtImporte.Text) * -1
rsCaja!concepto = "REINTEGRO"

rsCaja.Update
rsCaja.Requery
dtgArchivo.Refresh

'calcular totales para la pantalla
Call Recalculo

txtImporte.SetFocus
SendKeys "{end}+{home}", 2

End Sub

Private Sub cmdReintegrar_Click()
If rsCaja.EOF = True Then
    Exit Sub
End If
rsCaja!Importe = (rsCaja!Importe) * -1
rsCaja!concepto = "REINTEGRO"

rsCaja.Update
rsCaja.Requery
dtgArchivo.Refresh

'calcular totales para la pantalla
Call Recalculo

txtImporte.SetFocus
SendKeys "{end}+{home}", 2
End Sub

Private Sub cmdSalir_Click()
rsCaja.Close
Unload Me
End Sub
Private Sub dtpFecha_Change()
rsCaja.Close
rsCaja.Open "select * from movimientoscaja where fecha = #" & Format(dtpFecha.Value, "mm/dd/yyyy") & "#", cn, adOpenDynamic, adLockOptimistic, adCmdText
Set dtgArchivo.DataSource = rsCaja
dtgArchivo.Refresh

Call Recalculo

txtImporte.SetFocus
SendKeys "{end}+{home}", 2
End Sub
Private Sub Form_Load()
Me.Height = 8880
Me.Width = 11190
Me.Top = 200
Me.Left = 2000

dtpFecha.Value = Date
rsCaja.Open "select * from movimientoscaja where fecha = #" & Format(dtpFecha.Value, "mm/dd/yyyy") & "#", cn, adOpenDynamic, adLockOptimistic, adCmdText
Set dtgArchivo.DataSource = rsCaja
dtgArchivo.Refresh

End Sub
Private Sub GrabaImporte()
If Len(txtImporte.Text) = 0 Then
    Exit Sub
End If
SioNo = MsgBox(" 'SI' PARA EFECTIVO - 'NO' PARA CREDITO - 'CANCELAR' PARA ANULAR", vbYesNoCancel + vbInformation, "SELECCIONE CONCEPTO DEL IMPORTE ...")
If SioNo = vbCancel Then
    Exit Sub
End If

rsCaja.AddNew
rsCaja!fecha = dtpFecha.Value
rsCaja!Hora = Time()
rsCaja!Importe = txtImporte.Text

If SioNo = vbYes Then
    rsCaja!concepto = "EFECTIVO"
Else
    rsCaja!concepto = "CREDITO"
End If

rsCaja.Update
rsCaja.Requery
dtgArchivo.Refresh

'calcular totales para la pantalla
Call Recalculo


End Sub
Private Sub txtImporte_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        
        Call GrabaImporte
        
        txtImporte.SetFocus
        SendKeys "{end}+{home}", 2
    Case 44 'para que no acepte la coma
        KeyAscii = 0
    Case 45
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select

End Sub

Private Sub Recalculo()
If rsCaja.RecordCount > 0 Then
    rsCaja.MoveFirst
    vEfe = 0
    vCre = 0
    vExt = 0
    vRei = 0
    rsCaja.MoveFirst
    Do While rsCaja.EOF = False
            If rsCaja!concepto = "EFECTIVO" Then
                vEfe = vEfe + rsCaja!Importe
            End If
            If rsCaja!concepto = "CREDITO" Then
                vCre = vCre + rsCaja!Importe
            End If
            If rsCaja!concepto = "EXTRACCION" Then
                vExt = vExt + rsCaja!Importe
            End If
            If rsCaja!concepto = "REINTEGRO" Then
                vRei = vRei + rsCaja!Importe
            End If
            rsCaja.MoveNext
    Loop
    txtEfectivo.Text = vEfe
    txtCredito.Text = vCre
    txtExt.Text = vExt
    txtRein.Text = vRei
    
    'calcula el resultado final de caja
    txtVentas.Text = (vEfe + vCre) - vRei
    
    vIni = Val(txtInicio.Text)
    
    Totefe = vIni + vEfe
    
    TotResto = Abs(vExt) ' + Abs(vRei) + vCre
    
    vtotC = Totefe - TotResto
    
    txtTotCaja.Text = vtotC

End If


End Sub
Private Sub txtInicio_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        
        Call Recalculo

        txtImporte.SetFocus
        SendKeys "{end}+{home}", 2
    Case 44 'para que no acepte la coma
        KeyAscii = 0
    Case 45
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub

