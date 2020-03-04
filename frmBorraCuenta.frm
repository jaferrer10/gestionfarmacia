VERSION 5.00
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBorraCuenta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reducir Cuenta Corriente..."
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   Icon            =   "frmBorraCuenta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameOpc 
      Caption         =   "Opciones de ELIMINACION"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7215
      Begin VB.OptionButton optTodo 
         Caption         =   "Toda la Cuenta ..."
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
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
      Begin MSComCtl2.DTPicker dtpIni 
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
         Format          =   69009409
         CurrentDate     =   39392
      End
      Begin MSComCtl2.DTPicker dtpFin 
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
         Format          =   69009409
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
   Begin AIFCmp1.asxPowerButton cmdEliminar 
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      Picture         =   "frmBorraCuenta.frx":0442
      Caption         =   "&Eliminar"
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
      Picture         =   "frmBorraCuenta.frx":0894
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
Attribute VB_Name = "frmBorraCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdEliminar_Click()
Err.Clear
On erro GoTo SolucionErr
SioNo = MsgBox("ESTA SEGURO DE BORRAR DATOS DE LA CUENTA DE ESTE CLIENTE ???", vbCritical + vbYesNo, "ATENCION, BORRANDO CUENTA CORRIENTE ...")
If SioNo = vbYes Then
    If optTodo.Value = True Then
        cn.Execute "delete from CuentasCorrientes where idcliente = " & vIdCliente
    Else
        cn.Execute "delete from CuentasCorrientes where idcliente = " & vIdCliente & _
                   " and CuentasCorrientes.Fecha between #" & Format(dtpIni.Value, "mm/dd/yyyy") & _
                   "# and #" & Format(dtpFin.Value, "mm/dd/yyyy") & "#"
    End If
    Unload Me
End If
Exit Sub

SolucionErr:
        MsgBox Err.Number & " " & Err.Description, vbInformation, "ATENCION - ERROR ..."
End Sub
Private Sub Form_Load()

If TempNivel <> 1 Then
    MsgBox "NO POSEE AUTORIZACION PARA EJECUTAR ESE PROCESO ...", vbExclamation, "Atencion !!!"
    Exit Sub
End If
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
Private Sub optFechas_Click()
optTodo.Value = False
frameFechas.Visible = True
End Sub
Private Sub optTodo_Click()
optFechas.Value = False
frameFechas.Visible = False
End Sub
