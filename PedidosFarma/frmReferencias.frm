VERSION 5.00
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReferencias 
   Caption         =   "Corrigiendo Referencias..."
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10800
   Icon            =   "frmReferencias.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   10800
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Proceso"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   10575
      Begin MSComctlLib.ProgressBar pbBarra 
         Height          =   735
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   1296
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Referencia"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10575
      Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
         Height          =   615
         Left            =   2040
         Top             =   840
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   1085
         FormatString    =   "Actualizando id de RUBROS"
         Orientation     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextColor       =   8388863
      End
      Begin VB.Label Label1 
         Caption         =   $"frmReferencias.frx":0442
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   10335
      End
   End
   Begin AIFCmp1.asxPowerButton cmdSalir 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   3480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      Picture         =   "frmReferencias.frx":04CD
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
      PictureAlignment=   6
   End
   Begin AIFCmp1.asxPowerButton cmdok 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      Picture         =   "frmReferencias.frx":0A59
      Caption         =   "&Procesar"
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
Attribute VB_Name = "frmReferencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsCps As New ADODB.Recordset

Private Sub cmdok_Click()

SioNo = MsgBox("ESTA SEGURO DE PROCESAR LA INFORMACION ???!", vbCritical + vbYesNo, "ATENCION !!!")
If SioNo = vbNo Then
    Exit Sub
End If

rsCps.Open "select * from facturascompras order by rubro", cn, adOpenDynamic, adLockOptimistic, adCmdText

pbBarra.Min = 0
pbBarra.Max = rsCps.RecordCount


Dim vRubro As String

Dim c As Integer
c = 0
rsCps.MoveFirst
Do While rsCps.EOF = False
    vRubro = rsCps!rubro
    If vRubro = "Medicamentos" Then
        rsCps!idRubro = 2
        c = c + 1
    End If
    If vRubro = "Perfumeria" Then
        rsCps!idRubro = 1
        c = c + 1
    End If
    If vRubro = "Fragancias" Then
        rsCps!idRubro = 3
        c = c + 1
    End If
    pbBarra.Value = c
    rsCps.Update
    rsCps.MoveNext

Loop


MsgBox "PROCESO TERMINADO - " & c & " Registros actualizados....", vbInformation, "PROCESO TERMINADO !!!"

rsCps.Close
Set rsCps = Nothing
Unload Me

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

