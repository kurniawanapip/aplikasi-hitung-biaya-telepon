VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "hitung waktu"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form2"
   ScaleHeight     =   5475
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5880
      Top             =   3360
   End
   Begin VB.CommandButton cmdtutup 
      Caption         =   "Tutup Telepon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdangkat 
      Caption         =   "Angkat Telepon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblbayar 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label Label8 
      Caption         =   "Jumlah Bayar"
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
      Left            =   1080
      TabIndex        =   9
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Jam"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lbljam 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   3360
      Width           =   3615
   End
   Begin VB.Label lbljamakhir 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lbljamawal 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Selesai Bicara"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Mulai Bicara"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "Contoh Program Waktu"
      BeginProperty Font 
         Name            =   "Brush Script Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vjam, vmenit, vdetik, vjamawal As Single
Dim vjamakhir, vdurasi, vpulsa, vbayar As Single
Private Sub cmdangkat_Click()
lbljamawal = Time
vjam = DatePart("h", lbljamawal)
vmenit = DatePart("n", lbljamawal)
vdetik = DatePart("s", lbljamawal)

vjamawal = (vjam * 3600) + (vmenit * 60) + vdetik
End Sub

Private Sub cmdtutup_Click()
lbljamakhir = Time
vjam = DatePart("h", lbljamakhir)
vmenit = DatePart("n", lbljamakhir)
vdetik = DatePart("s", lbljamakhir)

vjamakhir = (vjam * 3600) + (vmenit * 60) + vdetik
vdurasi = vjamakhir - vjamawal
vpulsa = vdurasi / 5
vbayar = vpulsa * 150

lblbayar = Format(vbayar, "currency")
End Sub

Private Sub Timer1_Timer()
lbljam = Time
End Sub
