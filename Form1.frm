VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5160
      Top             =   3000
   End
   Begin VB.CommandButton cmdTutup 
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
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton cmdAngkat 
      Caption         =   "Angkat Telepon "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblBayar 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label8 
      Caption         =   "Jumlah Bayar"
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
      TabIndex        =   9
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblJam 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label6 
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
      TabIndex        =   7
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblJamakhir 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblJamawal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
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
      Top             =   1920
      Width           =   1575
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
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "PROGRAM APLIKASI HITUNG WAKTU BIAYA TELEPON"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vjam, vmenit, vdetik, vjamawal As Single
Dim vjamakhir, vdurasi, vpulsa, vbayar As Single
Private Sub cmdAngkat_Click()
lblJamawal = Time
vjam = DatePart("h", lblJamawal)
vmenit = DatePart("n", lblJamawal)
vdetik = DatePart("s", lblJamawal)

vjamawal = (vjam * 3600) + (vmenit * 60) + vdetik

End Sub

Private Sub cmdTutup_Click()
lblJamakhir = Time
vjam = DatePart("h", lblJamakhir)
vmenit = DatePart("n", lblJamakhir)
vdetik = DatePart("s", lblJamakhir)

vjamakhir = (vjam * 3600) + (vmenit * 60) + vdetik
vdurasi = vjamakhir - vjamawal
vpulsa = vdurasi / 5
vbayar = vpulsa * 150

lblBayar = Format(vbayar, "currency")

End Sub

Private Sub Timer1_Timer()
lblJam = Time
End Sub
