VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form welcome 
   Caption         =   "SELAMAT DATANG ( APLIKASI PENGANTAR SURAT DESA KROYO, KABUPATEN SRAGEN )"
   ClientHeight    =   9165
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   20370
   LinkTopic       =   "Form2"
   Picture         =   "welcome.frx":0000
   ScaleHeight     =   9165
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   16680
      Top             =   0
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   16320
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Admnistrator"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   17520
      TabIndex        =   6
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Admnistrator"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   17520
      TabIndex        =   5
      Top             =   240
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   16440
      Picture         =   "welcome.frx":CA741
      Stretch         =   -1  'True
      Top             =   240
      Width           =   960
   End
   Begin VB.Label Label88 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "--/--/----"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   -480
      TabIndex        =   4
      Top             =   2160
      Width           =   9015
   End
   Begin VB.Label Label77 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "--:--:--"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   5160
      TabIndex        =   3
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DESA KROYO, KARANGMALANG, SRAGEN"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   3240
      TabIndex        =   2
      Top             =   5880
      Width           =   13935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SISTEM INFORMASIS DATA PENDUDUK"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   3240
      TabIndex        =   1
      Top             =   5160
      Width           =   13935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SELAMAT DATANG..."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   48
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1575
      Left            =   3240
      TabIndex        =   0
      Top             =   3840
      Width           =   11295
   End
   Begin VB.Menu master 
      Caption         =   "MASTER"
      Begin VB.Menu datpen 
         Caption         =   "DATA PENDUDUK"
      End
      Begin VB.Menu report 
         Caption         =   "LAPORAN SURAT"
      End
      Begin VB.Menu surat 
         Caption         =   "TRANSAKSI SURAT"
      End
   End
   Begin VB.Menu user 
      Caption         =   "USER"
      Begin VB.Menu admin 
         Caption         =   "DATA ADMIN"
      End
      Begin VB.Menu keluar 
         Caption         =   "LOG-OUT"
      End
   End
   Begin VB.Menu setting 
      Caption         =   "PENGATURAN"
   End
End
Attribute VB_Name = "welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






'======================= FORM HALAMAN DEPAN CODE===========================
     '======================= BUDIHARTO===========================
Private Sub Form_Load()
Label5 = login.Text1.Text
End Sub

'TAMPILKAN FORM TRANSAKSI
Private Sub surat_Click()
menu_utama.Show
End Sub

Private Sub datpen_Click()
penduduk.Show
End Sub

'TAMPILKAN LAPORAN KESELURUHAN DATA
Private Sub report_Click()
CR1.Reset
With CR1
    .ReportFileName = App.Path & "\Data_surat.rpt"
    .WindowState = crptMaximized
    .RetrieveDataFiles
    .Action = 1
End With
End Sub

'TAMPILKAN FORM ADMN
Private Sub admin_Click()
form1.Show
End Sub

'KELUAR APLIKASI
Private Sub keluar_Click()
xx = MsgBox("Apakah Anda yakin akan kelua dari aplikasi pengantar surat desa kroyo ?", vbOKCancel, "Peringatan")
            If xx = vbOK Then
                    Unload Me
                Else
                    'NO NOTIF
            End If
End Sub

'TAMPILKAN PENGATURAN (NAMA LURAH & CAMAT)
Private Sub setting_Click()
pengaturan.Show
End Sub

'SOURCE JAM BERJALAN
Private Sub Timer1_Timer()
Label77.Caption = Format(Now, "hh : mm : ss")
Label88.Caption = Format(Now, "dd MMMM yyyy")
End Sub
