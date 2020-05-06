VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form menu_utama 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Transaksi Surat (Aplikasi Pengantar Surat Desa Kroyo )"
   ClientHeight    =   10755
   ClientLeft      =   8580
   ClientTop       =   1620
   ClientWidth     =   20370
   LinkTopic       =   "Form2"
   PaletteMode     =   2  'Custom
   Picture         =   "menuutama.frx":0000
   ScaleHeight     =   10755
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   8400
      TabIndex        =   57
      Top             =   5760
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   151584768
      CurrentDate     =   43194
   End
   Begin VB.ComboBox text7 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8400
      TabIndex        =   56
      Top             =   4320
      Width           =   3615
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FFFF&
      Caption         =   "CARI"
      Height          =   495
      Left            =   18360
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   3000
      Width           =   1695
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1380
      Left            =   240
      OLEDropMode     =   1  'Manual
      Picture         =   "menuutama.frx":3F95
      ScaleHeight     =   1350
      ScaleWidth      =   1035
      TabIndex        =   53
      Top             =   0
      Width           =   1065
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   19800
      Top             =   0
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   -360
      Picture         =   "menuutama.frx":5846
      ScaleHeight     =   1455
      ScaleWidth      =   21135
      TabIndex        =   45
      Top             =   0
      Width           =   21135
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "DESA KROYO, KECAMATAN KARANGMALANG"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Index           =   0
         Left            =   5760
         TabIndex        =   49
         Top             =   720
         Width           =   9375
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "APLIKASI PENGANTAR SURAT"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   615
         Index           =   0
         Left            =   6840
         TabIndex        =   48
         Top             =   120
         Width           =   7695
      End
      Begin VB.Label Label88 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "--/--/----"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   16800
         TabIndex        =   47
         Top             =   240
         Width           =   4095
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
         Left            =   16800
         TabIndex        =   46
         Top             =   600
         Width           =   4095
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   9480
      Top             =   10320
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=surat_kroyo"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "surat_kroyo"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "combo"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   12240
      Top             =   8880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2640
      TabIndex        =   43
      Top             =   5160
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   151584768
      CurrentDate     =   43140
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6480
      Top             =   10320
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=surat_kroyo"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "surat_kroyo"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "datamas"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Surat Bukti Diri"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6720
      TabIndex        =   30
      Top             =   2520
      Width           =   5535
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   35
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   34
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label10 
         Caption         =   "No KK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "No KTP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      Picture         =   "menuutama.frx":21C6F
      ScaleHeight     =   615
      ScaleWidth      =   22410
      TabIndex        =   29
      Top             =   10200
      Width           =   22410
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H000080FF&
      Caption         =   "CETAK DATA SURAT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   9480
      UseMaskColor    =   -1  'True
      Width           =   20175
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Menu Utama"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7815
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   12135
      Begin VB.ComboBox Combo4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   54
         Top             =   6480
         Width           =   3615
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   -120
         Picture         =   "menuutama.frx":3E099
         ScaleHeight     =   735
         ScaleWidth      =   12375
         TabIndex        =   44
         Top             =   0
         Width           =   12375
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "FORM PENGISIAN SURAT"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   18
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   495
            Index           =   1
            Left            =   240
            TabIndex        =   51
            Top             =   120
            Width           =   5655
         End
      End
      Begin VB.ComboBox Combo3 
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   42
         Top             =   5880
         Width           =   3615
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H8000000D&
         Caption         =   "UBAH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   7080
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000D&
         Caption         =   "SIMPAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   6360
         Width           =   2535
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H8000000D&
         Caption         =   "HAPUS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   7080
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "CARI DATA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   6360
         Width           =   2655
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8280
         TabIndex        =   31
         Top             =   5640
         Width           =   3615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Perempuan"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   25
         Top             =   2400
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Laki-Laki"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2520
         TabIndex        =   24
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2520
         TabIndex        =   9
         Top             =   2880
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   1800
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox textkab 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Text            =   "Sragen Provinsi Jawa Tengah"
         Top             =   7200
         Width           =   3615
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   5
         Top             =   3480
         Width           =   3615
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   4
         Top             =   4920
         Width           =   3615
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   3
         Top             =   4440
         Width           =   3615
      End
      Begin VB.ComboBox Combo2 
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   2
         Top             =   5160
         Width           =   3615
      End
      Begin VB.Label Label4 
         Caption         =   "Tanggal Lahir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   27
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Nomor Surat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Nama Pemohon"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Jenis Kelamin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   21
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Tempat Lahir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Kewarganegaraa /Agama"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   4440
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Pekerjaan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Tempat tinggal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   6600
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Kabupaten"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   7200
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Status perkawinan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   5880
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Keperluan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   14
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Berlaku mulai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   13
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Keperluan Lain*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   12
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Kepala desa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   11
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Camat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   10
         Top             =   5760
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   12480
      TabIndex        =   0
      Top             =   1560
      Width           =   7815
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "menuutama.frx":5A4C3
         Height          =   5415
         Left            =   120
         TabIndex        =   50
         Top             =   2160
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   9551
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Enabled         =   -1  'True
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
            Name            =   "Tahoma"
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
               LCID            =   1057
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
               LCID            =   1057
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
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   360
         TabIndex        =   41
         Top             =   1440
         Width           =   5415
      End
      Begin VB.Frame Frame4 
         Caption         =   "CARI DATA BERDASARKAN NAMA/NOMOR KTP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   40
         Top             =   1080
         Width           =   7575
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   0
         Picture         =   "menuutama.frx":5A4D8
         ScaleHeight     =   735
         ScaleWidth      =   7815
         TabIndex        =   28
         Top             =   0
         Width           =   7815
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "DATA SURAT"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   18
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   495
            Index           =   2
            Left            =   240
            TabIndex        =   52
            Top             =   120
            Width           =   3015
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   9480
      Top             =   5160
      Width           =   1215
   End
End
Attribute VB_Name = "menu_utama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================FORM TRANSAKASI CODE===========================
     '======================= BUDIHARTO===========================


'MENDEKLARASIKAN DATACOMBO
Dim DataCombo As New ADODB.Recordset

'SCRIPT AUTONUMBER SURAT


'MENAMPILKAN DATA PADA DATABASE KE COMBO
Sub tambahcom()
Adodc2.ConnectionString = conn.ConnectionString
Adodc2.RecordSource = "select* from combo"

For Each gosong In Me.Controls
If TypeOf gosong Is ComboBox Then
gosong.Text = ""
With Adodc2.Recordset
    Do While Not .EOF
    On Error Resume Next
    Combo1.AddItem !kewarganegaraan_dan_agama
    Combo2.AddItem !pekerjaan
    Combo3.AddItem !status_perkawinan
    Combo4.AddItem !alamat
    Text9.Text = !lurah
    Text10.Text = !camat
    text7.AddItem !keperluan
    .MoveNext
    Loop
End With
End If
Next
End Sub

'CLEAR FORM
Sub bersih()
Text1 = ""
Text2 = ""
Text3 = ""
Combo4 = ""
Text5 = ""
Text6 = ""
text7 = ""
Text8 = ""
Combo1 = ""
Combo2 = ""
Combo3 = ""
Option1.Value = False
Option2.Value = False
DTPicker1.Value = Now
DTPicker2.Value = Now
End Sub

'ENABLE TRUE FORM
Sub tambah()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Combo4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
text7.Enabled = True
Text8.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
DTPicker1.Enabled = True
Text1.SetFocus
End Sub

'FORMAT WAKTU DATAGRID
Sub pormat()
DataGrid1.Columns(4).NumberFormat = ("DD/MM/YYYY")
DataGrid1.Columns(14).NumberFormat = ("DD/MM/YYYY")
End Sub


'JIKA TOMBOL TAMBAH DI KLIK PANGGIL FUNGSI TERSEBUT
Private Sub Command1_Click()
Call tambahcom
Call tambah
Call bersih
cari_penduduk.Show
'Call AutoNumber
End Sub

'SIMPAN DATA
Private Sub Command2_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or (Option1.Value = False And Option2.Value = False) Or Combo1 = "" Or Combo2 = "" Or Combo3 = "" Or Combo4 = "" Or textkab = "" Or Text5 = "" Or Text6 = "" Or Text9 = "" Or Text10 = "" Then
MsgBox "LENGKAPI DAHULU DATA YANG AKAN ANDA INPUTKAN !", vbInformation, "PERHATIAN !"
Else
Adodc1.Recordset.AddNew
Adodc1.Recordset!nomor = Text1.Text
Adodc1.Recordset!nama = Text2.Text
If Option1.Value = True Then
    Adodc1.Recordset!jenis_kelamin = Option1.Caption
ElseIf Option2.Value = True Then
    Adodc1.Recordset!jenis_kelamin = Option2.Caption
End If
Adodc1.Recordset!tempat_lahir = Text3.Text
Adodc1.Recordset!tanggal_lahir = DTPicker1.Value
Adodc1.Recordset!kewarganegaraan_dan_agama = Combo1.Text
Adodc1.Recordset!pekerjaan = Combo2.Text
Adodc1.Recordset!status_perkawinan = Combo3.Text
Adodc1.Recordset!tempat_tinggal = Combo4.Text
Adodc1.Recordset!kabupaten = textkab.Text
Adodc1.Recordset!no_ktp = Text5.Text
Adodc1.Recordset!no_kk = Text6.Text
Adodc1.Recordset!keperluan = text7.Text
Adodc1.Recordset!berlaku_mulai = DTPicker2.Value
Adodc1.Recordset!keterangan_lain_lain = Text8.Text
Adodc1.Recordset!kepala_desa = Text9.Text
Adodc1.Recordset!camat = Text10.Text
Adodc1.Recordset.Update
Call bersih
MsgBox "DATA ANDA BERHASIL KAMI SIMPAN !", vbInformation, "INFORMASI !"
Adodc1.Refresh
End If
Call pormat
End Sub

'UBAH DATA
Private Sub Command3_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or (Option1.Value = False And Option2.Value = False) Or Combo1 = "" Or Combo2 = "" Or Combo3 = "" Or Combo4 = "" Or textkab = "" Or Text5 = "" Or Text6 = "" Or Text9 = "" Or Text10 = "" Then
MsgBox "PILIH DAHULU DATA YANG AKAN ANDA HAPUS !", vbInformation, "PERHATIAN !"
Else
Adodc1.Recordset!nomor = Text1.Text
Adodc1.Recordset!nama = Text2.Text
If Option1.Value = True Then
    Adodc1.Recordset!jenis_kelamin = Option1.Caption
ElseIf Option2.Value = True Then
    Adodc1.Recordset!jenis_kelamin = Option2.Caption
End If
Adodc1.Recordset!tempat_lahir = Text3.Text
Adodc1.Recordset!tanggal_lahir = DTPicker1.Value
Adodc1.Recordset!kewarganegaraan_dan_agama = Combo1.Text
Adodc1.Recordset!pekerjaan = Combo2.Text
Adodc1.Recordset!status_perkawinan = Combo3.Text
Adodc1.Recordset!tempat_tinggal = Combo4.Text
Adodc1.Recordset!kabupaten = textkab.Text
Adodc1.Recordset!no_ktp = Text5.Text
Adodc1.Recordset!no_kk = Text6.Text
Adodc1.Recordset!keperluan = text7.Text
Adodc1.Recordset!berlaku_mulai = DTPicker2.Value
Adodc1.Recordset!keterangan_lain_lain = Text8.Text
Adodc1.Recordset!kepala_desa = Text9.Text
Adodc1.Recordset!camat = Text10.Text
Adodc1.Recordset.Update
Call bersih
MsgBox "DATA ANDA BERHASIL KAMI UBAH !", vbInformation, "INFORMASI !"
Adodc1.Refresh
End If
End Sub

'HAPUS DATA
Private Sub Command4_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or (Option1.Value = False And Option2.Value = False) Or Combo1 = "" Or Combo2 = "" Or Combo3 = "" Or Combo4 = "" Or textkab = "" Or Text5 = "" Or Text6 = "" Or Text9 = "" Or Text10 = "" Then
MsgBox "PILIH DAHULU DATA YANG AKAN ANDA HAPUS !", vbInformation, "PERHATIAN !"
Else
xx = MsgBox("Apakah Anda yakin akan menghapus data?", vbOKCancel, "Peringatan")
            If xx = vbOK Then
               Adodc1.Recordset.Delete
               Call bersih
MsgBox "DATA ANDA BERHASIL KAMI HAPUS !", vbInformation, "INFORMASI !"
Adodc1.Refresh
            End If
           
End If
End Sub

'CARI DATA
Private Sub Command5_Click()
Adodc1.Recordset.Filter = "nama like '%" + Me.Text14.Text + "%' or NO_ktp like '%" + Me.Text14.Text + "%'"
End Sub

'MUNCULKAN DATA SAAT PENCARIAN BERAKHIR
Private Sub Text14_Change()
If Text14.Text = "" Then
Adodc1.Refresh
Else
'wkwk
End If
End Sub

'PINDAH DATA DARI DATAGRIDVIEW KE TEXTBOX
Private Sub DataGrid1_Click()
Text1 = Adodc1.Recordset!nomor
Text2 = Adodc1.Recordset!nama
If Adodc1.Recordset!jenis_kelamin = "Laki-Laki" Then
    Option1.Value = True
ElseIf Adodc1.Recordset!jenis_kelamin = "Perempuan" Then
    Option2.Value = True
End If
Text3 = Adodc1.Recordset!tempat_lahir
DTPicker1 = Adodc1.Recordset!tanggal_lahir
Combo1 = Adodc1.Recordset!kewarganegaraan_dan_agama
Combo2 = Adodc1.Recordset!pekerjaan
Combo3 = Adodc1.Recordset!status_perkawinan
Combo4 = Adodc1.Recordset!tempat_tinggal
textkab = Adodc1.Recordset!kabupaten
Text5 = Adodc1.Recordset!no_ktp
Text6 = Adodc1.Recordset!no_kk
text7 = Adodc1.Recordset!keperluan
DTPicker2 = Adodc1.Recordset!berlaku_mulai
Text8 = Adodc1.Recordset!keterangan_lain_lain
Text9 = Adodc1.Recordset!kepala_desa
Text10 = Adodc1.Recordset!camat
End Sub

'MUNCULKAN LAPORAN
Private Sub Command7_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or (Option1.Value = False And Option2.Value = False) Or Combo1 = "" Or Combo2 = "" Or Combo3 = "" Or Combo4 = "" Or textkab = "" Or Text5 = "" Or Text6 = "" Or Text9 = "" Or Text10 = "" Then
MsgBox "PILIH DAHULU DATA YANG AKAN ANDA HAPUS !", vbInformation, "PERHATIAN !"
Else
With CR1
    .SelectionFormula = "{datamas.nomor}='" & Text1.Text & "'"
    .ReportFileName = App.Path & "\lap_surat.rpt"
    .WindowState = crptMaximized
    .RetrieveDataFiles
    .Action = 1
End With
End If
'cr1.Connect = "dsn=database"
'cr1.ReportFileName = App.Path & "\lap_surat.rpt"
'cr1.Action = 1
End Sub

Private Sub DATA_ADMIN_Click()
admin.Show
End Sub

'FUNGSI AKTIF OTOMATIS SAAT FORM DIBUKA
Private Sub Form_Load()
'combo 1
Call tambahcom
Call bersih
Call pormat
End Sub

'WAKTU
Private Sub Timer1_Timer()
Label77.Caption = Format(Now, "hh : mm : ss")
Label88.Caption = Format(Now, "dd MMMM yyyy")
End Sub
