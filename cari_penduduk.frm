VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form cari_penduduk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CARI DATA PENDUDUK"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13845
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   13845
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
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
      RecordSource    =   "penduduk"
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
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FFFF&
      Caption         =   "CARI"
      Height          =   495
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   1320
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
      Left            =   6120
      TabIndex        =   3
      Top             =   960
      Width           =   7575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "cari_penduduk.frx":0000
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   7223
      _Version        =   393216
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      Picture         =   "cari_penduduk.frx":0015
      ScaleHeight     =   735
      ScaleWidth      =   13815
      TabIndex        =   0
      Top             =   0
      Width           =   13815
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "CARI DATA PENDUDUK"
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
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   9135
      End
   End
End
Attribute VB_Name = "cari_penduduk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CARI DATA
Private Sub Command5_Click()
Adodc1.Recordset.Filter = "nama like '%" + Me.Text14.Text + "%' or nik like '%" + Me.Text14.Text + "%'"
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
menu_utama.Text2 = Adodc1.Recordset!nama
If Adodc1.Recordset!jk = "Laki-Laki" Then
    menu_utama.Option1.Value = True
ElseIf Adodc1.Recordset!jk = "Perempuan" Then
    menu_utama.Option2.Value = True
End If
menu_utama.Text3 = Adodc1.Recordset!tempat_lahir
menu_utama.DTPicker1 = Adodc1.Recordset!tanggal_lahir
menu_utama.Combo1 = Adodc1.Recordset!kewarganegaraan
menu_utama.Combo2 = Adodc1.Recordset!pekerjaan
menu_utama.Combo3 = Adodc1.Recordset!status_perkawinan
menu_utama.Combo4 = Adodc1.Recordset!alamat
menu_utama.Text5 = Adodc1.Recordset!nik
menu_utama.Text6 = Adodc1.Recordset!no_kk
menu_utama.Show
End Sub
