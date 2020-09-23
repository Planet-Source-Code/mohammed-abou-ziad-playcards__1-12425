VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmplaycards 
   Caption         =   "Form1"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox Listhcards 
      Columns         =   1
      Height          =   840
      ItemData        =   "playcards.frx":0000
      Left            =   9240
      List            =   "playcards.frx":0007
      TabIndex        =   23
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox Listccards 
      Columns         =   1
      Height          =   840
      ItemData        =   "playcards.frx":0017
      Left            =   9000
      List            =   "playcards.frx":001E
      TabIndex        =   22
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   2760
   End
   Begin VB.PictureBox Imagec 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   2
      Left            =   3480
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox Imagec 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   3
      Left            =   4800
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox Imagec 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   1
      Left            =   2160
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox Imagec 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   0
      Left            =   840
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox nullp 
      Height          =   1455
      Left            =   8760
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Top             =   5280
      Width           =   855
   End
   Begin VB.PictureBox Imageh 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   0
      Left            =   960
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1125
   End
   Begin VB.PictureBox Imageh 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   1
      Left            =   2280
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1125
   End
   Begin VB.PictureBox Imageh 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   3
      Left            =   4920
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1125
   End
   Begin VB.PictureBox Imageh 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   2
      Left            =   3600
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1125
   End
   Begin VB.PictureBox image1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   3
      Left            =   3720
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   4
      Top             =   2520
      Width           =   1125
   End
   Begin VB.PictureBox image1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   2
      Left            =   2520
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   3
      Top             =   1800
      Width           =   1125
   End
   Begin VB.PictureBox image1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   1
      Left            =   1320
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   2
      Top             =   2520
      Width           =   1125
   End
   Begin VB.PictureBox image1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1500
      Index           =   0
      Left            =   120
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   1
      Top             =   1800
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Draw"
      Height          =   375
      Left            =   9000
      TabIndex        =   0
      Top             =   5520
      Width           =   1095
   End
   Begin MSComctlLib.ImageList ImageListb 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   70
      ImageHeight     =   95
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":002E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListc 
      Left            =   8400
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   71
      ImageHeight     =   97
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   52
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":4F2E
            Key             =   ""
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":A15A
            Key             =   ""
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":F2AE
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":14402
            Key             =   ""
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":19556
            Key             =   ""
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":1E6AA
            Key             =   ""
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":237FE
            Key             =   ""
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":28952
            Key             =   ""
            Object.Tag             =   "8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":2DAA6
            Key             =   ""
            Object.Tag             =   "9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":32A7A
            Key             =   ""
            Object.Tag             =   "10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":37BCE
            Key             =   ""
            Object.Tag             =   "11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":3CD22
            Key             =   ""
            Object.Tag             =   "12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":41E76
            Key             =   ""
            Object.Tag             =   "13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":46E4A
            Key             =   ""
            Object.Tag             =   "14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":4BF9E
            Key             =   ""
            Object.Tag             =   "15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":50F72
            Key             =   ""
            Object.Tag             =   "16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":560C6
            Key             =   ""
            Object.Tag             =   "17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":5B21A
            Key             =   ""
            Object.Tag             =   "18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":601EE
            Key             =   ""
            Object.Tag             =   "19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":65342
            Key             =   ""
            Object.Tag             =   "20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":6A496
            Key             =   ""
            Object.Tag             =   "21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":6F46A
            Key             =   ""
            Object.Tag             =   "22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":745BE
            Key             =   ""
            Object.Tag             =   "23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":79712
            Key             =   ""
            Object.Tag             =   "24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":7E866
            Key             =   ""
            Object.Tag             =   "25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":839BA
            Key             =   ""
            Object.Tag             =   "26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":8898E
            Key             =   ""
            Object.Tag             =   "27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":8DAE2
            Key             =   ""
            Object.Tag             =   "28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":92AB6
            Key             =   ""
            Object.Tag             =   "29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":97A8A
            Key             =   ""
            Object.Tag             =   "30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":9CBDE
            Key             =   ""
            Object.Tag             =   "31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":A1D32
            Key             =   ""
            Object.Tag             =   "32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":A6E86
            Key             =   ""
            Object.Tag             =   "33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":ABFDA
            Key             =   ""
            Object.Tag             =   "34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":B112E
            Key             =   ""
            Object.Tag             =   "35"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":B6102
            Key             =   ""
            Object.Tag             =   "36"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":BB256
            Key             =   ""
            Object.Tag             =   "37"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":C022A
            Key             =   ""
            Object.Tag             =   "38"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":C537E
            Key             =   ""
            Object.Tag             =   "39"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":CA4D2
            Key             =   ""
            Object.Tag             =   "40"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":CF626
            Key             =   ""
            Object.Tag             =   "41"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":D477A
            Key             =   ""
            Object.Tag             =   "42"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":D98CE
            Key             =   ""
            Object.Tag             =   "43"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":DEA22
            Key             =   ""
            Object.Tag             =   "44"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":E3B76
            Key             =   ""
            Object.Tag             =   "45"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":E8CCA
            Key             =   ""
            Object.Tag             =   "46"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":EDE1E
            Key             =   ""
            Object.Tag             =   "47"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":F2F72
            Key             =   ""
            Object.Tag             =   "48"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":F7F46
            Key             =   ""
            Object.Tag             =   "49"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":FD098
            Key             =   ""
            Object.Tag             =   "50"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":1021EC
            Key             =   ""
            Object.Tag             =   "51"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "playcards.frx":107268
            Key             =   ""
            Object.Tag             =   "52"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Listcomp 
      Left            =   0
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList Listhuman 
      Left            =   0
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Label Hscore 
      Alignment       =   2  'Center
      Caption         =   "Your Score"
      Height          =   375
      Left            =   7680
      TabIndex        =   21
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label hfinalscore 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   20
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label cscore 
      Alignment       =   2  'Center
      Caption         =   "Computer Score"
      Height          =   375
      Left            =   7560
      TabIndex        =   19
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label cfinalscore 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      TabIndex        =   18
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label clam 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   375
      Left            =   6240
      TabIndex        =   17
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   5880
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label hlam 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   5880
      Width           =   615
   End
   Begin VB.Image Imageb 
      Height          =   1455
      Index           =   0
      Left            =   960
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image Imageb 
      Height          =   1455
      Index           =   1
      Left            =   2280
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image Imageb 
      Height          =   1455
      Index           =   2
      Left            =   3600
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image Imageb 
      Height          =   1455
      Index           =   3
      Left            =   4920
      Top             =   0
      Width           =   1095
   End
   Begin VB.Menu menufile 
      Caption         =   "&Game"
      Begin VB.Menu itmnew_game 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu itmundo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu itmexit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmplaycards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim endgame  As Boolean
Dim CARDSELECTED As Integer
Dim global_last_lam As String
Dim drawflag As Boolean
Dim comp_play_ok, humanselect As Boolean
Dim comp_basra(4), human_basra(4) As Boolean
Dim collection() As Integer
Sub colect_cards_for_this(index As Integer)
Dim j, i, ix As Integer
Dim DEFF, proccess_no As Integer
Dim ok_lem, ok_lemx As Boolean
Dim ccot, LEMARRAY(52), LEMCOUNT As Integer

ok_lem = False
ok_lemx = False

i = index
  proccess_no = Val(Imagec(i).Tag) Mod 10
  If proccess_no = 0 Then
    proccess_no = 10
  End If
ccot = count_cards_on_table
If ccot = 1 Then
  For ix = 0 To image1.Count - 1
    If image1(ix).Picture.Type <> 0 Then
      'draw the selected card on the table
      Exit For
    End If
  Next ix
  If image1(ix).Tag = 27 Then
    collection(i, ix) = 1
    ok_lem = True
    GoTo ok_exit
  End If
End If
If Val(Imagec(i).Tag) > 48 Or Val(Imagec(i).Tag) = 27 Then
  For ix = 0 To image1.Count - 1
    If image1(ix).Picture.Type <> 0 Then
      collection(i, ix) = 1
      ok_lem = True
    End If
  Next ix
  GoTo ok_exit
End If
If Val(Imagec(i).Tag) >= 41 And Val(Imagec(i).Tag) <= 44 Then
  For ix = 0 To image1.Count - 1
    If image1(ix).TabStop Then
      If Val(image1(ix).Tag) >= 41 And Val(image1(ix).Tag) <= 44 Then
        collection(i, ix) = 1
        ok_lem = True
      End If
    End If
  Next ix
  GoTo ok_exit
End If
If Val(Imagec(i).Tag) >= 45 And Val(Imagec(i).Tag) <= 48 Then
  For ix = 0 To image1.Count - 1
    If image1(ix).TabStop Then
      If Val(image1(ix).Tag) >= 45 And Val(image1(ix).Tag) <= 48 Then
        collection(i, ix) = 1
        ok_lem = True
      End If
    End If
  Next ix
GoTo ok_exit
End If
If Imagec(i).Tag <> "" And Val(Imagec(i).Tag) <= 40 Then
  For ix = 0 To image1.Count - 1
     If image1(ix).TabStop Then
       If Val(image1(ix).Tag) <= 40 And Val(image1(ix).Tag) Mod 10 = Val(Imagec(i).Tag) Mod 10 Then
         collection(i, ix) = 1
         ok_lem = True
       End If
     End If
    Next ix
''''''''''''''''''''''''''''''''
  For ix = 0 To image1.Count - 1
     If Val(image1(ix).Tag) <= 40 And image1(ix).TabStop And with_waht_card(image1(ix).Tag) < proccess_no Then
       DEFF = proccess_no - with_waht_card(image1(ix).Tag)
       LEMCOUNT = -1
       For j = ix + 1 To image1.Count - 1
          If Val(image1(j).Tag) <= 40 And with_waht_card(image1(j).Tag) = DEFF And image1(j).TabStop Then
            LEMCOUNT = LEMCOUNT + 1
            LEMARRAY(LEMCOUNT) = j
            ok_lemx = True
            DEFF = 0
            Exit For
           End If
       Next j
       If ok_lemx And DEFF = 0 Then
         collection(i, ix) = 1
         For j = 0 To LEMCOUNT
            collection(i, LEMARRAY(j)) = 1
            ok_lem = True
         Next j
       End If
       LEMCOUNT = -1
       ok_lemx = True
     End If
  Next ix
''''''''''''''''''''''''''''''''''''''''''
End If
  
If CARDSELECTED > 0 Then
  Select Case proccess_no
    Case 2 To 10:
      For ix = 0 To image1.Count - 1
         LEMCOUNT = -1
         ok_lemx = False
         If Val(image1(ix).Tag) <= 40 And image1(ix).TabStop And with_waht_card(image1(ix).Tag) < proccess_no Then
           DEFF = proccess_no - with_waht_card(image1(ix).Tag)
           For j = 0 To image1.Count - 1
              If j <> ix Then
                If Val(image1(j).Tag) <= 40 And with_waht_card(image1(j).Tag) = DEFF And image1(j).TabStop Then
                  LEMCOUNT = LEMCOUNT + 1
                  LEMARRAY(LEMCOUNT) = j
                  ok_lemx = True
                  DEFF = DEFF - with_waht_card(image1(j).Tag)
                  Exit For
                ElseIf Val(image1(j).Tag) <= 40 And with_waht_card(image1(j).Tag) < DEFF And image1(j).TabStop Then
                  DEFF = DEFF - with_waht_card(image1(j).Tag)
                  LEMCOUNT = LEMCOUNT + 1
                  LEMARRAY(LEMCOUNT) = j
                End If
              End If
            Next j
            If ok_lemx And DEFF = 0 Then
              collection(i, ix) = 1
              For j = 0 To LEMCOUNT
                 collection(i, LEMARRAY(j)) = 1
                 ok_lem = True
                 LEMCOUNT = -1
                 ok_lemx = False
                 'Exit For
              Next j
              'If ok_lem And DEFF = 0 Then Exit For
            Else
              LEMCOUNT = -1
              ok_lemx = False
            End If
          'End If
         Else
           If image1(ix).TabStop Then
             'image1(ix).PaintPicture image1(ix).Picture, 0, 0, , , , , , , vbDstInvert
             CARDSELECTED = CARDSELECTED - 1
           End If
         End If
      Next ix
  End Select
End If
  
ok_exit:
'If Listwhat.ListImages.Count > 0 Then
'  If ok_lem And ccot - (Listwhat.ListImages.Count - labellam.Caption) = 0 And Val(Imagec(i).Tag) <= 48 Or _
'     ok_lem And ccot - (Listwhat.ListImages.Count - labellam.Caption) = 0 And Val(Imagec(i).Tag) > 48 And ccot = 1 And (Listwhat.ListImages.Item(Listwhat.ListImages.Count).Picture = ImageListc.ListImages(49).Picture Or Listwhat.ListImages.Item(Listwhat.ListImages.Count).Picture = ImageListc.ListImages(50).Picture Or Listwhat.ListImages.Item(Listwhat.ListImages.Count).Picture = ImageListc.ListImages(51).Picture Or Listwhat.ListImages.Item(Listwhat.ListImages.Count).Picture = ImageListc.ListImages(52).Picture) Then
'     If (Val(Imagec(i).Tag) = 27 And sum_as_one_number) Or Val(Imagec(i).Tag) <> 27 Then
'       'MsgBox "basra comp"
'       comp_basra(i) = True
'       If Val(Imagec(i).Tag) > 48 Then
'         cfinalscore.Caption = Val(cfinalscore.Caption) + 20
'       Else
'         cfinalscore.Caption = Val(cfinalscore.Caption) + 10
'       End If
'     Else
'       comp_basra(i) = False
'     End If
'  End If
'Else
'  comp_basra(i) = False
'End If
End Sub

Sub lem_h(Listwhat As ImageList, labellam As Label)
Dim j, i, ix As Integer
Dim DEFF, proccess_no As Integer
Dim ok_lem, ok_lemx As Boolean
Dim ccot, LEMARRAY(52), LEMCOUNT As Integer
ok_lem = False
ok_lemx = False

For i = 0 To 3
  If Imageh(i).TabStop = True Then
    'draw the selected card on the table
    Exit For
  End If
Next i
If i > 3 Then
  Exit Sub
End If
  proccess_no = Val(Imageh(i).Tag) Mod 10
  If proccess_no = 0 Then
    proccess_no = 10
  End If
ccot = count_cards_on_table
If ccot = 1 Then
  For ix = 0 To image1.Count - 1
    If image1(ix).Picture.Type <> 0 Then
      'draw the selected card on the table
      Exit For
    End If
  Next ix
  If image1(ix).Tag = 27 Then
    Listwhat.ListImages.Add Listwhat.ListImages.Count + 1, , image1(ix).Picture
    Listhcards.AddItem "T-" + CStr(ix) + "-" + image1(ix).Tag, Listhcards.ListCount
    image1(ix).Picture = nullp.Picture
    image1(ix).Enabled = False
    image1(ix).Tag = ""
    'LABEL3(ix).Caption = image1(ix).Tag
    'image1(ix).TabStop = False
    'CARDSELECTED = CARDSELECTED - 1
    ok_lem = True
    GoTo ok_exit
  End If
End If
If Val(Imageh(i).Tag) > 48 Or Val(Imageh(i).Tag) = 27 Then
  For ix = 0 To image1.Count - 1
    If image1(ix).Picture.Type <> 0 Then
      Listwhat.ListImages.Add Listwhat.ListImages.Count + 1, , image1(ix).Picture
      Listhcards.AddItem "T-" + CStr(ix) + "-" + image1(ix).Tag, Listhcards.ListCount
      image1(ix).Picture = nullp.Picture
      image1(ix).Enabled = False
      image1(ix).Tag = ""
      'LABEL3(ix).Caption = image1(ix).Tag
      'image1(ix).TabStop = False
      'CARDSELECTED = CARDSELECTED - 1
      ok_lem = True
    End If
  Next ix
  GoTo ok_exit
End If
If Val(Imageh(i).Tag) >= 41 And Val(Imageh(i).Tag) <= 44 Then
  For ix = 0 To image1.Count - 1
    If image1(ix).TabStop Then
      If Val(image1(ix).Tag) >= 41 And Val(image1(ix).Tag) <= 44 Then
        Listwhat.ListImages.Add Listwhat.ListImages.Count + 1, , image1(ix).Picture
        Listhcards.AddItem "T-" + CStr(ix) + "-" + image1(ix).Tag, Listhcards.ListCount
        image1(ix).Picture = nullp.Picture
        image1(ix).Enabled = False
        image1(ix).Tag = ""
        'LABEL3(ix).Caption = image1(ix).Tag
        'image1(ix).TabStop = False
        CARDSELECTED = CARDSELECTED - 1
        ok_lem = True
      End If
    End If
  Next ix
GoTo ok_exit
End If
If Val(Imageh(i).Tag) >= 45 And Val(Imageh(i).Tag) <= 48 Then
  For ix = 0 To image1.Count - 1
    If image1(ix).TabStop Then
      If Val(image1(ix).Tag) >= 45 And Val(image1(ix).Tag) <= 48 Then
        Listwhat.ListImages.Add Listwhat.ListImages.Count + 1, , image1(ix).Picture
        Listhcards.AddItem "T-" + CStr(ix) + "-" + image1(ix).Tag, Listhcards.ListCount
        image1(ix).Picture = nullp.Picture
        image1(ix).Enabled = False
        image1(ix).Tag = ""
        'LABEL3(ix).Caption = image1(ix).Tag
        'image1(ix).TabStop = False
        CARDSELECTED = CARDSELECTED - 1
        ok_lem = True
      End If
    End If
  Next ix
GoTo ok_exit
End If
If Imageh(i).Tag <> "" And Val(Imageh(i).Tag) <= 40 Then
  For ix = 0 To image1.Count - 1
     If image1(ix).TabStop Then
       If Val(image1(ix).Tag) <= 40 And Val(image1(ix).Tag) Mod 10 = Val(Imageh(i).Tag) Mod 10 Then
         Listwhat.ListImages.Add Listwhat.ListImages.Count + 1, , image1(ix).Picture
         Listhcards.AddItem "T-" + CStr(ix) + "-" + image1(ix).Tag, Listhcards.ListCount
         image1(ix).Picture = nullp.Picture
         image1(ix).Enabled = False
         image1(ix).Tag = ""
         'LABEL3(ix).Caption = image1(ix).Tag
         'image1(ix).TabStop = False
         CARDSELECTED = CARDSELECTED - 1
         ok_lem = True
       End If
     End If
    Next ix
''''''''''''''''''''''''''''''''
  For ix = 0 To image1.Count - 1
     If Val(image1(ix).Tag) <= 40 And image1(ix).TabStop And with_waht_card(image1(ix).Tag) < proccess_no Then
       DEFF = proccess_no - with_waht_card(image1(ix).Tag)
       LEMCOUNT = -1
       For j = ix + 1 To image1.Count - 1
          If Val(image1(j).Tag) <= 40 And with_waht_card(image1(j).Tag) = DEFF And image1(j).TabStop Then
            LEMCOUNT = LEMCOUNT + 1
            LEMARRAY(LEMCOUNT) = j
            ok_lemx = True
            DEFF = 0
            Exit For
           End If
       Next j
       If ok_lemx And DEFF = 0 Then
         Listwhat.ListImages.Add Listwhat.ListImages.Count + 1, , image1(ix).Picture
         Listccards.AddItem "T-" + CStr(ix) + "-" + image1(ix).Tag, Listccards.ListCount
         image1(ix).Picture = nullp.Picture
         image1(ix).Enabled = False
         image1(ix).Tag = ""
         'LABEL3(ix).Caption = image1(ix).Tag
         'image1(ix).TabStop = False
         CARDSELECTED = CARDSELECTED - 1
         For j = 0 To LEMCOUNT
            Listwhat.ListImages.Add Listwhat.ListImages.Count + 1, , image1(LEMARRAY(j)).Picture
            Listccards.AddItem "T-" + CStr(LEMARRAY(j)) + "-" + image1(LEMARRAY(j)).Tag, Listccards.ListCount
            image1(LEMARRAY(j)).Picture = nullp.Picture
            image1(LEMARRAY(j)).Enabled = False
            image1(LEMARRAY(j)).Tag = ""
            'LABEL3(LEMARRAY(J)).Caption = image1(LEMARRAY(J)).Tag
            'image1(LEMARRAY(j)).TabStop = False
            CARDSELECTED = CARDSELECTED - 1
            ok_lem = True
         Next j
       End If
       LEMCOUNT = -1
       ok_lemx = True
     End If
  Next ix
''''''''''''''''''''''''''''''''''''''''''
End If
  
If CARDSELECTED > 0 Then
  Select Case proccess_no
    Case 2 To 10:
      For ix = 0 To image1.Count - 1
         LEMCOUNT = -1
         ok_lemx = False
'         If Val(image1(ix).Tag) <= 40 And image1(ix).TabStop And with_waht_card(image1(ix).Tag) < proccess_no Then
'           DEFF = proccess_no - with_waht_card(image1(ix).Tag)
'           LEMCOUNT = -1
'           For J = ix + 1 To image1.Count - 1
'              If Val(image1(J).Tag) <= 40 And with_waht_card(image1(J).Tag) = DEFF And image1(J).TabStop Then
'                LEMCOUNT = LEMCOUNT + 1
'                LEMARRAY(LEMCOUNT) = J
'                ok_lemx = True
'                Exit For
'              End If
'           Next J
'           If ok_lemx And DEFF = 0 Then
'             Listwhat.ListImages.Add Listwhat.ListImages.Count +1, , image1(ix).Picture
'             ListHcards.AddItem "T-" + CStr(ix) + "-" + image1(ix).Tag, ListHcards.ListCount
'             image1(ix).Picture = nullp.Picture
'             image1(ix).Enabled = False
'             image1(ix).Tag = ""
'             'LABEL3(ix).Caption = image1(ix).Tag
'             image1(ix).TabStop = False
'             CARDSELECTED = CARDSELECTED - 1
'             For J = 0 To LEMCOUNT
'                Listwhat.ListImages.Add Listwhat.ListImages.Count +1, , image1(LEMARRAY(J)).Picture
'                ListHcards.AddItem "T-" + CStr(LEMARRAY(J)) + "-" + image1(LEMARRAY(J)).Tag, ListHcards.ListCount
'                image1(LEMARRAY(J)).Picture = nullp.Picture
'                image1(LEMARRAY(J)).Enabled = False
'                image1(LEMARRAY(J)).Tag = ""
'                'LABEL3(LEMARRAY(J)).Caption = image1(LEMARRAY(J)).Tag
'                image1(LEMARRAY(J)).TabStop = False
'                CARDSELECTED = CARDSELECTED - 1
'                DEFF = 0
'                ok_lem = True
'             Next J
'             LEMCOUNT = -1
'             ok_lemx = True
'             DEFF = 0
'             'If ok_lem And DEFF = 0 Then Exit For
'          Else
         If Val(image1(ix).Tag) <= 40 And image1(ix).TabStop And with_waht_card(image1(ix).Tag) < proccess_no Then
             DEFF = proccess_no - with_waht_card(image1(ix).Tag)
             For j = 0 To image1.Count - 1
             If j <> ix Then
                If Val(image1(j).Tag) <= 40 And with_waht_card(image1(j).Tag) = DEFF And image1(j).TabStop Then
                  LEMCOUNT = LEMCOUNT + 1
                  LEMARRAY(LEMCOUNT) = j
                  ok_lemx = True
                  DEFF = DEFF - with_waht_card(image1(j).Tag)
                  Exit For
                ElseIf Val(image1(j).Tag) <= 40 And with_waht_card(image1(j).Tag) < DEFF And image1(j).TabStop Then
                  DEFF = DEFF - with_waht_card(image1(j).Tag)
                  LEMCOUNT = LEMCOUNT + 1
                  LEMARRAY(LEMCOUNT) = j
                End If
             End If
             Next j
             If ok_lemx And DEFF = 0 Then
               Listwhat.ListImages.Add Listwhat.ListImages.Count + 1, , image1(ix).Picture
               Listhcards.AddItem "T-" + CStr(ix) + "-" + image1(ix).Tag, Listhcards.ListCount
               image1(ix).Picture = nullp.Picture
               image1(ix).Enabled = False
               image1(ix).Tag = ""
               'LABEL3(ix).Caption = image1(ix).Tag
               'image1(ix).TabStop = False
               CARDSELECTED = CARDSELECTED - 1
               For j = 0 To LEMCOUNT
                  Listwhat.ListImages.Add Listwhat.ListImages.Count + 1, , image1(LEMARRAY(j)).Picture
                  Listhcards.AddItem "T-" + CStr(LEMARRAY(j)) + "-" + image1(LEMARRAY(j)).Tag, Listhcards.ListCount
                  image1(LEMARRAY(j)).Picture = nullp.Picture
                  image1(LEMARRAY(j)).Enabled = False
                  image1(LEMARRAY(j)).Tag = ""
                  'LABEL3(LEMARRAY(J)).Caption = image1(LEMARRAY(J)).Tag
                  'image1(LEMARRAY(j)).TabStop = False
                  CARDSELECTED = CARDSELECTED - 1
                  ok_lem = True
                  LEMCOUNT = -1
                  ok_lemx = False
                  'Exit For
               Next j
               'If ok_lem And DEFF = 0 Then Exit For
             Else
               LEMCOUNT = -1
               ok_lemx = False
             End If
           'End If
         Else
           If image1(ix).TabStop Then
             'image1(ix).PaintPicture image1(ix).Picture, 0, 0, , , , , , , vbDstInvert
             'image1(ix).TabStop = False
             CARDSELECTED = CARDSELECTED - 1
           End If
         End If
      Next ix
  End Select
End If
  
ok_exit:
If Listwhat.ListImages.Count > 0 Then
  If ok_lem And ccot - (Listwhat.ListImages.Count - labellam.Caption) = 0 And Val(Imageh(i).Tag) <= 48 Or _
     ok_lem And ccot - (Listwhat.ListImages.Count - labellam.Caption) = 0 And Val(Imageh(i).Tag) > 48 And ccot = 1 And (Listwhat.ListImages.Item(Listwhat.ListImages.Count).Picture = ImageListc.ListImages(49).Picture Or Listwhat.ListImages.Item(Listwhat.ListImages.Count).Picture = ImageListc.ListImages(50).Picture Or Listwhat.ListImages.Item(Listwhat.ListImages.Count).Picture = ImageListc.ListImages(51).Picture Or Listwhat.ListImages.Item(Listwhat.ListImages.Count).Picture = ImageListc.ListImages(52).Picture) Then
     If (Val(Imageh(i).Tag) = 27 And sum_as_one_number) Or Val(Imageh(i).Tag) <> 27 Then
       'MsgBox "basra human"
       human_basra(i) = True
       If Val(Imageh(i).Tag) > 48 Then
         hfinalscore.Caption = Val(hfinalscore.Caption) + 20
       Else
         hfinalscore.Caption = Val(hfinalscore.Caption) + 10
       End If
     Else
       human_basra(i) = False
     End If
  End If
Else
  human_basra(i) = False
End If
  If ok_lem Then
    Listwhat.ListImages.Add Listwhat.ListImages.Count + 1, , Imageh(i).Picture
    Listhcards.AddItem "H-" + CStr(i) + "-" + Imageh(i).Tag, Listhcards.ListCount
    Imageh(i).Picture = nullp.Picture
    Imageh(i).Enabled = False
    Imageh(i).TabStop = False
    Imageh(i).Tag = ""
    global_last_lam = "HUMAN"
  Else
    For ix = 0 To image1.Count - 1
       If image1(ix).Tag = "" Then
         Listhcards.AddItem CStr(i) + "-" + CStr(ix) + "-" + Imageh(i).Tag, Listhcards.ListCount
         image1(ix).Picture = Imageh(i).Picture
         image1(ix).Tag = Imageh(i).Tag
         'LABEL3(ix).Caption = image1(ix).Tag
         image1(ix).Enabled = True
         'image1(ix).TabStop = False
         Imageh(i).Enabled = False
         Imageh(i).TabStop = False
         Imageh(i).Tag = ""
         Imageh(i).Picture = nullp.Picture
         ok_lem = True
         Exit For
       End If
    Next ix
    If Not ok_lem Then
      Listhcards.AddItem CStr(i) + "-" + CStr(ix) + "-" + Imageh(i).Tag, Listhcards.ListCount
      Load image1(ix)
      image1(ix).Visible = True
      image1(ix).Left = image1(ix - 1).Left + image1(ix - 1).Width + 120
      If ix Mod 2 = 0 Then
        image1(ix).Top = image1(ix - 2).Top '+ image1(ix - 1).Height
      Else
       image1(ix).Left = image1(ix - 1).Left + image1(ix - 1).Width + 120
       image1(ix).Top = image1(ix - 1).Top + image1(ix - 1).Height / 2
      End If
      image1(ix).Picture = Imageh(i).Picture
      image1(ix).Tag = Imageh(i).Tag
      'LABEL3(ix).Caption = image1(ix).Tag
      image1(ix).Enabled = True
      'image1(ix).TabStop = False
      Imageh(i).Enabled = False
      Imageh(i).TabStop = False
      Imageh(i).Tag = ""
      Imageh(i).Picture = nullp.Picture
    End If
  End If
labellam.Caption = Listwhat.ListImages.Count
Label1.Caption = CARDSELECTED
End Sub


Function lem_c(Listwhat As ImageList, labellam As Label)
Dim j, i, ix As Integer
Dim DEFF, proccess_no As Integer
Dim ok_lem, ok_lemx As Boolean
Dim ccot, LEMARRAY(52), LEMCOUNT As Integer
ok_lem = False
ok_lemx = False

For i = 0 To 3
  If Imagec(i).TabStop = True Then
    'draw the selected card on the table
    Exit For
  End If
Next i
If i > 3 Then
  Exit Function
End If
  proccess_no = Val(Imagec(i).Tag) Mod 10
  If proccess_no = 0 Then
    proccess_no = 10
  End If
ccot = count_cards_on_table
If ccot = 1 Then
  For ix = 0 To image1.Count - 1
    If image1(ix).Picture.Type <> 0 Then
      'draw the selected card on the table
      Exit For
    End If
  Next ix
  If image1(ix).Tag = 27 Then
    Listwhat.ListImages.Add Listwhat.ListImages.Count + 1, , image1(ix).Picture
    Listhcards.AddItem "T-" + CStr(ix) + "-" + image1(ix).Tag, Listhcards.ListCount
    image1(ix).Picture = nullp.Picture
    image1(ix).Enabled = False
    image1(ix).Tag = ""
    'LABEL3(ix).Caption = image1(ix).Tag
    image1(ix).TabStop = False
    'CARDSELECTED = CARDSELECTED - 1
    ok_lem = True
    GoTo ok_exit
  End If
End If
If Val(Imagec(i).Tag) > 48 Or Val(Imagec(i).Tag) = 27 Then
  For ix = 0 To image1.Count - 1
    If image1(ix).Picture.Type <> 0 Then
      Listwhat.ListImages.Add Listwhat.ListImages.Count + 1, , image1(ix).Picture
      Listccards.AddItem "T-" + CStr(ix) + "-" + image1(ix).Tag, Listccards.ListCount
      image1(ix).Picture = nullp.Picture
      image1(ix).Enabled = False
      image1(ix).Tag = ""
      'LABEL3(ix).Caption = image1(ix).Tag
      image1(ix).TabStop = False
'      CARDSELECTED = CARDSELECTED - 1
      ok_lem = True
    End If
  Next ix
GoTo ok_exit
End If
If Val(Imagec(i).Tag) >= 41 And Val(Imagec(i).Tag) <= 44 Then
  For ix = 0 To image1.Count - 1
    If image1(ix).TabStop Then
      If Val(image1(ix).Tag) >= 41 And Val(image1(ix).Tag) <= 44 Then
        Listwhat.ListImages.Add Listwhat.ListImages.Count + 1, , image1(ix).Picture
        Listccards.AddItem "T-" + CStr(ix) + "-" + image1(ix).Tag, Listccards.ListCount
        image1(ix).Picture = nullp.Picture
        image1(ix).Enabled = False
        image1(ix).Tag = ""
        'LABEL3(ix).Caption = image1(ix).Tag
        image1(ix).TabStop = False
        CARDSELECTED = CARDSELECTED - 1
        ok_lem = True
      End If
    End If
  Next ix
GoTo ok_exit
End If
If Val(Imagec(i).Tag) >= 45 And Val(Imagec(i).Tag) <= 48 Then
  For ix = 0 To image1.Count - 1
    If image1(ix).TabStop Then
      If Val(image1(ix).Tag) >= 45 And Val(image1(ix).Tag) <= 48 Then
        Listwhat.ListImages.Add Listwhat.ListImages.Count + 1, , image1(ix).Picture
        Listccards.AddItem "T-" + CStr(ix) + "-" + image1(ix).Tag, Listccards.ListCount
        image1(ix).Picture = nullp.Picture
        image1(ix).Enabled = False
        image1(ix).Tag = ""
        'LABEL3(ix).Caption = image1(ix).Tag
        image1(ix).TabStop = False
        CARDSELECTED = CARDSELECTED - 1
        ok_lem = True
      End If
    End If
  Next ix
GoTo ok_exit
End If
If Imagec(i).Tag <> "" And Val(Imagec(i).Tag) <= 40 Then
  For ix = 0 To image1.Count - 1
     If image1(ix).TabStop Then
       If Val(image1(ix).Tag) <= 40 And Val(image1(ix).Tag) Mod 10 = Val(Imagec(i).Tag) Mod 10 Then
         Listwhat.ListImages.Add Listwhat.ListImages.Count + 1, , image1(ix).Picture
         Listccards.AddItem "T-" + CStr(ix) + "-" + image1(ix).Tag, Listccards.ListCount
         image1(ix).Picture = nullp.Picture
         image1(ix).Enabled = False
         image1(ix).Tag = ""
         'LABEL3(ix).Caption = image1(ix).Tag
         image1(ix).TabStop = False
         CARDSELECTED = CARDSELECTED - 1
         ok_lem = True
       End If
     End If
    Next ix
''''''''''''''''''''''''''''''''
  For ix = 0 To image1.Count - 1
     If Val(image1(ix).Tag) <= 40 And image1(ix).TabStop And with_waht_card(image1(ix).Tag) < proccess_no Then
       DEFF = proccess_no - with_waht_card(image1(ix).Tag)
       LEMCOUNT = -1
       For j = ix + 1 To image1.Count - 1
          If Val(image1(j).Tag) <= 40 And with_waht_card(image1(j).Tag) = DEFF And image1(j).TabStop Then
            LEMCOUNT = LEMCOUNT + 1
            LEMARRAY(LEMCOUNT) = j
            ok_lemx = True
            DEFF = 0
            Exit For
           End If
       Next j
       If ok_lemx And DEFF = 0 Then
         Listwhat.ListImages.Add Listwhat.ListImages.Count + 1, , image1(ix).Picture
         Listccards.AddItem "T-" + CStr(ix) + "-" + image1(ix).Tag, Listccards.ListCount
         image1(ix).Picture = nullp.Picture
         image1(ix).Enabled = False
         image1(ix).Tag = ""
         'LABEL3(ix).Caption = image1(ix).Tag
         image1(ix).TabStop = False
         CARDSELECTED = CARDSELECTED - 1
         For j = 0 To LEMCOUNT
            Listwhat.ListImages.Add Listwhat.ListImages.Count + 1, , image1(LEMARRAY(j)).Picture
            Listccards.AddItem "T-" + CStr(LEMARRAY(j)) + "-" + image1(LEMARRAY(j)).Tag, Listccards.ListCount
            image1(LEMARRAY(j)).Picture = nullp.Picture
            image1(LEMARRAY(j)).Enabled = False
            image1(LEMARRAY(j)).Tag = ""
            'LABEL3(LEMARRAY(J)).Caption = image1(LEMARRAY(J)).Tag
            image1(LEMARRAY(j)).TabStop = False
            CARDSELECTED = CARDSELECTED - 1
            ok_lem = True
         Next j
       End If
       LEMCOUNT = -1
       ok_lemx = True
     End If
  Next ix
''''''''''''''''''''''''''''''''''''''''''
End If
  
If CARDSELECTED > 0 Then
  Select Case proccess_no
    Case 2 To 10:
      For ix = 0 To image1.Count - 1
         LEMCOUNT = -1
         ok_lemx = False
         If Val(image1(ix).Tag) <= 40 And image1(ix).TabStop And with_waht_card(image1(ix).Tag) < proccess_no Then
             DEFF = proccess_no - with_waht_card(image1(ix).Tag)
             For j = 0 To image1.Count - 1
             If j <> ix Then
               If Val(image1(j).Tag) <= 40 And with_waht_card(image1(j).Tag) = DEFF And image1(j).TabStop Then
                  LEMCOUNT = LEMCOUNT + 1
                  LEMARRAY(LEMCOUNT) = j
                  ok_lemx = True
                  DEFF = DEFF - with_waht_card(image1(j).Tag)
                  Exit For
                ElseIf Val(image1(j).Tag) <= 40 And with_waht_card(image1(j).Tag) < DEFF And image1(j).TabStop Then
                  DEFF = DEFF - with_waht_card(image1(j).Tag)
                  LEMCOUNT = LEMCOUNT + 1
                  LEMARRAY(LEMCOUNT) = j
                End If
             End If
             Next j
             If ok_lemx And DEFF = 0 Then
               Listwhat.ListImages.Add Listwhat.ListImages.Count + 1, , image1(ix).Picture
               Listccards.AddItem "T-" + CStr(ix) + "-" + image1(ix).Tag, Listccards.ListCount
               image1(ix).Picture = nullp.Picture
               image1(ix).Enabled = False
               image1(ix).Tag = ""
               'LABEL3(ix).Caption = image1(ix).Tag
               image1(ix).TabStop = False
               CARDSELECTED = CARDSELECTED - 1
               For j = 0 To LEMCOUNT
                  Listwhat.ListImages.Add Listwhat.ListImages.Count + 1, , image1(LEMARRAY(j)).Picture
                  Listccards.AddItem "T-" + CStr(LEMARRAY(j)) + "-" + image1(LEMARRAY(j)).Tag, Listccards.ListCount
                  image1(LEMARRAY(j)).Picture = nullp.Picture
                  image1(LEMARRAY(j)).Enabled = False
                  image1(LEMARRAY(j)).Tag = ""
                  'LABEL3(LEMARRAY(J)).Caption = image1(LEMARRAY(J)).Tag
                  image1(LEMARRAY(j)).TabStop = False
                  CARDSELECTED = CARDSELECTED - 1
                  ok_lem = True
                  LEMCOUNT = -1
                  ok_lemx = False
                  'Exit For
               Next j
               'If ok_lem And DEFF = 0 Then Exit For
             Else
               LEMCOUNT = -1
               ok_lemx = False
             End If
           'End If
         Else
           If image1(ix).TabStop Then
             'image1(ix).PaintPicture image1(ix).Picture, 0, 0, , , , , , , vbDstInvert
             image1(ix).TabStop = False
             CARDSELECTED = CARDSELECTED - 1
           End If
         End If
      Next ix
  End Select
End If
  
ok_exit:
If Listwhat.ListImages.Count > 0 Then
  If ok_lem And ccot - (Listwhat.ListImages.Count - labellam.Caption) = 0 And Val(Imagec(i).Tag) <= 48 Or _
     ok_lem And ccot - (Listwhat.ListImages.Count - labellam.Caption) = 0 And Val(Imagec(i).Tag) > 48 And ccot = 1 And (Listwhat.ListImages.Item(Listwhat.ListImages.Count).Picture = ImageListc.ListImages(49).Picture Or Listwhat.ListImages.Item(Listwhat.ListImages.Count).Picture = ImageListc.ListImages(50).Picture Or Listwhat.ListImages.Item(Listwhat.ListImages.Count).Picture = ImageListc.ListImages(51).Picture Or Listwhat.ListImages.Item(Listwhat.ListImages.Count).Picture = ImageListc.ListImages(52).Picture) Then
     If (Val(Imagec(i).Tag) = 27 And sum_as_one_number) Or Val(Imagec(i).Tag) <> 27 Then
       'MsgBox "basra comp"
       comp_basra(i) = True
       If Val(Imagec(i).Tag) > 48 Then
         cfinalscore.Caption = Val(cfinalscore.Caption) + 20
       Else
         cfinalscore.Caption = Val(cfinalscore.Caption) + 10
       End If
     Else
       comp_basra(i) = False
     End If
  End If
Else
  comp_basra(i) = False
End If
  If ok_lem Then
    Listwhat.ListImages.Add Listwhat.ListImages.Count + 1, , Imagec(i).Picture
    Listccards.AddItem "C-" + CStr(i) + "-" + Imagec(i).Tag, Listccards.ListCount
    Imagec(i).Picture = nullp.Picture
    Imagec(i).Enabled = False
    Imagec(i).TabStop = False
    Imagec(i).Tag = ""
    global_last_lam = "COMP"
  Else
    For ix = 0 To image1.Count - 1
       If image1(ix).Tag = "" Then
         Listccards.AddItem CStr(i) + "-" + CStr(ix) + "-" + Imagec(i).Tag, Listccards.ListCount
         image1(ix).Picture = Imagec(i).Picture
         image1(ix).Tag = Imagec(i).Tag
         'LABEL3(ix).Caption = image1(ix).Tag
         image1(ix).Enabled = True
         image1(ix).TabStop = False
         Imagec(i).Enabled = False
         Imagec(i).TabStop = False
         Imagec(i).Tag = ""
         Imagec(i).Picture = nullp.Picture
         ok_lem = True
         Exit For
       End If
    Next ix
    If Not ok_lem Then
      Load image1(ix)
      image1(ix).Visible = True
      image1(ix).Left = image1(ix - 1).Left + image1(ix - 1).Width + 120
      If ix Mod 2 = 0 Then
        image1(ix).Top = image1(ix - 2).Top '+ image1(ix - 1).Height
      Else
       image1(ix).Left = image1(ix - 1).Left + image1(ix - 1).Width + 120
       image1(ix).Top = image1(ix - 1).Top + image1(ix - 1).Height / 2
      End If
      Listccards.AddItem CStr(i) + "-" + CStr(ix) + "-" + Imagec(i).Tag, Listccards.ListCount
      image1(ix).Picture = Imagec(i).Picture
      image1(ix).Tag = Imagec(i).Tag
      'LABEL3(ix).Caption = image1(ix).Tag
      image1(ix).Enabled = True
      image1(ix).TabStop = False
      Imagec(i).Enabled = False
      Imagec(i).TabStop = False
      Imagec(i).Tag = ""
      Imagec(i).Picture = nullp.Picture
    End If
  End If
labellam.Caption = Listwhat.ListImages.Count
Label1.Caption = CARDSELECTED
For j = 0 To image1.Count - 1
   If image1(j).Picture.Type <> 0 Then
     image1(j).TabStop = False
    End If
Next j
lem_c = True
End Function


Sub play_comp()
  Dim i As Integer
  For i = 0 To image1.Count - 1
     image1(i).TabStop = False
     If image1(i).Picture.Type <> 0 Then
       image1(i).TabStop = True
       CARDSELECTED = CARDSELECTED + 1
      End If
  Next i
  'For i = 0 To Imagec.Count - 1
  '   If Imagec(i).Picture.Type <> 0 Then
       i = thinking
 '      If play_this(i) Then
         Imagec(i).TabStop = True
         Imagec(i).Visible = True
         Timer1.Enabled = True
         Do While Not comp_play_ok
           DoEvents
           If comp_play_ok Then
             comp_play_ok = False
             Imagec(i).Visible = False
             Imageb(i).Visible = False
             Exit Do
           End If
         Loop
 '      End If
  '   End If
  'Next i
End Sub

Function play_this(index As Integer) As Boolean
  If Imagec(index).Tag > 48 Or Imagec(index).Tag = "27" Then
    Dim vr_count_cards_on_table, i, count_card_for_comp As Integer
    count_card_for_comp = 0
    vr_count_cards_on_table = count_cards_on_table
      For i = 0 To Imagec.Count - 1
         If Imagec(i).Picture.Type <> 0 And Val(Imagec(i).Tag) <= 48 And Imagec(i).Tag <> "27" Then
           count_card_for_comp = count_card_for_comp + 1
         End If
      Next i
    If count_card_for_comp <= 0 Then
      play_this = True
    Else
      play_this = False
    End If
  Else
    play_this = True
  End If
End Function

Function count_cards_on_table() As Integer
  Dim i As Integer
  count_cards_on_table = 0
  For i = 0 To image1.Count - 1
     If image1(i).Picture.Type <> 0 Then
       count_cards_on_table = count_cards_on_table + 1
     End If
  Next i
End Function

Function sum_as_one_number() As Boolean
  sum_as_one_number = False
End Function

' To decide which card can be played
' The return is the card number
Function thinking() As Integer
  Dim i, j As Integer
  ReDim collection(3, 51)
  Dim MaxCards, CardsCount As Integer
    
  For i = 0 To Imagec.Count - 1
     If Imagec(i).Picture.Type <> 0 Then
       colect_cards_for_this (i)
     End If
  Next i
  
MaxCards = 0
thinking = -1

For i = 0 To 3
   CardsCount = 0
   For j = 0 To 51
      If collection(i, j) = 1 Then
        CardsCount = CardsCount + 1
      End If
   Next j
   If MaxCards < CardsCount Then
      MaxCards = CardsCount
      thinking = i
   End If
Next i
  
If thinking = -1 Then
  For j = 0 To Imagec.Count - 1
     If Imagec(j).Picture.Type <> 0 Then
       If play_this(j) Then
         thinking = j
         Exit For
       End If
     End If
  Next j
End If
  
  For i = 0 To image1.Count - 1
     image1(i).TabStop = False
     If image1(i).Picture.Type <> 0 Then
       image1(i).TabStop = True
       CARDSELECTED = CARDSELECTED + 1
      End If
  Next i
End Function

Sub undo_comp()
   Dim i, start_on, tmplen1 As Integer
   Dim tmps1 As String
   start_on = Listccards.ListCount - 1
   If Mid(Listccards.List(start_on), 1, 1) = "C" And start_on <> -1 Then
   'Listccards.AddItem "C-" + CStr(i) + "-" + Imagec(i).Tag, Listccards.ListCount
   For i = 1 To ImageListc.ListImages.Count + 1 - 1
      tmplen1 = Len(Listccards.List(start_on))
      tmps1 = InStr(1, Mid(Listccards.List(start_on), 3, tmplen1), "-") + 1
      If ImageListc.ListImages(i).index = Mid(Mid(Listccards.List(start_on), 3, tmplen1), tmps1, tmplen1) Then
        If comp_basra(Val(Mid(Listccards.List(start_on), 3, InStr(1, Mid(Listccards.List(start_on), 3, tmplen1), "-")))) = True Then
          If ImageListc.ListImages(i).index > 48 Then
            cfinalscore.Caption = Val(cfinalscore.Caption) - 20
          Else
            cfinalscore.Caption = Val(cfinalscore.Caption) - 10
          End If
        End If
        Imagec(Val(Mid(Listccards.List(start_on), 3, InStr(1, Mid(Listccards.List(start_on), 3, tmplen1), "-")))).Picture = ImageListc.ListImages(i).Picture
        Imageb(Val(Mid(Listccards.List(start_on), 3, InStr(1, Mid(Listccards.List(start_on), 3, tmplen1), "-")))).Visible = True
        Imagec(Val(Mid(Listccards.List(start_on), 3, InStr(1, Mid(Listccards.List(start_on), 3, tmplen1), "-")))).Tag = ImageListc.ListImages(i).index
        Imagec(Val(Mid(Listccards.List(start_on), 3, InStr(1, Mid(Listccards.List(start_on), 3, tmplen1), "-")))).Enabled = True
        Imagec(Val(Mid(Listccards.List(start_on), 3, InStr(1, Mid(Listccards.List(start_on), 3, tmplen1), "-")))).TabStop = False
        clam.Caption = Val(clam.Caption) - 1
        Listcomp.ListImages.Remove (Listcomp.ListImages.Count)
        Exit For
      End If
   Next i
   Listccards.RemoveItem (start_on)
   start_on = Listccards.ListCount - 1
   While Mid(Listccards.List(start_on), 1, 1) = "T" And start_on <> -1
     For i = 1 To ImageListc.ListImages.Count + 1 - 1
        tmplen1 = Len(Listccards.List(start_on))
        tmps1 = InStr(1, Mid(Listccards.List(start_on), 3, tmplen1), "-") + 1
        If ImageListc.ListImages(i).index = Mid(Mid(Listccards.List(start_on), 3, tmplen1), tmps1, tmplen1) Then
          image1(Val(Mid(Listccards.List(start_on), 3, InStr(1, Mid(Listccards.List(start_on), 3, tmplen1), "-")))).Picture = ImageListc.ListImages(i).Picture
          image1(Val(Mid(Listccards.List(start_on), 3, InStr(1, Mid(Listccards.List(start_on), 3, tmplen1), "-")))).Tag = ImageListc.ListImages(i).index
          image1(Val(Mid(Listccards.List(start_on), 3, InStr(1, Mid(Listccards.List(start_on), 3, tmplen1), "-")))).Enabled = True
          image1(Val(Mid(Listccards.List(start_on), 3, InStr(1, Mid(Listccards.List(start_on), 3, tmplen1), "-")))).TabStop = False
          Exit For
        End If
     Next i
     Listccards.RemoveItem (start_on)
     clam.Caption = Val(clam.Caption) - 1
     Listcomp.ListImages.Remove (Listcomp.ListImages.Count)
     start_on = Listccards.ListCount - 1
   Wend
   ElseIf Mid(Listccards.List(start_on), 1, 1) <> "C" And start_on <> -1 Then
      tmplen1 = Len(Listccards.List(start_on))
      tmps1 = InStr(1, Mid(Listccards.List(start_on), 3, tmplen1), "-") + 1
      Imagec(Mid(Listccards.List(start_on), 1, 1)).Picture = image1(Val(Mid(Listccards.List(start_on), 3, InStr(1, Mid(Listccards.List(start_on), 3, tmplen1), "-")))).Picture
      Imagec(Mid(Listccards.List(start_on), 1, 1)).Enabled = True
      Imagec(Mid(Listccards.List(start_on), 1, 1)).TabStop = False
      Imagec(Mid(Listccards.List(start_on), 1, 1)).Tag = image1(Val(Mid(Listccards.List(start_on), 3, InStr(1, Mid(Listccards.List(start_on), 3, tmplen1), "-")))).Tag
      image1(Val(Mid(Listccards.List(start_on), 3, InStr(1, Mid(Listccards.List(start_on), 3, tmplen1), "-")))).Enabled = False
      image1(Val(Mid(Listccards.List(start_on), 3, InStr(1, Mid(Listccards.List(start_on), 3, tmplen1), "-")))).TabStop = False
      image1(Val(Mid(Listccards.List(start_on), 3, InStr(1, Mid(Listccards.List(start_on), 3, tmplen1), "-")))).Tag = ""
      image1(Val(Mid(Listccards.List(start_on), 3, InStr(1, Mid(Listccards.List(start_on), 3, tmplen1), "-")))).Picture = nullp.Picture
      Imageb(Mid(Listccards.List(start_on), 1, 1)).Visible = True
      Listccards.RemoveItem (start_on)
   End If
End Sub

Sub undo_human()
   Dim i, start_on, tmplen1 As Integer
   Dim tmps1 As String
   start_on = Listhcards.ListCount - 1
   If Mid(Listhcards.List(start_on), 1, 1) = "H" And start_on <> -1 Then
   'ListHcards.AddItem "H-" + CStr(i) + "-" + ImageH(i).Tag, ListHcards.ListCount
   For i = 1 To ImageListc.ListImages.Count + 1 - 1
      tmplen1 = Len(Listhcards.List(start_on))
      tmps1 = InStr(1, Mid(Listhcards.List(start_on), 3, tmplen1), "-") + 1
      If ImageListc.ListImages(i).index = Mid(Mid(Listhcards.List(start_on), 3, tmplen1), tmps1, tmplen1) Then
        If human_basra(Val(Mid(Listhcards.List(start_on), 3, InStr(1, Mid(Listhcards.List(start_on), 3, tmplen1), "-")))) = True Then
          If ImageListc.ListImages(i).index > 48 Then
            hfinalscore.Caption = Val(hfinalscore.Caption) - 20
          Else
            hfinalscore.Caption = Val(hfinalscore.Caption) - 10
          End If
        End If
        Imageh(Val(Mid(Listhcards.List(start_on), 3, InStr(1, Mid(Listhcards.List(start_on), 3, tmplen1), "-")))).Picture = ImageListc.ListImages(i).Picture
        Imageh(Val(Mid(Listhcards.List(start_on), 3, InStr(1, Mid(Listhcards.List(start_on), 3, tmplen1), "-")))).Tag = ImageListc.ListImages(i).index
        Imageh(Val(Mid(Listhcards.List(start_on), 3, InStr(1, Mid(Listhcards.List(start_on), 3, tmplen1), "-")))).Enabled = True
        Imageh(Val(Mid(Listhcards.List(start_on), 3, InStr(1, Mid(Listhcards.List(start_on), 3, tmplen1), "-")))).TabStop = False
        hlam.Caption = Val(hlam.Caption) - 1
        Listhuman.ListImages.Remove (Listhuman.ListImages.Count)
        Exit For
      End If
   Next i
   Listhcards.RemoveItem (start_on)
   start_on = Listhcards.ListCount - 1
   While Mid(Listhcards.List(start_on), 1, 1) = "T" And start_on <> -1
     For i = 1 To ImageListc.ListImages.Count + 1 - 1
        tmplen1 = Len(Listhcards.List(start_on))
        tmps1 = InStr(1, Mid(Listhcards.List(start_on), 3, tmplen1), "-") + 1
        If ImageListc.ListImages(i).index = Mid(Mid(Listhcards.List(start_on), 3, tmplen1), tmps1, tmplen1) Then
          image1(Val(Mid(Listhcards.List(start_on), 3, InStr(1, Mid(Listhcards.List(start_on), 3, tmplen1), "-")))).Picture = ImageListc.ListImages(i).Picture
          image1(Val(Mid(Listhcards.List(start_on), 3, InStr(1, Mid(Listhcards.List(start_on), 3, tmplen1), "-")))).Tag = ImageListc.ListImages(i).index
          image1(Val(Mid(Listhcards.List(start_on), 3, InStr(1, Mid(Listhcards.List(start_on), 3, tmplen1), "-")))).Enabled = True
          image1(Val(Mid(Listhcards.List(start_on), 3, InStr(1, Mid(Listhcards.List(start_on), 3, tmplen1), "-")))).TabStop = False
          Exit For
        End If
     Next i
     Listhcards.RemoveItem (start_on)
     hlam.Caption = Val(hlam.Caption) - 1
     Listhuman.ListImages.Remove (Listhuman.ListImages.Count)
     start_on = Listhcards.ListCount - 1
   Wend
   ElseIf Mid(Listhcards.List(start_on), 1, 1) <> "H" And start_on <> -1 Then
      tmplen1 = Len(Listhcards.List(start_on))
      tmps1 = InStr(1, Mid(Listhcards.List(start_on), 3, tmplen1), "-") + 1
      Imageh(Mid(Listhcards.List(start_on), 1, 1)).Picture = image1(Val(Mid(Listhcards.List(start_on), 3, InStr(1, Mid(Listhcards.List(start_on), 3, tmplen1), "-")))).Picture
      Imageh(Mid(Listhcards.List(start_on), 1, 1)).Tag = image1(Val(Mid(Listhcards.List(start_on), 3, InStr(1, Mid(Listhcards.List(start_on), 3, tmplen1), "-")))).Tag
      image1(Val(Mid(Listhcards.List(start_on), 3, InStr(1, Mid(Listhcards.List(start_on), 3, tmplen1), "-")))).Enabled = False
      image1(Val(Mid(Listhcards.List(start_on), 3, InStr(1, Mid(Listhcards.List(start_on), 3, tmplen1), "-")))).TabStop = False
      Imageh(Mid(Listhcards.List(start_on), 1, 1)).Enabled = True
      Imageh(Mid(Listhcards.List(start_on), 1, 1)).TabStop = False
      image1(Val(Mid(Listhcards.List(start_on), 3, InStr(1, Mid(Listhcards.List(start_on), 3, tmplen1), "-")))).Tag = ""
      image1(Val(Mid(Listhcards.List(start_on), 3, InStr(1, Mid(Listhcards.List(start_on), 3, tmplen1), "-")))).Picture = nullp.Picture
      Listhcards.RemoveItem (start_on)
   End If
End Sub

Function with_waht_card(proccess_with As String) As Integer
  with_waht_card = Val(proccess_with) Mod 10
  If with_waht_card = 0 Then
    with_waht_card = 10
  End If
End Function

Private Sub new_draw()
Dim X, i, limit As Integer
endgame = False
clam.Caption = 0
hlam.Caption = 0

Listcomp.ListImages.Clear
Listccards.Clear
Listhuman.ListImages.Clear
Listhcards.Clear
CARDSELECTED = 0
For i = 1 To 52
  ImageListc.ListImages(i).Tag = 0
Next i
Randomize

ImageListc.Tag = 0

For i = 0 To 3
  Imageh(i).TabStop = False
  Imageh(i).Enabled = True
Next i
limit = 0
While limit < 4 And Val(ImageListc.Tag) <> 52
  X = Int((52 * Rnd) + 1)
  If ImageListc.ListImages(X).Tag <> 1 Then
    Imageh(limit).Picture = ImageListc.ListImages(X).Picture
    Imageh(limit).Tag = X
    ImageListc.ListImages(X).Tag = 1
    limit = limit + 1
    ImageListc.Tag = Val(ImageListc.Tag) + 1
  End If
Wend
For i = 0 To image1.Count - 1
  image1(i).TabStop = False
  image1(i).Enabled = True
  image1(i).Picture = nullp.Picture
Next i
limit = 0
While limit < 4 And Val(ImageListc.Tag) <> 52
  X = Int((48 * Rnd) + 1)
  If ImageListc.ListImages(X).Tag <> 1 And X <> 27 Then
    image1(limit).Picture = ImageListc.ListImages(X).Picture
    image1(limit).Tag = X
    'LABEL3(limit).Caption = image1(limit).Tag
    ImageListc.ListImages(X).Tag = 1
    limit = limit + 1
    ImageListc.Tag = Val(ImageListc.Tag) + 1
  End If
Wend
For i = 0 To 3
  Imagec(i).TabStop = False
  Imagec(i).Enabled = True
  Imageb(i).Visible = True
Next i
limit = 0
While limit < 4 And Val(ImageListc.Tag) <> 52
  X = Int((52 * Rnd) + 1)
  If ImageListc.ListImages(X).Tag <> 1 Then
    Imagec(limit).Picture = ImageListc.ListImages(X).Picture
    Imagec(limit).Tag = X
    ImageListc.ListImages(X).Tag = 1
    limit = limit + 1
    ImageListc.Tag = Val(ImageListc.Tag) + 1
  End If
Wend
End Sub

Private Sub continue_game()
If Val(ImageListc.Tag) <> 52 Then
Dim X, i, limit As Integer
Dim ok_lem As Boolean
Listccards.Clear
Listhcards.Clear

CARDSELECTED = 0
Randomize

For i = 0 To 3
  Imageh(i).TabStop = False
  Imageh(i).Enabled = True
Next i
limit = 0
While limit < 4 And Val(ImageListc.Tag) <> 52
  X = Int((52 * Rnd) + 1)
  If ImageListc.ListImages(X).Tag <> 1 Then
    Imageh(limit).Picture = ImageListc.ListImages(X).Picture
    Imageh(limit).Tag = X
    ImageListc.ListImages(X).Tag = 1
    limit = limit + 1
    ImageListc.Tag = Val(ImageListc.Tag) + 1
  End If
Wend
For i = 0 To 3
  Imagec(i).TabStop = False
  Imagec(i).Enabled = True
  Imageb(i).Visible = True
Next i
limit = 0
While limit < 4 And Val(ImageListc.Tag) <> 52
  X = Int((52 * Rnd) + 1)
  If ImageListc.ListImages(X).Tag <> 1 Then
    Imagec(limit).Picture = ImageListc.ListImages(X).Picture
    Imagec(limit).Tag = X
    ImageListc.ListImages(X).Tag = 1
    limit = limit + 1
    ImageListc.Tag = Val(ImageListc.Tag) + 1
  End If
Wend
Else
  MsgBox "There is no more cards", vbCritical, "Warning"
  endgame = True
  If global_last_lam = "HUMAN" Then
  For i = 0 To image1.Count - 1
    If image1(i).Picture.Type <> 0 Then
      Listhuman.ListImages.Add Listhuman.ListImages.Count + 1, , image1(i).Picture
      Listhcards.AddItem "T-" + CStr(i) + "-" + image1(i).Tag, Listhcards.ListCount
      image1(i).Picture = nullp.Picture
      image1(i).Enabled = False
      image1(i).Tag = ""
      'LABEL3(i).Caption = image1(i).Tag
      image1(i).TabStop = False
      ok_lem = True
      hlam.Caption = Val(hlam.Caption) + 1
    End If
  Next i
  Else
    For i = 0 To image1.Count - 1
    If image1(i).Picture.Type <> 0 Then
      Listcomp.ListImages.Add Listcomp.ListImages.Count + 1, , image1(i).Picture
      Listccards.AddItem "T-" + CStr(i) + "-" + image1(i).Tag, Listccards.ListCount
      image1(i).Picture = nullp.Picture
      image1(i).Enabled = False
      image1(i).Tag = ""
      'LABEL3(i).Caption = image1(i).Tag
      image1(i).TabStop = False
      ok_lem = True
      clam.Caption = Val(clam.Caption) + 1
    End If
  Next i
  End If
  If endgame Then
    Command3.Enabled = False
    If Val(hlam.Caption) > Val(clam.Caption) Then
      MsgBox "Congratulations.... Your Game", vbInformation, "MedoCards"
      hfinalscore.Caption = Val(hfinalscore.Caption) + 30 + (30 * -1 * drawflag)
      drawflag = False
    ElseIf Val(hlam.Caption) < Val(clam.Caption) Then
      MsgBox "Hard Luck.... My Game", vbInformation, "MedoCards"
      cfinalscore.Caption = Val(cfinalscore.Caption) + 30 + (30 * -1 * drawflag)
      drawflag = False
    Else
      MsgBox "Draw Game next Game 'will be from 60 point'", vbInformation, "MedoCards"
      drawflag = True
    End If
  End If
  If Val(hfinalscore.Caption) > Val(cfinalscore.Caption) And Val(hfinalscore.Caption) >= 210 Then
    MsgBox "Your Game Man", vbInformation, "Game Over"
  ElseIf Val(hfinalscore.Caption) < Val(cfinalscore.Caption) And Val(cfinalscore.Caption) >= 210 Then
    MsgBox "My Game Man", vbInformation, "Game Over"
  ElseIf Val(hfinalscore.Caption) = Val(cfinalscore.Caption) And Val(cfinalscore.Caption) >= 210 Then
    MsgBox "It is Draw Man .... Play Until One Win", vbInformation, "No One"
    new_draw
  Else
    new_draw
  End If
End If
End Sub




Private Sub Command1_Click()
new_draw
End Sub

Private Sub Command3_Click()
  Dim i As Integer
  Dim cont As Boolean
  If Not humanselect Then Exit Sub
  Command3.Enabled = False
  cont = False
  lem_h Listhuman, hlam
  Call play_comp
  For i = 0 To Imageh.Count - 1
     If Imageh(i).Picture.Type <> 0 Then
       cont = False
       Exit For
     Else
       cont = True
     End If
  Next i
  If cont Then
    continue_game
  End If
End Sub



Private Sub Form_Load()
Imageb(0).Picture = ImageListb.ListImages(1).Picture
Imageb(1).Picture = ImageListb.ListImages(1).Picture
Imageb(2).Picture = ImageListb.ListImages(1).Picture
Imageb(3).Picture = ImageListb.ListImages(1).Picture
CARDSELECTED = 0
End Sub


Private Sub Image1_Click(index As Integer)
If image1(index).Picture.Type <> 0 Then
  image1(index).PaintPicture image1(index).Picture, 0, 0, , , , , , , vbDstInvert
  image1(index).TabStop = Not image1(index).TabStop
  If image1(index).TabStop Then
    CARDSELECTED = CARDSELECTED + 1
  Else
    CARDSELECTED = CARDSELECTED - 1
  End If
  Label1.Caption = CARDSELECTED
End If
End Sub


Private Sub Imagec_Click(index As Integer)
Dim i As Integer
For i = 0 To 3
  If Imagec(i).TabStop = True And i <> index Then
    Exit Sub
  End If
Next i
Imagec(index).PaintPicture Imagec(index).Picture, 0, 0, , , , , , , vbDstInvert
Imagec(index).TabStop = Not Imagec(index).TabStop
End Sub

Private Sub Imageh_Click(index As Integer)
Dim i As Integer
For i = 0 To 3
  If Imageh(i).TabStop = True And i <> index Then
    Exit Sub
  End If
Next i
Imageh(index).PaintPicture Imageh(index).Picture, 0, 0, , , , , , , vbDstInvert
Imageh(index).TabStop = Not Imageh(index).TabStop
If Imageh(index).TabStop Then
  humanselect = True
  Command3.Enabled = True
Else
  humanselect = False
  Command3.Enabled = False
End If
End Sub


Private Sub itmnew_game_Click()
cfinalscore.Caption = ""
hfinalscore.Caption = ""
new_draw
End Sub

Private Sub itmundo_Click()
  undo_comp
  undo_human
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
'MsgBox "ok"
comp_play_ok = lem_c(Listcomp, clam)
End Sub


