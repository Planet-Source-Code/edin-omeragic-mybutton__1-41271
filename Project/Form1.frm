VERSION 5.00
Begin VB.Form MyButtonDemo 
   Caption         =   "MyButton Demo Project (Updated)"
   ClientHeight    =   6015
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10230
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   401
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   682
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3540
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   35
      Top             =   4800
      Visible         =   0   'False
      Width           =   240
   End
   Begin MyButtonProject.MyButton MyButton23 
      Height          =   495
      Left            =   6420
      TabIndex        =   34
      ToolTipText     =   "Tool tip is working"
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      SPN             =   "Skn"
      Text            =   "Alone in group"
      FillWithColor   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Group           =   101
      Pressed         =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   4155
      Left            =   6420
      ScaleHeight     =   4155
      ScaleWidth      =   2235
      TabIndex        =   27
      Top             =   120
      Width           =   2235
      Begin MyButtonProject.MyButton MyButton16 
         Height          =   675
         Left            =   60
         TabIndex        =   28
         Top             =   3480
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   1191
         SizeCW          =   1
         SizeCH          =   1
         SPN             =   "Skn"
         Text            =   "MyButton16"
         FillWithColor   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Form1.frx":058A
         Group           =   1
      End
      Begin MyButtonProject.MyButton MyButton15 
         Height          =   675
         Left            =   60
         TabIndex        =   29
         Top             =   2760
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   1191
         SizeCW          =   1
         SizeCH          =   1
         SPN             =   "Skn"
         Text            =   "MyButton15"
         FillWithColor   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Form1.frx":0E64
         Group           =   1
      End
      Begin MyButtonProject.MyButton MyButton14 
         Height          =   675
         Left            =   60
         TabIndex        =   30
         Top             =   2085
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   1191
         SizeCW          =   1
         SizeCH          =   1
         SPN             =   "Skn"
         Text            =   "MyButton14"
         FillWithColor   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Form1.frx":173E
         Group           =   1
      End
      Begin MyButtonProject.MyButton MyButton17 
         Height          =   675
         Left            =   60
         TabIndex        =   31
         Top             =   1410
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   1191
         SPN             =   "Skn"
         Text            =   "MyButton16"
         FillWithColor   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Form1.frx":2018
         Group           =   1
      End
      Begin MyButtonProject.MyButton MyButton18 
         Height          =   675
         Left            =   60
         TabIndex        =   32
         Top             =   735
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   1191
         SPN             =   "Skn"
         Text            =   "MyButton15"
         FillWithColor   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Form1.frx":28F2
         Group           =   1
      End
      Begin MyButtonProject.MyButton MyButton19 
         Height          =   675
         Left            =   60
         TabIndex        =   33
         ToolTipText     =   "What do you wont"
         Top             =   60
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   1191
         SPN             =   "Skn"
         Text            =   "MyButton14"
         FillWithColor   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Form1.frx":31CC
         Group           =   1
      End
   End
   Begin MyButtonProject.MyButton MyButton22 
      Height          =   615
      Left            =   6420
      TabIndex        =   26
      ToolTipText     =   "Clear screen"
      Top             =   5040
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      SPN             =   "MyButtonDefSkin"
      Text            =   "Tips for better picture"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form1.frx":3AA6
   End
   Begin MyButtonProject.MyButton MyButton5 
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2820
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      SPN             =   "MyButtonDefSkin"
      Text            =   "Pita od jabuka"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "CountryBlueprint"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form1.frx":4380
   End
   Begin MyButtonProject.MyButton cmdShowDlg 
      Height          =   435
      Left            =   3540
      TabIndex        =   18
      Top             =   5220
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   767
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Show Dialog"
      AccessKey       =   "S"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Text            =   "Joj  ja sam pijan ko ce me popeti"
      Top             =   5040
      Width           =   3135
   End
   Begin MyButtonProject.MyButton MyButton21 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   8820
      TabIndex        =   20
      Top             =   540
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      SPN             =   "MyButtonDefSkin"
      Text            =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MyButtonProject.MyButton MyButton20 
      Default         =   -1  'True
      Height          =   375
      Left            =   8820
      TabIndex        =   19
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      SPN             =   "MyButtonDefSkin"
      Text            =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkDisplace 
      Caption         =   "Displace"
      Height          =   315
      Left            =   4680
      TabIndex        =   17
      Top             =   4440
      Width           =   1575
   End
   Begin VB.PictureBox Skn 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   3540
      Picture         =   "Form1.frx":4C5A
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   25
      Top             =   1140
      Width           =   2250
   End
   Begin MyButtonProject.MyButton MyButton1 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Click to disable/enable gradient fill"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form1.frx":6FEC
   End
   Begin MyButtonProject.MyButton MyButton12 
      Height          =   435
      Left            =   5760
      TabIndex        =   16
      ToolTipText     =   "Mute"
      Top             =   3900
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   767
      SPN             =   "MyButtonDefSkin"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MyButtonProject.MyButton MyButton11 
      Height          =   435
      Left            =   5205
      TabIndex        =   15
      ToolTipText     =   "Fullscreen"
      Top             =   3900
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   767
      SPN             =   "MyButtonDefSkin"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form1.frx":78C6
   End
   Begin MyButtonProject.MyButton MyButton10 
      Height          =   435
      Left            =   4650
      TabIndex        =   14
      ToolTipText     =   "Stop"
      Top             =   3900
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   767
      SPN             =   "MyButtonDefSkin"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form1.frx":7EA8
   End
   Begin MyButtonProject.MyButton MyButton9 
      Height          =   435
      Left            =   4095
      TabIndex        =   13
      ToolTipText     =   "Pause"
      Top             =   3900
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   767
      SPN             =   "MyButtonDefSkin"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form1.frx":848A
   End
   Begin MyButtonProject.MyButton MyButton8 
      Height          =   435
      Left            =   3540
      TabIndex        =   12
      ToolTipText     =   "Play"
      Top             =   3900
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   767
      SPN             =   "MyButtonDefSkin"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form1.frx":8A6C
   End
   Begin VB.PictureBox Skin1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   3540
      Picture         =   "Form1.frx":904E
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   24
      Top             =   780
      Width           =   2250
   End
   Begin MyButtonProject.MyButton MyButton7 
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      SPN             =   "MyButtonDefSkin"
      Text            =   "Print Test Page"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form1.frx":95AB
   End
   Begin VB.PictureBox picNew 
      AutoSize        =   -1  'True
      Height          =   360
      Left            =   4260
      Picture         =   "Form1.frx":9E85
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   23
      Top             =   4440
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   3540
      Picture         =   "Form1.frx":A457
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   22
      Top             =   120
      Width           =   2250
   End
   Begin VB.PictureBox Standard 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3540
      Picture         =   "Form1.frx":C9AD
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   21
      Top             =   480
      Width           =   1200
   End
   Begin MyButtonProject.MyButton MyButton2 
      Height          =   570
      Left            =   120
      TabIndex        =   2
      Top             =   795
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1005
      SPN             =   "MyButtonDefSkin"
      Text            =   "Disabled Button"
      Enabled         =   0   'False
      TextColorDisabled=   -2147483632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColorDisabled2=   -2147483628
      Picture         =   "Form1.frx":D8EF
   End
   Begin MyButtonProject.MyButton MyButton3 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Click to test text justify"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form1.frx":E1C9
      TextAlign       =   0
      OwnerDraw       =   -1  'True
   End
   Begin MyButtonProject.MyButton MyButton4 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2145
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Picture can be on left side too"
      AccessKey       =   "P"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form1.frx":EAA3
      PicturePos      =   3
   End
   Begin MyButtonProject.MyButton MyButton6 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1296
      SPN             =   "MyButtonDefSkin"
      Text            =   "Shut down app with scream"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form1.frx":F37D
   End
   Begin MyButtonProject.MyButton cmdChangeSkin 
      Height          =   495
      Left            =   3540
      TabIndex        =   8
      Top             =   1560
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   873
      SPN             =   "MyButtonDefSkin"
      Text            =   "Change clothes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form1.frx":FC57
      PicturePos      =   3
   End
   Begin MyButtonProject.MyButton cmdChangeColor 
      Height          =   495
      Left            =   3540
      TabIndex        =   9
      Top             =   2100
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   873
      SPN             =   "MyButtonDefSkin"
      Text            =   "Change text color"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form1.frx":10239
      PicturePos      =   3
   End
   Begin MyButtonProject.MyButton cmdMoreStuff 
      Height          =   555
      Left            =   3540
      TabIndex        =   10
      Top             =   2640
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   979
      SPN             =   "MyButtonDefSkin"
      Text            =   "Show more stuff"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form1.frx":1081B
      PicturePos      =   3
   End
   Begin MyButtonProject.MyButton MyButton13 
      Height          =   495
      Left            =   3540
      TabIndex        =   11
      ToolTipText     =   "Smece"
      Top             =   3300
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   873
      SPN             =   "MyButtonDefSkin"
      Text            =   "Set skn picture to nothing"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DrawFocus       =   4
      BackColor       =   -2147483635
   End
   Begin MyButtonProject.MyButton MyButton24 
      Height          =   495
      Left            =   7680
      TabIndex        =   36
      ToolTipText     =   "Tool tip is working"
      Top             =   4440
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   873
      SPN             =   "Skn"
      Text            =   "Alone in group"
      FillWithColor   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Group           =   102
      Pressed         =   -1  'True
   End
   Begin MyButtonProject.MyButton MyButton25 
      Height          =   495
      Left            =   8880
      TabIndex        =   37
      ToolTipText     =   "Tool tip is working"
      Top             =   4440
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   873
      SPN             =   "Skn"
      Text            =   "Alone in group"
      FillWithColor   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Group           =   103
      Pressed         =   -1  'True
   End
   Begin VB.Image imgSoundOn 
      Height          =   300
      Left            =   3900
      Picture         =   "Form1.frx":10DFD
      Top             =   4440
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgSoundOff 
      Height          =   300
      Left            =   3540
      Picture         =   "Form1.frx":113CF
      Top             =   4440
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Menu mnuPopUP 
      Caption         =   "MenuPopup"
      Begin VB.Menu mnuLeft 
         Caption         =   "Left"
      End
      Begin VB.Menu mnuRight 
         Caption         =   "Right"
      End
      Begin VB.Menu mnuCenter 
         Caption         =   "Center"
      End
   End
End
Attribute VB_Name = "MyButtonDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DEMO for MyButton UserControl
Const BTN_NORMAL = 1
Const BTN_FOCUS = 2
Const BTN_HOVER = 3
Const BTN_DOWN = 4
Const BTN_DISABLED = 5

Dim Mute As Boolean

Private Sub chkDisplace_Click()
    Dim C As Object
    Dim x As Long
    If chkDisplace.Value Then
        x = 1
    Else
        x = 0
    End If
    
    For Each C In Me.Controls
        If TypeName(C) = "MyButton" Then
            C.DisplaceText = x
        End If
    Next
End Sub

Private Sub cmdChangeSkin_Click()
    Dim C As Object
    Dim B As MyButton
    
    For Each C In Me.Controls
        If TypeName(C) = "MyButton" Then
        Set B = C
        Select Case B.SkinPictureName
            Case "MyButtonDefSkin"
                Set B.SkinPicture = Standard
                B.DisableHover = True
                B.DrawFocus = 4
            Case "Standard"
                Set B.SkinPicture = Skin1
                B.DrawFocus = 0
                B.DisableHover = False
            Case "Skin1"
                Set C.SkinPicture = MyButtonDefSkin
                C.DisableHover = False
                C.DrawFocus = 0
        End Select
        End If
    Next
End Sub


Private Sub cmdMoreStuff_MouseHover()
    cmdMoreStuff.FontBold = True
    cmdMoreStuff.TextColorEnabled = vbHighlight
    Set cmdMoreStuff.Picture = picNew.Picture
End Sub

Private Sub cmdMoreStuff_MouseOut()
    cmdMoreStuff.FontBold = False
    cmdMoreStuff.TextColorEnabled = vbBlack
    Set cmdMoreStuff.Picture = cmdChangeColor.Picture
End Sub

Private Sub cmdShowDlg_Click()
    
    Dim IB As New InputBox
    
    IB.Show vbModal
    
    If IB.Response = vbOK Then
        Text1.Text = IB.Text
    Else
        Text1.Text = "UNNAMED"
    End If
    
End Sub

Private Sub Form_Load()
    chkDisplace.Value = vbChecked
    Me.Tag = Me.Caption
    Set MyButton12.Picture = imgSoundOff
    mnuPopUP.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Debug.Print "QUERY UNLOAD"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Dim f As Form
'    For Each f In Forms
'        Unload f
'        Set f = Nothing
'    Next
'
'    'End
End Sub




Private Sub mnuCenter_Click()
        MyButton3.TextAlign = vbCenter
End Sub

Private Sub mnuLeft_Click()
    MyButton3.TextAlign = vbLeftJustify
End Sub

Private Sub mnuRight_Click()
        MyButton3.TextAlign = vbRightJustify
End Sub

Private Sub MyButton1_Click()
    'enable / disable gradient fill
    MyButton1.DisableGradient = Not MyButton1.DisableGradient
    MyButton2.DisableGradient = Not MyButton2.DisableGradient
    MyButton3.DisableGradient = Not MyButton3.DisableGradient
    MyButton4.DisableGradient = Not MyButton4.DisableGradient
    MyButton5.DisableGradient = Not MyButton5.DisableGradient

End Sub

Private Sub MyButton11_Click()

If Me.WindowState = vbMaximized Then
    Me.WindowState = 0
Else
    Me.WindowState = vbMaximized
End If
End Sub

Private Sub MyButton12_Click()
    
    If Mute Then
        Set MyButton12.Picture = imgSoundOff
    Else
        Set MyButton12.Picture = imgSoundOn
    End If
    Mute = Not Mute
End Sub

Private Sub MyButton13_Click()
If Not MyButton13.SkinPicture Is Nothing Then
    Set MyButton13.SkinPicture = Nothing
    MyButton13.Text = "Standard button"
End If
    
End Sub

Private Sub MyButton20_Click()
    MsgBox "Ok"
End Sub

Private Sub MyButton21_Click()
    MsgBox "Cancel", vbInformation
End Sub

Private Sub MyButton22_Click()
    MsgBox "Not evrything can be done by software...", vbInformation
End Sub

Private Sub MyButton3_Click()
    PopupMenu mnuPopUP, vbPopupMenuRightAlign, MyButton3.Left + MyButton3.Width, MyButton3.Top + MyButton3.Height
End Sub



Private Sub MyButton3_OnDrawButton(ByVal State As Integer)
   'subclassing hihihi (easy way)
   '====================
   'OWNER DRAW BUTTON - Only for special efects
   '====================
   Dim W As Long
   Dim H As Long
   Dim D As Long
   
   Me.ScaleMode = vbPixels
   W = MyButton3.Width
   H = MyButton3.Height
   
   If State = BTN_DOWN Then
     D = MyButton3.DisplaceText * 2
   Else
     D = 0
   End If
   
   'Draw image on right side
   MyButton3.PaintPicture picIcon, W - picIcon.Width - 10 + D, (H - picIcon.Height) / 2 + D

End Sub

Private Sub MyButton4_Click()
    MsgBox "Hello world!!"
End Sub

Private Sub MyButton5_Click()
    MsgBox "in bosnien it means 'Apple pie'", vbInformation
End Sub

Private Sub MyButton6_Click()
    'you scream
    Unload Me
End Sub

Private Sub cmdChangeColor_Click()
    Dim C As Object
    Dim B As MyButton
    Dim TextColor As Long
    
    If MyButton1.TextColorEnabled = vbButtonShadow Then
        TextColor = vbBlack
    Else
        TextColor = vbButtonShadow
    End If
    
    For Each C In Me.Controls
        If TypeName(C) = "MyButton" Then
            Set B = C
            B.TextColorEnabled = TextColor
        End If
    Next
End Sub

Private Sub MyButton7_Click()
' MsgBox "Cool! Isn't it?", vbCritical, "On print"
' Debug.Print Printer.DeviceName

' only a test (not so important at the time)
   On Error Resume Next

' this checks do you have Jaws PDF Creator, and then prints somthing using
' it instead of default printer

   Dim p As Printer
   For Each p In Printers
        If p.DeviceName = "Jaws PDF Creator" Then
            Set Printer = p
            'Printer.DriverName = "Jaws PDF Creator"
            Printer.FontSize = 14
            Printer.ScaleMode = vbMillimeters
            Printer.CurrentX = 25
            Printer.CurrentY = 10
            
            Printer.Print "Testing the Click event of MyButton Control."
            
            Printer.Circle (210 / 2, 297 / 2), 50
            Printer.Line (20, 5)-(205, 292), , B
            Printer.EndDoc
            Exit Sub
        End If
   Next

MsgBox "Cool! Isn't it?", vbCritical, "On print"

End Sub

Function AppPath(Optional FileName As String = "") As String
Dim TempPath As String
If Right(App.Path, 1) <> "\" Then
   TempPath = App.Path + "\"
Else
   TempPath = App.Path
End If

AppPath = TempPath + FileName

End Function

Private Sub MyButton8_Click()
    'MsgBox "If you think its good, please vote", vbInformation
End Sub

Private Sub MyButton8_MouseHover()
    Me.Caption = "Press to play"
End Sub

Private Sub MyButton8_MouseOut()
    Me.Caption = Me.Tag
End Sub
