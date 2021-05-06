VERSION 5.00
Object = "*\A..\control-source\TabExCtl.vbp"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test SSTabEx properties"
   ClientHeight    =   7380
   ClientLeft      =   1488
   ClientTop       =   1548
   ClientWidth     =   12156
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   12156
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAutoTabHeight 
      Caption         =   "AutoTabHeight"
      Height          =   264
      Left            =   8010
      TabIndex        =   62
      Top             =   4188
      Width           =   1740
   End
   Begin VB.ComboBox cboBackStyle 
      Height          =   336
      ItemData        =   "frmTest.frx":0000
      Left            =   6780
      List            =   "frmTest.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   456
      Width           =   4872
   End
   Begin VB.CommandButton cmdChangeTabSelForeColor 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   11685
      TabIndex        =   95
      Top             =   5592
      Width           =   330
   End
   Begin VB.PictureBox picTabSelForeColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8CCB1&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   11070
      ScaleHeight     =   276
      ScaleWidth      =   564
      TabIndex        =   94
      Top             =   5592
      Width           =   588
   End
   Begin VB.PictureBox picTabSelBackColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8CCB1&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6780
      ScaleHeight     =   276
      ScaleWidth      =   564
      TabIndex        =   92
      Top             =   5544
      Width           =   588
   End
   Begin VB.CommandButton cmdChangeTabSelBackColor 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   7440
      TabIndex        =   91
      Top             =   5556
      Width           =   330
   End
   Begin VB.ComboBox cboTabHoverHighlight 
      Height          =   336
      ItemData        =   "frmTest.frx":0004
      Left            =   6780
      List            =   "frmTest.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   80
      Top             =   5916
      Width           =   2700
   End
   Begin VB.CheckBox chkShowFocusRect 
      Caption         =   "ShowFocusRect"
      Height          =   264
      Left            =   9840
      TabIndex        =   59
      Top             =   3900
      Width           =   1740
   End
   Begin VB.CheckBox chkSoftEdges 
      Caption         =   "SoftEdges"
      Height          =   264
      Left            =   9840
      TabIndex        =   56
      Top             =   3612
      Width           =   1740
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   5070
      Top             =   4236
   End
   Begin VB.ComboBox cboTabPictureAlignment 
      Height          =   336
      ItemData        =   "frmTest.frx":0008
      Left            =   6780
      List            =   "frmTest.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   82
      Top             =   6276
      Width           =   2712
   End
   Begin VB.CommandButton cmdChangeBackColor 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   11685
      TabIndex        =   88
      Top             =   4860
      Width           =   330
   End
   Begin VB.PictureBox picBackColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8CCB1&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   11070
      ScaleHeight     =   276
      ScaleWidth      =   564
      TabIndex        =   87
      Top             =   4860
      Width           =   588
   End
   Begin VB.CommandButton cmdAllTabsVisibleEnabled 
      Caption         =   "All tabs visible and enabled"
      Height          =   336
      Left            =   2916
      TabIndex        =   27
      Top             =   6852
      Width           =   2400
   End
   Begin VB.PictureBox picMaskColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8CCB1&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6780
      ScaleHeight     =   276
      ScaleWidth      =   564
      TabIndex        =   67
      Top             =   4836
      Width           =   588
   End
   Begin VB.CommandButton cmdChangeMaskColor 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   7440
      TabIndex        =   68
      Top             =   4848
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tab specific data (current tab):"
      Height          =   1668
      Left            =   288
      TabIndex        =   17
      Top             =   4356
      Width           =   4584
      Begin VB.CommandButton cmdRemoveTabPicture 
         Height          =   336
         Left            =   4140
         Picture         =   "frmTest.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   25
         Tag             =   "E"
         ToolTipText     =   "Remove Tab"
         Top             =   1116
         Width           =   275
      End
      Begin VB.CommandButton cmdDisableTab 
         Caption         =   "Disable tab"
         Height          =   336
         Left            =   1548
         TabIndex        =   23
         Top             =   1116
         Width           =   1176
      End
      Begin VB.CommandButton cmdHideTab 
         Caption         =   "Hide tab"
         Height          =   336
         Left            =   180
         TabIndex        =   22
         Top             =   1116
         Width           =   1176
      End
      Begin VB.CommandButton cmdChangeTabPicture 
         Caption         =   "Load picture"
         Height          =   336
         Left            =   2916
         TabIndex        =   24
         Top             =   1116
         Width           =   1176
      End
      Begin VB.TextBox txtTabToolTipText 
         Height          =   300
         Left            =   1584
         TabIndex        =   21
         Top             =   684
         Width           =   2820
      End
      Begin VB.TextBox txtTabCaption 
         Height          =   300
         Left            =   1584
         TabIndex        =   19
         Top             =   324
         Width           =   2820
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "TabToolTipText:"
         Height          =   228
         Left            =   144
         TabIndex        =   20
         Top             =   720
         Width           =   1344
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "TabCaption:"
         Height          =   228
         Left            =   144
         TabIndex        =   18
         Top             =   360
         Width           =   1344
      End
   End
   Begin VB.TextBox txtTabHeight 
      Height          =   300
      Left            =   6780
      TabIndex        =   52
      Top             =   3396
      Width           =   588
   End
   Begin VB.TextBox txtTabs 
      Height          =   300
      Left            =   8940
      MaxLength       =   3
      TabIndex        =   35
      Top             =   804
      Width           =   588
   End
   Begin VB.CommandButton cmdChangeFont 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   11685
      TabIndex        =   78
      Top             =   5232
      Width           =   330
   End
   Begin VB.TextBox txtFont 
      Height          =   300
      Left            =   8940
      Locked          =   -1  'True
      TabIndex        =   77
      Top             =   5220
      Width           =   2724
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   336
      Left            =   6780
      TabIndex        =   83
      Top             =   6852
      Width           =   1000
   End
   Begin VB.PictureBox picForeColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8CCB1&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   8940
      ScaleHeight     =   276
      ScaleWidth      =   564
      TabIndex        =   74
      Top             =   4860
      Width           =   588
   End
   Begin VB.CommandButton cmdChangeForeColor 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   9555
      TabIndex        =   75
      Top             =   4860
      Width           =   330
   End
   Begin VB.CommandButton cmdChangeTabBackColor 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   7440
      TabIndex        =   72
      Top             =   5208
      Width           =   330
   End
   Begin VB.PictureBox picTabBackColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8CCB1&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6780
      ScaleHeight     =   276
      ScaleWidth      =   564
      TabIndex        =   71
      Top             =   5196
      Width           =   588
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2952
      Top             =   6516
   End
   Begin VB.ComboBox cboTabAppearance 
      Height          =   336
      ItemData        =   "frmTest.frx":0156
      Left            =   8940
      List            =   "frmTest.frx":0158
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   1560
      Width           =   2712
   End
   Begin VB.ComboBox cboShowRowsInPerspective 
      Height          =   336
      ItemData        =   "frmTest.frx":015A
      Left            =   8940
      List            =   "frmTest.frx":015C
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   2280
      Width           =   2712
   End
   Begin VB.TextBox txtTabMinWidth 
      Height          =   324
      Left            =   6780
      TabIndex        =   64
      Top             =   4476
      Width           =   588
   End
   Begin VB.TextBox txtTabMaxWidth 
      Height          =   300
      Left            =   6780
      TabIndex        =   61
      Top             =   4116
      Width           =   588
   End
   Begin VB.CheckBox chkShowDisabledState 
      Caption         =   "ShowDisabledState"
      Height          =   264
      Left            =   9840
      TabIndex        =   50
      Top             =   3036
      Width           =   2028
   End
   Begin VB.CheckBox chkChangeControlsBackColor 
      Caption         =   "ChangeControlsBackColor"
      Height          =   264
      Left            =   8010
      TabIndex        =   65
      Top             =   4476
      Width           =   2892
   End
   Begin VB.CheckBox chkTabSelHighlight 
      Caption         =   "TabSelHighlight"
      Height          =   264
      Left            =   8010
      TabIndex        =   55
      Top             =   3612
      Width           =   1740
   End
   Begin VB.CheckBox chkUseMaskColor 
      Caption         =   "UseMaskColor"
      Height          =   264
      Left            =   8010
      TabIndex        =   69
      Top             =   3900
      Width           =   1740
   End
   Begin VB.ComboBox cboTabSelFontBold 
      Height          =   336
      ItemData        =   "frmTest.frx":015E
      Left            =   8940
      List            =   "frmTest.frx":0160
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Top             =   2640
      Width           =   2712
   End
   Begin VB.TextBox txtTabSelExtraHeight 
      Height          =   300
      Left            =   6780
      TabIndex        =   58
      Top             =   3756
      Width           =   588
   End
   Begin VB.TextBox txtTabSeparation 
      Height          =   300
      Left            =   6780
      MaxLength       =   2
      TabIndex        =   48
      Top             =   3036
      Width           =   588
   End
   Begin VB.CheckBox chkEnabled 
      Caption         =   "Enabled"
      Height          =   264
      Left            =   8010
      TabIndex        =   49
      Top             =   3036
      Width           =   1740
   End
   Begin VB.CheckBox chkWordWrap 
      Caption         =   "WordWrap"
      Height          =   264
      Left            =   9840
      TabIndex        =   54
      Top             =   3324
      Width           =   1740
   End
   Begin VB.CheckBox chkVisualStyles 
      Caption         =   "VisualStyles"
      Height          =   264
      Left            =   8010
      TabIndex        =   53
      Top             =   3324
      Width           =   1740
   End
   Begin VB.ComboBox cboTabWidthStyle 
      Height          =   336
      ItemData        =   "frmTest.frx":0162
      Left            =   8940
      List            =   "frmTest.frx":0164
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   1920
      Width           =   2712
   End
   Begin VB.ComboBox cboStyle 
      Height          =   336
      ItemData        =   "frmTest.frx":0166
      Left            =   6780
      List            =   "frmTest.frx":0168
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   1164
      Width           =   4872
   End
   Begin VB.TextBox txtTabsPerRow 
      Height          =   300
      Left            =   6780
      MaxLength       =   2
      TabIndex        =   33
      Top             =   804
      Width           =   588
   End
   Begin VB.ComboBox cboOrientation 
      Height          =   336
      ItemData        =   "frmTest.frx":016A
      Left            =   6780
      List            =   "frmTest.frx":016C
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   108
      Width           =   4872
   End
   Begin TabExCtl.SSTabEx SSTabEx1 
      Height          =   4116
      Left            =   252
      TabIndex        =   0
      Top             =   108
      Width           =   4644
      _ExtentX        =   8192
      _ExtentY        =   7260
      Tabs            =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabsPerRow      =   4
      Tab             =   1
      TabHeight       =   601
      AutoTabHeight   =   -1  'True
      TabPic16(0)     =   "frmTest.frx":016E
      TabPic20(0)     =   "frmTest.frx":04C0
      TabPic24(0)     =   "frmTest.frx":09C2
      TabCaption(0)   =   "Theme"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Picture2(0)"
      TabPic16(1)     =   "frmTest.frx":10D4
      TabPic20(1)     =   "frmTest.frx":1426
      TabPic24(1)     =   "frmTest.frx":1928
      TabCaption(1)   =   "Frame"
      Tab(1).ControlCount=   1
      Tab(1).Control(0)=   "Frame1"
      TabPic16(2)     =   "frmTest.frx":203A
      TabPic20(2)     =   "frmTest.frx":238C
      TabPic24(2)     =   "frmTest.frx":288E
      TabCaption(2)   =   "Other"
      Tab(2).ControlCount=   1
      Tab(2).Control(0)=   "Picture1"
      TabPic16(3)     =   "frmTest.frx":2FA0
      TabPic20(3)     =   "frmTest.frx":32F2
      TabPic24(3)     =   "frmTest.frx":37F4
      TabCaption(3)   =   "Cmd"
      Tab(3).ControlCount=   6
      Tab(3).Control(0)=   "Option4"
      Tab(3).Control(1)=   "Check2"
      Tab(3).Control(2)=   "Option3"
      Tab(3).Control(3)=   "Command2"
      Tab(3).Control(4)=   "Command1"
      Tab(3).Control(5)=   "Label8"
      TabCaption(4)   =   "Label"
      Tab(4).ControlCount=   2
      Tab(4).Control(0)=   "Picture3"
      Tab(4).Control(1)=   "Label22"
      TabCaption(5)   =   "Tab 5"
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Tab 6"
      Tab(6).ControlCount=   0
      Begin VB.PictureBox Picture3 
         Height          =   12
         Left            =   -84972
         ScaleHeight     =   12
         ScaleWidth      =   12
         TabIndex        =   90
         Top             =   240
         Width           =   12
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1596
         Index           =   0
         Left            =   -74820
         ScaleHeight     =   1596
         ScaleWidth      =   3720
         TabIndex        =   85
         Top             =   960
         Width           =   3720
         Begin VB.Label lblThemedIDE 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "The SSTabEx is themed now in the IDE just for testing, normally it won't be themed in the IDE"
            Height          =   984
            Left            =   468
            TabIndex        =   86
            Top             =   216
            Width           =   2784
         End
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Option4"
         Height          =   300
         Left            =   -74568
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2724
         Width           =   1900
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   480
         Left            =   -74568
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3156
         Width           =   1900
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   300
         Left            =   -74568
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2364
         Value           =   -1  'True
         Width           =   1900
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command graphical"
         Height          =   588
         Left            =   -74568
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1644
         Width           =   1900
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command standard"
         Height          =   588
         Left            =   -74568
         TabIndex        =   10
         Top             =   960
         Width           =   1900
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2928
         Left            =   -74784
         ScaleHeight     =   2928
         ScaleWidth      =   3720
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   852
         Width           =   3720
         Begin VB.ListBox List1 
            Height          =   528
            Left            =   396
            TabIndex        =   8
            Top             =   1224
            Width           =   1848
         End
         Begin VB.ComboBox Combo1 
            Height          =   336
            ItemData        =   "frmTest.frx":3F06
            Left            =   396
            List            =   "frmTest.frx":3F13
            TabIndex        =   7
            Text            =   "1234567890"
            Top             =   648
            Width           =   1164
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "These controls are inside a PictureBox"
            Height          =   240
            Left            =   432
            TabIndex        =   9
            Top             =   2340
            Width           =   2988
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Just a Label"
            Height          =   372
            Left            =   396
            TabIndex        =   6
            Top             =   180
            Width           =   1668
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   2784
         Left            =   360
         TabIndex        =   1
         Top             =   852
         Width           =   2640
         Begin VB.TextBox Text1 
            Height          =   336
            Left            =   360
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   1728
            Width           =   1056
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   408
            Left            =   360
            TabIndex        =   4
            Top             =   1188
            Width           =   1344
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Option2"
            Height          =   300
            Left            =   324
            TabIndex        =   3
            Top             =   756
            Value           =   -1  'True
            Width           =   1164
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   300
            Left            =   324
            TabIndex        =   2
            Top             =   324
            Width           =   1488
         End
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTest.frx":3F41
         Height          =   1164
         Left            =   -74280
         TabIndex        =   16
         Top             =   1212
         Width           =   2784
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "The painting of the contained controls must be tested compiled, when they have the visual style applied"
         Height          =   1524
         Left            =   -72516
         TabIndex        =   15
         Top             =   1140
         Width           =   1380
      End
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      Caption         =   "BackStyle:"
      Height          =   228
      Left            =   5208
      TabIndex        =   30
      Top             =   504
      Width           =   1488
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      Caption         =   "TabSelForeColor:"
      Height          =   228
      Left            =   9600
      TabIndex        =   96
      Top             =   5628
      Width           =   1368
   End
   Begin VB.Label lblTabSelBackColor 
      Alignment       =   1  'Right Justify
      Caption         =   "TabSelBackColor:"
      Height          =   228
      Left            =   5208
      TabIndex        =   93
      Top             =   5604
      Width           =   1488
   End
   Begin VB.Label lblTabHoverHighlight 
      Alignment       =   1  'Right Justify
      Caption         =   "TabHoverHighlight:"
      Height          =   228
      Left            =   5088
      TabIndex        =   79
      Top             =   5952
      Width           =   1596
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      Caption         =   "TabPictureAlignment:"
      Height          =   228
      Left            =   4632
      TabIndex        =   81
      Top             =   6312
      Width           =   2052
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      Caption         =   "BackColor:"
      Height          =   228
      Left            =   9900
      TabIndex        =   89
      Top             =   4896
      Width           =   1032
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      Caption         =   "MaskColor:"
      Height          =   228
      Left            =   5208
      TabIndex        =   66
      Top             =   4872
      Width           =   1488
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "TabHeight:"
      Height          =   228
      Index           =   2
      Left            =   5676
      TabIndex        =   51
      Top             =   3432
      Width           =   1020
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Tab Count:"
      Height          =   228
      Left            =   7800
      TabIndex        =   34
      Top             =   852
      Width           =   1056
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Caption         =   "Font:"
      Height          =   228
      Left            =   7848
      TabIndex        =   76
      Top             =   5256
      Width           =   984
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Caption         =   "ForeColor:"
      Height          =   228
      Left            =   7812
      TabIndex        =   73
      Top             =   4896
      Width           =   1008
   End
   Begin VB.Label lblTabBackColor 
      Alignment       =   1  'Right Justify
      Caption         =   "TabBackColor:"
      Height          =   228
      Left            =   5208
      TabIndex        =   70
      Top             =   5256
      Width           =   1488
   End
   Begin VB.Label lblFocus 
      ForeColor       =   &H00C00000&
      Height          =   264
      Left            =   324
      TabIndex        =   26
      Top             =   6888
      Width           =   2568
   End
   Begin VB.Label lblTabAppearance 
      Alignment       =   1  'Right Justify
      Caption         =   "TabAppearance:"
      Height          =   228
      Left            =   7032
      TabIndex        =   39
      Top             =   1596
      Width           =   1812
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "ShowRowsInPerspective:"
      Height          =   228
      Left            =   6552
      TabIndex        =   43
      Top             =   2316
      Width           =   2292
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "TabMinWidth:"
      Height          =   228
      Left            =   5208
      TabIndex        =   63
      Top             =   4512
      Width           =   1488
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "TabMaxWidth:"
      Height          =   228
      Left            =   5208
      TabIndex        =   60
      Top             =   4152
      Width           =   1488
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "TabSelFontBold:"
      Height          =   228
      Left            =   7032
      TabIndex        =   45
      Top             =   2676
      Width           =   1812
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "TabSelExtraHeight:"
      Height          =   228
      Left            =   5028
      TabIndex        =   57
      Top             =   3792
      Width           =   1668
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "TabSeparation:"
      Height          =   228
      Left            =   5208
      TabIndex        =   47
      Top             =   3072
      Width           =   1488
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "TabWidthStyle:"
      Height          =   228
      Left            =   7032
      TabIndex        =   41
      Top             =   1956
      Width           =   1812
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Customize Style:"
      Height          =   228
      Left            =   5100
      TabIndex        =   38
      Top             =   1596
      Width           =   1596
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Style:"
      Height          =   228
      Left            =   5208
      TabIndex        =   36
      Top             =   1212
      Width           =   1488
   End
   Begin VB.Label lblTabsPerRow 
      Alignment       =   1  'Right Justify
      Caption         =   "TabsPerRow:"
      Height          =   228
      Left            =   5208
      TabIndex        =   32
      Top             =   852
      Width           =   1488
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Orientation:"
      Height          =   225
      Left            =   5205
      TabIndex        =   28
      Top             =   150
      Width           =   1485
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cboOrientation_Click()
    SSTabEx1.TabOrientation = cboOrientation.ListIndex
End Sub

Private Sub cboBackStyle_Click()
    SSTabEx1.BackStyle = cboBackStyle.ListIndex
    chkVisualStyles.Enabled = (SSTabEx1.BackStyle = ssOpaque)
    chkVisualStyles_Click
End Sub

Private Sub cboShowRowsInPerspective_Click()
    SSTabEx1.ShowRowsInPerspective = cboShowRowsInPerspective.ListIndex
End Sub

Private Sub cboStyle_Click()
    SSTabEx1.Style = cboStyle.ListIndex
    lblTabsPerRow.Enabled = SSTabEx1.Style <> ssStyleTabStrip
    txtTabsPerRow.Enabled = lblTabsPerRow.Enabled
End Sub

Private Sub cboTabAppearance_Click()
    SSTabEx1.TabAppearance = cboTabAppearance.ListIndex
End Sub

Private Sub cboTabHoverHighlight_Click()
    SSTabEx1.TabHoverHighlight = cboTabHoverHighlight.ListIndex
End Sub

Private Sub cboTabPictureAlignment_Click()
    SSTabEx1.TabPictureAlignment = cboTabPictureAlignment.ListIndex
End Sub

Private Sub cboTabSelFontBold_Click()
    SSTabEx1.TabSelFontBold = cboTabSelFontBold.ListIndex
End Sub

Private Sub cboTabWidthStyle_Click()
    SSTabEx1.TabWidthStyle = cboTabWidthStyle.ListIndex
End Sub

Private Sub chkAutoTabHeight_Click()
    SSTabEx1.AutoTabHeight = chkAutoTabHeight.Value = 1
    txtTabHeight.Text = SSTabEx1.TabHeight
End Sub

Private Sub chkChangeControlsBackColor_Click()
    SSTabEx1.ChangeControlsBackColor = chkChangeControlsBackColor.Value = 1
End Sub

Private Sub chkEnabled_Click()
    SSTabEx1.Enabled = chkEnabled.Value = 1
End Sub

Private Sub chkShowDisabledState_Click()
    SSTabEx1.ShowDisabledState = chkShowDisabledState.Value = 1
End Sub

Private Sub chkShowFocusRect_Click()
    SSTabEx1.ShowFocusRect = chkShowFocusRect.Value = 1
End Sub

Private Sub chkSoftEdges_Click()
    SSTabEx1.SoftEdges = chkSoftEdges.Value = 1
End Sub

Private Sub chkTabSelHighlight_Click()
    SSTabEx1.TabSelHighlight = chkTabSelHighlight.Value = 1
End Sub

Private Sub chkUseMaskColor_Click()
    SSTabEx1.UseMaskColor = chkUseMaskColor.Value = 1
End Sub

Private Sub chkVisualStyles_Click()
    SSTabEx1.VisualStyles = chkVisualStyles.Value = 1
'    If Not SSTabEx1.IsVisualStyleApplied Then
'        chkVisualStyles.Value = 0
'    End If
    chkTabSelHighlight.Enabled = Not SSTabEx1.IsVisualStyleApplied
    cboTabAppearance.Enabled = chkTabSelHighlight.Enabled
    chkShowDisabledState.Enabled = chkTabSelHighlight.Enabled
    lblTabBackColor.Enabled = chkTabSelHighlight.Enabled
    lblTabSelBackColor.Enabled = chkTabSelHighlight.Enabled
    picTabBackColor.Enabled = chkTabSelHighlight.Enabled
    picTabSelBackColor.Enabled = chkTabSelHighlight.Enabled
    cmdChangeTabBackColor.Enabled = chkTabSelHighlight.Enabled
    cmdChangeTabSelBackColor.Enabled = chkTabSelHighlight.Enabled
    lblThemedIDE.Visible = InIde And SSTabEx1.IsVisualStyleApplied
End Sub

Private Sub chkWordWrap_Click()
    SSTabEx1.WordWrap = chkWordWrap.Value = 1
End Sub

Private Sub cmdChangeBackColor_Click()
    Dim iDlg As New cDlg
    
    iDlg.ShowColor
    If Not iDlg.Canceled Then
        picBackColor.BackColor = iDlg.Color
        SSTabEx1.BackColor = picBackColor.BackColor
    End If

End Sub

Private Sub cmdChangeFont_Click()
    Dim iDlg As New cDlg
    
    Set iDlg.Font = SSTabEx1.Font
    iDlg.ShowFont
    If Not iDlg.Canceled Then
        SSTabEx1.Font = iDlg.Font
        ShowFont
        txtTabHeight.Text = SSTabEx1.TabHeight
    End If

End Sub

Private Sub ShowFont()
    txtFont.Text = SSTabEx1.Font.Name & "  " & SSTabEx1.Font.Size
End Sub

Private Sub cmdChangeForeColor_Click()
    Dim iDlg As New cDlg
    
    iDlg.ShowColor
    If Not iDlg.Canceled Then
        picForeColor.BackColor = iDlg.Color
        SSTabEx1.ForeColor = picForeColor.BackColor
        picTabSelForeColor.BackColor = SSTabEx1.TabSelForeColor
        lblThemedIDE.ForeColor = SSTabEx1.TabSelForeColor
        Option1.ForeColor = SSTabEx1.TabSelForeColor
        Option2.ForeColor = SSTabEx1.TabSelForeColor
        Check1.ForeColor = SSTabEx1.TabSelForeColor
        Label1.ForeColor = SSTabEx1.TabSelForeColor
        Label22.ForeColor = SSTabEx1.TabSelForeColor
        Option3.ForeColor = SSTabEx1.TabSelForeColor
        Option4.ForeColor = SSTabEx1.TabSelForeColor
        Check2.ForeColor = SSTabEx1.TabSelForeColor
    End If
End Sub

Private Sub cmdChangeMaskColor_Click()
    Dim iDlg As New cDlg
    
    iDlg.ShowColor
    If Not iDlg.Canceled Then
        picMaskColor.BackColor = iDlg.Color
        SSTabEx1.MaskColor = picMaskColor.BackColor
    End If

End Sub

Private Sub cmdChangeTabBackColor_Click()
    Dim iDlg As New cDlg
    
    iDlg.ShowColor
    If Not iDlg.Canceled Then
        picTabBackColor.BackColor = iDlg.Color
        SSTabEx1.TabBackColor = picTabBackColor.BackColor
        picTabSelBackColor.BackColor = SSTabEx1.TabSelBackColor
    End If
End Sub

Private Sub cmdChangeTabSelBackColor_Click()
    Dim iDlg As New cDlg
    
    iDlg.ShowColor
    If Not iDlg.Canceled Then
        picTabSelBackColor.BackColor = iDlg.Color
        SSTabEx1.TabSelBackColor = picTabSelBackColor.BackColor
    End If
End Sub

Private Sub cmdChangeTabSelForeColor_Click()
    Dim iDlg As New cDlg
    
    iDlg.ShowColor
    If Not iDlg.Canceled Then
        picTabSelForeColor.BackColor = iDlg.Color
        SSTabEx1.TabSelForeColor = picTabSelForeColor.BackColor
        lblThemedIDE.ForeColor = SSTabEx1.TabSelForeColor
        Option1.ForeColor = SSTabEx1.TabSelForeColor
        Option2.ForeColor = SSTabEx1.TabSelForeColor
        Check1.ForeColor = SSTabEx1.TabSelForeColor
        Label1.ForeColor = SSTabEx1.TabSelForeColor
        Label22.ForeColor = SSTabEx1.TabSelForeColor
        Option3.ForeColor = SSTabEx1.TabSelForeColor
        Option4.ForeColor = SSTabEx1.TabSelForeColor
        Check2.ForeColor = SSTabEx1.TabSelForeColor
    End If
End Sub

Private Sub cmdDisableTab_Click()
    SSTabEx1.TabEnabled(SSTabEx1.TabSel) = Not SSTabEx1.TabEnabled(SSTabEx1.TabSel)
    If SSTabEx1.TabEnabled(SSTabEx1.TabSel) Then
        cmdDisableTab.Caption = "Disable tab"
    Else
        cmdDisableTab.Caption = "Enable tab"
    End If
End Sub

Private Sub cmdHelp_Click()
    Const SW_SHOWMAXIMIZED = 3
    
    If FileExists(App.Path & "\..\docs\tabexctl_reference.html") Then
        If ShellExecute(0&, "OPEN", App.Path & "\..\docs\tabexctl_reference.html", "", "", SW_SHOWMAXIMIZED) <= 32 Then
            MsgBox "Html document could not be displayed.", vbExclamation
        End If
    Else
        MsgBox "Html file not found", vbExclamation
    End If
End Sub

Private Function FileExists(ByVal strPathName As String) As Boolean
    Dim intFileNum As Integer

    On Error Resume Next

    '
    'Attempt to open the file, return value of this function is False
    'if an error occurs on open, True otherwise
    '
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum
    
'    Debug.Print Err.Number, Err.Description
    FileExists = (Err.Number = 0) Or (Err.Number = 70)
    
    Close intFileNum

    Err.Clear
End Function

Private Sub cmdHideTab_Click()
    SSTabEx1.TabVisible(SSTabEx1.TabSel) = False
    cmdHideTab.Enabled = SSTabEx1.TabSel > -1
    cmdDisableTab.Enabled = cmdHideTab.Enabled
    cmdChangeTabPicture.Enabled = cmdHideTab.Enabled
    cmdRemoveTabPicture.Enabled = cmdHideTab.Enabled
End Sub

Private Sub cmdChangeTabPicture_Click()
    Dim iDlg As New cDlg
    
    iDlg.Filter = "Image files (*.bmp, *.ico)|*.bmp;*.ico"
    iDlg.ShowOpen
    If Not iDlg.Canceled Then
        On Error Resume Next
        Set SSTabEx1.TabPicture(SSTabEx1.TabSel) = LoadPicture(iDlg.FileName)
        txtTabHeight.Text = SSTabEx1.TabHeight
        If Err.Number <> 0 Then
            MsgBox "Error: " & Err.Number & ", " & Err.Description, vbCritical
        End If
        On Error GoTo 0
    End If
End Sub

Private Sub cmdAllTabsVisibleEnabled_Click()
    Dim t As Long
    
    For t = 0 To SSTabEx1.Tabs - 1
        SSTabEx1.TabVisible(t) = True
        SSTabEx1.TabEnabled(t) = True
    Next t
    
    cmdHideTab.Enabled = True
    cmdDisableTab.Enabled = True
    cmdChangeTabPicture.Enabled = True
    cmdRemoveTabPicture.Enabled = True
End Sub

Private Sub cmdRemoveTabPicture_Click()
    Set SSTabEx1.TabPicture(SSTabEx1.TabSel) = Nothing
    Set SSTabEx1.TabPic16(SSTabEx1.TabSel) = Nothing
    Set SSTabEx1.TabPic20(SSTabEx1.TabSel) = Nothing
    Set SSTabEx1.TabPic24(SSTabEx1.TabSel) = Nothing
    txtTabHeight.Text = SSTabEx1.TabHeight
End Sub

Private Sub Form_Load()
    Me.Move Screen.Width / 2 - Me.Width / 2, Screen.Height / 2 - Me.Height / 2
    SSTabEx1.ForceVisualStyles = InIde
    lblThemedIDE.Visible = InIde And SSTabEx1.IsVisualStyleApplied
    
    cboOrientation.Clear
    cboOrientation.AddItem ssTabOrientationTop & " - ssTabOrientationTop"
    cboOrientation.AddItem ssTabOrientationBottom & " - ssTabOrientationBottom"
    cboOrientation.AddItem ssTabOrientationLeft & " - ssTabOrientationLeft"
    cboOrientation.AddItem ssTabOrientationRight & " - ssTabOrientationRight"
    
    cboBackStyle.Clear
    cboBackStyle.AddItem ssTransparent & " - ssTransparent"
    cboBackStyle.AddItem ssOpaque & " - ssOpaque"
    
    cboStyle.Clear
    cboStyle.AddItem ssStyleTabbedDialog & " - ssStyleTabbedDialog"
    cboStyle.AddItem ssStylePropertyPage & " - ssStylePropertyPage"
    cboStyle.AddItem ssStyleTabStrip & " - ssStyleTabStrip"
    
    cboShowRowsInPerspective.Clear
    cboShowRowsInPerspective.AddItem ssNo & " - No"
    cboShowRowsInPerspective.AddItem ssYes & " - Yes"
    cboShowRowsInPerspective.AddItem ssYNAuto & " - Automatic"
    
    cboTabWidthStyle.Clear
    cboTabWidthStyle.AddItem ssTWSJustified & " - ssTWSJustified"
    cboTabWidthStyle.AddItem ssTWSNonJustified & " - ssTWSNonJustified"
    cboTabWidthStyle.AddItem ssTWSFixed & " - ssTWSFixed"
    cboTabWidthStyle.AddItem ssTWSAuto & " - Automatic"
    
    cboTabAppearance.Clear
    cboTabAppearance.AddItem ssTAAuto & " - Automatic"
    cboTabAppearance.AddItem ssTATabbedDialog & " - ssTATabbedDialog"
    cboTabAppearance.AddItem ssTATabbedDialogRounded & " - ssTATabbedDialogRounded"
    cboTabAppearance.AddItem ssTAPropertyPage & " - ssTAPropertyPage"
    cboTabAppearance.AddItem ssTAPropertyPageRounded & " - ssTAPropertyPageRounded"
    
    cboTabSelFontBold.Clear
    cboTabSelFontBold.AddItem ssNo & " - No"
    cboTabSelFontBold.AddItem ssYes & " - Yes"
    cboTabSelFontBold.AddItem ssYNAuto & " - Automatic"
    
    cboTabHoverHighlight.Clear
    cboTabHoverHighlight.AddItem ssTHHNo & " - ssTHHNo"
    cboTabHoverHighlight.AddItem ssTHHInstant & " - ssTHHInstant"
    cboTabHoverHighlight.AddItem ssTHHEffect & " - ssTHHEffect"
    
    cboTabPictureAlignment.Clear
    cboTabPictureAlignment.AddItem "ssPicAlignBeforeCaption"
    cboTabPictureAlignment.AddItem "ssPicAlignCenteredBeforeCaption"
    cboTabPictureAlignment.AddItem "ssPicAlignAfterCaption"
    cboTabPictureAlignment.AddItem "ssPicAlignCenteredAfterCaption"
    
    cboOrientation.ListIndex = SSTabEx1.TabOrientation
    chkVisualStyles.Value = Abs(SSTabEx1.VisualStyles)
    cboBackStyle.ListIndex = SSTabEx1.BackStyle
    txtTabsPerRow.Text = SSTabEx1.TabsPerRow
    cboStyle.ListIndex = SSTabEx1.Style
    cboTabAppearance.ListIndex = SSTabEx1.TabAppearance
    cboTabWidthStyle.ListIndex = SSTabEx1.TabWidthStyle
    cboShowRowsInPerspective.ListIndex = SSTabEx1.ShowRowsInPerspective
    cboTabSelFontBold.ListIndex = SSTabEx1.TabSelFontBold
    txtTabSeparation.Text = SSTabEx1.TabSeparation
    txtTabSelExtraHeight.Text = SSTabEx1.TabSelExtraHeight
    txtTabMaxWidth.Text = SSTabEx1.TabMaxWidth
    txtTabMinWidth.Text = SSTabEx1.TabMinWidth
    chkEnabled.Value = Abs(SSTabEx1.Enabled)
    chkTabSelHighlight.Value = Abs(SSTabEx1.TabSelHighlight)
    chkShowFocusRect.Value = Abs(SSTabEx1.ShowFocusRect)
    chkChangeControlsBackColor.Value = Abs(SSTabEx1.ChangeControlsBackColor)
    chkWordWrap.Value = Abs(SSTabEx1.WordWrap)
    chkUseMaskColor.Value = Abs(SSTabEx1.UseMaskColor)
    chkAutoTabHeight.Value = Abs(SSTabEx1.AutoTabHeight)
    chkShowDisabledState.Value = Abs(SSTabEx1.ShowDisabledState)
    chkSoftEdges.Value = Abs(SSTabEx1.SoftEdges)
    picMaskColor.BackColor = SSTabEx1.MaskColor
    picTabBackColor.BackColor = SSTabEx1.TabBackColor
    picTabSelBackColor.BackColor = SSTabEx1.TabSelBackColor
    picForeColor.BackColor = SSTabEx1.ForeColor
    picTabSelForeColor.BackColor = SSTabEx1.TabSelForeColor
    picBackColor.BackColor = SSTabEx1.BackColor
    ShowFont
    txtTabs.Text = SSTabEx1.Tabs
    txtTabHeight.Text = SSTabEx1.TabHeight
    cboTabHoverHighlight.ListIndex = SSTabEx1.TabHoverHighlight
    cboTabPictureAlignment.ListIndex = SSTabEx1.TabPictureAlignment
    
    txtTabCaption.Text = SSTabEx1.TabCaption(SSTabEx1.TabSel)
    txtTabToolTipText.Text = SSTabEx1.TabToolTipText(SSTabEx1.TabSel)
    
    cmdHelp.Visible = FileExists(App.Path & "\..\docs\tabexctl_reference.html")
End Sub

Private Sub picBackColor_DblClick()
    cmdChangeBackColor_Click
End Sub

Private Sub picForeColor_DblClick()
    cmdChangeForeColor_Click
End Sub

Private Sub picMaskColor_DblClick()
    cmdChangeMaskColor_Click
End Sub

Private Sub picTabBackColor_DblClick()
    cmdChangeTabBackColor_Click
End Sub

Private Sub picTabSelBackColor_DblClick()
    cmdChangeTabSelBackColor_Click
End Sub

Private Sub picTabSelForeColor_DblClick()
    cmdChangeTabSelForeColor_Click
End Sub

Private Sub SSTabEx1_Click(PreviousTab As Integer)
    txtTabCaption.Text = SSTabEx1.TabCaption(SSTabEx1.TabSel)
    txtTabToolTipText.Text = SSTabEx1.TabToolTipText(SSTabEx1.TabSel)
    If SSTabEx1.TabEnabled(SSTabEx1.TabSel) Then
        cmdDisableTab.Caption = "Disable tab"
    Else
        cmdDisableTab.Caption = "Enable tab"
    End If
End Sub

Private Sub Timer1_Timer()
    lblFocus.Caption = "Focus: " & Me.ActiveControl.Name
End Sub

Private Sub tmrUpdate_Timer()
    tmrUpdate.Enabled = False
    SSTabEx1.TabMaxWidth = Val(txtTabMaxWidth.Text)
    SSTabEx1.TabMinWidth = Val(txtTabMinWidth.Text)
    
    txtTabMaxWidth.Text = SSTabEx1.TabMaxWidth
    txtTabMinWidth.Text = SSTabEx1.TabMinWidth

    UpdateTabCount
End Sub

Private Sub UpdateTabCount()
    Dim iLng As Long
    
    iLng = Val(txtTabs.Text)
    If iLng < SSTabEx1.Tabs Then
        If MsgBox("It will erase the tab data of the last " & SSTabEx1.Tabs - iLng & " tabs, proceed?", vbYesNo + vbExclamation) = vbNo Then
            txtTabs.Text = SSTabEx1.Tabs
            txtTabs.SelStart = Len(txtTabs.Text)
            Exit Sub
        End If
    End If
    On Error GoTo ErrH
    SSTabEx1.Tabs = iLng
    Exit Sub

ErrH:
    MsgBox Err.Number & " " & Err.Description, vbCritical
    txtTabs.Text = SSTabEx1.Tabs
    txtTabs.SelStart = Len(txtTabs.Text)
End Sub

Private Sub txtFont_DblClick()
    cmdChangeFont_Click
End Sub

Private Sub txtTabCaption_Change()
    SSTabEx1.TabCaption(SSTabEx1.TabSel) = txtTabCaption.Text
End Sub

Private Sub txtTabHeight_Change()
    SSTabEx1.TabHeight = Val(txtTabHeight.Text)
    chkAutoTabHeight.Value = Abs(SSTabEx1.AutoTabHeight)
End Sub

Private Sub txtTabMaxWidth_Change()
    tmrUpdate.Enabled = True
End Sub

Private Sub txtTabMaxWidth_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tmrUpdate.Enabled Then
            tmrUpdate_Timer
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtTabMinWidth_Change()
    tmrUpdate.Enabled = True
End Sub

Private Sub txtTabMinWidth_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tmrUpdate.Enabled Then
            tmrUpdate_Timer
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtTabs_Change()
    tmrUpdate.Enabled = True
End Sub

Private Sub txtTabs_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tmrUpdate.Enabled Then
            tmrUpdate_Timer
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtTabSelExtraHeight_Change()
    SSTabEx1.TabSelExtraHeight = Val(txtTabSelExtraHeight.Text)
    If SSTabEx1.TabSelExtraHeight <> Val(txtTabSelExtraHeight.Text) Then
        txtTabSelExtraHeight.Text = SSTabEx1.TabSelExtraHeight ' there is a limit, so get the actual value that it took
        txtTabSelExtraHeight.SelStart = Len(txtTabSelExtraHeight.Text)
    End If
End Sub

Private Sub txtTabSeparation_Change()
    SSTabEx1.TabSeparation = Val(txtTabSeparation.Text)
End Sub

Private Sub txtTabsPerRow_Change()
    SSTabEx1.TabsPerRow = Val(txtTabsPerRow.Text)
End Sub

Private Function InIde() As Boolean
    Static sValue As Long
    
    If sValue = 0 Then
        On Error Resume Next
        Debug.Print 1 / 0
        If Err.Number Then
            sValue = 1
        Else
            sValue = 2
        End If
    End If
    InIde = (sValue = 1)
End Function
    
Private Sub txtTabToolTipText_Change()
    SSTabEx1.TabToolTipText(SSTabEx1.TabSel) = txtTabToolTipText.Text
End Sub
