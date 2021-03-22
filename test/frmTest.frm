VERSION 5.00
Object = "*\ATabExCtl.vbp"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test SSTabEx"
   ClientHeight    =   7020
   ClientLeft      =   1176
   ClientTop       =   720
   ClientWidth     =   12156
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   7020
   ScaleWidth      =   12156
   Begin VB.CommandButton cmdChangeTabSelForeColor 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   11685
      TabIndex        =   93
      Top             =   5256
      Width           =   330
   End
   Begin VB.PictureBox picTabSelForeColor 
      Height          =   300
      Left            =   11070
      ScaleHeight     =   252
      ScaleWidth      =   540
      TabIndex        =   92
      Top             =   5256
      Width           =   588
   End
   Begin VB.PictureBox picTabSelBackColor 
      Height          =   300
      Left            =   6780
      ScaleHeight     =   252
      ScaleWidth      =   540
      TabIndex        =   90
      Top             =   5208
      Width           =   588
   End
   Begin VB.CommandButton cmdChangeTabSelBackColor 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   7440
      TabIndex        =   89
      Top             =   5220
      Width           =   330
   End
   Begin VB.ComboBox cboTabHoverHighlight 
      Height          =   300
      ItemData        =   "frmTest.frx":0000
      Left            =   6780
      List            =   "frmTest.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   77
      Top             =   5580
      Width           =   2700
   End
   Begin VB.CheckBox chkShowFocusRect 
      Caption         =   "ShowFocusRect"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   9840
      TabIndex        =   57
      Top             =   3564
      Width           =   1740
   End
   Begin VB.CheckBox chkSoftEdges 
      Caption         =   "SoftEdges"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   9840
      TabIndex        =   54
      Top             =   3276
      Width           =   1740
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   5070
      Top             =   3900
   End
   Begin VB.ComboBox cboTabPictureAlignment 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmTest.frx":0004
      Left            =   6780
      List            =   "frmTest.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   79
      Top             =   5940
      Width           =   2712
   End
   Begin VB.CommandButton cmdChangeBackColor 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   11685
      TabIndex        =   86
      Top             =   4524
      Width           =   330
   End
   Begin VB.PictureBox picBackColor 
      Height          =   300
      Left            =   11070
      ScaleHeight     =   252
      ScaleWidth      =   540
      TabIndex        =   85
      Top             =   4524
      Width           =   588
   End
   Begin VB.CommandButton cmdAllTabsVisibleEnabled 
      Caption         =   "All tabs visible and enabled"
      Height          =   336
      Left            =   2916
      TabIndex        =   27
      Top             =   6516
      Width           =   2400
   End
   Begin VB.PictureBox picMaskColor 
      Height          =   300
      Left            =   6780
      ScaleHeight     =   252
      ScaleWidth      =   540
      TabIndex        =   64
      Top             =   4500
      Width           =   588
   End
   Begin VB.CommandButton cmdChangeMaskColor 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   7440
      TabIndex        =   65
      Top             =   4512
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
         Picture         =   "frmTest.frx":0008
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
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1584
         TabIndex        =   21
         Top             =   684
         Width           =   2820
      End
      Begin VB.TextBox txtTabCaption 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1584
         TabIndex        =   19
         Top             =   324
         Width           =   2820
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "TabToolTipText:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   144
         TabIndex        =   20
         Top             =   720
         Width           =   1344
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "TabCaption:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      TabIndex        =   50
      Top             =   3060
      Width           =   588
   End
   Begin VB.TextBox txtTabs 
      Height          =   300
      Left            =   8940
      MaxLength       =   3
      TabIndex        =   33
      Top             =   468
      Width           =   588
   End
   Begin VB.CommandButton cmdChangeFont 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   11685
      TabIndex        =   75
      Top             =   4896
      Width           =   330
   End
   Begin VB.TextBox txtFont 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8940
      Locked          =   -1  'True
      TabIndex        =   74
      Top             =   4884
      Width           =   2724
   End
   Begin VB.CommandButton cmdTestSBS 
      Caption         =   "Test side by side with SSTab"
      Height          =   336
      Left            =   7860
      TabIndex        =   81
      Top             =   6516
      Width           =   2544
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   336
      Left            =   6780
      TabIndex        =   80
      Top             =   6516
      Width           =   1000
   End
   Begin VB.PictureBox picForeColor 
      Height          =   300
      Left            =   8940
      ScaleHeight     =   252
      ScaleWidth      =   540
      TabIndex        =   71
      Top             =   4524
      Width           =   588
   End
   Begin VB.CommandButton cmdChangeForeColor 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   9555
      TabIndex        =   72
      Top             =   4524
      Width           =   330
   End
   Begin VB.CommandButton cmdChangeTabBackColor 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   7440
      TabIndex        =   69
      Top             =   4872
      Width           =   330
   End
   Begin VB.PictureBox picTabBackColor 
      Height          =   300
      Left            =   6780
      ScaleHeight     =   252
      ScaleWidth      =   540
      TabIndex        =   68
      Top             =   4860
      Width           =   588
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2952
      Top             =   6516
   End
   Begin VB.ComboBox cboTabAppearance 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmTest.frx":0152
      Left            =   8940
      List            =   "frmTest.frx":0154
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   1224
      Width           =   2712
   End
   Begin VB.ComboBox cboShowRowsInPerspective 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmTest.frx":0156
      Left            =   8940
      List            =   "frmTest.frx":0158
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   1944
      Width           =   2712
   End
   Begin VB.TextBox txtTabMinWidth 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6804
      TabIndex        =   61
      Top             =   4140
      Width           =   588
   End
   Begin VB.TextBox txtTabMaxWidth 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6780
      TabIndex        =   59
      Top             =   3780
      Width           =   588
   End
   Begin VB.CheckBox chkShowDisabledState 
      Caption         =   "ShowDisabledState"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   9840
      TabIndex        =   48
      Top             =   2700
      Width           =   2028
   End
   Begin VB.CheckBox chkChangeControlsBackColor 
      Caption         =   "ChangeControlsBackColor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   8010
      TabIndex        =   62
      Top             =   3852
      Width           =   2892
   End
   Begin VB.CheckBox chkTabSelHighlight 
      Caption         =   "TabSelHighlight"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   8010
      TabIndex        =   53
      Top             =   3276
      Width           =   1740
   End
   Begin VB.CheckBox chkUseMaskColor 
      Caption         =   "UseMaskColor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   8010
      TabIndex        =   66
      Top             =   3564
      Width           =   1740
   End
   Begin VB.ComboBox cboTabSelFontBold 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmTest.frx":015A
      Left            =   8940
      List            =   "frmTest.frx":015C
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   2304
      Width           =   2712
   End
   Begin VB.TextBox txtTabSelExtraHeight 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6780
      TabIndex        =   56
      Top             =   3420
      Width           =   588
   End
   Begin VB.TextBox txtTabSeparation 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6780
      MaxLength       =   2
      TabIndex        =   46
      Top             =   2700
      Width           =   588
   End
   Begin VB.CheckBox chkEnabled 
      Caption         =   "Enabled"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   8010
      TabIndex        =   47
      Top             =   2700
      Width           =   1740
   End
   Begin VB.CheckBox chkWordWrap 
      Caption         =   "WordWrap"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   9840
      TabIndex        =   52
      Top             =   2988
      Width           =   1740
   End
   Begin VB.CheckBox chkVisualStyles 
      Caption         =   "VisualStyles"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   8010
      TabIndex        =   51
      Top             =   2988
      Width           =   1740
   End
   Begin VB.ComboBox cboTabWidthStyle 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmTest.frx":015E
      Left            =   8940
      List            =   "frmTest.frx":0160
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   1584
      Width           =   2712
   End
   Begin VB.ComboBox cboStyle 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmTest.frx":0162
      Left            =   6780
      List            =   "frmTest.frx":0164
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   828
      Width           =   4872
   End
   Begin VB.TextBox txtTabsPerRow 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6780
      MaxLength       =   2
      TabIndex        =   31
      Top             =   468
      Width           =   588
   End
   Begin VB.ComboBox cboOrientation 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmTest.frx":0166
      Left            =   6780
      List            =   "frmTest.frx":0168
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
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabsPerRow      =   4
      Tab             =   1
      TabHeight       =   582
      Themed          =   -1  'True
      TabPic16(0)     =   "frmTest.frx":016A
      TabPic20(0)     =   "frmTest.frx":04BC
      TabPic24(0)     =   "frmTest.frx":09BE
      TabCaption(0)   =   "Theme"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Picture2(0)"
      TabPic16(1)     =   "frmTest.frx":10D0
      TabPic20(1)     =   "frmTest.frx":1422
      TabPic24(1)     =   "frmTest.frx":1924
      TabCaption(1)   =   "Frame"
      Tab(1).ControlCount=   1
      Tab(1).Control(0)=   "Frame1"
      TabPic16(2)     =   "frmTest.frx":2036
      TabPic20(2)     =   "frmTest.frx":2388
      TabPic24(2)     =   "frmTest.frx":288A
      TabCaption(2)   =   "Other"
      Tab(2).ControlCount=   1
      Tab(2).Control(0)=   "Picture1"
      TabPic16(3)     =   "frmTest.frx":2F9C
      TabPic20(3)     =   "frmTest.frx":32EE
      TabPic24(3)     =   "frmTest.frx":37F0
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
         TabIndex        =   88
         Top             =   216
         Width           =   12
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1596
         Index           =   0
         Left            =   -74820
         ScaleHeight     =   1596
         ScaleWidth      =   3720
         TabIndex        =   83
         Top             =   936
         Width           =   3720
         Begin VB.Label lblThemedIDE 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "The SSTabEx is themed now in the IDE just for testing, normally it won't be themed in the IDE"
            Height          =   984
            Left            =   468
            TabIndex        =   84
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
         Top             =   2700
         Width           =   1900
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   480
         Left            =   -74568
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3132
         Width           =   1900
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   300
         Left            =   -74568
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2340
         Value           =   -1  'True
         Width           =   1900
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command graphical"
         Height          =   588
         Left            =   -74568
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1620
         Width           =   1900
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command standard"
         Height          =   588
         Left            =   -74568
         TabIndex        =   10
         Top             =   936
         Width           =   1900
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2928
         Left            =   -74784
         ScaleHeight     =   2928
         ScaleWidth      =   3720
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   828
         Width           =   3720
         Begin VB.ListBox List1 
            Height          =   660
            Left            =   396
            TabIndex        =   8
            Top             =   1224
            Width           =   1848
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            ItemData        =   "frmTest.frx":3F02
            Left            =   396
            List            =   "frmTest.frx":3F0F
            TabIndex        =   7
            Text            =   "1234567890"
            Top             =   648
            Width           =   1164
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "These controls are inside a PictureBox"
            Height          =   204
            Left            =   432
            TabIndex        =   9
            Top             =   2340
            Width           =   3012
         End
         Begin VB.Label Label1 
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
         Top             =   828
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
         Caption         =   $"frmTest.frx":3F3D
         Height          =   1164
         Left            =   -74280
         TabIndex        =   16
         Top             =   1188
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
         Top             =   1116
         Width           =   1380
      End
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      Caption         =   "TabSelForeColor:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   9600
      TabIndex        =   94
      Top             =   5292
      Width           =   1368
   End
   Begin VB.Label lblTabSelBackColor 
      Alignment       =   1  'Right Justify
      Caption         =   "TabSelBackColor:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   5208
      TabIndex        =   91
      Top             =   5268
      Width           =   1488
   End
   Begin VB.Label lblTabHoverHighlight 
      Alignment       =   1  'Right Justify
      Caption         =   "TabHoverHighlight:"
      Height          =   228
      Left            =   5088
      TabIndex        =   76
      Top             =   5616
      Width           =   1596
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      Caption         =   "TabPictureAlignment:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4635
      TabIndex        =   78
      Top             =   5970
      Width           =   2055
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      Caption         =   "BackColor:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   9900
      TabIndex        =   87
      Top             =   4560
      Width           =   1032
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      Caption         =   "MaskColor:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5205
      TabIndex        =   63
      Top             =   4530
      Width           =   1485
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "TabHeight:"
      Height          =   225
      Index           =   2
      Left            =   5670
      TabIndex        =   49
      Top             =   3090
      Width           =   1020
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Tab Count:"
      Height          =   228
      Left            =   7800
      TabIndex        =   32
      Top             =   516
      Width           =   1056
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Caption         =   "Font:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   7848
      TabIndex        =   73
      Top             =   4920
      Width           =   984
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Caption         =   "ForeColor:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   7812
      TabIndex        =   70
      Top             =   4560
      Width           =   1008
   End
   Begin VB.Label lblTabBackColor 
      Alignment       =   1  'Right Justify
      Caption         =   "TabBackColor:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   5208
      TabIndex        =   67
      Top             =   4920
      Width           =   1488
   End
   Begin VB.Label lblFocus 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   264
      Left            =   324
      TabIndex        =   26
      Top             =   6552
      Width           =   2568
   End
   Begin VB.Label lblTabAppearance 
      Alignment       =   1  'Right Justify
      Caption         =   "TabAppearance:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7035
      TabIndex        =   37
      Top             =   1260
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "ShowRowsInPerspective:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6555
      TabIndex        =   41
      Top             =   1980
      Width           =   2295
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "TabMinWidth:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5205
      TabIndex        =   60
      Top             =   4170
      Width           =   1485
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "TabMaxWidth:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5205
      TabIndex        =   58
      Top             =   3810
      Width           =   1485
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "TabSelFontBold:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7035
      TabIndex        =   43
      Top             =   2340
      Width           =   1815
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "TabSelExtraHeight:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5025
      TabIndex        =   55
      Top             =   3450
      Width           =   1665
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "TabSeparation:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5205
      TabIndex        =   45
      Top             =   2730
      Width           =   1485
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "TabWidthStyle:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7035
      TabIndex        =   39
      Top             =   1620
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Customize Style:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5100
      TabIndex        =   36
      Top             =   1260
      Width           =   1590
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Style:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5205
      TabIndex        =   34
      Top             =   870
      Width           =   1485
   End
   Begin VB.Label lblTabsPerRow 
      Alignment       =   1  'Right Justify
      Caption         =   "TabsPerRow:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5205
      TabIndex        =   30
      Top             =   510
      Width           =   1485
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Orientation:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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

Private Sub cboOrientation_Click()
    SSTabEx1.TabOrientation = cboOrientation.ListIndex
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
    If Not SSTabEx1.IsVisualStyleApplied Then
        chkVisualStyles.Value = 0
    End If
    chkTabSelHighlight.Enabled = Not SSTabEx1.VisualStyles
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
    Dim iDlg As New CDLg
    
    iDlg.ShowColor
    If Not iDlg.Canceled Then
        picBackColor.BackColor = iDlg.Color
        SSTabEx1.BackColor = picBackColor.BackColor
    End If

End Sub

Private Sub cmdChangeFont_Click()
    Dim iDlg As New CDLg
    
    Set iDlg.Font = SSTabEx1.Font
    iDlg.ShowFont
    If Not iDlg.Canceled Then
        SSTabEx1.Font = iDlg.Font
        ShowFont
    End If

End Sub

Private Sub ShowFont()
    txtFont.Text = SSTabEx1.Font.Name & "  " & SSTabEx1.Font.Size
End Sub

Private Sub cmdChangeForeColor_Click()
    Dim iDlg As New CDLg
    
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
    Dim iDlg As New CDLg
    
    iDlg.ShowColor
    If Not iDlg.Canceled Then
        picMaskColor.BackColor = iDlg.Color
        SSTabEx1.MaskColor = picMaskColor.BackColor
    End If

End Sub

Private Sub cmdChangeTabBackColor_Click()
    Dim iDlg As New CDLg
    
    iDlg.ShowColor
    If Not iDlg.Canceled Then
        picTabBackColor.BackColor = iDlg.Color
        SSTabEx1.TabBackColor = picTabBackColor.BackColor
        picTabSelBackColor.BackColor = SSTabEx1.TabSelBackColor
    End If
End Sub

Private Sub cmdChangeTabSelBackColor_Click()
    Dim iDlg As New CDLg
    
    iDlg.ShowColor
    If Not iDlg.Canceled Then
        picTabSelBackColor.BackColor = iDlg.Color
        SSTabEx1.TabSelBackColor = picTabSelBackColor.BackColor
    End If
End Sub

Private Sub cmdChangeTabSelForeColor_Click()
    Dim iDlg As New CDLg
    
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
    frmHelp.Show
End Sub

Private Sub cmdHideTab_Click()
    SSTabEx1.TabVisible(SSTabEx1.TabSel) = False
    cmdHideTab.Enabled = SSTabEx1.TabSel > -1
    cmdDisableTab.Enabled = cmdHideTab.Enabled
    cmdChangeTabPicture.Enabled = cmdHideTab.Enabled
    cmdRemoveTabPicture.Enabled = cmdHideTab.Enabled
End Sub

Private Sub cmdChangeTabPicture_Click()
    Dim iDlg As New CDLg
    
    iDlg.Filter = "Image files (*.bmp, *.ico)|*.bmp;*.ico"
    iDlg.ShowOpen
    If Not iDlg.Canceled Then
        On Error Resume Next
        Set SSTabEx1.TabPicture(SSTabEx1.TabSel) = LoadPicture(iDlg.FileName)
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
End Sub

Private Sub cmdTestSBS_Click()
    frmTestSBS.Show 1
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
    chkVisualStyles.Value = Abs(SSTabEx1.VisualStyles)
    chkEnabled.Value = Abs(SSTabEx1.Enabled)
    chkTabSelHighlight.Value = Abs(SSTabEx1.TabSelHighlight)
    chkShowFocusRect.Value = Abs(SSTabEx1.ShowFocusRect)
    chkChangeControlsBackColor.Value = Abs(SSTabEx1.ChangeControlsBackColor)
    chkWordWrap.Value = Abs(SSTabEx1.WordWrap)
    chkUseMaskColor.Value = Abs(SSTabEx1.UseMaskColor)
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
