VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "*\ATabExCtl.vbp"
Begin VB.Form frmTestSBS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test side by side with SSTab"
   ClientHeight    =   7020
   ClientLeft      =   1740
   ClientTop       =   1044
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4116
      Left            =   5184
      TabIndex        =   20
      Top             =   108
      Width           =   4644
      _ExtentX        =   8192
      _ExtentY        =   7260
      _Version        =   393216
      Tabs            =   7
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Theme"
      TabPicture(0)   =   "frmTestSBS.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture2(1)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Frame"
      TabPicture(1)   =   "frmTestSBS.frx":13E2
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Other"
      TabPicture(2)   =   "frmTestSBS.frx":27C4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Cmd"
      TabPicture(3)   =   "frmTestSBS.frx":3BA6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label10"
      Tab(3).Control(1)=   "Option7"
      Tab(3).Control(2)=   "Check4"
      Tab(3).Control(3)=   "Option8"
      Tab(3).Control(4)=   "Command3"
      Tab(3).Control(5)=   "Command4"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Label"
      TabPicture(4)   =   "frmTestSBS.frx":4F88
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label11"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Tab 5"
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Tab 6"
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      Begin VB.CommandButton Command4 
         Caption         =   "Command standard"
         Height          =   588
         Left            =   -74568
         TabIndex        =   32
         Top             =   936
         Width           =   1900
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command graphical"
         Height          =   588
         Left            =   -74568
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1620
         Width           =   1900
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Option3"
         Height          =   300
         Left            =   -74568
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2340
         Value           =   -1  'True
         Width           =   1900
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check2"
         Height          =   480
         Left            =   -74568
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   3132
         Width           =   1900
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Option4"
         Height          =   300
         Left            =   -74568
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2700
         Width           =   1900
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   2928
         Left            =   -74784
         ScaleHeight     =   2928
         ScaleWidth      =   3720
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   828
         Width           =   3720
         Begin VB.ComboBox Combo2 
            Height          =   300
            ItemData        =   "frmTestSBS.frx":4FA4
            Left            =   396
            List            =   "frmTestSBS.frx":4FB1
            TabIndex        =   29
            Text            =   "1234567890"
            Top             =   648
            Width           =   1164
         End
         Begin VB.ListBox List2 
            Height          =   660
            Left            =   396
            TabIndex        =   30
            Top             =   1224
            Width           =   1848
         End
         Begin VB.Label Label9 
            Caption         =   "Just a Label"
            Height          =   372
            Left            =   396
            TabIndex        =   28
            Top             =   180
            Width           =   2352
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "These controls are inside a PictureBox"
            Height          =   204
            Left            =   432
            TabIndex        =   31
            Top             =   2340
            Width           =   3012
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame1"
         Height          =   2784
         Left            =   360
         TabIndex        =   22
         Top             =   828
         Width           =   2640
         Begin VB.OptionButton Option6 
            Caption         =   "Option1"
            Height          =   300
            Left            =   324
            TabIndex        =   23
            Top             =   324
            Width           =   1488
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Option2"
            Height          =   300
            Left            =   324
            TabIndex        =   24
            Top             =   756
            Value           =   -1  'True
            Width           =   1164
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Check1"
            Height          =   408
            Left            =   360
            TabIndex        =   25
            Top             =   1188
            Width           =   1344
         End
         Begin VB.TextBox Text2 
            Height          =   336
            Left            =   360
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   1728
            Width           =   1056
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1596
         Index           =   1
         Left            =   -74820
         ScaleHeight     =   1596
         ScaleWidth      =   3720
         TabIndex        =   21
         Top             =   936
         Width           =   3720
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "This label is placed directly into the SSTab."
         Height          =   1164
         Left            =   -74280
         TabIndex        =   38
         Top             =   1188
         Width           =   2784
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Yeah..."
         Height          =   1524
         Left            =   -72516
         TabIndex        =   34
         Top             =   1116
         Width           =   1380
      End
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   6120
      Top             =   5868
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
      Left            =   9540
      TabIndex        =   73
      Top             =   6372
      Width           =   330
   End
   Begin VB.PictureBox picBackColor 
      Height          =   300
      Left            =   8928
      ScaleHeight     =   252
      ScaleWidth      =   540
      TabIndex        =   72
      Top             =   6372
      Width           =   588
   End
   Begin VB.CommandButton cmdAllTabsVisibleEnabled 
      Caption         =   "All tabs visible and enabled"
      Height          =   336
      Left            =   1260
      TabIndex        =   48
      Top             =   6156
      Width           =   2400
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tab specific data (current tab):"
      Height          =   1236
      Left            =   288
      TabIndex        =   41
      Top             =   4824
      Width           =   4584
      Begin VB.CommandButton cmdRemoveTabPicture 
         Height          =   336
         Left            =   4140
         Picture         =   "frmTestSBS.frx":4FDF
         Style           =   1  'Graphical
         TabIndex        =   47
         Tag             =   "E"
         ToolTipText     =   "Remove Tab"
         Top             =   720
         Width           =   275
      End
      Begin VB.CommandButton cmdDisableTab 
         Caption         =   "Disable tab"
         Height          =   336
         Left            =   1548
         TabIndex        =   45
         Top             =   720
         Width           =   1176
      End
      Begin VB.CommandButton cmdHideTab 
         Caption         =   "Hide tab"
         Height          =   336
         Left            =   180
         TabIndex        =   44
         Top             =   720
         Width           =   1176
      End
      Begin VB.CommandButton cmdChangeTabPicture 
         Caption         =   "Load picture"
         Height          =   336
         Left            =   2916
         TabIndex        =   46
         Top             =   720
         Width           =   1176
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
         TabIndex        =   43
         Top             =   324
         Width           =   2820
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
         TabIndex        =   42
         Top             =   360
         Width           =   1344
      End
   End
   Begin VB.TextBox txtTabHeight 
      Height          =   300
      Left            =   8460
      TabIndex        =   57
      Top             =   5040
      Width           =   588
   End
   Begin VB.TextBox txtTabs 
      Height          =   300
      Left            =   10908
      MaxLength       =   3
      TabIndex        =   64
      Top             =   5364
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
      Left            =   9504
      TabIndex        =   62
      Top             =   5400
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
      Left            =   6624
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   5400
      Width           =   2820
   End
   Begin VB.PictureBox picForeColor 
      Height          =   300
      Left            =   6624
      ScaleHeight     =   252
      ScaleWidth      =   540
      TabIndex        =   69
      Top             =   6372
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
      Left            =   7236
      TabIndex        =   70
      Top             =   6372
      Width           =   330
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2952
      Top             =   6516
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
      Left            =   10908
      TabIndex        =   59
      Top             =   5040
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
      Left            =   6624
      TabIndex        =   65
      Top             =   5760
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
      Left            =   6624
      TabIndex        =   67
      Top             =   6048
      Width           =   1740
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
      Left            =   8928
      TabIndex        =   66
      Top             =   5724
      Width           =   1950
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
      Left            =   288
      TabIndex        =   40
      Top             =   4392
      Width           =   1740
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
      Height          =   312
      ItemData        =   "frmTestSBS.frx":5129
      Left            =   6624
      List            =   "frmTestSBS.frx":512B
      Style           =   2  'Dropdown List
      TabIndex        =   53
      Top             =   4680
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
      Left            =   6624
      MaxLength       =   2
      TabIndex        =   55
      Top             =   5040
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
      Height          =   312
      ItemData        =   "frmTestSBS.frx":512D
      Left            =   6624
      List            =   "frmTestSBS.frx":512F
      Style           =   2  'Dropdown List
      TabIndex        =   51
      Top             =   4320
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
      TabPicture(0)   =   "frmTestSBS.frx":5131
      TabCaption(0)   =   "Theme"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Picture2(0)"
      TabPicture(1)   =   "frmTestSBS.frx":6513
      TabCaption(1)   =   "Frame"
      Tab(1).ControlCount=   1
      Tab(1).Control(0)=   "Frame1"
      TabPicture(2)   =   "frmTestSBS.frx":78F5
      TabCaption(2)   =   "Other"
      Tab(2).ControlCount=   1
      Tab(2).Control(0)=   "Picture1"
      TabPicture(3)   =   "frmTestSBS.frx":8CD7
      TabCaption(3)   =   "Cmd"
      Tab(3).ControlCount=   6
      Tab(3).Control(0)=   "Option4"
      Tab(3).Control(1)=   "Check2"
      Tab(3).Control(2)=   "Option3"
      Tab(3).Control(3)=   "Command2"
      Tab(3).Control(4)=   "Command1"
      Tab(3).Control(5)=   "Label8"
      TabCaption(4)   =   "Label"
      Tab(4).ControlCount=   1
      Tab(4).Control(0)=   "Label22"
      TabCaption(5)   =   "Tab 5"
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Tab 6"
      Tab(6).ControlCount=   0
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1596
         Index           =   0
         Left            =   -74820
         ScaleHeight     =   1596
         ScaleWidth      =   3720
         TabIndex        =   1
         Top             =   936
         Width           =   3720
         Begin VB.Label lblThemedIDE 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "The SSTabEx is themed now in the IDE just for testing, normally it won't be themed in the IDE"
            Height          =   984
            Left            =   468
            TabIndex        =   2
            Top             =   216
            Width           =   2784
         End
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Option4"
         Height          =   300
         Left            =   -74568
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2700
         Width           =   1900
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   480
         Left            =   -74568
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3132
         Width           =   1900
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   300
         Left            =   -74568
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2340
         Value           =   -1  'True
         Width           =   1900
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command graphical"
         Height          =   588
         Left            =   -74568
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1620
         Width           =   1900
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command standard"
         Height          =   588
         Left            =   -74568
         TabIndex        =   13
         Top             =   936
         Width           =   1900
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2928
         Left            =   -74784
         ScaleHeight     =   2928
         ScaleWidth      =   3720
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   828
         Width           =   3720
         Begin VB.ListBox List1 
            Height          =   660
            Left            =   396
            TabIndex        =   11
            Top             =   1224
            Width           =   1848
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            ItemData        =   "frmTestSBS.frx":A0B9
            Left            =   396
            List            =   "frmTestSBS.frx":A0C6
            TabIndex        =   10
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
            TabIndex        =   12
            Top             =   2340
            Width           =   3012
         End
         Begin VB.Label Label1 
            Caption         =   "Just a Label"
            Height          =   372
            Left            =   396
            TabIndex        =   9
            Top             =   180
            Width           =   2352
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   2784
         Left            =   360
         TabIndex        =   3
         Top             =   828
         Width           =   2640
         Begin VB.TextBox Text1 
            Height          =   336
            Left            =   360
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   1728
            Width           =   1056
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   408
            Left            =   360
            TabIndex        =   6
            Top             =   1188
            Width           =   1344
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Option2"
            Height          =   300
            Left            =   324
            TabIndex        =   5
            Top             =   756
            Value           =   -1  'True
            Width           =   1164
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   300
            Left            =   324
            TabIndex        =   4
            Top             =   324
            Width           =   1488
         End
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTestSBS.frx":A0F4
         Height          =   1164
         Left            =   -74280
         TabIndex        =   19
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "The right is the SSTab"
      ForeColor       =   &H00FF0000&
      Height          =   480
      Left            =   10224
      TabIndex        =   39
      Top             =   432
      Width           =   1344
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
      Left            =   7884
      TabIndex        =   71
      Top             =   6408
      Width           =   912
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "TabHeight:"
      Height          =   228
      Index           =   2
      Left            =   7344
      TabIndex        =   56
      Top             =   5076
      Width           =   1020
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Tab Count:"
      Height          =   228
      Left            =   9324
      TabIndex        =   63
      Top             =   5400
      Width           =   1488
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
      Left            =   5040
      TabIndex        =   60
      Top             =   5436
      Width           =   1488
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
      Height          =   225
      Left            =   5505
      TabIndex        =   68
      Top             =   6405
      Width           =   1065
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
      TabIndex        =   49
      Top             =   6552
      Width           =   2568
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
      Height          =   228
      Left            =   9324
      TabIndex        =   58
      Top             =   5076
      Width           =   1488
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
      Height          =   228
      Left            =   5040
      TabIndex        =   52
      Top             =   4716
      Width           =   1488
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
      Height          =   228
      Left            =   5040
      TabIndex        =   54
      Top             =   5076
      Width           =   1488
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
      Height          =   228
      Left            =   5040
      TabIndex        =   50
      Top             =   4356
      Width           =   1488
   End
End
Attribute VB_Name = "frmTestSBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboOrientation_Click()
    SSTabEx1.TabOrientation = cboOrientation.ListIndex
    SSTab1.TabOrientation = cboOrientation.ListIndex
End Sub

Private Sub cboStyle_Click()
    SSTabEx1.Style = cboStyle.ListIndex
    SSTab1.Style = cboStyle.ListIndex
End Sub

Private Sub chkEnabled_Click()
    SSTabEx1.Enabled = chkEnabled.Value
    SSTab1.Enabled = chkEnabled.Value
End Sub

Private Sub chkShowFocusRect_Click()
    SSTabEx1.ShowFocusRect = chkShowFocusRect.Value
    SSTab1.ShowFocusRect = chkShowFocusRect.Value
End Sub

Private Sub chkTabEnabled_Click()
    
End Sub

Private Sub chkVisualStyles_Click()
    SSTabEx1.VisualStyles = chkVisualStyles.Value
    If Not SSTabEx1.IsVisualStyleApplied Then
        chkVisualStyles.Value = 0
    End If
    lblThemedIDE.Visible = InIde And SSTabEx1.IsVisualStyleApplied
End Sub

Private Sub chkWordWrap_Click()
    SSTabEx1.WordWrap = chkWordWrap.Value
    SSTab1.WordWrap = chkWordWrap.Value
End Sub

Private Sub cmdChangeBackColor_Click()
    Dim iDlg As New CDLg
    
    iDlg.ShowColor
    If Not iDlg.Canceled Then
        picBackColor.BackColor = iDlg.Color
        SSTabEx1.BackColor = picBackColor.BackColor
        SSTab1.BackColor = picBackColor.BackColor
    End If

End Sub

Private Sub cmdChangeFont_Click()
    Dim iDlg As New CDLg
    
    Set iDlg.Font = SSTabEx1.Font
    iDlg.ShowFont
    If Not iDlg.Canceled Then
        Set SSTabEx1.Font = iDlg.Font
        Set SSTab1.Font = iDlg.Font
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
        SSTab1.ForeColor = picForeColor.BackColor
        lblThemedIDE.ForeColor = SSTabEx1.ForeColor
        Option1.ForeColor = SSTabEx1.ForeColor: Option6.ForeColor = SSTabEx1.ForeColor
        Option2.ForeColor = SSTabEx1.ForeColor: Option5.ForeColor = SSTabEx1.ForeColor
        Check1.ForeColor = SSTabEx1.ForeColor: Check3.ForeColor = SSTabEx1.ForeColor
        Label1.ForeColor = SSTabEx1.ForeColor: Label9.ForeColor = SSTabEx1.ForeColor
        Label22.ForeColor = SSTabEx1.ForeColor: Label11.ForeColor = SSTabEx1.ForeColor
        Option3.ForeColor = SSTabEx1.ForeColor: Option8.ForeColor = SSTabEx1.ForeColor
        Option4.ForeColor = SSTabEx1.ForeColor: Option7.ForeColor = SSTabEx1.ForeColor
        Check2.ForeColor = SSTabEx1.ForeColor: Check4.ForeColor = SSTabEx1.ForeColor
    End If
End Sub

Private Sub cmdDisableTab_Click()
    SSTabEx1.TabEnabled(SSTabEx1.TabSel) = Not SSTabEx1.TabEnabled(SSTabEx1.TabSel)
    SSTab1.TabEnabled(SSTabEx1.TabSel) = SSTabEx1.TabEnabled(SSTabEx1.TabSel)
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
    Dim iLng As Long
    
    iLng = SSTabEx1.TabSel
    SSTabEx1.TabVisible(iLng) = False
    SSTab1.TabVisible(iLng) = False
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
        cmdRemoveTabPicture_Click
        On Error Resume Next
        Set SSTabEx1.TabPicture(SSTabEx1.TabSel) = LoadPicture(iDlg.FileName)
        Set SSTab1.TabPicture(SSTabEx1.TabSel) = LoadPicture(iDlg.FileName)
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
        SSTab1.TabVisible(t) = True
        SSTab1.TabEnabled(t) = True
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
    Set SSTab1.TabPicture(SSTabEx1.TabSel) = Nothing
End Sub

Private Sub Form_Load()
    Me.Move Screen.Width / 2 - Me.Width / 2, Screen.Height / 2 - Me.Height / 2
    
    SSTabEx1.ForceVisualStyles = InIde
    lblThemedIDE.Visible = InIde And SSTabEx1.IsVisualStyleApplied
    
    lblThemedIDE.Visible = InIde And SSTabEx1.VisualStyles
    
    cboOrientation.Clear
    cboOrientation.AddItem ssTabOrientationTop & " - ssTabOrientationTop"
    cboOrientation.AddItem ssTabOrientationBottom & " - ssTabOrientationBottom"
    cboOrientation.AddItem ssTabOrientationLeft & " - ssTabOrientationLeft"
    cboOrientation.AddItem ssTabOrientationRight & " - ssTabOrientationRight"
    
    cboStyle.Clear
    cboStyle.AddItem ssStyleTabbedDialog & " - ssStyleTabbedDialog"
    cboStyle.AddItem ssStylePropertyPage & " - ssStylePropertyPage"
    
    cboOrientation.ListIndex = SSTabEx1.TabOrientation
    txtTabsPerRow.Text = SSTabEx1.TabsPerRow
    cboStyle.ListIndex = SSTabEx1.Style
    txtTabMaxWidth.Text = SSTabEx1.TabMaxWidth
    chkVisualStyles.Value = Abs(SSTabEx1.VisualStyles)
    chkEnabled.Value = Abs(SSTabEx1.Enabled)
    chkShowFocusRect.Value = Abs(SSTabEx1.ShowFocusRect)
    chkWordWrap.Value = Abs(SSTabEx1.WordWrap)
    picForeColor.BackColor = SSTabEx1.ForeColor
    picBackColor.BackColor = SSTabEx1.BackColor
    ShowFont
    txtTabs.Text = SSTabEx1.Tabs
    txtTabHeight.Text = SSTabEx1.TabHeight
    
    txtTabCaption.Text = SSTabEx1.TabCaption(SSTabEx1.TabSel)
    SSTab1.Tab = SSTabEx1.TabSel
End Sub

Private Sub Label6_Click()

End Sub

Private Sub picBackColor_DblClick()
    cmdChangeBackColor_Click
End Sub

Private Sub picForeColor_DblClick()
    cmdChangeForeColor_Click
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    SSTabEx1.TabSel = SSTab1.Tab
End Sub

Private Sub SSTabEx1_Click(PreviousTab As Integer)
    SSTab1.Tab = SSTabEx1.TabSel
    txtTabCaption.Text = SSTabEx1.TabCaption(SSTabEx1.TabSel)
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
    SSTab1.TabMaxWidth = Val(txtTabMaxWidth.Text)
    
    txtTabMaxWidth.Text = SSTabEx1.TabMaxWidth

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
    SSTab1.Tabs = iLng
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
    SSTab1.TabCaption(SSTabEx1.TabSel) = txtTabCaption.Text
End Sub

Private Sub txtTabHeight_Change()
    SSTabEx1.TabHeight = Val(txtTabHeight.Text)
    SSTab1.TabHeight = Val(txtTabHeight.Text)
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


Private Sub txtTabsPerRow_Change()
    SSTabEx1.TabsPerRow = Val(txtTabsPerRow.Text)
    SSTab1.TabsPerRow = Val(txtTabsPerRow.Text)
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
    
