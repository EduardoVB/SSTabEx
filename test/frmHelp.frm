VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "Help"
   ClientHeight    =   5580
   ClientLeft      =   6360
   ClientTop       =   2640
   ClientWidth     =   6384
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   6384
   Begin VB.TextBox txtHelp 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5484
      Left            =   36
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmHelp.frx":0000
      Top             =   36
      Width           =   6312
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    txtHelp.Move 40, 40, Me.ScaleWidth - 80, Me.ScaleHeight - 80
End Sub
