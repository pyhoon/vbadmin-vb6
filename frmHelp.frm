VERSION 5.00
Begin VB.Form frmHelp 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   8790
   ClientLeft      =   14865
   ClientTop       =   1005
   ClientWidth     =   17295
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   17295
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraContainer1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   3960
      TabIndex        =   9
      Top             =   1080
      Width           =   12975
      Begin VB.Frame fraContainerTitle1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   12975
         Begin VB.Label lblContainerTitle1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CONTAINER TITLE 1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   2565
         End
      End
   End
   Begin VB.Frame fraMenuContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   3615
      Begin VB.Frame fraMenu2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   1200
         Width           =   3375
         Begin VB.Label lblMenu2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Users"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame fraMenu1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   240
         Width           =   3375
         Begin VB.Label lblMenu1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dashboard"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   1275
         End
      End
   End
   Begin VB.Frame fraTitle 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17295
      Begin VB.Label lblUserIcon 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "user"
         BeginProperty Font 
            Name            =   "Font Awesome 5 Free Regular"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   14400
         TabIndex        =   12
         Top             =   180
         Width           =   315
      End
      Begin VB.Label lblUserName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Administrator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   14880
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   16680
         TabIndex        =   2
         Top             =   0
         Width           =   615
      End
      Begin VB.Shape shpX 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   16680
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "APPLICATION TITLE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MoveStartX As Single
Dim MoveStartY As Single
Dim MoveEndX As Single
Dim MoveEndY As Single

Private Sub Form_Initialize()
    ' Source: http://www.vbforums.com/showthread.php?432036-Classic-VB-How-can-I-set-my-exe-icon-using-a-resource-file
    Me.Icon = LoadResPicture("APPICON", vbResIcon)
End Sub

Private Sub Form_Load()
    Me.Caption = "HELP"
    lblTitle.Caption = Me.Caption
    lblUserName.Caption = gstrUserName
    LoadMousePointer
    SetContainerTitle
End Sub

Private Sub LoadMousePointer()
On Error Resume Next
    fraMenu1.MouseIcon = LoadPicture(App.Path & "\Resources\Icon\hand.ico")
    fraMenu2.MouseIcon = LoadPicture(App.Path & "\Resources\Icon\hand.ico")
End Sub

Private Sub SetContainerTitle()
    lblContainerTitle1.Caption = "HELP"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'frmDashboard.Show
End Sub

Private Sub fraMenu1_Click()
    Unload Me
    frmDashboard.Show
End Sub

Private Sub fraMenu2_Click()
    Unload Me
    frmUsers.Show
End Sub

Private Sub fraMenuContainer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraMenu1.BackColor = &H80000010
    fraMenu2.BackColor = &H80000010
End Sub

Private Sub fraMenu1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraMenu1.BackColor = &HE0E0E0
End Sub

Private Sub fraMenu2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraMenu2.BackColor = &HE0E0E0
End Sub

Private Sub fraTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetMouseMove Button, X, Y
End Sub

Private Sub fraTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseMove Button, X, Y
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetMouseMove Button, X, Y
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseMove Button, X, Y
End Sub

Private Sub GetMouseMove(Button As Integer, X As Single, Y As Single)
    MoveStartX = X
    MoveStartY = Y
End Sub

Private Sub SetMouseMove(Button As Integer, X As Single, Y As Single)
    MoveEndX = X - MoveStartX
    MoveEndY = Y - MoveStartY
    If Button = 1 Then
        Me.Left = Me.Left + MoveEndX
        Me.Top = Me.Top + MoveEndY
    End If
End Sub

Private Sub lblX_Click()
    Unload Me
    frmDashboard.Show
End Sub
