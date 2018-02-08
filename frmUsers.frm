VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsers 
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
      Begin VB.Frame fraButton1 
         Appearance      =   0  'Flat
         BackColor       =   &H00B7D736&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   9000
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   6480
         Width           =   3495
         Begin VB.Label lblButton1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BUTTON LABEL 1"
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
            Left            =   225
            TabIndex        =   13
            Top             =   240
            Width           =   3075
         End
      End
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   5295
         Left            =   360
         TabIndex        =   15
         Top             =   1080
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   9340
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
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
         TabIndex        =   14
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
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DB As New OmlDatabase
Dim strAppDataPath As String
Dim strAppDataFile As String

Dim MoveStartX As Single
Dim MoveStartY As Single
Dim MoveEndX As Single
Dim MoveEndY As Single

Private Sub Form_Initialize()
    ' Source: http://www.vbforums.com/showthread.php?432036-Classic-VB-How-can-I-set-my-exe-icon-using-a-resource-file
    Me.Icon = LoadResPicture("APPICON", vbResIcon)
End Sub

Private Sub Form_Load()
    Me.Caption = "USERS"
    lblTitle.Caption = Me.Caption
    lblUserName.Caption = gstrUserName
    lblButton1.Caption = "CHANGE PASSWORD"
    LoadMousePointer
    SetContainerTitle
    LoadList
End Sub

Private Sub AddColHeader()
    Dim intWidth(0 To 4) As Integer
    'Dim strText As String
    On Error GoTo CheckErr
    With ListView1
        .ColumnHeaders.Clear
        intWidth(0) = .Width * 0.1
        intWidth(1) = .Width * 0.25
        intWidth(2) = .Width * 0.35
        intWidth(3) = .Width * 0.15
        intWidth(4) = .Width * 0.15
        .ColumnHeaders.Add , "ID", "ID", intWidth(0)
        .ColumnHeaders.Add , "User ID", "User ID", intWidth(1), lvwColumnLeft
        .ColumnHeaders.Add , "User Name", "User Name", intWidth(2), lvwColumnLeft
        .ColumnHeaders.Add , "Role", "Role", intWidth(3), lvwColumnLeft
        .ColumnHeaders.Add , "Active", "Active", intWidth(4), lvwColumnCenter
        .Refresh
    End With
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, "AddColHeader"
End Sub

Private Sub LoadList()
    Dim con As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim strSQL As String
    Dim List As ListItem
    Dim i As Integer
    Dim r As Integer
    
    strAppDataPath = App.Path & "\Storage\"
    strAppDataFile = "Data.mdb"
    With DB
        .DataPath = strAppDataPath
        .DataFile = strAppDataFile
        '.DataPassword = ""
        .OpenMdb
        If .ErrorDesc <> "" Then
            MsgBox "Error: " & .ErrorDesc, vbExclamation, "Open Database"
            Exit Sub
        End If
        strSQL = "SELECT *"
        strSQL = strSQL & " FROM Users"
        Set rst = .OpenRs(strSQL)
        If .ErrorDesc <> "" Then
            MsgBox "Error: " & .ErrorDesc, vbExclamation, "Query Database"
            Exit Sub
        End If
        ListView1.ListItems.Clear
        AddColHeader
        While Not rst.EOF
            Set List = ListView1.ListItems.Add(, "U" & rst!ID, rst!ID, , 0)
            List.SubItems(1) = rst!UserID
            List.SubItems(2) = rst!UserName
            List.SubItems(3) = rst!UserRole
            List.SubItems(4) = rst!Active
            If rst!Active = False Then
                List.ForeColor = vbRed
                For r = 1 To List.ListSubItems.Count
                    List.ListSubItems(r).ForeColor = vbRed
                Next
            Else
                List.ForeColor = vbBlack
                For r = 1 To List.ListSubItems.Count
                    List.ListSubItems(r).ForeColor = vbBlack
                Next
            End If
            rst.MoveNext
            i = i + 1
        Wend
        .CloseRs rst
        .CloseMdb
    End With
End Sub

Private Sub LoadMousePointer()
On Error Resume Next
    fraMenu1.MouseIcon = LoadPicture(App.Path & "\Resources\Icon\hand.ico")
    'fraMenu2.MouseIcon = LoadPicture(App.Path & "\Resources\Icon\hand.ico")
End Sub

Private Sub SetContainerTitle()
    lblContainerTitle1.Caption = "USERS"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetMouseMove Button, X, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseMove Button, X, Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmDashboard.Show
End Sub

Private Sub fraContainer1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetMouseMove Button, X, Y
End Sub

Private Sub fraContainer1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseMove Button, X, Y
End Sub

Private Sub fraContainerTitle1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetMouseMove Button, X, Y
End Sub

Private Sub fraContainerTitle1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseMove Button, X, Y
End Sub

Private Sub fraMenu1_Click()
    'MsgBox "" & lblMenu1.Caption, vbInformation, "Click"
    Unload Me
    frmDashboard.Show
End Sub

'Private Sub fraMenu2_Click()
'    MsgBox "" & lblMenu2.Caption, vbInformation, "Click"
'End Sub

Private Sub fraMenuContainer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetMouseMove Button, X, Y
End Sub

Private Sub fraMenuContainer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseMove Button, X, Y
    fraMenu1.BackColor = &H80000010
    fraMenu2.BackColor = &H80000010
End Sub

Private Sub fraMenu1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraMenu1.BackColor = &HE0E0E0
End Sub

'Private Sub fraMenu2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    fraMenu2.BackColor = &HE0E0E0
'End Sub

Private Sub fraTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetMouseMove Button, X, Y
End Sub

Private Sub fraTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseMove Button, X, Y
End Sub

Private Sub lblContainerTitle1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetMouseMove Button, X, Y
End Sub

Private Sub lblContainerTitle1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
