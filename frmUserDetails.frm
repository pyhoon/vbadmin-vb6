VERSION 5.00
Begin VB.Form frmUserDetails 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   8595
   ClientLeft      =   14865
   ClientTop       =   1005
   ClientWidth     =   13365
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
   ScaleHeight     =   8595
   ScaleWidth      =   13365
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboUserRole 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   4320
      Width           =   6975
   End
   Begin VB.Frame fraContainer1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   12615
      Begin VB.Frame fraButton2 
         Appearance      =   0  'Flat
         BackColor       =   &H00B7D736&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   9360
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   6000
         Width           =   2775
         Begin VB.Label lblButton2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BUTTON LABEL 2"
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
            TabIndex        =   18
            Top             =   240
            Width           =   2235
         End
      End
      Begin VB.OptionButton optActive 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   1
         Left            =   5400
         TabIndex        =   16
         Top             =   4200
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optActive 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   0
         Left            =   3000
         TabIndex        =   15
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Frame fraButton1 
         Appearance      =   0  'Flat
         BackColor       =   &H00B7D736&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   5280
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   6000
         Width           =   2775
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
            TabIndex        =   2
            Top             =   240
            Width           =   2235
         End
      End
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   3000
         MaxLength       =   20
         TabIndex        =   1
         Top             =   2280
         Width           =   6975
      End
      Begin VB.TextBox txtUserID 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   615
         Left            =   3000
         MaxLength       =   20
         TabIndex        =   0
         Top             =   1320
         Width           =   6975
      End
      Begin VB.Frame fraContainerTitle1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   12615
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
            ForeColor       =   &H00404040&
            Height          =   300
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   2565
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LABEL 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   960
         TabIndex        =   14
         Top             =   4200
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LABEL 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   960
         TabIndex        =   12
         Top             =   3240
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LABEL 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   960
         TabIndex        =   11
         Top             =   2280
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LABEL 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   960
         TabIndex        =   10
         Top             =   1320
         Width           =   1080
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
      TabIndex        =   3
      Top             =   0
      Width           =   13365
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
         Left            =   12730
         TabIndex        =   5
         Top             =   0
         Width           =   615
      End
      Begin VB.Shape shpX 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   12750
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
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmUserDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DB As New OmlDatabase
Dim strAppDataPath As String
Dim strAppDataFile As String
Dim strSQL As String
Dim strUserID As String
Dim strUserName As String
Dim strUserRole As String
Dim strActive As String
Dim strSalt As String
Dim strPassword As String

Dim MoveStartX As Single
Dim MoveStartY As Single
Dim MoveEndX As Single
Dim MoveEndY As Single

Private Sub Form_Initialize()
    Me.Icon = LoadResPicture("APPICON", vbResIcon)
End Sub

Private Sub Form_Load()
    Me.Caption = "USER DETAILS"
    lblTitle.Caption = Me.Caption
    Label1.Caption = "USER ID"
    Label2.Caption = "USER NAME"
    Label3.Caption = "USER ROLE"
    Label4.Caption = "ACTIVE"
    lblButton1.Caption = "UPDATE"
    lblButton2.Caption = "SECRET"
    With cboUserRole
        .AddItem "Admin"
        .AddItem "Manager"
        .AddItem "User"
    End With
    LoadMousePointer
    SetContainerTitle
End Sub

Private Sub LoadMousePointer()
On Error Resume Next
    fraButton1.MouseIcon = LoadPicture(App.Path & "\Resources\Icon\hand.ico")
End Sub

Private Sub SetContainerTitle()
    lblContainerTitle1.Caption = "USER DETAILS"
End Sub

Private Sub fraButton1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With fraButton1
        .BackColor = &HE0E0E0
    End With
    With lblButton1
        .ForeColor = &H404040
    End With
End Sub

Private Sub fraButton2_Click()
    If txtUserID.Text = "" Then
        MsgBox "User ID is empty!", vbExclamation, "User ID"
        Exit Sub
    End If
    With frmUserUpdateSaltPassword
        .Show
        .cboUserID.Text = txtUserID.Text
    End With
    Unload Me
End Sub

Private Sub fraContainer1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With fraButton1
        .BackColor = &HB7D736
    End With
    With lblButton1
        .ForeColor = &HFFFFFF
    End With
End Sub

Private Sub fraTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetMouseMove Button, X, Y
End Sub

Private Sub fraTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseMove Button, X, Y
End Sub

Private Sub fraButton1_Click()
    If txtUserID.Text = "" Then
        MsgBox "User ID is empty!", vbExclamation, "User ID"
        Exit Sub
    End If
    If FindUser Then
        UpdateUser
    Else
        AddUser
    End If
End Sub

Private Sub lblButton1_Click()
    fraButton1_Click
End Sub

Private Sub lblButton2_Click()
    fraButton2_Click
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
    frmUsers.Show
End Sub

Public Sub PopulateValues(strUserID As String)
    Dim rst As ADODB.Recordset
    
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
        strSQL = "SELECT"
        strSQL = strSQL & " UserID,"
        strSQL = strSQL & " UserName,"
        strSQL = strSQL & " UserRole,"
        strSQL = strSQL & " Active"
        strSQL = strSQL & " FROM Users"
        strSQL = strSQL & " WHERE ID = " & strUserID
        Set rst = .OpenRs(strSQL)
        If .ErrorDesc <> "" Then
            MsgBox "Error: " & .ErrorDesc, vbExclamation, "Query Database"
            Exit Sub
        End If
        With rst
            If Not .EOF Then
                txtUserID.Text = !UserID
                txtUserName.Text = !UserName
                cboUserRole.Text = !UserRole
                If !Active Then
                    optActive(0).Value = True
                Else
                    optActive(1).Value = True
                End If
                txtUserID.Enabled = False
                With txtUserName
                    .SelStart = Len(.Text)
                    .SelLength = 0
                End With
            Else
                txtUserID.Text = ""
                txtUserName.Text = ""
                cboUserRole.ListIndex = -1
                optActive(1).Value = True ' Default = No
                With txtUserID
                    .Enabled = True
                    .SetFocus
                End With
            End If
        End With
        .CloseRs rst
        .CloseMdb
    End With
End Sub

Private Function FindUser() As Boolean
    Dim rst As ADODB.Recordset
        
    strAppDataPath = App.Path & "\Storage\"
    strAppDataFile = "Data.mdb"
    With DB
        .DataPath = strAppDataPath
        .DataFile = strAppDataFile
        '.DataPassword = ""
        .OpenMdb
        If .ErrorDesc <> "" Then
            MsgBox "Error: " & .ErrorDesc, vbExclamation, "Open Database"
            Exit Function
        End If
        strUserID = Trim(txtUserID.Text)
        strSQL = "SELECT ID"
        strSQL = strSQL & " FROM Users"
        strSQL = strSQL & " WHERE UserID = '" & strUserID & "'"
        Set rst = .OpenRs(strSQL)
        If .ErrorDesc <> "" Then
            MsgBox "Error: " & .ErrorDesc, vbExclamation, "Query Database"
            Exit Function
        End If
        If Not rst.EOF Then
            FindUser = True
        Else
            FindUser = False
        End If
        .CloseRs rst
        .CloseMdb
    End With
End Function

Private Sub AddUser()
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
        strUserID = Trim(txtUserID.Text)
        strUserName = Trim(txtUserName.Text)
        strUserRole = cboUserRole.Text
        If optActive(0).Value = True Then
            strActive = "Yes"
        Else
            strActive = "No"
        End If
        strSalt = GenerateSalt("SECRET")    ' default
        strPassword = strUserID             ' default
        strPassword = MD5(strPassword & strSalt)
        strSQL = "INSERT INTO Users"
        strSQL = strSQL & " (UserID,"
        strSQL = strSQL & " UserName,"
        strSQL = strSQL & " UserRole,"
        strSQL = strSQL & " Salt,"
        strSQL = strSQL & " UserPassword,"
        strSQL = strSQL & " Active)"
        strSQL = strSQL & " VALUES"
        strSQL = strSQL & " ('" & strUserID & "',"
        strSQL = strSQL & " '" & strUserName & "',"
        strSQL = strSQL & " '" & strUserRole & "',"
        strSQL = strSQL & " '" & strSalt & "',"
        strSQL = strSQL & " '" & strPassword & "',"
        strSQL = strSQL & " " & strActive & ")"
        .Execute strSQL
        If .ErrorDesc <> "" Then
            MsgBox "Error: " & .ErrorDesc, vbExclamation, "Add User"
            Exit Sub
        End If
        MsgBox "User added", vbInformation, "Add User"
        .CloseMdb
    End With
End Sub

Private Sub UpdateUser()
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
        strUserID = Trim(txtUserID.Text)
        strUserName = Trim(txtUserName.Text)
        strUserRole = cboUserRole.Text
        If optActive(0).Value = True Then
            strActive = "Yes"
        Else
            strActive = "No"
        End If
        strSQL = "UPDATE Users SET"
        strSQL = strSQL & " UserName = '" & strUserName & "',"
        strSQL = strSQL & " UserRole = '" & strUserRole & "',"
        strSQL = strSQL & " Active = " & strActive
        strSQL = strSQL & " WHERE UserID = '" & strUserID & "'"
        .Execute strSQL
        If .ErrorDesc <> "" Then
            MsgBox "Error: " & .ErrorDesc, vbExclamation, "Update User"
            Exit Sub
        End If
        MsgBox "User updated", vbInformation, "Update User"
        .CloseMdb
    End With
End Sub

Private Function GenerateSalt(ByVal strPlain As String) As String
    ' A better way is to generate a random string for salt
    GenerateSalt = MD5(strPlain)
End Function
