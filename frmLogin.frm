VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   7110
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
   ScaleHeight     =   7110
   ScaleWidth      =   13365
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraContainer1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   12615
      Begin VB.Frame fraButton1 
         Appearance      =   0  'Flat
         BackColor       =   &H00B7D736&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   5280
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   4440
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
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "n"
         TabIndex        =   1
         ToolTipText     =   "Password is case sensitive"
         Top             =   3120
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
         Left            =   3120
         TabIndex        =   0
         Top             =   1920
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
         Left            =   3120
         TabIndex        =   11
         Top             =   2760
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
         Left            =   3120
         TabIndex        =   10
         Top             =   1560
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
         Left            =   12750
         TabIndex        =   5
         Top             =   0
         Width           =   615
      End
      Begin VB.Shape shpX 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   12730
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
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DB As New OmlDatabase
Dim strAppDataPath As String
Dim strAppDataFile As String
Dim strSQL As String

Dim MoveStartX As Single
Dim MoveStartY As Single
Dim MoveEndX As Single
Dim MoveEndY As Single

Private Sub Form_Initialize()
    ' Source: http://www.vbforums.com/showthread.php?432036-Classic-VB-How-can-I-set-my-exe-icon-using-a-resource-file
    Me.Icon = LoadResPicture("APPICON", vbResIcon)
End Sub

Private Sub Form_Load()
    Me.Caption = "LOGIN"
    lblTitle.Caption = Me.Caption
    Label1.Caption = "USER ID"
    Label2.Caption = "PASSWORD"
    lblButton1.Caption = "SUBMIT"
    txtUserID.MaxLength = 20
    txtPassword.MaxLength = 20
    LoadMousePointer
    SetContainerTitle
End Sub

Private Sub LoadMousePointer()
On Error Resume Next
    fraButton1.MouseIcon = LoadPicture(App.Path & "\Resources\Icon\hand.ico")
End Sub

Private Sub SetContainerTitle()
    lblContainerTitle1.Caption = "PLEASE ENTER YOUR LOGIN CREDENTIALS"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'frmDashboard.Show
End Sub

Private Sub fraButton1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With fraButton1
        .BackColor = &HE0E0E0
    End With
    With lblButton1
        .ForeColor = &H404040
    End With
End Sub

Private Sub fraContainer1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With fraButton1
        .BackColor = &HB7D736
    End With
    With lblButton1
        .ForeColor = &HFFFFFF
    End With
End Sub

Private Sub fraTitle_Click()
    ' Shortcut
    txtUserID.Text = "Aeric"
    txtPassword.Text = "aeric"
End Sub

Private Sub fraTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetMouseMove Button, X, Y
End Sub

Private Sub fraTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseMove Button, X, Y
End Sub

Private Sub fraButton1_Click()
    If Not AuthenticateUser Then
        MsgBox "Wrong User ID or Password!", vbExclamation, "Access Denied"
    Else
        Unload Me
        frmDashboard.Show
    End If
End Sub

Private Sub lblButton1_Click()
    If Not AuthenticateUser Then
        MsgBox "Wrong User ID or Password!", vbExclamation, "Access Denied"
    Else
        Unload Me
        frmDashboard.Show
    End If
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
    'frmDashboard.Show
End Sub

Private Sub txtPassword_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Not AuthenticateUser Then
            MsgBox "Wrong User ID or Password!", vbExclamation, "Access Denied"
        Else
            Unload Me
            frmDashboard.Show
        End If
    End If
End Sub

Private Sub txtUserID_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        With txtPassword
            .SetFocus
            .SelStart = Len(.Text)
            .SelLength = 0
        End With
    End If
End Sub

Private Function AuthenticateUser() As Boolean
    Dim rst As ADODB.Recordset    
    Dim strUserID As String
    Dim strPassword As String
    Dim strSalt As String
        
    strAppDataPath = App.Path & "\Storage\"
    strAppDataFile = "Data.mdb"
    
    strUserID = Trim(txtUserID.Text)
    strPassword = Trim(txtPassword.Text)
    strSalt = GetSalt(strUserID)
    
    With DB
        .DataPath = strAppDataPath
        .DataFile = strAppDataFile
        '.DataPassword = ""
        .OpenMdb
        If .ErrorDesc <> "" Then
            MsgBox "Error: " & .ErrorDesc, vbExclamation, "Open Database"
            Exit Function
        End If
        strSQL = "SELECT UserName"
        strSQL = strSQL & " FROM Users"
        strSQL = strSQL & " WHERE UserID = '" & strUserID & "'"
        strSQL = strSQL & " AND UserPassword = '" & MD5(strPassword & strSalt) & "'"
        strSQL = strSQL & " AND Active = Yes"
        Set rst = .OpenRs(strSQL)
        If .ErrorDesc <> "" Then
            MsgBox "Error: " & .ErrorDesc, vbExclamation, "Query Database"
            Exit Function
        End If
        If Not rst.EOF Then
            gstrUserName = rst!UserName
            AuthenticateUser = True
        Else
            gstrUserName = ""
            AuthenticateUser = False
        End If
        .CloseRs rst
        .CloseMdb
    End With
End Function

Private Function GetSalt(ByVal strUserID As String) As String
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
        strSQL = "SELECT Salt"
        strSQL = strSQL & " FROM Users"
        strSQL = strSQL & " WHERE UserID = '" & strUserID & "'"
        Set rst = .OpenRs(strSQL)
        If .ErrorDesc <> "" Then
            MsgBox "Error: " & .ErrorDesc, vbExclamation, "Query Database"
            Exit Function
        End If
        With rst
            If Not .EOF Then
                GetSalt = !Salt
            Else
                GetSalt = ""
            End If
        End With
        .CloseRs rst
        .CloseMdb
    End With
End Function
