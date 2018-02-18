VERSION 5.00
Begin VB.Form frmUserUpdateSaltPassword 
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
   Begin VB.Frame fraContainer1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   12615
      Begin VB.ComboBox cboUserID 
         Appearance      =   0  'Flat
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
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1920
         Width           =   6975
      End
      Begin VB.OptionButton optType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "PASSWORD"
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
         Left            =   6360
         TabIndex        =   12
         Top             =   4200
         Width           =   2895
      End
      Begin VB.OptionButton optType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SALT"
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
         TabIndex        =   11
         Top             =   4200
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.Frame fraButton1 
         Appearance      =   0  'Flat
         BackColor       =   &H00B7D736&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   5280
         MousePointer    =   99  'Custom
         TabIndex        =   8
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
            TabIndex        =   1
            Top             =   240
            Width           =   2235
         End
      End
      Begin VB.TextBox txtSecretWord 
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
         TabIndex        =   0
         Top             =   3120
         Width           =   6975
      End
      Begin VB.Frame fraContainerTitle1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   0
         TabIndex        =   6
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
            TabIndex        =   7
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
         Left            =   3000
         TabIndex        =   10
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
         Left            =   3000
         TabIndex        =   9
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
      TabIndex        =   2
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmUserUpdateSaltPassword"
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
    Me.Icon = LoadResPicture("APPICON", vbResIcon)
End Sub

Private Sub Form_Load()
    Me.Caption = "UPDATE SALT && PASSWORD"
    lblTitle.Caption = Me.Caption
    Label1.Caption = "USER ID"
    Label2.Caption = "SECRET WORD"
    lblButton1.Caption = "UPDATE"
    LoadMousePointer
    SetContainerTitle
    LoadCombo
End Sub

Private Sub LoadMousePointer()
On Error Resume Next
    fraButton1.MouseIcon = LoadPicture(App.Path & "\Resources\Icon\hand.ico")
End Sub

Private Sub SetContainerTitle()
    lblContainerTitle1.Caption = "PLEASE UPDATE PASSWORD AGAIN AFTER UPDATE SALT"
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

Private Sub fraTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetMouseMove Button, X, Y
End Sub

Private Sub fraTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseMove Button, X, Y
End Sub

Private Sub fraButton1_Click()
    Dim strUser As String
    Dim strSecret As String
    strUser = Trim(cboUserID.Text)
    strSecret = Trim(txtSecretWord.Text)
    If optType(0).Value = True Then
        UpdateSalt strUser, strSecret
    Else
        UpdatePassword strUser, strSecret
    End If
End Sub

Private Sub lblButton1_Click()
    fraButton1_Click
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

Private Sub LoadCombo()
    Dim con As ADODB.Connection
    Dim rst As ADODB.Recordset
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
        strSQL = "SELECT"
        strSQL = strSQL & " UserID"
        strSQL = strSQL & " FROM Users"
        'strSQL = strSQL & " WHERE Active = Yes"
        Set rst = .OpenRs(strSQL)
        If .ErrorDesc <> "" Then
            MsgBox "Error: " & .ErrorDesc, vbExclamation, "Query Database"
            Exit Sub
        End If
        cboUserID.Clear
        While Not rst.EOF
            cboUserID.AddItem rst!UserID
            rst.MoveNext
        Wend
        .CloseRs rst
        .CloseMdb
    End With
End Sub

Private Function GetSalt(ByVal strUserID As String) As String
    Dim con As ADODB.Connection
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

Private Sub UpdateSalt(ByVal strUserID As String, ByVal strSalt As String)
    Dim con As ADODB.Connection
    
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
        strSQL = "UPDATE Users SET"
        strSQL = strSQL & " Salt = '" & GenerateSalt(strSalt) & "'"
        strSQL = strSQL & " WHERE UserID = '" & strUserID & "'"
        .Execute strSQL
        If .ErrorDesc <> "" Then
            MsgBox "Error: " & .ErrorDesc, vbExclamation, "Update Database"
            Exit Sub
        End If
        MsgBox "Salt updated", vbInformation, "UpdateSalt"
        .CloseMdb
    End With
End Sub

Private Sub UpdatePassword(ByVal strUserID As String, ByVal strPassword As String)
    Dim con As ADODB.Connection
    Dim strSalt As String
    
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
        strSalt = GetSalt(strUserID)
        strSQL = "UPDATE Users SET"
        strSQL = strSQL & " UserPassword = '" & MD5(strPassword & strSalt) & "'"
        strSQL = strSQL & " WHERE UserID = '" & strUserID & "'"
        .Execute strSQL
        If .ErrorDesc <> "" Then
            MsgBox "Error: " & .ErrorDesc, vbExclamation, "Update Database"
            Exit Sub
        End If
        MsgBox "Password updated", vbInformation, "UpdatePassword"
        .CloseMdb
    End With
End Sub

Private Function GenerateSalt(ByVal strPlain As String) As String
    GenerateSalt = MD5(strPlain)
End Function
