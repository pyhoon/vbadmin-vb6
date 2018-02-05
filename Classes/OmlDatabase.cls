VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OmlDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'**************************************************************************
'   Library Name: Omelette Database Class (OmlDatabase)
'**************************************************************************
' Version: 0.0.0.1
' Created on: 01 Feb 2018
' Updated on:
' Created by: Aeric Poon Yip Hoon
' Description: A high level API written for MS Visual Basic 6.0
' Company Name: Computerise System Solutions
' Copyright: (c) 2015-2018 Computerise System Solutions, All rights reserved.
' Company Website: http://computerise.my
'**************************************************************************
'
'**************************************************************************
'   Disclaimer
'**************************************************************************
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' You are free to use this code within your own applications, but you are '
' expressly forbidden from selling or otherwise distributing this source  '
' code without prior written consent.                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' References:
' Microsoft ActiveX Data Objects 6.1 Library
' C:\Program Files (x86)\Common Files\System\ado\msado15.dll
' Can use older version such as 2.8
' Recommend to install VB6SP6 for latest version of libraries
'
'**************************************************************************
'   Global Constants
'**************************************************************************
Private mstrErrorDesc As String
Private mstrConnectionString As String
Private mstrProvider As String
Private mstrDataPath As String
Private mstrDataFile As String
Private mstrDataSource As String
Private mstrDataPassword As String
Private mconADODB As ADODB.Connection

Public Property Let ErrorDesc(ByVal strErrorDesc As String)
    mstrErrorDesc = strErrorDesc
End Property
Public Property Get ErrorDesc() As String
    ErrorDesc = mstrErrorDesc
End Property

Public Property Let ConnectionString(ByVal strConStr As String)
    mstrConnectionString = strConStr
End Property
Public Property Get ConnectionString() As String
    ConnectionString = mstrConnectionString
End Property

Public Property Let Provider(ByVal strProvider As String)
    mstrProvider = strProvider
End Property
Public Property Get Provider() As String
    Provider = mstrProvider
End Property

Public Property Let DataPath(ByVal strDataPath As String)
    mstrDataPath = strDataPath
End Property
Public Property Get DataPath() As String
    DataPath = mstrDataPath
End Property

Public Property Let DataFile(ByVal strDataFile As String)
    mstrDataFile = strDataFile
End Property
Public Property Get DataFile() As String
    DataFile = mstrDataFile
End Property

Public Property Let DataSource(ByVal strDataSource As String)
    mstrDataSource = strDataSource
End Property
Public Property Get DataSource() As String
    DataSource = mstrDataSource
End Property

Public Property Let DataPassword(ByVal strDataPassword As String)
    mstrDataPassword = strDataPassword
End Property
Public Property Get DataPassword() As String
    DataPassword = mstrDataPassword
End Property

Public Property Let Connection(ByVal conADODB As ADODB.Connection)
    mconADODB = conADODB
End Property
Public Property Get Connection() As ADODB.Connection
    Connection = mconADODB
End Property

Private Sub Class_Initialize()
    Provider = "Microsoft.Jet.OLEDB.4.0"
    'DataSource = DataPath & DataFile ' & "\" & mstrDataFile
    'ConnectionString = "Data Source=" & DataSource
    'DataPassword = ""
End Sub

Public Sub OpenMdb()
On Error GoTo Catch
Try: ' Optional
    Set mconADODB = New ADODB.Connection
    DataSource = DataPath & DataFile ' & "\" & mstrDataFile
    ConnectionString = "Data Source=" & DataSource
    With mconADODB
        .Provider = mstrProvider
        .ConnectionString = mstrConnectionString
        .Properties("Jet OLEDB:Database Password") = mstrDataPassword
        .Open
    End With
    Exit Sub
Catch:
    ErrorDesc = Err.Description
End Sub

Public Sub CloseMdb()
On Error GoTo Catch
Try:
    If mconADODB Is Nothing Then
        ' No action
    Else
        With mconADODB
            If .State = adStateOpen Then
                .Close
            End If
        End With
        Set mconADODB = Nothing
    End If
    Exit Sub
Catch:
    ErrorDesc = Err.Description
End Sub

Public Sub Execute(ByVal strSQL As String, Optional ByRef lngRecordsAffected As Long)
On Error GoTo Catch
Try: ' Optional
    With mconADODB
        .BeginTrans
        .Execute strSQL, lngRecordsAffected
        .CommitTrans
    End With
    Exit Sub
Catch:
    ErrorDesc = Err.Description
    On Error Resume Next
    mconADODB.RollbackTrans
    ErrorDesc = Err.Description & " " & strSQL
End Sub

Public Function OpenRs(ByVal strSQL As String) As ADODB.Recordset
Dim rst As New ADODB.Recordset
On Error GoTo Catch
Try:
    rst.Open strSQL, mconADODB, adOpenStatic, adLockPessimistic, adCmdText
    Set OpenRs = rst
    Exit Function
Catch:
    ErrorDesc = Err.Description & " " & strSQL
End Function

Public Sub CloseRs(ByVal rst As ADODB.Recordset)
On Error GoTo Catch
Try:
    If rst Is Nothing Then
    Else
        If rst.State = adStateOpen Then
            rst.Close
        End If
        Set rst = Nothing
    End If
    Exit Sub
Catch:
    ErrorDesc = Err.Description
End Sub
