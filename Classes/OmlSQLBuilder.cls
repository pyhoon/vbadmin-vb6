VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OmlSQLBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************
'   Library Name: Omelette SQL Builder Class (OmlSQLBuilder)
'**************************************************************************
' Version: 1.0.0.0
' Created on: 19 Jan 2019
' Updated on: 24 Jan 2019
' Created by: Aeric Poon Yip Hoon
' Description: A way to simplify SQL query written for MS Visual Basic 6.0
'              This class replaced modSQLQuery which uses gstrSQL
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
' No dependencies
'
'**************************************************************************
'   Global Constants
'**************************************************************************
Option Explicit
Private mstrError As String
Private mstrText As String

Public Property Let Error(ByVal strError As String)
    mstrError = strError
End Property
Public Property Get Error() As String
    Error = mstrError
End Property

Public Property Let Text(ByVal strText As String)
    mstrText = strText
End Property
Public Property Get Text() As String
    Text = mstrText
End Property

Private Sub Class_Initialize()
    mstrText = ""
End Sub

' Set a Text value for Update
Public Sub UTX(strField As String, strText As String, Optional blnEndComma As Boolean = True)
    mstrText = mstrText & " " & strField & " = '" & strText & "'"
    If blnEndComma = True Then
        mstrText = mstrText & ","
    End If
End Sub

' Set a Double value for Update
Public Sub UDB(strField As String, dblNumber As Double, Optional blnEndComma As Boolean = True)
    mstrText = mstrText & " " & strField & " = " & dblNumber
    If blnEndComma = True Then
        mstrText = mstrText & ","
    End If
End Sub

' Set a Long value for Update
Public Sub ULN(strField As String, lngNumber As Long, Optional blnEndComma As Boolean = True)
    mstrText = mstrText & " " & strField & " = " & lngNumber
    If blnEndComma = True Then
        mstrText = mstrText & ","
    End If
End Sub

' Set a Boolean value for Update
Public Sub UBL(strField As String, blnValue As Boolean, Optional blnEndComma As Boolean = True)
    mstrText = mstrText & " " & strField
    If blnValue Then
        mstrText = mstrText & " = TRUE"
    Else
        mstrText = mstrText & " = FALSE"
    End If
    If blnEndComma = True Then
        mstrText = mstrText & ","
    End If
End Sub
' Set a YesNo value for Update
Public Sub UYN(strField As String, blnValue As Boolean, Optional blnEndComma As Boolean = True)
    mstrText = mstrText & " " & strField
    If blnValue Then
        mstrText = mstrText & " = Yes"
    Else
        mstrText = mstrText & " = No"
    End If
    If blnEndComma = True Then
        mstrText = mstrText & ","
    End If
End Sub

' Set a DateTime value for Update
Public Sub UDT(strField As String, strDateTime As String, Optional blnEndComma As Boolean = True)
    mstrText = mstrText & " " & strField & " = #" & strDateTime & "#"
    If blnEndComma = True Then
        mstrText = mstrText & ","
    End If
End Sub

''' Append text to SQL statement string
Public Sub SQL(strText As String, Optional blnEndComma As Boolean = True, Optional blnBeginSpace As Boolean = True)
    If blnBeginSpace = True Then
        mstrText = mstrText & " "
    End If
    mstrText = mstrText & strText
    If blnEndComma = True Then
        mstrText = mstrText & ","
    End If
End Sub

''' Open Bracket when building INSERT statement
''' blnSpaceAfter is adding a blankspace after open bracket "( "
Public Sub SQL_Open_Bracket(strText As String, Optional blnEndComma As Boolean = True, Optional blnSpaceAfter As Boolean = False)
    mstrText = mstrText & "("
    If blnSpaceAfter = True Then
        mstrText = mstrText & " "
    End If
    mstrText = mstrText & strText
    If blnEndComma = True Then
        mstrText = mstrText & ","
    End If
End Sub
Public Sub SOB(strText As String, Optional blnEndComma As Boolean = True, Optional blnSpaceAfter As Boolean = False)
    SQL_Open_Bracket strText, blnEndComma, blnSpaceAfter
End Sub

''' Close Bracket when building INSERT statement
''' blnSpaceBefore is adding a blankspace before close bracket " )"
Public Sub SQL_Close_Bracket(strText As String, Optional blnEndComma As Boolean = False, Optional blnSpaceBefore As Boolean = False)
    If blnSpaceBefore = True Then
        mstrText = mstrText & " "
    End If
    mstrText = mstrText & strText
    If blnEndComma = True Then
        mstrText = mstrText & ","
    End If
    mstrText = mstrText & ")"
End Sub
Public Sub SCB(strText As String, Optional blnEndComma As Boolean = False, Optional blnSpaceBefore As Boolean = False)
    SQL_Close_Bracket strText, blnEndComma, blnSpaceBefore
End Sub

''' Single Field inside Brackets
''' blnBeginSpace is adding a blankspace before open bracket " ("
''' blnSpaceAfter is adding a blankspace after open bracket "( "
''' blnSpaceBefore is adding a blankspace before close bracket " )"
Public Sub SQL_Single_Field(strField As String, Optional blnSpaceAfter As Boolean = True, Optional blnSpaceBefore As Boolean = True, Optional blnBeginSpace As Boolean = True)
    If blnBeginSpace = True Then
        mstrText = " " & mstrText
    End If
    mstrText = mstrText & "("
    If blnSpaceAfter = True Then
        mstrText = mstrText & " "
    End If
    mstrText = mstrText & strField
    If blnSpaceBefore = True Then
        mstrText = mstrText & " "
    End If
    mstrText = mstrText & ")"
End Sub
Public Sub SSF(strField As String, Optional blnSpaceAfter As Boolean = True, Optional blnSpaceBefore As Boolean = True, Optional blnBeginSpace As Boolean = True)
    SQL_Single_Field strField, blnSpaceAfter, blnSpaceBefore, blnBeginSpace
End Sub

''' Append WHERE to SQL statement with first condition field is a Text type
Public Sub WHERE_Text(strField As String, strText As String)
    mstrText = mstrText & " WHERE " & strField & " = '" & strText & "'"
End Sub

''' Append WHERE to SQL statement with first condition field is a Long type
Public Sub WHERE_Long(strField As String, lngNumber As Long)
    mstrText = mstrText & " WHERE " & strField & " = " & lngNumber
End Sub

''' Append WHERE to SQL statement with first condition field is a Integer type
Public Sub WHERE_Integer(strField As String, intNumber As Integer)
    mstrText = mstrText & " WHERE " & strField & " = " & intNumber
End Sub

''' Append WHERE to SQL statement with first condition field is a Boolean type
Public Sub WHERE_Boolean(strField As String, blnBoolean As Boolean)
    'mstrText = mstrText & " WHERE " & strField & " = " & blnBoolean
    If blnBoolean Then
        mstrText = mstrText & " WHERE " & strField & " = Yes"
    Else
        mstrText = mstrText & " WHERE " & strField & " = No"
    End If
End Sub

''' Append WHERE to SQL statement with first condition field is using BETWEEN
Public Sub WHERE_BETWEEN(strField As String, strLeftValue As String, strRightValue As String)
    mstrText = mstrText & " WHERE " & strField & " BETWEEN " & strLeftValue & " AND " & strRightValue
End Sub

''' Append WHERE to SQL statement with first condition field is using LIKE
Public Sub WHERE_LIKE_Text(strField As String, strText As String)
    mstrText = mstrText & " WHERE " & strField & " LIKE '%" & strText & "%'"
End Sub

''' Append OR to SQL statement with condition field is using LIKE
Public Sub OR_LIKE_Text(strField As String, strText As String)
    mstrText = mstrText & " OR " & strField & " LIKE '%" & strText & "%'"
End Sub

''' Append AND to SQL statement with condition field is a Text type
Public Sub AND_Text(strField As String, strText As String)
    mstrText = mstrText & " AND " & strField & " = '" & strText & "'"
End Sub

''' Append AND to SQL statement with condition field is Yes/No
Public Sub AND_Boolean(strField As String, strYesNo As String)
    mstrText = mstrText & " AND " & strField & " = " & strYesNo
End Sub

''' Append ORDER BY to SQL statement
Public Sub ORDER_BY(strField As String, Optional blnAscending As Boolean = True)
    mstrText = mstrText & " ORDER BY " & strField
    If blnAscending = False Then
        mstrText = mstrText & " DESC"
    End If
End Sub

''' Append INNER JOIN to SQL statement
Public Sub INNER_JOIN(strTable1 As String, strTable2 As String, strCommonField1 As String, strCommonField2 As String)
    'SQL "FROM " & strTable1, False
    SQL "INNER JOIN " & strTable2, False
    SQL "ON " & strTable1 & "." & strCommonField1 & " = " & strTable2 & "." & strCommonField2, False
End Sub

''' Append LEFT JOIN to SQL statement
Public Sub LEFT_JOIN(strTable1 As String, strTable2 As String, strCommonField1 As String, strCommonField2 As String)
    'SQL "FROM " & strTable1, False
    SQL "LEFT JOIN " & strTable2, False
    SQL "ON " & strTable1 & "." & strCommonField1 & " = " & strTable2 & "." & strCommonField2, False
End Sub

''' Start building SELECT SQL statement
Public Sub SELECT_(Optional strField As String = "")
    mstrText = "SELECT"
    If strField <> "" Then mstrText = mstrText & " " & strField
End Sub

''' Start building SELECT SQL statement with all fields
Public Sub SELECT_ALL(strTable As String)
    mstrText = "SELECT * FROM " & strTable
End Sub

''' Start building SELECT SQL statement with Top n rows
Public Sub SELECT_TOP(strField As String, strTable As String, Optional intTop As Integer = 1)
    mstrText = "SELECT TOP " & intTop & " " & strField & " FROM " & strTable
End Sub

''' Start building SELECT SQL statement with first field name ID, optional Top n rows
Public Sub SELECT_ID(strTable As String, Optional intTop As Integer = 1)
    If intTop > 0 Then
        mstrText = "SELECT TOP " & intTop & " ID FROM " & strTable
    Else
        mstrText = "SELECT ID FROM " & strTable
    End If
End Sub

''' Append FROM to SQL statement
Public Sub FROM(strTable As String)
    mstrText = mstrText & " FROM " & strTable
End Sub

''' Build INSERT INTO Table SQL statement
Public Sub INSERT(strTable As String, Optional blnSpaceBeforeBracket As Boolean = True, Optional blnOpenBracket As Boolean = False)
    mstrText = "INSERT INTO " & strTable '& " ("
    If blnSpaceBeforeBracket Then mstrText = mstrText & " "
    If blnOpenBracket Then mstrText = mstrText & "("
End Sub

' Result = [ ][)] VALUES [(]
Public Sub VALUES(Optional blnCloseBracket As Boolean = False, Optional blnOpenBracket As Boolean = False, Optional blnBeginSpace As Boolean = True)
    If blnBeginSpace = True Then mstrText = mstrText & " "
    If blnCloseBracket Then mstrText = mstrText & ")"
    mstrText = mstrText & " VALUES " ' Space BEFORE and AFTER is set as default
    If blnOpenBracket Then mstrText = mstrText & "("
End Sub

''' Build UPDATE Table SET SQL statement
Public Sub UPDATE(strTable As String)
    mstrText = "UPDATE " & strTable & " SET"
End Sub

''' Build DELETE FROM Table SQL statement
Public Sub DELETE(strTable As String)
    mstrText = "DELETE FROM " & strTable
End Sub

''' Build DROP Table SQL statement
Public Sub DROP(strTable As String)
    mstrText = "DROP TABLE [" & strTable & "]"
End Sub

''' Build ALTER Table SQL statement
Public Sub ALTER_TABLE(strTable As String)
    mstrText = "ALTER TABLE " & strTable
End Sub

' References:
' http://allenbrowne.com/ser-49.html
' https://docs.microsoft.com/en-us/sql/odbc/microsoft/microsoft-access-data-types
' https://msdn.microsoft.com/en-us/library/aa140015(office.10).aspx (VERY GOOD)
''' Build CREATE Table SQL statement
Public Sub CREATE(strTable As String, Optional strPrefix As String = "")
    mstrText = "CREATE TABLE " & strPrefix & strTable
    mstrText = mstrText & " ("
End Sub

''' Append ID column in CREATE Table SQL statement
Public Sub COLUMN_ID(Optional strColumnName As String = "ID", Optional blnPrimaryKey As Boolean = True, Optional blnAutoIncrement As Boolean = True, Optional blnEndComma As Boolean = True)
    'mstrText = mstrText & "[" & strColumnName & "]"
    mstrText = mstrText & strColumnName
    'If blnAutoIncrement Then mstrText = mstrText & " AUTOINCREMENT"
    If blnAutoIncrement Then
        mstrText = mstrText & " AUTOINCREMENT"
    Else
        mstrText = mstrText & " LONG"
    End If
    If blnPrimaryKey Then mstrText = mstrText & " PRIMARY KEY"
    If blnEndComma = True Then mstrText = mstrText & "," 'mstrText = mstrText & ","
End Sub

''' Append Text type column in CREATE Table SQL statement
Public Sub COLUMN_TEXT(strColumnName As String, Optional intLength As Integer = 255, Optional strDefault As String = "", Optional blnNullable As Boolean = True, Optional blnBeginSpace As Boolean = True, Optional blnEndComma As Boolean = True)
    If blnBeginSpace Then mstrText = mstrText & " "
    mstrText = mstrText & strColumnName & " TEXT(" & intLength & ")"
    If strDefault <> "" Then mstrText = mstrText & " DEFAULT """ & strDefault & """"
    If Not blnNullable Then mstrText = mstrText & " NOT NULL"
    If blnEndComma = True Then mstrText = mstrText & ","
End Sub

''' Append Memo type column in CREATE Table SQL statement
Public Sub COLUMN_MEMO(strColumnName As String, Optional strDefault As String = "", Optional blnNullable As Boolean = True, Optional blnBeginSpace As Boolean = True, Optional blnEndComma As Boolean = True)
    If blnBeginSpace Then mstrText = mstrText & " "
    mstrText = mstrText & strColumnName & " MEMO"
    If strDefault <> "" Then mstrText = mstrText & " DEFAULT " & strDefault
    If Not blnNullable Then mstrText = mstrText & " NOT NULL"
    If blnEndComma = True Then mstrText = mstrText & ","
End Sub

' NOTE: Not yet used or tested
''' Append Number type column in CREATE Table SQL statement
Public Sub COLUMN_NUMBER(strColumnName As String, Optional strFieldSize As String = "LONG", Optional strDefault As String = "", Optional blnNullable As Boolean = True, Optional blnBeginSpace As Boolean = True, Optional blnEndComma As Boolean = True)
    If blnBeginSpace Then mstrText = mstrText & " "
    Select Case strFieldSize
        Case "BYTE"
            mstrText = mstrText & strColumnName & " BYTE"
        Case "SHORT"
            mstrText = mstrText & strColumnName & " SHORT"
        Case "INTEGER" ' Same as SHORT ?
            mstrText = mstrText & strColumnName & " INTEGER"
        Case "LONG" ' Default
            mstrText = mstrText & strColumnName & " LONG"
        Case "SINGLE"
            mstrText = mstrText & strColumnName & " SINGLE"
        Case "DOUBLE"
            mstrText = mstrText & strColumnName & " DOUBLE"
        Case "REPLICA", "GUID"
            mstrText = mstrText & strColumnName & " GUID"
        Case "DECIMAL"
            mstrText = mstrText & strColumnName & " DECIMAL (18, 0)" ' (precision, scale) 9, 4
        Case Else ' LONG
            mstrText = mstrText & strColumnName & " LONG"
    End Select
    If strDefault <> "" Then mstrText = mstrText & " DEFAULT " & strDefault
    If Not blnNullable Then mstrText = mstrText & " NOT NULL"
    If blnEndComma = True Then mstrText = mstrText & ","
End Sub

''' Append Bit type column in CREATE Table SQL statement
Public Sub COLUMN_BIT(strColumnName As String, Optional strDefault As String = "-1", Optional blnNullable As Boolean = True, Optional blnBeginSpace As Boolean = True, Optional blnEndComma As Boolean = True)
    If blnBeginSpace Then mstrText = mstrText & " "
    mstrText = mstrText & strColumnName & " BIT"
    If strDefault <> "" Then mstrText = mstrText & " DEFAULT " & strDefault
    If Not blnNullable Then mstrText = mstrText & " NOT NULL"
    If blnEndComma = True Then mstrText = mstrText & ","
End Sub

' Same as COLUMN_BIT
''' Append YesNo type column in CREATE Table SQL statement
Public Sub COLUMN_YESNO(strColumnName As String, Optional strDefault As String = "Yes", Optional blnNullable As Boolean = True, Optional blnBeginSpace As Boolean = True, Optional blnEndComma As Boolean = True)
    If blnBeginSpace Then mstrText = mstrText & " "
    mstrText = mstrText & strColumnName & " YESNO"
    If strDefault <> "" Then mstrText = mstrText & " DEFAULT " & strDefault
    If Not blnNullable Then mstrText = mstrText & " NOT NULL"
    If blnEndComma = True Then mstrText = mstrText & ","
End Sub

''' Append DateTime type column in CREATE Table SQL statement
Public Sub COLUMN_DATETIME(strColumnName As String, Optional strDefault As String = "", Optional blnNullable As Boolean = True, Optional blnBeginSpace As Boolean = True, Optional blnEndComma As Boolean = True)
    If blnBeginSpace Then mstrText = mstrText & " "
    mstrText = mstrText & strColumnName & " DATETIME"
    If strDefault <> "" Then mstrText = mstrText & " DEFAULT " & strDefault
    If Not blnNullable Then mstrText = mstrText & " NOT NULL"
    If blnEndComma = True Then mstrText = mstrText & ","
End Sub

''' Append a comma character to SQL statement
Public Sub Comma()
    mstrText = mstrText & ","
End Sub

''' Append an open bracket character to SQL statement
Public Sub Open_Bracket()
    mstrText = mstrText & "("
End Sub

''' Append a close bracket character to SQL statement
Public Sub Close_Bracket()
    mstrText = mstrText & ")"
End Sub

''' Appending value when building INSERT statement, default with single quote
' blnAddQuotes = True for Text
' blnAddQuotes = False for Number
Public Sub VAL(strValue As String, Optional blnAddQuotes As Boolean = True, Optional blnEndComma As Boolean = True, Optional blnBeginSpace As Boolean = True)
    If blnBeginSpace = True Then
        mstrText = mstrText & " "
    End If
    If blnAddQuotes Then
        mstrText = mstrText & "'" & strValue & "'"
    Else
        mstrText = mstrText & strValue
    End If
    If blnEndComma = True Then
        mstrText = mstrText & ","
    End If
End Sub
Public Sub VTX(strValue As String, Optional blnAddQuotes As Boolean = True, Optional blnEndComma As Boolean = True, Optional blnBeginSpace As Boolean = True)
    If blnBeginSpace = True Then
        mstrText = mstrText & " "
    End If
    If blnAddQuotes Then
        mstrText = mstrText & "'" & strValue & "'"
    Else
        mstrText = mstrText & strValue
    End If
    If blnEndComma = True Then
        mstrText = mstrText & ","
    End If
End Sub

''' Append First Value in Values when building INSERT SQL statement
Public Sub VOB(strValue As String, Optional blnAddQuotes As Boolean = True, Optional blnEndComma As Boolean = True, Optional blnBeginSpace As Boolean = False, Optional blnOpenBracket As Boolean = True)
    If blnBeginSpace = True Then mstrText = mstrText & " "
    If blnOpenBracket = True Then mstrText = mstrText & "("
    If blnAddQuotes Then
        mstrText = mstrText & "'" & strValue & "'"
    Else
        mstrText = mstrText & strValue
    End If
    If blnEndComma = True Then mstrText = mstrText & ","
End Sub

''' Append Last Value in Values when building INSERT SQL statement
Public Sub VCB(strValue As String, Optional blnAddQuotes As Boolean = True, Optional blnBeginSpace As Boolean = True, Optional blnCloseBracket As Boolean = True)
    If blnBeginSpace = True Then
        mstrText = mstrText & " "
    End If
    If blnAddQuotes Then
        mstrText = mstrText & "'" & strValue & "'"
    Else
        mstrText = mstrText & strValue
    End If
    If blnCloseBracket = True Then mstrText = mstrText & ")"
End Sub

''' Append Single Value in Values when building INSERT SQL statement
Public Sub VSV(strValue As String, Optional blnAddQuotes As Boolean = True, Optional blnBeginSpace As Boolean = True)
    If blnBeginSpace = True Then mstrText = mstrText & " "
    If blnAddQuotes Then
        mstrText = mstrText & "('" & strValue & "')"
    Else
        mstrText = mstrText & "(" & strValue & ")"
    End If
End Sub

''' Appending Double type value when building INSERT statement
Public Sub VDB(dblNumber As Double, Optional blnEndComma As Boolean = True)
    mstrText = mstrText & " " & dblNumber
    If blnEndComma = True Then
        mstrText = mstrText & ","
    End If
End Sub

''' Appending Long type value when building INSERT statement
Public Sub VLN(lngNumber As Long, Optional blnEndComma As Boolean = True)
    mstrText = mstrText & " " & lngNumber
    If blnEndComma = True Then
        mstrText = mstrText & ","
    End If
End Sub

''' Appending Integer type value when building INSERT statement
Public Sub VIN(intNumber As Integer, Optional blnEndComma As Boolean = True)
    mstrText = mstrText & " " & intNumber
    If blnEndComma = True Then
        mstrText = mstrText & ","
    End If
End Sub

''' Appending Boolean type value when building INSERT statement
Public Sub VBL(blnValue As Boolean, Optional blnEndComma As Boolean = True)
    If blnValue Then
        mstrText = mstrText & " TRUE"
    Else
        mstrText = mstrText & " FALSE"
    End If
    If blnEndComma = True Then
        mstrText = mstrText & ","
    End If
End Sub

''' Appending DateTime type value when building INSERT statement
Public Sub VDT(strDateTime As String, Optional blnEndComma As Boolean = True)
    mstrText = mstrText & " #" & strDateTime & "#"
    If blnEndComma = True Then
        mstrText = mstrText & ","
    End If
End Sub
Public Sub DAT(strValue As String, Optional blnAddHatches As Boolean = True, Optional blnEndComma As Boolean = True, Optional blnBeginSpace As Boolean = True)
    If blnBeginSpace = True Then
        mstrText = mstrText & " "
    End If
    If blnAddHatches Then
        mstrText = mstrText & "#" & strValue & "#"
    Else
        mstrText = mstrText & strValue
    End If
    If blnEndComma = True Then
        mstrText = mstrText & ","
    End If
End Sub
