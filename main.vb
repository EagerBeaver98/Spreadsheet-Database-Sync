Option Explicit

Dim intStartRows As Integer 'the start row number in the excel sheet
Dim intEndRows As Integer 'the end row number in the excel sheet
Dim intStartCols As Integer 'the start column number in the sheet'
Dim intEndCols As Integer   'the end column number in the sheet'
Dim strSQLInsert As String
Dim strSQLUpdate As String
Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim Values As String
Dim TableHeaders As String
Dim TableNames As Variant

Dim Counter As Integer


' Variables for database access.
' DAO would be similar. This is ADO
Dim adoConnection, adoRecordset
Public Const strConn = "PROVIDER=SQLOLEDB;SERVER=SERVER;" & _
                        "DATABASE=DATABASE;UID=USER;PWD=PASSWORD"


Public Function CountTableRows(TableName As String) As Integer
    'Returns integer for number of rows in Table
    'MsgBox (TableName)
    Sheets(TableName).Activate
    CountTableRows = Range("A3").End(xlDown).Row - 1 '-1 because of table
    
    
End Function

Public Function CountTableCols(TableName As String) As Integer
    'Returns integer for number of columns in Table
    Sheets(TableName).Activate
    CountTableCols = Range("A2").End(xlToRight).Column
    
End Function


'open the database connection
Public Sub OpenDatabase()
    
    ' Initialize the database connection.
    Set adoConnection = CreateObject("ADODB.Connection")
    ' Open the database, use Microsoft Jet OLEDB data provider
    adoConnection.Open strConn
    
End Sub

'used to close the database connection
Public Sub CloseDatabase()
    
    'close the connections
    'adoRecordset.Close
    adoConnection.Close
    'destroy these variables
    'Set adoRecordset = Nothing
    Set adoConnection = Nothing
    
End Sub

Public Function FormatTableName(TableName As String) As String
    'Formats table name to sql syntax
    FormatTableName = Replace(TableName, " ", "_")
End Function

Public Function DateFormat(ColumnName As String) As String
    'Formats date columns into wanted format
    DateFormat = Format(ColumnName, "mmm/yy")
End Function

Public Function ClearTableQuery(TableName As String) As String
    ClearTableQuery = "DELETE FROM " & FormatTableName(TableName) & ";"
End Function

Public Sub ProvForecastWriter(TableName As String)
    Dim SQLQuery As String
    Dim ProvName As String
    Dim MoneyData As String
    
    ProvName = Replace(Left(TableName, 3), " ", "")
    intStartRows = 3
    intEndRows = CountTableRows(TableName)
    intStartCols = 1
    intEndCols = CountTableCols(TableName)
    Sheets(TableName).Activate
    
    adoConnection.Execute ("DELETE FROM PBI_Data WHERE [PROV] = '" & ProvName) & "'"
    For i = 3 To intEndRows
    Values = ""
        For j = 1 To intEndCols
            If j = 1 Then
                Values = "'" & ProvName & "', " & _
                Left(Cells(i, j).Value, 4) & ", " & _
                Mid(Cells(i, j).Value, 6, 4) & ", " & _
                Left(Cells(i, j).Value, 9) & ", " & _
                "'" & Mid(Cells(i, j).Value, 11) & "', " & _
                "'" & Mid(Cells(i, j).Value, 6, 4) & " - " & Mid(Cells(i, j).Value, 11) & "', "
            ElseIf VarType(Cells(i, j).Value) = 0 Or VarType(Cells(i, j).Value) = 1 Then
                
            ElseIf Right(Cells(2, j).Value, 5) <> "Total" Then
                MoneyData = Cells(i, j).Value & ", " & Right(Cells(2, j).Value, 4) & ", " & _
                "'" & Left(DateFormat(Cells(2, j).Value), 3) & "'"
                SQLQuery = "INSERT INTO PBI_Data VALUES(" & Values & MoneyData & ");"
                adoConnection.Execute (SQLQuery)
            End If
        Next j
        
    Next i
End Sub


Public Sub AFEForecastWriter(TableName As String)
    Dim SQLQuery As String
    Dim MoneyData As String
    intStartRows = 3
    intEndRows = CountTableRows(TableName)
    intStartCols = 1
    intEndCols = CountTableCols(TableName)
    Sheets(TableName).Activate
    For x = 3 To intEndRows + 1
        adoConnection.Execute ("DELETE FROM PBI_AFE_Data WHERE [TYPE] = '" & Cells(x, 1).Value & "'")
    Next x
    For i = 3 To intEndRows
    Values = ""
        For j = 1 To intEndCols
            If j = 5 Then
                Values = Values & "'" & Cells(i, j).Value & "', " & Cells(i, 4).Value & "." & Cells(i, j).Value & ", "
            ElseIf j = 6 Then
                Values = Values & "'" & Cells(i, 5).Value & " - " & Cells(i, j).Value & "', '" & Cells(i, j).Value & "', "
            ElseIf j < 7 Then
                Values = Values & "'" & Cells(i, j).Value & "', "
            ElseIf VarType(Cells(i, j).Value) = 0 Or VarType(Cells(i, j).Value) = 1 Then
                
            ElseIf Right(Cells(2, j).Value, 5) <> "Total" Then
                MoneyData = Cells(i, j).Value & ", " & Right(Cells(2, j).Value, 4) & ", " & _
                "'" & Left(DateFormat(Cells(2, j).Value), 3) & "'"
                SQLQuery = "INSERT INTO PBI_AFE_Data VALUES(" & Values & MoneyData & ");"
                adoConnection.Execute (SQLQuery)
            End If
        Next j
        
    Next i
End Sub

Public Sub Button()
    OpenDatabase
    Select Case ActiveSheet.Name
    Case "AB Forecast", "SK Forecast", "GEN Forecast"
        ProvForecastWriter (ActiveSheet.Name)
    Case "AR AFE Forecast", "WE AFE Forecast", "WC AFE Forecast", "RA AFE Forecast", "RC AFE Forecast", _
        "EQ AFE Forecast", "DR_CM AFE Forecast", "PL_FC AFE Forecast", "AQ_DP_LN_GG AFE Forecast"
        AFEForecastWriter (ActiveSheet.Name)
    End Select
    CloseDatabase
End Sub

Public Sub SyncAll()
    TableNames = Array("AB Forecast", "SK Forecast", "GEN Forecast", "AR AFE Forecast", "WE AFE Forecast", _
        "WC AFE Forecast", "RA AFE Forecast", "RC AFE Forecast", "EQ AFE Forecast", "DR_CM AFE Forecast", "PL_FC AFE Forecast", _
        "AQ_DP_LN_GG AFE Forecast")
    Dim item As Variant
    OpenDatabase
    For Each item In TableNames
        Select Case item
    Case "AB Forecast", "SK Forecast", "GEN Forecast"
        ProvForecastWriter (item)
    Case "AR AFE Forecast", "WE AFE Forecast", "WC AFE Forecast", "RA AFE Forecast", "RC AFE Forecast", _
        "EQ AFE Forecast", "DR_CM AFE Forecast", "PL_FC AFE Forecast", "AQ_DP_LN_GG AFE Forecast"
        AFEForecastWriter (item)
    End Select
    Next item
    CloseDatabase
    Sheets("Control Panel").Activate
End Sub

