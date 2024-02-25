Option Explicit
Const qryStr$ = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="
Const prefix$ = "Power Query - "

'''''''''''''''''''''
' LOADING
'''''''''''''''''''''

'load a query to a sheet and/or PowerPivot data model based on the ticked checkboxes in the Introduction sheet
'Sub LoadQuery(q$)
'ConfirmQueryExists q
'If ToPP() Then LoadQueryToPP q
'If ToSheet() Then LoadQueryToSheet q
'End Sub

'load a query to the PowerPivot data model
'based on the approach by Gil Raviv: https://gallery.technet.microsoft.com/VBA-to-automate-Power-956a52d1
Sub LoadQueryToPP(q$)
    Dim name$
    name = prefix & q
    If ConnExists(name) Then
        If ActiveWorkbook.Connections(name).InModel Then
            RefreshPPTable name
            Exit Sub
        Else
            ' recreate, since InModel is read-only
            ActiveWorkbook.Connections(name).Delete
        End If
    End If
    ActiveWorkbook.Connections.Add2 name, "PQ-PP connection: " & q, qryStr & q, """" & q & """", xlCmdTableCollection, True, False
End Sub

'load a query to a given range (default: A1) and sheet (default: new sheet named after the query)
Sub LoadQueryToSheet(q$, Optional destRange As Range, Optional ws As Worksheet)
    If SheetExists(q) Then 'And ConnExists(prefix & q) Then
        Dim cn As WorkbookConnection
        For Each cn In ActiveWorkbook.Connections
            If ConnMatchesQuery(cn, q) Then RefreshQueryConn cn.name
        Next
    Else
        Dim lo As ListObject
        If ws Is Nothing Then
            Set ws = Sheets.Add(After:=lastSheet(ActiveWorkbook))
            On Error Resume Next
            ws.name = Left$(q, 31)
            On Error GoTo 0
        End If
        If destRange Is Nothing Then Set destRange = ws.[A1]
        Set lo = ws.ListObjects.Add(xlSrcQuery, qryStr & q, Destination:=destRange)
        With lo.QueryTable
            .CommandType = xlCmdDefault
            .CommandText = Array("SELECT * FROM [" & q & "]")
            .BackgroundQuery = True
            On Error Resume Next
            .ListObject.DisplayName = q
            On Error GoTo 0
            .Refresh
        End With
    End If
End Sub


'''''''''''''''''''''
' DELETING
'''''''''''''''''''''

'Delete all queries in the active workbook
Public Sub DeleteQueries()
    If tryMsgBox("Are you sure you wish to delete all queries from the active workbook?", vbYesNo, "DELETE QUERIES?") <> vbYes Then Exit Sub
    Dim q
    For Each q In ActiveWorkbook.Queries
    '    If toSheet() Then ActiveWorkbook.Sheets(q.name).Delete
        If ConnExists(prefix & q.name) Then DeleteQueryConn (prefix & q.name)    'If ToPP() Then
    '    DeleteQuery (q.name)
        q.Delete
    Next
End Sub

'Delete the given query in the active workbook, as well as associated sheets/connections (if boxes checked)
Sub DeleteQuery(name$)
    ConfirmQueryExists name
    ActiveWorkbook.Queries(name).Delete
    'If ToSheet() And SheetExists(name, ActiveWorkbook) Then ActiveWorkbook.Sheets(name).Delete
    DeleteQueryConn (prefix & name)  'If ToPP() Then
End Sub

'Delete the connection for a given query from the active workbook
Sub DeleteQueryConn(name$)
    Dim cn As WorkbookConnection
    For Each cn In ActiveWorkbook.Connections
        If ConnMatchesQuery(cn, name) Then cn.Delete
    Next
End Sub


'''''''''''''''''''''
' REFRESHING
'''''''''''''''''''''

'Refreshes all Power Query scripts
Sub RefreshQueries()
    'If ToSheet() Then RefreshQueriesConn 'RefreshQueriesToWB
    'If ToPP() Then RefreshPowerPivot
    RefreshQueriesConn
    RefreshPowerPivot
End Sub

'Refreshes a given Power Query script
Sub RefreshQuery(q$)
    ConfirmQueryExists q
    'If ToSheet() Then RefreshQueryConn q 'RefreshQueryToWB q
    'If ToSheet() Then RefreshQueriesConn 'RefreshQueriesToWB
    'If ToPP() Then RefreshPPTable q
    RefreshQueriesConn
    On Error Resume Next    'might not be loaded to powerpivot yet
    RefreshPPTable q
    On Error GoTo 0
End Sub

'Refresh whole PowerPivot model
'see the efforts of Tom 'Goban Saor' Gleeson: http://www.tomgleeson.ie/posts/201404/PowerPivotVBARefresh.html
Sub RefreshPowerPivot()
    ActiveWorkbook.Model.Initialize
    ActiveWorkbook.Model.Refresh
End Sub

'Refresh particular PowerPivot table
Sub RefreshPPTable(name$)
    ActiveWorkbook.Model.Initialize
    ActiveWorkbook.Connections(prefix & name).Refresh
End Sub

'Refreshes all Power Query scripts (by Connection)
'Adapted from original by Ken Puls: http://www.excelguru.ca/blog/2014/10/22/refresh-power-query-with-vba/
Public Sub RefreshQueriesConn()
    Dim cn As WorkbookConnection
    For Each cn In ActiveWorkbook.Connections
        If startsWith(cn.name, prefix) Then cn.Refresh
    Next
End Sub

'Refreshes a given Power Query script (by Connection)
Sub RefreshQueryConn(q$)
    Dim cn As WorkbookConnection
    For Each cn In ActiveWorkbook.Connections
        If ConnMatchesQuery(cn, q) Then
            'If cn.InModel Then
            ' skip separately created PowerPivot connections here to avoid getting stuck due to simultaneous access
            If cn.OLEDBConnection.CommandType = xlCmdDefault Then _
                cn.Refresh
            End If
    Next
End Sub

'Alternate way, seems not currently compatible with Excel 2016's transient VBA-loaded tables:
'Refreshes all Power Query sheet connections (by ListObject)
'Sub RefreshQueriesToWB()
'Dim ws As Worksheet, lo As ListObject
'For Each ws In ActiveWorkbook.Worksheets
'    For Each lo In ws.ListObjects
'        If startsWith(lo.QueryTable.Connection, qryStr) Then lo.QueryTable.Refresh True
'    Next
'Next
'End Sub

'Refreshes a given Power Query sheet connection (by ListObject)
'Sub RefreshQueryToWB(q$)
'Dim ws As Worksheet, lo As ListObject, qt As QueryTable
'For Each ws In ActiveWorkbook.Worksheets
'    For Each lo In ws.ListObjects
'        'sometimes an extra ;Extended Properties="" at the end
'        'If startsWith(lo.QueryTable.Connection, qryStr & q) Then lo.QueryTable.Refresh True
'        If lo.name = q Then lo.QueryTable.Refresh True  'If hasSubstr(lo.QueryTable.Connection, q)
''        If lo.name = q Then
''            lo.Refresh
''            'Set qt = lo.QueryTable
''            'qt.Refresh True
''        End If
'    Next
'Next
'End Sub


'''''''''''''''''''''
' MOVING
'''''''''''''''''''''

'import all queries from the given workbooks to the active workbook
Public Sub ImportQueriesFromExcel()
    Dim overwrite As Boolean, fd As FileDialog, file, fromWB As Workbook, toWB As Workbook
    Set toWB = ActiveWorkbook
    'overwrite = tryMsgBox("Would you like to overwrite any existing queries?", vbYesNo, "OVERWRITE QUERIES?") = vbYes
    overwrite = doOverwrite()
    Set fd = openFiles(, True, "Please select any Excel workbooks to import queries from", ThisWorkbook.Sheets(1).[LoadPath])
    For Each file In fd.SelectedItems
        Set fromWB = getWB(CStr(file))
        TransferQueries fromWB, toWB, overwrite
    Next
    'MsgBox "Done!"
End Sub

'export all queries from the active workbook to the given workbooks
Public Sub ExportQueriesToExcel()
    Dim overwrite As Boolean, fd As FileDialog, file, fromWB As Workbook, toWB As Workbook
    Set fromWB = ActiveWorkbook
    'overwrite = tryMsgBox("Would you like to overwrite any existing queries?", vbYesNo, "OVERWRITE QUERIES?") = vbYes
    overwrite = doOverwrite()
    Set fd = openFiles(, True, "Please select any Excel workbooks to export queries to")
    For Each file In fd.SelectedItems
        Set toWB = getWB(CStr(file))
        TransferQueries fromWB, toWB, overwrite
    Next
    'MsgBox "Done!"
End Sub

'export the specified query from the active workbook to the given workbooks
Public Sub ExportQueryToExcel(name$)
    Dim overwrite As Boolean, fd As FileDialog, file, fromWB As Workbook, toWB As Workbook
    Set fromWB = ActiveWorkbook
    ConfirmQueryExists name, fromWB
    'overwrite = tryMsgBox("Would you like to overwrite any existing queries?", vbYesNo, "OVERWRITE QUERIES?") = vbYes
    overwrite = doOverwrite()
    Set fd = openFiles(, True, "Please select any Excel workbooks to export queries to")
    For Each file In fd.SelectedItems
        Set toWB = getWB(CStr(file))
        TransferQuery name, fromWB, toWB, overwrite
    Next
    'MsgBox "Done!"
End Sub

'transfer all Power Query queries from one workbook to another
Sub TransferQueries(Optional fromWB As Workbook, Optional toWB As Workbook, Optional overwrite As Boolean = False)
    Dim q
    If fromWB Is Nothing Then Set fromWB = ThisWorkbook
    If toWB Is Nothing Then Set toWB = ActiveWorkbook
    If fromWB.FullName = toWB.FullName Then Exit Sub
    For Each q In fromWB.Queries
        If QueryExists(q.name, toWB) Then
            If overwrite Then
                toWB.Queries(q.name).Delete
            Else
                GoTo skip:
            End If
        End If
        toWB.Queries.Add q.name, q.Formula, q.description
    skip:
    Next
End Sub

'transfer a specific Power Query query (by name) from one workbook to another
Sub TransferQuery(name$, Optional fromWB As Workbook, Optional toWB As Workbook, Optional overwrite As Boolean = False)
    Dim q
    If fromWB Is Nothing Then Set fromWB = ThisWorkbook
    If toWB Is Nothing Then Set toWB = ActiveWorkbook
    If fromWB.FullName = toWB.FullName Then Exit Sub
    If QueryExists(name, toWB) Then
        If overwrite Then
            toWB.Queries(name).Delete
        Else
            Exit Sub
        End If
    End If
    Set q = fromWB.Queries(name)
    toWB.Queries.Add q.name, q.Formula, q.description
End Sub

'import Power Query queries from selected files into the active workbook
Sub ImportQueriesFromFiles()
    Dim fd As FileDialog, file$, i&, name$, overwrite As Boolean, wb As Workbook
    Set wb = ActiveWorkbook
    'overwrite = tryMsgBox("Would you like to overwrite any existing queries?", vbYesNo, "OVERWRITE QUERIES?") = vbYes
    overwrite = doOverwrite()
    Set fd = openFiles(, , "Please select any M queries to import")
    For i = 1 To fd.SelectedItems.Count
        file = fd.SelectedItems(i)
        name = Replace(fileIn(file, False), ".", "_")
        If QueryExists(name, wb) Then
            If overwrite Then
                wb.Queries(name).Delete
            Else
                GoTo skip:
            End If
        End If
        wb.Queries.Add name, ReadFile(file)
    skip:
    Next
End Sub

'export Power Query queries from the active workbook into files in the selected folder -- overwrites existing files with the same names!
Public Sub ExportQueriesToFiles()
    Dim overwrite As Boolean, path$, full$, q
    overwrite = doOverwrite()
    path = getPath("Please select a folder to export the queries to", ThisWorkbook.Sheets(1).[LoadPath])
    For Each q In ActiveWorkbook.Queries
        full = path & Replace(q.name, "_", ".") & ".pq"
        If overwrite Or Not PathExists(full) Then write2File q.Formula, full
    Next
End Sub

'export specified Power Query query from the active workbook into a file in the selected folder -- overwrites existing files with the same names!
Public Sub ExportQueryToFile(name$)
    Dim overwrite As Boolean, path$, full$, q As WorkbookQuery
    overwrite = doOverwrite()
    path = getPath("Please select a folder to export the queries to", ThisWorkbook.Sheets(1).[LoadPath])
    ConfirmQueryExists (name)
    Set q = ActiveWorkbook.Queries(name)
    full = path & Replace(q.name, "_", ".") & ".pq"
    If overwrite Or Not PathExists(full) Then write2File q.Formula, full
End Sub

''''''''''''''''''''''''''''''''''
' INFO
''''''''''''''''''''''''''''''''''

' confirm a given query exists before continuing
Sub ConfirmQueryExists(q$, Optional wb As Workbook)
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If Not QueryExists(q, wb) Then
        MsgBox "No query named " & q & " in workbook " & wb.name & "!"
        End
    End If
End Sub

' check if a given query exists in the given workbook
Function QueryExists(q$, Optional wb As Workbook) As Boolean
    If wb Is Nothing Then Set wb = ActiveWorkbook
    On Error Resume Next
    QueryExists = CBool(Len(wb.Queries(q).name))
End Function

' check if a given connection exists in the given workbook
Function ConnExists(name$, Optional wb As Workbook) As Boolean
    If wb Is Nothing Then Set wb = ActiveWorkbook
    On Error Resume Next
    ConnExists = CBool(Len(wb.Connections(name).name))
End Function

' check whether loading to (/ deleting) sheets is enabled
'Function ToSheet() As Boolean
'ToSheet = ThisWorkbook.Sheets(1).[load2sheet] > 0
'End Function

' check whether loading to PowerPivot (/ deleting connections) is enabled
'Function ToPP() As Boolean
'ToPP = ThisWorkbook.Sheets(1).[load2pp] > 0
'End Function

' check whether to overwrite queries during import/export
Function doOverwrite() As Boolean
    doOverwrite = ThisWorkbook.Sheets(1).[overwrite] > 0
End Function

' check whether a given connection corresponds to a certain query
Function ConnMatchesQuery(con As WorkbookConnection, name$) As Boolean
    Dim conString$
    ConnMatchesQuery = False
    If con.name = prefix & name Then
        ConnMatchesQuery = True
    ElseIf con.Type = xlConnectionTypeOLEDB Then 'Not IsNull(con.OLEDBConnection) Then
        conString = con.OLEDBConnection.Connection
        ' don't just check for containment -- some a query name "a" could be contained in a query name "abc"!
        If endsWith(conString, "Location=" & name) Or hasSubstr(conString, "Location=" & name & ";") Then ConnMatchesQuery = True
    End If
End Function

'Sub checkconns()
'Dim con, ole
'For Each con In ActiveWorkbook.Connections
'    If con.Type = xlConnectionTypeOLEDB Then Set ole = con.OLEDBConnection
'Next
'End Sub

''''''''''''''''''''''''''''''''''
' HELPER FUNCTIONS
''''''''''''''''''''''''''''''''''

' check if a given workbook is open
Function wbExists(name$) As Boolean
    name = fileIn(name)
    On Error Resume Next
    wbExists = CBool(Len(Workbooks(name).name))
End Function

' check if a sheet exists in a workbook
Function SheetExists(sName$, Optional wb As Workbook = Nothing) As Boolean
    If wb Is Nothing Then Set wb = ActiveWorkbook
    On Error Resume Next
    SheetExists = CBool(Len(wb.Sheets(sName).name))
End Function

' check if a given file path exists
Public Function PathExists(path$) As Boolean
    On Error Resume Next
    PathExists = path <> "" And Dir(path, vbDirectory) <> ""
End Function

'Use or open a workbook with a given full path that may or may not be open already
Function getWB(name$) As Workbook
    If Not wbExists(name) Then Workbooks.Open name
    Set getWB = Workbooks(fileIn(name))
End Function

'return the results of a file dialog
Function openFiles(Optional multiple As Boolean = True, Optional filterXL As Boolean = False, Optional title$ = "Please select the files", Optional default$) As FileDialog
    Set openFiles = Application.FileDialog(msoFileDialogFilePicker)
    openFiles.AllowMultiSelect = multiple
    openFiles.title = title
    If default <> "" Then openFiles.InitialFileName = default
    openFiles.Filters.Clear
    If filterXL Then openFiles.Filters.Add "Excel files", "*.xl*;*.csv"
    If openFiles.Show = 0 Then End
End Function

'Reference: Microsoft ActiveX Data Objects 6.1 Library
'return the contents of a file at a given path as a string
Function ReadFile$(Optional path$ = "", Optional encoding$ = "UTF-8")
    With New ADODB.Stream
        .Charset = encoding
        .Open
        .LoadFromFile path
        ReadFile = .ReadText
    End With
End Function

'Reference: Microsoft ActiveX Data Objects 6.1 Library
'write a string to a file at a given path -- it will be created if it does not exist yet
Function write2File(s$, Optional file$ = "", Optional encoding$ = "UTF-8")
    'If file = "" Then file = pickFile
    checkFolder PathIn(file)
    With New ADODB.Stream
        .Type = adTypeText
        .Charset = encoding
        .Open
        .WriteText (s)
        .SaveToFile file, adSaveCreateOverWrite
    End With
End Function

'confirms the existence of a folder, creating it if it hadn't been yet
Sub checkFolder(path$)  'path name should include a slash (optionally with file name)
    Dim parent$
    parent = stringUntil(path, "\", True, , -1)
    If Not PathExists(parent) Then
        checkFolder parent
        MkDir parent
    End If
End Sub

'get the file of a path/url
Public Function fileIn$(path$, Optional inclExt As Boolean = True)
    fileIn = path
    fileIn = stringFrom(fileIn, "/", True, , 1)
    fileIn = stringFrom(fileIn, "\", True, , 1)
    If Not inclExt Then fileIn = stringUntil(fileIn, ".", True, , -1)
End Function

'get the directory of a path/url
Public Function PathIn$(path$, Optional inclSlash As Boolean = True)
    Dim s$: s = path
    s = stringUntil(s, "/", True, , IIf(inclSlash, 0, -1))
    s = stringUntil(s, "\", True, , IIf(inclSlash, 0, -1))
    PathIn = s
End Function

'select part of a string starting from a given substring up to the end
Public Function stringFrom$(stack$, needle$, Optional rev As Boolean = False, Optional from& = 1, Optional Offset& = 0, Optional ignore As Boolean = False)
    On Error GoTo Handler
    If rev Then
        stringFrom = Mid$(stack, InStrRev(stack, needle) + Offset)
    Else
        stringFrom = Mid$(stack, InStr(from, stack, needle) + Offset)
    End If
    Exit Function
    Handler:
    If ignore = True Then stringFrom = stack
End Function

'select part of a string from the start and ending with a given substring
Public Function stringUntil$(stack$, needle$, Optional rev As Boolean = False, Optional from& = 1, Optional Offset& = 0, Optional noBlank As Boolean = False)
    On Error GoTo Handler
    Dim Pos&
    If rev Then
        Pos = InStrRev(stack, needle)
    Else
        Pos = InStr(from, stack, needle)
    End If
    If Pos > 0 And (Pos > 1 Or noBlank = False) Then
        stringUntil = Mid$(stack, from, Pos + Offset)
    Else
        stringUntil = Mid$(stack, from)
    End If
    Exit Function
    Handler:
    If noBlank = True Then stringUntil = stack
End Function

'check for substring containment
Public Function hasSubstr(haystack$, needle$, Optional compareMode& = vbTextCompare) As Boolean
    hasSubstr = False
    If needle = "" Then Exit Function
    On Error Resume Next
    hasSubstr = InStr(1, haystack, needle, compareMode) > 0
End Function

'return a folder selected by the user from a dialog
Function getPath$(Optional title$ = "Please pick a folder", Optional default$)
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = title
        If default <> "" Then .InitialFileName = default
        If .Show = 0 Then End
        getPath = .SelectedItems(1)
    End With
    getPath = getPath & "\"
End Function

'returns the last sheet of a given workbook
Function lastSheet(Optional wb As Workbook = Nothing) As Worksheet
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Set lastSheet = wb.Sheets(wb.Sheets.Count)
End Function

'check whether a string starts with a certain substring
Function startsWith(haystack$, needle$) As Boolean
    startsWith = Left$(haystack, Len(needle)) = needle
End Function

'check whether a string ends with a certain substring
Function endsWith(haystack$, needle$) As Boolean
    endsWith = Right$(haystack, Len(needle)) = needle
End Function

'show a given message box and return the result, terminating macro execution in case the user clicks close instead
Function tryMsgBox(Prompt$, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional title$ = "?") As Variant
    tryMsgBox = MsgBox(Prompt, Buttons, title)
    If tryMsgBox = vbCancel Then End
End Function

