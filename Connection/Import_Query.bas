Option Explicit

Public Sub ImportQueryLPQ()
    Dim CurrentWorkbook As Excel.Workbook
    Dim Query As Excel.WorkbookQuery
    Dim Name, Code, ReplaceQuery
    Name = "LPQ"
    Code = "Expression.Evaluate(Text.FromBinary(Web.Contents(Json.Document(" & _
            "Web.Contents(""https://api.github.com/gists/700a6d65e098189881ecd77e585b233a""))[files][LPQ.pq][raw_url])), #shared)"
    Set CurrentWorkbook = ActiveWorkbook
    Set Query = Nothing
    On Error Resume Next
    Set Query = CurrentWorkbook.Queries(Name)
    On Error GoTo 0
    If Query Is Nothing Then
        Set Query = CurrentWorkbook.Queries.Add(Name, Code)
    Else
        Query.Formula = Code
    End If
    Set ReplaceQuery = Query
End Sub
