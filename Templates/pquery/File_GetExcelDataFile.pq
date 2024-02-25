/**
Call table from excel file in INNO folder
name = price_Latest file name without extension.
result = get_ExcelDataFile(name)

Arguments:

fileName: File name without extension

tableName: Table name (table if null)
**/

(fileName as text, optional tableName as text) as table =>
let
    folder = "C:\Users\Laktionov.I\OneDrive\INNO\",
    filepath = folder & fileName &".xlsx",
    tblName = tableName ?? "table",

    data = Excel.Workbook(File.Contents(filepath),null,true),
    dataTable = data{[ Item=tblName , Kind="Table" ]}[Data],

    fnDetectAndTransformTypes = (SourceTable as table) =>
    let
        moneyList = {"amount", "price", "deal", "cost", "rebate", "msrp"},

        FnDetectType = (SourceTable, ColumnName) =>
        let ListColumn = Table.Column(SourceTable, ColumnName),
            Money = List.Contains(moneyList, Text.Trim(ColumnName), (x,y) => Text.StartsWith(y,x,Comparer.OrdinalIgnoreCase)),

            NumberPercentage =
                List.Count(
                    List.Select(ListColumn , each _ is number)
                ) / List.Count(ListColumn),
            IntPercentage =
                List.Count(
                    List.Select(ListColumn , each _ is number and _ = Number.Round(_, 0))
                ) / List.Count(ListColumn),
            DatePercentage =
                List.Count(
                    List.Select(ListColumn , each _ is date)
                ) / List.Count(ListColumn),
            DateTimePercentage =
                List.Count(
                    List.Select(ListColumn , each _ is datetime)
                ) / List.Count(ListColumn),
            Max = List.Max(
                {NumberPercentage, DatePercentage, DateTimePercentage, IntPercentage}
            ),
            Result =
                if Money then
                    {ColumnName, Currency.Type}
                else if Max < .95 then
                    {ColumnName, type any}
                else if DateTimePercentage = Max then
                    {ColumnName, type datetime}
                else if DatePercentage = Max then
                    {ColumnName, type date}
                else if IntPercentage = Max then
                    {ColumnName, Int64.Type}
                else
                    {ColumnName, type number}
        in  Result,

        ConvertTypesList =
            List.Transform(
                Table.ColumnNames(SourceTable),
                each FnDetectType(SourceTable, _)
            ),

        ConvertedTable = Table.TransformColumnTypes(SourceTable, ConvertTypesList)
    in
        ConvertedTable,

    Result = fnDetectAndTransformTypes(dataTable)
in
    Result

