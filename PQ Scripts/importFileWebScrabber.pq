(filepath as text) =>
    let
        // Видалення [] із тексту
        fnCleanListGrp = (txt) => Text.Remove(txt, {"[", "]"}),
        // Переіменування стовбчиків
        fxRename = (tbl as table, listMatch as list, newName as list) =>
            let s = List.Select( Table.ColumnNames(tbl), (x) => List.AnyTrue(List.Transform(listMatch, each Text.Contains(x, _))) ),
                Indices = List.PositionOf(Table.ColumnNames(tbl), s{0}),
                ColumnNames = Table.ColumnNames(tbl),
                ColumnNamesList = List.Transform({Indices}, each ColumnNames{_}),
                ZippedList = List.Zip({ColumnNamesList, newName})
                //ColumnNamesNew
            in  Table.RenameColumns(tbl, ZippedList),
        //
        fnHeaders = (tbl) =>
            let columns_filter = {"web-scraper-", "next-pagination"},
                columns_new = {
                    "gpu",
                    "brand",
                    "model",
                    "name",
                    "memory",
                    "mem",
                    "overclock",
                    "ocl",
                    "oc",
                    "avail",
                    "price",
                    "price_novat",
                    "cooling",
                    "fans",
                    "rgb",
                    "rating",
                    "reviews",
                    "url",
                    "link-href",
                    "partnumber",
                    "pn",
                    "kod"
                },
                columns_Filtered = List.Select(Table.ColumnNames(tbl), each
                    not List.Contains(columns_filter, _, (x,y) => Text.StartsWith(y,x))
                ),
                TBL = Table.SelectColumns(tbl, columns_Filtered),
                columns_src = Table.ColumnNames(TBL),
                Indexed_Generate = List.Generate(
                    () => [x = 0, Name = columns_src{x}, Pos = List.PositionOf(columns_new, Name)],
                    each [x] < List.Count(columns_src),
                    each [
                        x = [x] + 1,
                        Name = columns_src{x},
                        Pos = List.PositionOf(columns_new, columns_src{x})
                    ],
                    each ({[Name], [x], [Pos]})
                ),
                Filter_notFound = List.Select(Indexed_Generate, each _{2} >= 0),
                Sort_byNewIndex = List.Sort(Filter_notFound, {each _{2}, Order.Ascending}),
                Delete_Columns = Table.SelectColumns(TBL, List.Transform(Sort_byNewIndex, each _{0})),
                Rename_Columns = fxRename(Delete_Columns, {"partnumber"}, {"pn"}),
                Rename_Columns2 = fxRename(Rename_Columns, {"overclock", "ocl"}, {"oc"}),
                rating = Table.ReplaceValue(Rename_Columns2,
                    each [rating],
                    each
                        if Text.StartsWith([rating], "[") then
                            List.NonNullCount(Text.Split(fnCleanListGrp([rating]), ","))
                        else
                            Text.Select(Text.Replace([rating], ".", ","), {"0".."9", ","}),
                    Replacer.ReplaceValue,
                    {"rating"}
                ),
                cooling = Table.ReplaceValue(rating,
                    each [cooling],
                    each
                        if Text.Contains([cooling], ",") then
                            Text.Combine(
                                List.Transform(
                                    Text.Split(fnCleanListGrp([cooling]), ","),
                                    each Text.BetweenDelimiters(_, ":""", """}")
                                ),
                                ", "
                            )
                        else
                            null,
                    Replacer.ReplaceValue,
                    {"cooling"}
                ),
                resTable = Table.TransformColumns(cooling, {
                    {"brand", each Text.Upper(_), type text},
                    {"model", each Text.Clean(_), type text},
                    {"price", each fnPropNum(_), Currency.Type},
                    //{ "price_novat", each fnPropNum(_), Currency.Type },
                    {"memory", each Number.FromText(_), Int64.Type},
                    //{"fans", each Number.FromText(_), Int64.Type},
                    //{"kod", each _, type text},
                    //{"rating", each fnPropNum(_), type number},
                    {"reviews", each fnPropNum(_), Int64.Type} }
                ),
                result = Table.Sort(resTable, {
                    {"gpu", Order.Descending}, 
                    {"memory", Order.Descending},
                    //{"oc", Order.Ascending},
                    {"price", Order.Ascending}}
                )
            in  result,
        Source = Table.PromoteHeaders(
            Excel.Workbook(File.Contents(filepath),null,true){[Item="Sheet1",Kind="Sheet"]}[Data],
            [PromoteAllScalars=true]
        ),
        Function = fnHeaders(Source)
    in
        Function
