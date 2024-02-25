/**
*  Inspired by Power-BI-Desktop-Query-Extensions - https://github.com/tnclark8012/Power-BI-Desktop-Query-Extensions
*  Expression.Evaluate(Text.FromBinary(Web.Contents("https://gist.github.com/laktiv/700a6d65e098189881ecd77e585b233a")),#shared)
*
**/

let

MissingTests = List.Select(Record.FieldNames(_extensionLibrary), each not
    List.MatchesAny(Tests, (testRecord as record) => Text.Contains(testRecord[CaseName], _))),
MissingDocumentation = List.Select(Record.FieldNames(_extensionLibrary), each not
    Record.HasFields(Value.Metadata(Value.Type(Record.Field(_extensionLibrary, _))), "Documentation.Examples")),
Tests = {
    [
        CaseName = "List.From",
        Test = (library as record) =>
            TestUtils[AssertEqual](
                {"A", "B", "C"},
                library[List.From]("{A, B, C}"),
                "Text should be split on ,")
    ],
    [
        CaseName = "List.From - no braces",
        Test = (library as record) =>
            TestUtils[AssertEqual](
                {"A", "B", "C"},
                library[List.From]("A, B, C"),
                "Text should be split on ,")
    ],
    [
        CaseName = "List.Flatten",
        Test = (library as record) =>
            TestUtils[AssertEqual](
                { 1, 2, Table.FromRecords({[x=1]}), 3, 4, 5},
                library[List.Flatten]({ 1, 2, Table.FromRecords({[x=1]}), {3, 4, 5} }),
                "Flattened list")
    ],
    [
        CaseName = "Number.ToLetters",
        Test = (library as record) =>
            TestUtils[AssertEqual](
                { "A", "BC", "BAD" },
                {
                    library[Number.ToLetters](1),
                    library[Number.ToLetters](55),
                    library[Number.ToLetters](1382)
                },
                "1, 2, and 3 letter translations")
    ],
    [
        CaseName = "Table.RenameColumn",
        Test = (library as record) =>
            TestUtils[AssertEqual](
                {"NewName", "NumberCol"},
                Table.ColumnNames(library[Table.RenameColumn](TestUtils[SimpleTable], "TextCol", "NewName")),
                "Column should have new name")
    ],
    [
        CaseName = "Text.Until",
        Test = (library as record) =>
            TestUtils[AssertEqual](
                "abc ",
                library[Text.Until]("abc 123", "1"),
                "Proper substring should be found")
    ],
    [
        CaseName = "Text.Substring - infer end",
        Test = (library as record) =>
            TestUtils[AssertEqual](
                "BI Rules!",
                library[Text.Substring]("Power BI Rules!", 6),
                "Text.Substring('Power BI Rules', 6)")
    ],
    [
        CaseName = "Switch case pairs",
        Test = (library as record) =>
            TestUtils[AssertEqual](
                2,
                library[Switch]("B", {{"A", 1}, {"B", 2}, {"C", 3}}),
                "B is second case")
    ],
    [
        CaseName = "Switch separate case and result",
        Test = (library as record) =>
            TestUtils[AssertEqual](
                4,
                library[Switch]("TE", {"QB", "RB", "WR", "TE", "K", "D/ST"}, {1, 2, 3, 4, 5, 6}),
                "TE is 4th case")
    ]
},

TestUtils = [
    AssertEqual = (expected as any, actual as any, description as text) =>
        let
            expectedType = Value.Type(expected),
            actualType = Value.Type(actual),
            listAsText = (list as list) =>
                "{" & Text.Range(List.Accumulate(expected, "", (state, current) => state & ", " & Text.From(current)), 1) & " }",
            expectedAsText = if expected is list then listAsText(expected) else Text.From(expected),
            actualAsText = if expected is list then listAsText(actual) else Text.From(actual),
            typeAsText = (value as any) =>
                if value is binary then "binary" else
                if value is date then "date" else
                if value is datetime then "datetime" else
                if value is datetimezone then "datetimezone" else
                if value is duration then "duration" else
                if value is function then "function" else
                if value is list then "list" else
                if value is logical then "logical" else
                if value is none then "none" else
                if value is null then "null" else
                if value is number then "number" else
                if value is record then "record" else
                if value is table then "table" else
                if value is text then "text" else
                if value is time then "time" else
                if value is type then "type" else
                if value is any then "any"
                else "unknown -- not a primitive type!"
        in
            if not Value.Is(actualType, Value.Type(expectedType)) then
                error "Expected type " & typeAsText(expectedType) & " does not match actual type " & typeAsText(actualType)
            else if expected = actual then
                true
            else if expectedAsText = actualAsText then error "Cannot provide accurate failure message! Text versions of expected and actual are identical, but they failed the equality test!"
            else error "Expected: " & expectedAsText & "; Actual: " & actualAsText & "; Reason: " & description,
    SimpleTable = Table.FromRecords({[TextCol = "A", NumberCol = "1"], [TextCol = "B", NumberCol = 2], [TextCol = "C", NumberCol = 3]})
],

TestResults =
    let
        failedTests = List.Select(
            List.Transform(Tests, (suite as record) =>
                let
                    testResult = try suite[Test](_extensionLibrary)
                in
                    if testResult[HasError] then Error.Record(suite[CaseName], testResult[Error][Message], null) else true),
            each _ <> true)
    in
        if List.IsEmpty(failedTests) then "All " & Text.From(List.Count(Tests)) & " tests passed! :)" else failedTests,


_extensionLibrary = [

Document = (name as text, description as text, valueOrExample as any, optional valueIfExample as any) =>
    let
        value = if valueIfExample is null then valueOrExample else valueIfExample,
        examples = if valueIfExample is null then {} else valueOrExample
    in
        Value.ReplaceType(value, Value.Type(value) meta [
            Documentation.Name = name,
            Documentation.Description = description,
            // [Description = "", Code="", Result =""]
            Documentation.Examples = examples])
,


/////////////////////////
// INNO                //
/////////////////////////
Inno.SortTablePn = Document(
    "LPQ[Inno.SortTablePn]",
    "Returns sorted table by INNO partnumbers in proper order. Also could filter-out old models.",
    {[
        Description = "Inno.SortTablePn",
        Code = "LPQ[Inno.TableTypes](Table, ""pn"", true)",
        Result = "New sorted table by column 'pn' with old filtered."
    ]},
    (Table as table, PN as text, optional FilterOld as logical) =>
    let
        oldList = {"C16", "C20", "207", "208", "106", "107", "108", "R"},
        sortList = {"40", "30", "20", "16", "10", "73", "71", "21"},
        letterList = {"C", "N", "B", "M", "R"},

        fltrV = if (FilterOld ?? {FilterOld=""}{1}?) = null then false else FilterOld,

        Filtered = if fltrV
            then
            Table.RemoveColumns(
            Table.SelectRows(
            Table.AddColumn(Table, "Filter",
                (FindStrings) =>
                List.AnyTrue(List.Transform(oldList, each Text.Contains(Record.Field(FindStrings, PN), _, Comparer.OrdinalIgnoreCase))))
            , each ([Filter] = false))
            , {"Filter"})

            else Table,

        ReplaceTable = Table.ReplaceValue(
            Filtered,
            each Record.Field(_,PN),
            each if Record.Field(_,PN)="TBC"
                    or Record.Field(_,PN)=null
                    or Record.Field(_,PN)=""
                        then "A00003-01D5-1234552"
                        else Record.Field(_,PN),
            Replacer.ReplaceValue,
            {PN}
        ),

        Sorting = Table.Sort(ReplaceTable,
            {
                { each List.PositionOf(letterList, Text.Start(Record.Field(_,PN), 1)), Order.Ascending },   // by first letter of PN
                { each List.PositionOf(sortList, Text.Range(Record.Field(_,PN), 1,2)), Order.Ascending },   // by series
                { each Text.Range(Record.Field(_,PN), 1,4), Order.Descending },                             // by gpu
                { each Text.Middle(Text.Split(Record.Field(_,PN),"-"){0}, 5,1), Order.Descending },         // by fans
                { each Text.Start(Text.Split(Record.Field(_,PN),"-"){1}, 2), Order.Descending },            // by mem
                { each Text.Select(Text.Split(Record.Field(_,PN),"-"){1}, {"X"}), Order.Descending }        // by OC
            }
        )
    in
        Sorting)
,
Inno.TableTypes = Document(
    "LPQ[Inno.TableTypes]",
    "Converts types of columns in the table",
    {[  Description = "TableTypes",
        Code = "LPQ[Inno.TableTypes](Table)",
        Result = "New table"
    ]},
    (Table as table, optional ToOrder as logical) as table =>
    let
        toOrder = if (ToOrder ?? {ToOrder=""}{1}?) = null then false else ToOrder,

        lists = [
            txt = {"customer", "country", "pn", "product", "gpu", "partner", "url", "name"},
            num = {"quantity", "d_week", "d_month", "d_year", "stock", "VAT", "memory"},
            dec = {"k"},
            cur = {"price", "amount"},
            dat = {"date", "Invoice Date (YYYY/MM/DD)", "short_date", "d_date"},
            order = {"customer", "country", "pn", "product", "gpu", "quantity", "d_date", "d_week", "d_month", "d_year", "partner"}
        ],
        headers = Table.ColumnNames(Table),
        c = List.Combine({
                List.Transform(lists[txt], each {_, type text}),
                List.Transform(lists[num], each {_, Int64.Type}),
                List.Transform(lists[dec], each {_, type number}),
                List.Transform(lists[cur], each {_, Currency.Type}),
                List.Transform(lists[dat], each {_, type date})
            }),
        cT = List.Transform(c, each _{0}),

        Type = Table.TransformColumnTypes(Table,
            List.Select(c, each List.ContainsAny(headers, _))
        ),

        Result = if toOrder
            then Table.SelectColumns(Type,
                    List.RemoveMatchingItems(
                        lists[order],
                        List.Difference(lists[order],headers)
                    )
                )
            else Type
    in  Result)
,
Inno.Online = Document(
    "LPQ[Inno.Online]",
    "Table for calculations.",
    {[  Description = "Margin, VAT, Currencies",
        Code = "Tables",
        Result = "Table"
    ]},
    ImportPq.GitHubRaw("https://raw.githubusercontent.com/laktiv/LPQ/main/Inno.TablesOnline.pq"))
,
Inno.GetGpuCode = Document(
    "LPQ[Inno.GetGpuCode]",
    "Table for calculations.",
    {[  Description = "'3060, 3080' - add mem",
        Code = "LPQ[Inno.GetGpuCode](PN, *MemList list, *CoolerList list)",
        Result = "Get pn_code from PN, before first '-'"
    ]},
    (PN as text, optional AddMemoryList as text, optional AddCoolerList as text) =>
    let
        memList =
            if AddMemoryList = null or AddMemoryList = "" then
                false
            else
                List.Transform(Text.Split(AddMemoryList, ","), Text.Trim),
        coolList =
            if AddCoolerList = null or AddCoolerList = "" then
                false
            else
                List.Transform(Text.Split(AddCoolerList, ","), Text.Trim),

        goodsShort = File.GetExcelDataFile("goods", "shortCode"),

        func = [
            split = Text.Split(PN, "-"),
            pn_first = split{0},
            pn_firstEnds =
                if List.Contains({"T", "0", "S"}, Text.End(pn_first, 1)) then 1
                    else if Text.Length(pn_first) > 6 and Text.End(pn_first, 1) = "K" then 3
                else 2,
            range = Text.Range(split{0}, 1, Text.Length(pn_first) - pn_firstEnds),
            cooler = if Text.End(split{0}, 4) = range then "" else Text.Select(Text.End(split{0}, 2), {"1".."9", "B"}),
            addCool = try List.Contains(coolList, range) otherwise false,
            memory = Text.Start(split{1}, 2),
            addMem = try List.Contains(memList, range) otherwise false,
            code = if addCool then range & "-X" & cooler else if addMem then range & "-" & memory else range
        ],

        position = List.PositionOf(goodsShort[pn_code], func[code]),
        pos = if position < 0
                then List.PositionOf(goodsShort[pn_code], Text.Split(func[code],"-"){0})
                else position,

        result = if position >= 0
                    then goodsShort{pos}[gpu_short]
                    else if pos >= 0
                        then goodsShort{pos}[gpu]
                        else null
    in
        result)
,
Inno.GetCooler = (Product as text) =>
    let
        split = Text.Split(Product, " "),
        cool =  if List.ContainsAny({"Frostbite","Black"}, split, Comparer.OrdinalIgnoreCase) then "WB"
                else if List.ContainsAny({"X4"}, split, Comparer.OrdinalIgnoreCase) then "X4"
                else if List.ContainsAny({"X3","GAMING OC","iCHILL"}, split, Comparer.OrdinalIgnoreCase) then "X3"
                else if List.ContainsAny({"X2","Twin"}, split, Comparer.OrdinalIgnoreCase) then "X2"
                else if List.ContainsAny({"X1","Compact"}, split, Comparer.OrdinalIgnoreCase) then "X1"
                    else null,
        oc = if List.ContainsAny({"OC","iCHILL","Pro","Ultra","Frostbite","Black"}, split, Comparer.OrdinalIgnoreCase)
                then "OC" else null,
        color = if Text.Contains(Product, "White", Comparer.OrdinalIgnoreCase)
                then "White" else null,
        result = [
            oc = oc,
            fans = cool,
            color = color
        ]
    in
        result
,
Inno.ParseInnoCom = (url as text) =>
    let
        src = Table.FromColumns({Lines.FromBinary(Web.Contents(url))}),
        filtrTags = Table.SelectRows(src, each Text.Contains([Column1],"productBox") or Text.Contains([Column1],"<p class=""h4"">") or Text.Contains([Column1],"Part No.")),
        addCol = Table.AddColumn(filtrTags,"src", each if Text.Contains([Column1],"href") then Text.BetweenDelimiters([Column1],"href=""",""">") else Text.BetweenDelimiters([Column1], ">","</")),
        splitCols = Table.FromRows(List.Split(addCol[src], 3)),
        cleanCol3 = Table.ReplaceValue(splitCols, each [Column3], each Text.PowerTrim(Text.AfterDelimiter([Column3],":")), Replacer.ReplaceValue, {"Column3"} ),
        splitCol3 = Table.ExpandListColumn(Table.AddColumn(cleanCol3, "pn", each Text.SplitAny(Text.Remove([Column3], {" "}),"/")), "pn"),
        res = Table.RenameColumnsAdv(Table.SelectColumns(splitCol3,{"pn","Column2","Column1"}), {"Column1","Column2"},{"url","model"})
    in
        res
,
Inno.ParseInnoComSpecs = (url as text) =>
    let
        //url = "https://inno3d.com/en/PRODUCT_INNO3D_GEFORCE_RTX_3060_8GB_TWIN_X2_2",
        src = Web.Page(Web.Contents(url)){0}[Data],
        Rename = Table.RenameColumnsAdv(src,{"GPU Engine Specs:",""}, {"spec","value"}),
        Replace = Text.BulkTableReplace(Rename, MyFindReplace_InnoSpecs, {"spec","value"})
    in
        Replace
,
Inno.CleanTable = (Table as table) =>
    let
        order = {"Name","customer","country","pn","product","gpu","quantity","d_date","d_week","d_month","d_year","partner"},
        headers = Table.ColumnNames(Table),
        Quantity = Table.SelectRows(Table, each [quantity]<>0 and [quantity]<>null and [quantity]<>"" )
    in  Table.SelectColumns(Quantity, List.RemoveMatchingItems(order, List.Difference(order,headers)) )
,

// BULK REPLACE & TABLES
MyFindReplace_Sample =
    let table = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("TY7LDoIwFAV/hXQrxijInrY8GnkFKi4ICyQsGhUIYvh925JCdzOTe5NTVQCXBJgA1GYFQhxvXAauQpwpMg7GFnFuaWxrfNXYUezYkM1KMkSOnpLldA7gLhddLF3sXSjjRInkoPOHqe14CDw/zZEnK+n7wcI8kiRJOcivhfXi70GS9QaFJIp4YSuIVvzGbuKpuGdeLgsaPmPTiu0ojTMXUVnhu2lfvMHIRTdZ/Gn4zk82iyV+nhYUEsq31H8=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Find = _t, Replace = _t])
    in  Table.ReplaceValue(table,"",null,Replacer.ReplaceValue,{"Replace"})
,
MyFindReplace =
    let
        table = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("JcpBDoIwEIXhq5iuJMGF8QRQSpxYW0LHuCAsEFk0Ykuwxus7Lbs33/xdxxBYztCyPu8YOOdPFd2glKYRDX/WxeIOamv4GaQksduIZr7LtBKZWyPaJNy/l2EMZFxfm4Jj0nIexhdZKQt+SVKv/hMeNkykdasNloAiffbHg5l92ImnDda7jIJIUuNOVICgVcb6/g8=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Find = _t, Replace = _t])
    in  Table.ReplaceValue(table,"",null,Replacer.ReplaceValue,{"Replace"})
,
MyFindReplace_InnoSpecs =
    let
        table = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("bY/PboJAEIdfZcKpjcQo9gVaOOiBlMSaHoiHLYxKxJ3NsqTi03f/AAumt5nv983Mbp4H8SF5h5gkNkFom+AY5sEHUaMgrqm4wku6fbzqcGBXZ7AGnTDmjrh43wpR4w25YrKDjH5R6iucY6FImlNPyA6leCNt261aGdt+pWK8ZLKEgRM/VWfv7asHTtfsuEJ5YgV65asTM8X08F2V6uIdPzY7mlSNqFk3/8Rnq0SrrLhN0h1Ey7Wmppyy1c8zZFPQb85IKlgv30yWZP8lPtjcQVemhfvGkmhCovFOCM7oTVjA7BkujoaxMT7+AQ==", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Find = _t, Replace = _t])
    in  Table.ReplaceValue(table,"",null,Replacer.ReplaceValue,{"Replace"})
,
Text.BulkTableReplace = let BulkReplace = (DataTable as table, FindReplaceTable as table, DataTableColumn as list, optional ReplaceFullCell as logical) =>
        /*
        *   Replace values in column by table list of replaces
        *   Create extra table "MyFindReplace" with [Find] and [Replace] columns
        */
        let
            //Convert the FindReplaceTable to a list using the Table.ToRows function so we can reference the list with an index number
            FindReplaceList = Table.ToRows(FindReplaceTable),
            //Count number of rows in the FindReplaceTable to determine how many iterations are needed
            Counter = Table.RowCount(FindReplaceTable),
            //Define a function to iterate over our list with the Table.ReplaceValue function
            BulkReplaceValues = (DataTableTemp, n) =>
            let
                // set Replace type
                ReplaceType = if ReplaceFullCell=null or ReplaceFullCell=false then Replacer.ReplaceText else Replacer.ReplaceValue,

                //Replace values using nth item in FindReplaceList
                ReplaceTable = Table.ReplaceValue(
                    DataTableTemp,
                    //replace null with empty string
                    if FindReplaceList{n}{0} = null then "" else FindReplaceList{n}{0},
                    if FindReplaceList{n}{1} = null then "" else FindReplaceList{n}{1},
                    ReplaceType,
                    DataTableColumn
                )
            in
                //if we are not at the end of the FindReplaceList then iterate through Table.ReplaceValue again
                if n = Counter - 1
                    then ReplaceTable
                    else @BulkReplaceValues(ReplaceTable, n + 1),
            //Evaluate the sub-function at the first row
            Output = BulkReplaceValues(DataTable, 0)
        in
            Output
    in
        BulkReplace
,


/////////////////////////
// GPU & NV SERIES     //   Checks GPU
/////////////////////////

fxGpuSeries = (Product as text) =>
    let p = Text.Lower(Product), //Text.Remove(, {" "})
        posG = if Text.Contains(p, "RTX") then Text.PositionOf(p, "RTX")
                else if Text.Contains(p, "GTX") then Text.PositionOf(p, "GTX")
                else if Text.Contains(p, "GT") then Text.PositionOf(p, "GT")
                else 0,
        posE = if Text.Contains(p, "ti super", Comparer.OrdinalIgnoreCase) then 9
                else if Text.Contains(p, "ti", Comparer.OrdinalIgnoreCase) then 3
                else if Text.Contains(p, "super", Comparer.OrdinalIgnoreCase) then 6
                else 0,
        gpu = Text.Middle(Product, posG, 4 + 4 + posE)
    in  Text.Trim(Text.Replace(p,"SUPER","Super"))
,


/////////////////////////
// FANS & TABLES       //   SEARCH # OF FANS BY FILLED BRAND/SERIES TABLES
/////////////////////////
// Search full product name at Geizhals.eu, scrape result from specs table by coolers
fsGeizFans = (Product as text) =>
    let
        searchstring = "https://geizhals.eu/?fs=" & Uri.EscapeDataString(Product) & "&hloc=pl",
        Source = Web.BrowserContents(searchstring),
        Table = Html.Table(Source, {{"Column1","DT"}, {"Column2","DD"}, {"Column3", ".listview__toggle-details-button .visually-hidden"}, {"Column4",".listview__name-link"}, {"Column5",".listview__offercount-link"}}, [RowSelector="DT"]),
        Cooler = Table.SelectRows(Table, each ([Column1] = "Kühlung"))[Column2]{0},
        txt = Text.Split(Cooler," "){0},
        Split = if Text.EndsWith(txt,"x")
                    then "X" & Text.Select(Text.Split(Cooler," "){0},{"0".."9"})
                    else txt
    in
        try Split otherwise null
,
// Search by Matching Series/Brand in the Table 'tblBrandsFans'
fsGetFansAdv = (Product as text) =>
    let srcTable = Table.Buffer( tblBrandsFans ),
        // small replace
        listReplace =
        {   { "Storm X", "StormX" },
            { "GAMING BLOCK", "GAMING_BLOCK" },
            { "STRIX LC", "STRIX_LC" },
            { "GAMING OC", "GAMING_OC" },
            { "GAMING SLIM", "GAMING_SLIM" },
            { "GAMING X", "GAMING_X" } },
        mReplace = List.Accumulate(listReplace, Product, (s,x) => Text.Replace(s, x{0}, x{1})),
        mList = List.Transform(Text.Split(mReplace, " "), each Text.Upper(Text.Trim(_))),
        brand = mList{0},   // Get Brand
        fMatchFib = (ml, b, optional i) =>  //Main function
            let        i = i ?? 0,
                tbl      = Table.SelectRows(srcTable, each [brand]=b),
                listLook = List.Transform(Text.Split(tbl[name]{i}, ","), each Text.Upper(Text.Trim(_))),
                Matching = if List.ContainsAny(ml, listLook, (x,y) => x=y) then tbl[fans]{i} else null,
                Checking = Matching ?? {i<Table.RowCount(tbl)}{1}? ?? @fMatchFib(ml, b, i+1) //Recurse
            in  try Checking otherwise null
    in  fMatchFib(mList,"") ?? fMatchFib(mList,brand)
,
// Simplier search
fsGetFans = (Brand as text, Fans as text, Product as text) =>
	let srcTable = Table.Buffer( tblBrandsFans ),
        Table = Table.SelectRows(srcTable,
			each [brand] = Text.Lower(Brand) and Text.Lower([fans]) = Text.Lower(Fans) )
	in	Result = List.AnyTrue(
			List.Transform(Table[series], each Text.Contains(Product,_,Comparer.OrdinalIgnoreCase)) )
,
// Flat table # Fans by brands
tblBrandsFans =
    let Table = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("fZbRcuIgFEB/JeNzZ3aq+wMYMbImIRtIzW6nD9GyNVMNOxhn698vEE2TeuEFkzmHcLnAxefnCWIFmzxM8EcrVFMd9GNUHevmLdge5O59SF4eensz1w1rVf0RxKF5r7QTzG2PgVbOdJNUxtMPSCh5RzMlkWpdlPGclBCc6mZRoNjF1hQij2bAvRQNHI/BvMjnYN/pNxsRL5ZDGiGSblC+6KdTNe1eKK+y1448+pW7GIeKmWC0l6fWJdh5irfqdD6NFBKh+S+ObwtoV+2PVDsBLOFAtjEhnFPHxzpOc5sop4BjYl+cQoQSkkYBDX1SghjHuc94IozQ1GeUPMeJM5RpH4rLMOnVnPg4P6utdA7R7SWMotgdRudsSLpY0jwceSRN6WxxW8b5obIHtVvB1WWr6ldAW4mqPdXN+/Dx3loqva+2dQttid4sv9+a+5C6DM9gZnLL/9WNm9oGYCanoTz+rXatW/ghPNA2PVsv0bTPzOVVSSiBQylDjJEn/DV7V8VOeFYCI3QrXQa6qo7rwphHLrSiSxdizk6MQ8Q0UzBGW07PtvqDDI96JYzcsqV/Y/KzIAtgu4w1hlGwQpu1UxzWABaT5H7EbmpFljuhvi6oA12/DEzEceIHkOtTCKBHoCzeenWn9wmnfFwVM6TLYJ/yLzdYD82XWSvVsfTjAObX0c2Wy7tL3Ofoyz5T0ifpc6Uve1GNLq3flKOwP+tknIYR1P8glDgKJ+eqbur2AnJbDpLMx4JQKvjjduuGcGBQKeqhTTCNP3u+/Ac=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [brand = _t, fans = _t, series = _t, #"type" = _t]),
        Lower = Table.TransformColumns(Table, {
            {"brand", Text.Lower, type text},
            {"series", Text.Lower, type text} })
    in  Lower
,
//Grouped table by Brand/Series and # Fans
tblBrandsFansGrouped =
    let Table = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("fZbRcuIgFEB/JeNzZ3aq+wMYMbImIRtIzW6nD9GyNVMNOxhn698vEE2TeuEFkzmHcLnAxefnCWIFmzxM8EcrVFMd9GNUHevmLdge5O59SF4eensz1w1rVf0RxKF5r7QTzG2PgVbOdJNUxtMPSCh5RzMlkWpdlPGclBCc6mZRoNjF1hQij2bAvRQNHI/BvMjnYN/pNxsRL5ZDGiGSblC+6KdTNe1eKK+y1448+pW7GIeKmWC0l6fWJdh5irfqdD6NFBKh+S+ObwtoV+2PVDsBLOFAtjEhnFPHxzpOc5sop4BjYl+cQoQSkkYBDX1SghjHuc94IozQ1GeUPMeJM5RpH4rLMOnVnPg4P6utdA7R7SWMotgdRudsSLpY0jwceSRN6WxxW8b5obIHtVvB1WWr6ldAW4mqPdXN+/Dx3loqva+2dQttid4sv9+a+5C6DM9gZnLL/9WNm9oGYCanoTz+rXatW/ghPNA2PVsv0bTPzOVVSSiBQylDjJEn/DV7V8VOeFYCI3QrXQa6qo7rwphHLrSiSxdizk6MQ8Q0UzBGW07PtvqDDI96JYzcsqV/Y/KzIAtgu4w1hlGwQpu1UxzWABaT5H7EbmpFljuhvi6oA12/DEzEceIHkOtTCKBHoCzeenWn9wmnfFwVM6TLYJ/yLzdYD82XWSvVsfTjAObX0c2Wy7tL3Ofoyz5T0ifpc6Uve1GNLq3flKOwP+tknIYR1P8glDgKJ+eqbur2AnJbDpLMx4JQKvjjduuGcGBQKeqhTTCNP3u+/Ac=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [brand = _t, fans = _t, series = _t, #"type" = _t]),
        Lower = Table.TransformColumns(Table, {
            {"brand", Text.Lower, type text},
            {"series", Text.Lower, type text} })
    in  Lower
,


/////////////////////////
// Date                //
/////////////////////////
// Inno
Date.DateFromISOwn = Document(
    "Date.DateFromISOwn",
    "Compute date of the first day of the week given an ISO weeknumber. If the Year is not entered, it will default to 'this' year.",
    {[  Description = "Year must be a full year: eg 2021 vs 21. The latter will not be interpreted as the year 2021.",
        Code = "Date.DateFromISOwn(""week number"", opt. ""year"")",
        Result = "Date on start of week #"
    ]},
    (wn as number, optional year as number) as date =>
    let
        yr = if year = null then Date.Year(DateTime.LocalNow()) else year,
        wn1Start = Date.StartOfWeek(#date(yr,1,1), Day.Monday),
        w1 = if Date.AddDays(wn1Start,3) < #date(yr,1,1) then Date.AddDays(wn1Start,7) else wn1Start
        //-2 if not 01.01.01 start of the year(?)
    in Date.AddDays(w1, 7*(wn-1)) )
,
// Basic calendar
Date.Calendar = Document(
        "Date.Calendar",
        "Generate a calendar table for a given date span - can be text or proper dates. Current columns are Date, DayOfWeek, Month, MonthNum, WeekStartData, WeekStart, Year, YearMonth",
        {[ Description = "2016 calendar", Code ="PBI[Date.Calendar](""1/1/2016"", ""12/31/2016""", Result = "2016 calendar"]},
        (start as any, end as any) =>
            let
                StartDate = Date.From(start),
                EndDate = Date.From(end),
                Source = Date.DatesBetween(StartDate, EndDate),
                FromList = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
                Date = Table.RenameColumns(FromList,{{"Column1", "Date"}}),
                DayOfWeek = Table.AddColumn(Date, "Day of Week", each Date.DayName([Date])),
                Month = Table.AddColumn(DayOfWeek, "Month", each Date.MonthName([Date])),
                MonthNum = Table.AddColumn(Month, "MonthNumber", each Date.Month([Date])),
                WeekStartDate = Table.AddColumn(MonthNum, "WeekStartDate", each Date.StartOfWeek([Date])),
                WeekStart = Table.AddColumn(WeekStartDate, "Week Start", each [Month] & " " & Text.From(Date.Day([WeekStartDate]))),
                Year = Table.AddColumn(WeekStart, "Year", each Date.Year([Date])),
                YearMonth = Table.AddColumn(Year, "YearMonth", each Number.From(Text.From([Year]) & (if [MonthNumber] < 10 then "0" else "") & Text.From([MonthNumber]))),
                Result = YearMonth
           in
                Result)
,
Date.DatesBetween = Document(
    "Date.DatesBetween",
    "Returns a list of dates in a given span (inclusive). Start and end parameters can be any order",
    {[Description = "Date range", Code = "PBI[Date.DatesBetween](""1/1/2016"", ""1/3/2016"")", Result="{""1/1/2016"", ""1/2/2016"", ""1/3/2016""}" ]},
    (start as any, end as any) =>
      let
        StartDate = Date.From(start),
        EndDate = Date.From(end),
        adjustedStart = List.Min({StartDate, EndDate}),
        adjustedEnd = List.Max({StartDate, EndDate}),
        GetDates = (start as date, end as date, dates as list)=> if start > end then dates else @GetDates(Date.AddDays(start, 1), end, List.Combine({dates, {start}})),
        Dates = GetDates(adjustedStart, adjustedEnd, {})
      in Dates)
,
Date.DayName = Document(
    "Date.DayName",
    "Returns the English day of the week name for a date",
    {[ Description = "Get the day name", Code="Date.DayName(""9/9/2016"")", Result="Friday"]},
    (date as any) => Switch(Date.DayOfWeek(DateTime.From(date)), {0, 1, 2, 3, 4, 5, 6}, {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"}, null))
,
Date.MonthName = Document(
    "Date.MonthName",
    "Returns the English month name for a date",
    {[ Description = "Get the month name", Code="PBI[Date.MonthName](""9/9/2016"")", Result = "September" ]},
    (date as any) =>
        let
            monthNumber = if date is number then date else Date.Month(DateTime.From(date))
        in
            Switch(
                monthNumber,
                {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12},
                {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"}, null))
,

/////////////////////////
// List                //
/////////////////////////
List.Flatten = Document(
    "List.Flatten",
    "Recursively flattens list elements. The end result is a single list",
    { [ Description = "Flattening nested lists into one", Code = "List.Flatten({ 1, 2, Table.FromRecords({[x=1]}), {3, 4, 5} })", Result = "{ 1, 2, Table.FromRecords({[x=1]}), 3, 4, 5}"] },
    (list as list) => List.Accumulate(list, {}, (state, current) =>
        let
            currentListContent = if current is list then @List.Flatten(current) else {current}
        in
            List.Combine({state, currentListContent})
    ))
,
List.From = Document(
    "List.From",
    "Converts a text representation of a list into a list of the elements. Items are considered to be split by ,",
    { [ Description = "Convert a text list", Code = "List.From(""{A, B, C}"")", Result = "{ ""A"", ""B"", ""C"" }"] },
    (simpleTextList as text) =>
        let
            trimWhitespace = Text.Trim(simpleTextList),
            listToSplit = Text.TrimEnd(Text.TrimStart(trimWhitespace, "{"), "}"),
            Result = List.Transform(Text.Split(listToSplit, ","), each Text.Trim(_))
        in
            Result)
,
List.ToText = Document(
    "List.ToText",
    "Converts a list to a textual representation. Inverse of List.From",
    { [ Description = "Conver to text", Code = "List.ToText({ 1, 2, 3})", Result = """{1, 2, 3}"""] },
    (list as list) =>
            List.Accumulate(list, "{", (state, current) => current & Text.From(current)) & "}")
,

/////////////////////////
// Number              //
/////////////////////////
Number.Digits = {0,1,2,3,4,5,6,7,8,9}
,
Number.ParseText = Document(
    "Number.ParseText",
    "Returns the first number (1+ consecutive digits) in the given text. Optionally, provide allowed characters to ignore them",
    {[ Description = "Allow commas", Code="PBI[Number.ParseText](""It's over 9,000!"", 0, {"",""})", Result = "9000" ]},
    (text as text, optional startIndex as number, optional allowCharacters as list) =>
        let
            consider = if startIndex is null then text else Text.Range(text,startIndex),
            _allowCharacters = if allowCharacters is null then {} else allowCharacters,
            numberSeries = List.FirstN(List.Skip(Text.ToList(consider), each not Text.IsNumber(_)), each Text.IsNumber(_) or List.Contains(_allowCharacters, _))
        in
            if text is null then null else Text.FromList(numberSeries))
,
Number.ToLetters = Document(
        "Number.ToLetters",
        "Converts a number  (starting at 1) to an alphabet representation. Works like column headers in Excel.",
        {[
            Description = "Column 27",
            Code = "PBI[Number.ToLetters](27)",
            Result = "AB"
        ]},
        (value as number) =>
            let
                GetLetter = (num as number) =>
                    let
                        number = Number.Mod(num, 26),
                        val = if number = 0 then 26 else number,
                        valid = number < 26 and number > 0
                    in
                        if valid then Text.At(Text.Alphabet, val - 1) else error "Can't get letter for " & Text.From(num),
                func = (value as number, factor as number) =>
                    let
                        ThisLetter = GetLetter(Number.RoundDown(value/Number.Power(26, factor))),
                        Result = if value <= Number.Power(26, factor) then "" else @func(value, factor+1) & ThisLetter
                    in
                        Result
            in
                if value <= 26 then GetLetter(value) else func(value, 1) & GetLetter(value))
,
// MY
Number.ProperNumber = Document(
    "Number.ProperNumber",
    "Concerts text as number to proper number type. Also will erase inconvinient thouthands and decimal symbols to regular format. <br>" &
    "In case of error will return input text",
    {[
        Description = "Allow commas",
        Code = "LPQ[Number.ProperNumber](""It's over 9,000!"", null, {"",""}, 0)",
        Result = "9000"
    ]},
    (text as text, optional digitsAfter as number, optional allowCharacters as list, optional startIndex as number) =>
    let
        after = if digitsAfter is null then 0 else digitsAfter,
        consider = if startIndex is null then text else Text.Range(text,startIndex),
        _allowCharacters = if allowCharacters is null then {} else allowCharacters,
        numberSeries = List.FirstN(List.Skip(Text.ToList(consider), each not Text.IsNumber(_)), each Text.IsNumber(_) or List.Contains(_allowCharacters, _))
    in
        if text is null then null else Number.Round(Text.FromList(numberSeries), after))
,

/////////////////////////
// Splitters           //
/////////////////////////
Splitter.SplitTextByNonAlpha = Document(
    "Splitter.SplitTextByNonAlpha",
    "Splits text by characters that aren't [A-Za-z]",
    {[ Description = "Split text", Code="LPQ[Splitter.SplitTextByNonAlpha](""A1B,C"")", Result = "{ ""A"", ""B"", ""C"" }"]},
    (line as text) => Splitter.SplitTextByNotIn(Text.Alphabet))
,
Splitter.SplitTextByNotIn = Document(
    "Splitter.SplitTextByNotIn",
    "Splits text on any characters that aren't the provided 'safe' characters",
    {[
        Description = "Split on non-alphanumeric",
        Code = "LPQ[Splitter.SplitTextByNotIn](LPQ[Text.AlphaNumeric])(""Power BI is #1"")",
        Result = "{""Power BI is "", ""1""}"
    ]},
    (safeCharacters as text) => (line as nullable text) =>
        if line is null then
            {}
        else
            List.Accumulate(Text.ToList(line), {null} , (state, current) =>
            let
                doSkip = not Text.Contains(safeCharacters, current),
                lastItem = List.Last(state),
                appendLast = lastItem<>null
            in
                if doSkip then
                    if lastItem is null then
                        state
                    else
                        List.Combine({state, {null}})
                else if appendLast then
                    List.Combine({List.RemoveLastN(state, 1), {lastItem & current}})
                else
                    List.Combine({List.RemoveLastN(state, 1), {current}})))
,

/////////////////////////
// Switch             //
////////////////////////
Switch = Document(
    "Switch",
    "Given a value, find it's paired item <br>"&
    "Switch(value as any, cases as list, results as list, optional default as any) <br>" &
    "Switch(value as any, pairs as list, optional default as any)",
    {
        [Description = "Using separate lists", Code = "Switch(1, {1, 2, 3}, {""A"", ""B"", ""C""})", Result = "A"],
        [Description = "Using one paired list", Code = "Switch(1, {{1, ""A""}, {2, ""B""}, {3, ""C""}})", Result = "A"]
    },
    (value as any, casesOrPairs as list, optional resultsOrDefault as any, optional default as any) =>
        let
            hasPairs = List.First(casesOrPairs) is list,
            usingPairs =
                let
                    targetPosition = List.PositionOf(casesOrPairs, value, Occurrence.First, (case, theValue) => theValue = case{0})
                in
                    if targetPosition = -1 then resultsOrDefault else casesOrPairs{targetPosition}{1},
            usingCases =
                let
                    cases = casesOrPairs,
                    results = resultsOrDefault
                in
                    if List.IsEmpty(cases) or List.IsEmpty(results) then default else if value = List.First(cases) then List.First(results) else @Switch(value, List.Skip(cases, 1), List.Skip(results, 1), default)
        in
            if hasPairs then usingPairs else usingCases)
,

/////////////////////////
// Table               //
/////////////////////////
Table.RenameColumn = (Table as table, column as text, newName as text) => Table.RenameColumns(Table, {{column, newName}})
,
Table.RenameColumnsAdv = (Table as table, OldColumns as list, NewColumns as list, optional SelectHeaders as logical) =>
    let selectColumns = if (SelectHeaders ?? {SelectHeaders=""}{1}?) = null then false else SelectHeaders,
        // Merge Old names and New names
        columnsZip = List.Zip({ OldColumns, NewColumns }),
        // Get table headers and Select list ny Old names
        headers = List.Transform(Table.ColumnNames(Table), each _),
        selectHeaders = List.Select(headers, (x) => List.AnyTrue(List.Transform(OldColumns, each x = _))),
        // Get only picked columns
        SelectTable = if selectColumns then Table.SelectColumns(Table, selectHeaders) else Table.SelectColumns(Table, headers),
        // Rename columns
        RenameTable = Table.RenameColumns(SelectTable, columnsZip)
    in  try RenameTable otherwise null
,
Table.TrimHeader = (Table as table) =>
    let hRaw = Table.ColumnNames(Table),
        hTrim = List.Transform(hRaw, each Text.Trim(_))
    in  if List.Count(List.Difference(hRaw, hTrim)) > 0
            then Table.RenameColumns(Table, List.Zip({hRaw, hTrim}))
            else Table
,
// Replaces a value if it matches a predicate
Table.ReplaceValueIf = (table as table, replaceIf as function, after as any, columnNameOrList as any) =>
    Table.ReplaceValue(table, null,after, (text, old, new)=>if replaceIf(text) then new else text, if columnNameOrList is list then columnNameOrList else {columnNameOrList})
,
// Splits camelCased and PascalCased column names.
Table.SplitColumnNames = (table as table) => Table.RenameColumns(table, List.Transform(Table.ColumnNames(table), each {_, Text.SplitCamelCase(_)}))
,
// Splits camelCased and PascalCased text in a column.
Table.SplitColumnText = (table as table, columns as list) => if List.Count(columns) = 0 then table else Table.TransformColumns(@Table.SplitColumnText(table, List.Skip(columns, 1)), {{List.First(columns), Text.SplitCamelCase}})
,
Table.TransformColumn = (table as table, column as text, transform as function) => Table.TransformColumns(table, {{column, transform}})
,
Table.RenameAndTransformColumn = (table, currentName as text, newName as text, transform as function) => Table.TransformColumn(Table.RenameColumns(table, {currentName, newName}), newName, transform)
,
Table.FromListCrossJoin = (listColumnNamePairs as any) =>
   let remainingPairs = List.Skip(listColumnNamePairs, 1),
       current = List.First(listColumnNamePairs),
       theList = List.First(current),
       columnName = List.First(List.Skip(current),1),
       firstTable = Table.FromList(theList, null, {columnName}),
       doStuff = (table as table, remainingPairs as list) =>
          if List.Count(remainingPairs) <= 0 then table else
          let
             current = List.First(remainingPairs),
             theList = List.First(current),
             columnName = List.First(List.Skip(current), 1),
             nextTable = Table.ExpandListColumn(Table.AddColumn(table, columnName, each theList), columnName)
          in @doStuff(nextTable, List.Skip(remainingPairs, 1)),
       Result = doStuff(firstTable, remainingPairs)
   in  Result
,
Table.JsonDecode = (encoded as text) =>
    let Decompressed = Binary.Decompress(Binary.FromText(encoded, BinaryEncoding.Base64), Compression.Deflate),
        Decoded = Table.FromRows(Json.Document(Decompressed), let _t = ((type text) meta [Serialized.Text = true]) in type table [Name = _t, Number = _t])
    in  Decoded
,
Table.JsonEncode = (table as table) =>
    let Rows = Table.ToRows(table),
        Json = Json.FromValue(Rows),
        Compressed = Binary.Compress(Json, Compression.Deflate),
        Encoded = Binary.ToText(Compressed, BinaryEncoding.Base64)
    in  Encoded
,


Table.NumberColumns = (table as table, columns as list) as table =>
    let
        NumberTypes = List.Repeat({type number}, List.Count(columns)),
        DataWithTypes = Table.TransformColumnTypes( table, List.Zip({columns, NumberTypes}) ),
        NumberNotNull = (number) => if number is null then 0 else number,
        DataNoNulls = Table.TransformColumns( DataWithTypes, List.Zip({ columns, List.Repeat({NumberNotNull}, List.Count(columns)), NumberTypes }) )
    in  Table.ReplaceErrorValues(DataNoNulls, List.Transform(columns, each {_,0}))
,
Table.RenameColumnsByIndices = (Source as table, Indices as list, ColumnNamesNew as list) =>
    let
        ColumnNames = Table.ColumnNames(Source),
        ColumnNamesList = List.Transform(Indices, each ColumnNames{_} ),
        ZippedList = List.Zip({ ColumnNamesList, ColumnNamesNew } ),
        #"Renamed Columns" = Table.RenameColumns(Source, ZippedList )
    in
        #"Renamed Columns"
,
Table.RenameColumnsByOrderFiltered = (InputTable as table, ColumnNumbers as list, NewColumnNames as list) =>
    let
        OldColumnNames = Table.ColumnNames(InputTable),
        Indexed = List.Zip({OldColumnNames, {0..-1+List.Count(OldColumnNames)}}),
        Filtered = List.Select(Indexed, each List.Contains(ColumnNumbers,_{1})),
        IndexRemoved = List.Transform(Filtered, each _{0}),
        RenameList = List.Zip({IndexRemoved,NewColumnNames}),
        RenamedColumns = Table.RenameColumns(InputTable, RenameList)
    in
        RenamedColumns
,

Table.TrimHeaders = (Table as table) =>
    let hRaw = Table.ColumnNames(Table),
        hTrim = List.Transform(hRaw, each Text.Trim(_))
    in  if List.Count(List.Difference(hRaw, hTrim)) > 0 then Table.RenameColumns(Table, List.Zip({hRaw, hTrim})) else Table
,
Table.PromoteHeadersNonEmpty = (table as table) as table =>
    /**
    Promote the first table row as column names, but keep the current header
    if the value in the first row is empty.
    For simpler use cases check built-in function `Table.PromoteHeaders`
    **/
    let
        ExtraHeaders = Table.First(table),
        OldHeaders = Table.ColumnNames(table),
        NewHeaders = List.Transform( OldHeaders, (CurrentHeader) =>
            let
                Current = CurrentHeader,
                Candidate = Record.Field(ExtraHeaders, CurrentHeader)
            in
                if Candidate <> "" and Candidate <> null and Candidate <> 0
                then Text.From(Candidate)
                else Current
        ),
        RenamedColumns = Table.RenameColumns( Table.Skip(table, 1), List.Zip({OldHeaders, NewHeaders}) )
    in  RenamedColumns
,
Table.SelectHeaders = (List as list, ColumnStart as text, ColumnEnd as text, optional Index as text) =>
    let pos1 = List.PositionOf(List, ColumnStart),
        pos2 = List.PositionOf(List, ColumnEnd),
        range = List.Range(List, pos1 + 1, pos2 - pos1 - 1)
    in  if Index=null then range else List.Union({{Index}, range})
,

Table.RemoveEmptyRows = (tbl as table) =>
    // return a table with empty rows removed
    let
        empties_list = {null, ""}, //null = original null value, "" = after applying text.trim transformation
        col_names = Table.ColumnNames(tbl),
        tbl_no_erors =  Table.ReplaceErrorValues(tbl, List.Transform(col_names, each {_, null})),
        //Add helper column
        AddedCustom = Table.AddColumn(tbl_no_erors, "Custom",
            each
                let //convert each record to list and remove empty spaces in each list
                    record_values = List.Transform(Record.ToList(_), each try Text.Trim(_) otherwise _) //handles numbers
                in List.MatchesAll(record_values, each List.ContainsAny(empties_list, {_}) ) //evaluate empties in each record and return a boolean
            ),
        // filter out empty or null record
        Filter_out_empty_rows = Table.SelectRows(AddedCustom, each [Custom] = false)
    in
        Table.RemoveColumns(Filter_out_empty_rows, {"Custom"}) //Remove helper column added
,
Table.RemoveEmptyColumns = (tbl as table) =>
    // return table with blank columns removed
    let
        col_names = Table.ColumnNames(tbl), // get column names
        tbl_no_erors =  Table.ReplaceErrorValues(tbl, List.Transform(col_names, each {_, null})), //remove errors
    // fx to check if column has null or empty values
    Is_Null_Column_Values = (tbl as table, col_name)  =>
            let
                empty_values_list = {null, ""},
                col_values = List.Transform(Table.Column(tbl, col_name), each try Text.Trim(_) otherwise _), //remove spaces
                is_null_all = List.MatchesAll(col_values , each  List.ContainsAny(empty_values_list, {_}) )
            in
                is_null_all,
        // fx to get col names of empty or null columns
        non_empty_columns =  (tbl as table)=>
            let col_states =
                    List.Generate(
                        ()=> [col = col_names{count}, count = 0, is_null = Is_Null_Column_Values(tbl, col)],
                        each [count] < List.Count(col_names),
                        each [col = col_names{count}, count = [count] + 1, is_null = Is_Null_Column_Values(tbl, col)],
                        each [ col = if not [is_null] then [col] else null ][col]
                    )
            in  List.Select(col_states, each _ <> null )
    in
    Table.SelectColumns(tbl, non_empty_columns(tbl_no_erors))
,
// return a table with all errors replaced with specified value or null
Table.ReplaceAllErrorsInTable = (tbl as table, optional replacer_value) => Table.ReplaceErrorValues(tbl, List.Transform(Table.ColumnNames(tbl), each {_, replacer_value}))
,



Table.RunningTotalfunction = (TabName as table, RT_Amt_ColumnName as text, Added_Col_RT_Name as text)=>
    /**
    *   TabName = Name of the Table to implement running total on
    *   RT_Amt_ColumnName = the column with the amount to sum
    *   Added_Col_RT_Name = custom name for the added column
    */
    let
        // Amount List to apply List.Generate on
        Amt_List = List.Buffer(Table.Column(TabName, RT_Amt_ColumnName)),
        // List.Generate calculates the running total
        RT_List = List.Generate(
                    () => [RT = Amt_List{0}, counter = 0],
                    each [counter] < List.Count(Amt_List),
                    each [RT = List.Sum({[RT], Amt_List{counter}}), counter = [counter] + 1],
                    each [RT]
                ),
        // Consolidate brings the table together with the running total
        Consolidate = Table.FromColumns(Table.ToColumns(
                        TabName
                    ) & {RT_List}, Table.ColumnNames(TabName) & {Added_Col_RT_Name}
                    )
    in
        Consolidate
,
Table.GroupedRunningTotal = (values as list, grouping as list) as list =>
    /* Written by Philip Treacy
        https://www.myonlinetraininghub.com/grouped-running-totals-in-power-query
    */
    let
        GRTList = List.Generate
        (
            ()=> [ GRT = values{0}, i = 0 ],
            each [i] < List.Count(values),
            each try
                    if grouping{[i]} = grouping{[i] + 1}
                    then [GRT = [GRT] + values{[i] + 1}, i = [i] + 1]
                    else [GRT = values{[i] + 1}, i = [i] + 1]
                otherwise [i = [i] + 1]
        ,
            each [GRT]
        )
    in
        GRTList
,


Table.GetParameter = (ParameterName as text) =>
    // Query written by Ken Puls, FCPA, FCMA, MS MVP (Excel) (www.excelguru.ca)
    // Injected into this workbook via Monkey Tools at the model designer's request
    // For more information, Or to purchase a copy of this add-in, visit https://monkeytools.ca
    //// Original source can be found at the following URL:
    // http://www.excelguru.ca/blog/2014/11/26/building-a-parameter-table-for-power-query/
    // To use the function you must pass it a 2 column table which contains the Parameters
    // The header row must contain a "Parameter" and a "Value" column (with those names)
    // Data rows should list the parameter name in the first column and the value of that
    // parameter in the second column
    // Provided you save this query under the name "fnGetParameter", you can then reference
    // it from other queries as follows:
    //
    //  =fnGetParameter("your_parameter_name")
    // One useful sample to return the current folder:
    // Parameter Name: File Path
    // Parameter Value: =LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1),1)-1)
    // Call as  "fnGetParamater("File Path")
    let
        ParamSource = Excel.CurrentWorkbook(){[Name="XLGQuery_Parameters"]}[Content],
        ParamRow = Table.SelectRows(ParamSource, each ([Parameter]=ParameterName)),
        Value=
    if Table.IsEmpty(ParamRow)=true
    then null
    else Record.Field(ParamRow{0},"Value")
    in
        Value
,



/////////////////////////
// ARRAYS              //
/////////////////////////
/* VLOOKUP in other table */
fnVlookUp = Document(
    "fnVlookUp",
    "fnVlookUp in other table."
    {[
        Description = "Table,, index col name, search col name, value",
        Code = "fnVlookUp(lookupTable, lookupColumnName, returnColumnValue, lookupValue)",
        Result = "9000"
    ]},
    (lookupValue as any, lookupTable as table, lookupColumnName as text, returnColumnValue as text) =>
    let
        colLookup = Table.Column(lookupTable, lookupColumnName),
        colToReturn = Table.Column(lookupTable, returnColumnValue),
        lookup = List.PositionOf(colLookup, lookupValue, 0),
        Result = if lookup >=0 then colToReturn{lookup} else "Not found"
    in
        Result)
,
/* SEARCH value in other table and returns "1" if record found */
fnCheckUp = (lookupValue as any, lookupTable as table, lookupColumnName as text) =>
    let colLookup = Table.Column(lookupTable, lookupColumnName),
        // colToReturn = Table.Column(lookupTable, returnColumnValue),
        lookup = List.PositionOf(colLookup, lookupValue, 0)
    in  if lookup >=0 then 1 else null
,
// Search Name in List
fnCheckInList = Document(
    "fnCheckInList.",
    "Search word in the list. True - case not sesetive, False - case sensetive.",
    {[
        Description = "Using separate lists. Params: Word, List, Case sensetive.",
        Code = "fnCheckInList(""Word"", {""Text"", ""Word"", ""Letter""}, true)",
        Result = "TRUE."
    ]},
    (Name as text, List as list, optional Match as logical) =>
    let
        match = if (Match ?? {Match=""}{1}?) = null then true else Match,
        p = try Text.Trim(Text.BeforeDelimiter(Name,"(")) otherwise Name,
        pList = Text.Split(p, " "),
        result = if match
            then List.ContainsAny(pList, List, (x,y) => Text.Contains(x,y,Comparer.OrdinalIgnoreCase))
            else List.ContainsAny(pList, List, (x,y) => x=y)
    in
        try result otherwise null)
,

Text.RegExp = (txt as text, regex as text, delim as text) =>
    let
    RegExp =
        Web.Page(
        "<script>var x = '" &
        txt &
        "';var delim = '" &
        delim &
        "';var regex = /" &
        regex &
        "/gi;var result = x.match(regex).join(delim);document.write(result);</script>"
        )
        [Data]{0}[Children]{0}[Children]{1}[Text]{0},
        return = try RegExp otherwise null
    in
        return
,
Text.GetMatch = (Text as text, List as list) =>
    try List.Select(Text.Split(Text," "),
            each List.ContainsAny(List.LastN(Text.Split(_," "),1), List, (x,y) => Text.Contains(x,y,Comparer.OrdinalIgnoreCase))
        ){0}
    otherwise null
,

/////////////////////////
// Text                //
/////////////////////////
Text.Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz" & "АБВГДЕЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдежзийклмнопрстуфхцчшщъыьэюя"
,
Text.AlphaNumeric = Text.Alphabet & "0123456789"
,
Text.FromList = (list as list) => List.Accumulate(list, "", (state, current) => state & Text.From(current))
,
// Is text all uppercase? returns false if any non-alpha characters are present
Text.IsUpperCase = (text as text) => List.AllTrue(List.Transform(Text.ToList(text), (letter)=>Text.Contains(Text.Alphabet, letter) and letter = Text.Upper(letter)))
,
Text.IsAlpha = (text as text) => List.MatchesAll(Text.ToList(text), each Text.Contains(Text.Alphabet, _))
,
Text.RemoveExtraWhitespace = (text as text) => Text.Combine(Splitter.SplitTextByWhitespace()(text)," ")
,
// Splits camelCased and PascalCased text and separates by a space. Ex: "thisIsAColumn" -> "this Is A Column"
Text.SplitCamelCase = (text as nullable text) => if text is null then null else List.Accumulate(Text.ToList(text),"", (state, current) =>
   let
      PreviousLetter = Text.End(state, 1),
      Ignore = (text as text) => text = " " or text = "."
   in
      state &
      (if
         not Text.IsUpperCase(PreviousLetter) and
         not Ignore(PreviousLetter) and
         not Ignore(current) and
         Text.IsUpperCase(current)
       then
         " " else "" ) &
      current )
,
Text.Substring = (text as text, start as number, optional count as number) =>
   let
      start = if start >= 0 then start else error "start index should be >= 0",
      end = if count = null then Text.Length(text) else if count <= Text.Length(text) then count else error "count should be <= text length",
      textList = Text.ToList(text),
      substr = Text.FromList(List.FirstN(List.Skip(textList, start), end - start))
   in substr
,
Text.IsNumber = (text as text) => try Number.FromText(text) is number otherwise false
,
Text.PositionAfter = (text as nullable text, substring as text) =>
    let
        firstIndex = Text.PositionOf(text, substring),
        indexAfter = if firstIndex >=0 then firstIndex + Text.Length(substring) else -1
    in
        if text is null then -1 else if indexAfter >= 0 and indexAfter < Text.Length(text) then indexAfter else -1
,
Text.Until = (text as text, endDelimiter as text, optional startIndex as number) =>
    let
        start = if startIndex = null then 0 else startIndex,
        textFromStart = Text.Substring(text, start),
        delimPosition = if Text.PositionOf(textFromStart, endDelimiter) >= 0 then Text.PositionOf(textFromStart, endDelimiter) else Text.Length(textFromStart)
    in
        if text is null then null else Text.Range(textFromStart, 0, delimPosition)
,
Text.PowerTrim = (text as text, optional char_to_trim as text) =>
    let
        char = if char_to_trim = null then " " else char_to_trim,
		split = Text.Split(text, char),
		removeblanks = List.Select(split, each _ <> "")
	in
        Text.Combine(removeblanks, char)
,
Text.ReplaceByList = try (txt as text, replacementLists as list) as text =>
    List.Accumulate(replacementLists, txt, (s,x) => Text.Replace(s,x{0},x{1})) otherwise null
,
Text.RegexMatch = (txt as text, regex as text, delim as text) => try
    Web.Page("<script>var x='"&txt&"';var delim='"&delim&"';var regex=/"&regex&"/gi;var result=x.match(regex).join(delim);document.write(result);</script>")[Data]{0}[Children]{0}[Children]{1}[Text]{0}
    otherwise null
,


/////////////////////////
// Code AND WEB        //
/////////////////////////
Web.GetAllLinks = (SourceURL as text, optional AttributeDelimiter as text) =>
    let
        AttributeDelimiter = if AttributeDelimiter is null then """" else AttributeDelimiter,
        //Get File.GetExcelDataFile("goods", "shortCode") source
        Source = Text.FromBinary(Web.Contents(SourceURL)),
        //Function to find each link
        GetLink = (Counter as number) =>
            let CurrentLink = Text.BetweenDelimiters( Source, "href=" & AttributeDelimiter, AttributeDelimiter, Counter )
            in  if CurrentLink = "" then {}
                else List.Combine({{CurrentLink}, @GetLink(Counter + 1)})
    in  GetLink(0)
,
Web.StripHtmlList = (tbl as table) =>
	let
		t1 = Table.Column(tbl, "Column1"),
		l1 = List.Select(List.Transform(t1, each Text.Clean(Text.PowerTrim(_))), each _ <> ""),
		lBodyStart = List.PositionOf( l1, List.Select(l1, each Text.Contains(_, "<body", Comparer.OrdinalIgnoreCase)){0}
		),
		lBodyEnd = List.PositionOf(l1, List.Select(l1, each Text.Contains(_, "</body", Comparer.OrdinalIgnoreCase)){0} ),
		L = List.Range(l1, lBodyStart, lBodyEnd - lBodyStart - 1)
	in Table.FromList(L, Splitter.SplitByNothing(), null, null, ExtraValues.Error)
,
Web.RemoveHtmlTags = let func = (HTML) =>
        let
            Check = if Value.Is(Value.FromText(HTML), type text) then HTML else "",
            Source = Text.From(Check),
            //TextRemove = Text.Replace(Text.Replace(Source, "!--",""), "--",""),
            SplitAny = Text.SplitAny(Source,"<>"),
            ListAlternate = List.Alternate(SplitAny,1,1,1),
            ListSelect = List.Select(ListAlternate, each _<>"")
        in
            Text.Combine(ListSelect, ""),
            documentation = [
                Documentation.Name =  " Web.RemoveHtmlTags"
                , Documentation.Description = "Remove Html Tags"
                , Documentation.LongDescription = " Removes all Html tags from a text"
                , Documentation.Category = " Web.Modification"
                , Documentation.Source = " Inspired by a solution from Bill Szysz"
                , Documentation.Author = " Imke Feldmann: www.TheBIccountant.com & Mike Carlo: PowerBI.Tips"
                , Documentation.Examples = {[
                    Description = "Function that enables one to pass in a column that has HTML tags."
                    , Code = "<div>my bit of text</div>"
                    , Result = "my bit of text "
                    ]}]
    in
        Value.ReplaceType(func, Value.ReplaceMetadata(Value.Type(func), documentation))
,


// Web.RemoveShortCode??

/////////////////////////
// Import PQ           //
/////////////////////////

// Minimal script loader
ImportPq = Document(
    "ImportPq", "Most minimal external PQ script file loader.",
    {[ 
        Description="Load PQ file to Excel", 
        Code="LPQ[ImportPq](""C:\File.pq"")", 
        Result = "code from .pq file" 
    ]},
    (FilePath as text, optional Encoding as nullable number) as any =>
        try Expression.Evaluate(Text.FromBinary(File.Contents(FilePath),(Encoding ?? TextEncoding.Utf8)),#shared)
        otherwise null )
,
ImportPq.GitHubRaw = (URL as text) => try Expression.Evaluate(Text.FromBinary(Web.Contents(URL)),#shared) otherwise null
,
// Extended script loader with auto data types
File.GetExcelDataFile = Document(
    "LPQ[File.GetExcelDataFile]",
    "Call table from excel file in INNO folder.",
    {[
        Description = "File.GetExcelDataFile",
        Code = "LPQ[File.GetExcelDataFile](""File name without extension"", ""Table name (table if null)"")",
        Result = "Imported table with types formated."
    ]},
    (fileName as text, optional tableName as text) as table =>
    let
        folder = "C:\Users\Laktionov.I\OneDrive\INNO\",
        filepath = folder & fileName & ".xlsx",
        tblName = tableName ?? "table",
        data = Excel.Workbook(File.Contents(filepath), null, true),
        dataTable = data{[Item = tblName, Kind = "Table"]}[Data],
        fnDetectAndTransformTypes = (SourceTable as table) =>
        let
            moneyList = {"amount", "price", "deal", "cost", "rebate", "msrp"},
            FnDetectType = (SourceTable, ColumnName) =>
            let
                ListColumn = Table.Column(SourceTable, ColumnName),
                Money = List.Contains(moneyList, Text.Trim(ColumnName), (x, y) => Text.StartsWith(y, x, Comparer.OrdinalIgnoreCase)),
                NumberPercentage = List.Count(List.Select(ListColumn, each _ is number)) / List.Count(ListColumn),
                IntPercentage = List.Count(List.Select(ListColumn, each _ is number and _ = Number.Round(_, 0))) / List.Count(ListColumn),
                DatePercentage = List.Count(List.Select(ListColumn, each _ is date)) / List.Count(ListColumn),
                DateTimePercentage = List.Count(List.Select(ListColumn, each _ is datetime)) / List.Count(ListColumn),
                Max = List.Max({NumberPercentage, DatePercentage, DateTimePercentage, IntPercentage}),
                Result =
                    if Money then {ColumnName, Currency.Type}
                    else if Max < .95 then {ColumnName, type any}
                    else if DateTimePercentage = Max then {ColumnName, type datetime}
                    else if DatePercentage = Max then {ColumnName, type date}
                    else if IntPercentage = Max then {ColumnName, Int64.Type}
                    else {ColumnName, type number}
            in
                Result,
            ConvertTypesList = List.Transform(Table.ColumnNames(SourceTable), each FnDetectType(SourceTable, _)),
            ConvertedTable = Table.TransformColumnTypes(SourceTable, ConvertTypesList)
        in
            ConvertedTable,
        Result = fnDetectAndTransformTypes(dataTable)
    in
        Result)
,

/////////////////////////
// Value               //
/////////////////////////
Value.TypeText = (value as any) =>
    if value is binary then "binary" else
    if value is date then "date" else
    if value is datetime then "datetime" else
    if value is datetimezone then "datetimezone" else
    if value is duration then "duration" else
    if value is function then "function" else
    if value is list then "list" else
    if value is logical then "logical" else
    if value is none then "none" else
    if value is null then "null" else
    if value is number then "number" else
    if value is record then "record" else
    if value is table then "table" else
    if value is text then "text" else
    if value is time then "time" else
    if value is type then "type" else
    if value is any then "any"
    else error "unknown -- not a primitive type!"

]
,

Result = _extensionLibrary

in
    Result
