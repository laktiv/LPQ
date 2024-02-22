/**
*  Inspired by Power-BI-Desktop-Query-Extensions - https://github.com/tnclark8012/Power-BI-Desktop-Query-Extensions
*  Expression.Evaluate(Text.FromBinary(Web.Contents("https://raw.githubusercontent.com/laktiv/LPQ/main/LPQ.pq?token=GHSAT0AAAAAACI72RZ3KL2S5TZVMUGTD76GZNNNE2A")),#shared)
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
            Documentation.Examples = examples
        ]),


/////////////////////////
// INNO                //
/////////////////////////
Inno.SortTablePn = Document(
    "Inno.SortTablePn",
    "Returns sorted table by INNO partnumbers in proper order.",
    "Also could filter-out old models.",
    {[  Description = "Inno.SortTablePn",
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
        Sorting
),
Inno.TableTypes = Document(
    "Inno.TableTypes",
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
    in  Result
),


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
    )
),
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
            Result
),
List.ToText = Document(
    "List.ToText",
    "Converts a list to a textual representation. Inverse of List.From",
    { [ Description = "Conver to text", Code = "List.ToText({ 1, 2, 3})", Result = """{1, 2, 3}"""] },
    (list as list) =>
            List.Accumulate(list, "{", (state, current) => current & Text.From(current)) & "}"
),

/////////////////////////
// Number              //
/////////////////////////
Number.Digits = {0,1,2,3,4,5,6,7,8,9},
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
            if text is null then null else Text.FromList(numberSeries)
),
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
                if value <= 26 then GetLetter(value) else func(value, 1) & GetLetter(value)
),
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
        if text is null then null else Number.Round(Text.FromList(numberSeries), after)
),

/////////////////////////
// Splitters           //
/////////////////////////
Splitter.SplitTextByNonAlpha = Document(
    "Splitter.SplitTextByNonAlpha",
    "Splits text by characters that aren't [A-Za-z]",
    {[ Description = "Split text", Code="LPQ[Splitter.SplitTextByNonAlpha](""A1B,C"")", Result = "{ ""A"", ""B"", ""C"" }"]},
    (line as text) => Splitter.SplitTextByNotIn(Text.Alphabet)
),

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
                    List.Combine({List.RemoveLastN(state, 1), {current}}))
),

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
            if hasPairs then usingPairs else usingCases
),

/////////////////////////
// Table               //
/////////////////////////
Table.RenameColumn = (table as table, column as text, newName as text) => Table.RenameColumns(table, {{column, newName}}),
Table.RenameColumns = (Table as table, OldColumns as list, NewColumns as list, optional SelectHeaders as logical) =>
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
    in  try RenameTable otherwise null,

Table.TrimHeader = (Table as table) =>
    let hRaw = Table.ColumnNames(Table),
        hTrim = List.Transform(hRaw, each Text.Trim(_))
    in  if List.Count(List.Difference(hRaw, hTrim)) > 0 
            then Table.RenameColumns(Table, List.Zip({hRaw, hTrim}))
            else Table,

// Replaces a value if it matches a predicate
Table.ReplaceValueIf = (table as table, replaceIf as function, after as any, columnNameOrList as any) => 
    Table.ReplaceValue(table, null,after, (text, old, new)=>if replaceIf(text) then new else text, if columnNameOrList is list then columnNameOrList else {columnNameOrList}),
// Splits camelCased and PascalCased column names. 
Table.SplitColumnNames = (table as table) => Table.RenameColumns(table, List.Transform(Table.ColumnNames(table), each {_, Text.SplitCamelCase(_)})), 
// Splits camelCased and PascalCased text in a column. 
Table.SplitColumnText = (table as table, columns as list) => if List.Count(columns) = 0 then table else Table.TransformColumns(@Table.SplitColumnText(table, List.Skip(columns, 1)), {{List.First(columns), Text.SplitCamelCase}}),

Table.TransformColumn = (table as table, column as text, transform as function) => Table.TransformColumns(table, {{column, transform}}),
Table.RenameAndTransformColumn = (table, currentName as text, newName as text, transform as function) => Table.TransformColumn(Table.RenameColumns(table, {currentName, newName}), newName, transform),


/////////////////////////
// Text                //
/////////////////////////
Text.Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz" & "АБВГДЕЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдежзийклмнопрстуфхцчшщъыьэюя",
Text.AlphaNumeric = Text.Alphabet & "0123456789",
Text.FromList = (list as list) => List.Accumulate(list, "", (state, current) => state & Text.From(current)),
// Is text all uppercase? returns false if any non-alpha characters are present
Text.IsUpperCase = (text as text) => List.AllTrue(List.Transform(Text.ToList(text), (letter)=>Text.Contains(Text.Alphabet, letter) and letter = Text.Upper(letter))),
Text.IsAlpha = (text as text) => List.MatchesAll(Text.ToList(text), each Text.Contains(Text.Alphabet, _)),
Text.RemoveExtraWhitespace = (text as text) => Text.Combine(Splitter.SplitTextByWhitespace()(text)," "),
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
      current
),

Text.Substring = (text as text, start as number, optional count as number) =>
   let
      start = if start >= 0 then start else error "start index should be >= 0",
      end = if count = null then Text.Length(text) else if count <= Text.Length(text) then count else error "count should be <= text length",
      textList = Text.ToList(text),
      substr = Text.FromList(List.FirstN(List.Skip(textList, start), end - start))
   in substr,
Text.IsNumber = (text as text) => try Number.FromText(text) is number otherwise false,
Text.PositionAfter = (text as nullable text, substring as text) =>
    let
        firstIndex = Text.PositionOf(text, substring),
        indexAfter = if firstIndex >=0 then firstIndex + Text.Length(substring) else -1
    in
        if text is null then -1 else if indexAfter >= 0 and indexAfter < Text.Length(text) then indexAfter else -1,
Text.Until = (text as text, endDelimiter as text, optional startIndex as number) =>
    let
        start = if startIndex = null then 0 else startIndex,
        textFromStart = Text.Substring(text, start),
        delimPosition = if Text.PositionOf(textFromStart, endDelimiter) >= 0 then Text.PositionOf(textFromStart, endDelimiter) else Text.Length(textFromStart)
    in
        if text is null then null else Text.Range(textFromStart, 0, delimPosition),

Text.GetMatch = (Text, Lookup) => try List.Select(Text.Split(Text," "), each List.ContainsAny(List.LastN(Text.Split(_," "), 1), Lookup, (x,y) => Text.Contains(x,y,Comparer.OrdinalIgnoreCase))){0} otherwise null,

Text.PowerTrim = (text as text, optional char_to_trim as text) =>
    let
        char = if char_to_trim = null then " " else char_to_trim,
		split = Text.Split(text, char),
		removeblanks = List.Select(split, each _ <> "")
	in
        Text.Combine(removeblanks, char),

Text.ReplaceByList = try (txt as text, replacementLists as list) as text => List.Accumulate(replacementLists, txt, (s,x) => Text.Replace(s,x{0},x{1})) otherwise null,

Text.RegexMatch = (txt as text, regex as text, delim as text) => try
    Web.Page("<script>var x='"&txt&"';var delim='"&delim&"';var regex=/"&regex&"/gi;var result=x.match(regex).join(delim);document.write(result);</script>")[Data]{0}[Children]{0}[Children]{1}[Text]{0}
    otherwise null,

// Code                //
Text.RemoveHtmlTags = Document(
    "Text.RemoveHtmlTags",
    "Function that enables one to pass in a column that has HTML tags.",
    {[
        Description = "Removes all Html tags from a text",
        Code = "LPQ[Text.RemoveHtmlTags](""<div>my bit of text</div>"")",
        Result = "my bit of text"
    ]},
    (HTML) =>
        let
            Check = if Value.Is(Value.FromText(HTML), type text) then HTML else "",
            Source = Text.From(Check),
            SplitAny = Text.SplitAny(Source,"<>"),
            ListAlternate = List.Alternate(SplitAny,1,1,1),
            ListSelect = List.Select(ListAlternate, each _<>""),
            TextCombine = Text.Combine(ListSelect, "")
        in
            TextCombine
    ),


/////////////////////////
// Import              //
/////////////////////////
ImportPq.FromFile = Document(
    "ImportPq.FromFile",
	"Text.",
    {[
        Description = "Function to import an external .pq script",
        Code = "ImportPq.FromFile(""C:\docs\external-script.pq"")",
        Result = "code from .pq file"
    ]},
    (Filepath as text) =>
        try Expression.Evaluate(Text.FromBinary(File.Contents(Filepath)), #shared)
        otherwise null
    ),

ImportPq.GitHubRaw = (URL as text) =>
        try Expression.Evaluate(Text.FromBinary(Web.Contents(URL)), #shared)
        otherwise null,


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

],

Result = _extensionLibrary

in
    Result
