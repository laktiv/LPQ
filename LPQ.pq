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
    {[ Description = "Split on non-alphanumeric", Code = "LPQ[Splitter.SplitTextByNotIn](LPQ[Text.AlphaNumeric])(""Power BI is #1"")", Result = "{""Power BI is "", ""1""}"]},
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
            [ Description = "Using separate lists", Code = "Switch(1, {1, 2, 3}, {""A"", ""B"", ""C""})", Result = "A"],
            [ Description = "Using one paired list", Code = "Switch(1, {{1, ""A""}, {2, ""B""}, {3, ""C""}})", Result = "A"]
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


/////////////////////////
// Text                //
/////////////////////////
Text.Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz" & "АБВГДЕЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдежзийклмнопрстуфхцчшщъыьэюя",
Text.AlphaNumeric = Text.Alphabet & "0123456789",
Text.FromList = (list as list) => List.Accumulate(list, "", (state, current) => state & Text.From(current)),
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
    {[
        Description = "Function to import an external .pq script",
        Code = "ImportPq.FromFile(""C:\docs\external-script.pq"")",
        Result = "code from .pq file"
    ]},
    (filepath as text) as any => try Expression.Evaluate(Text.FromBinary(File.Contents(filepath), TextEncoding.Utf8), #shared) otherwise null
),
ImportPq.GitHubRaw = (URL as text) as any => try Expression.Evaluate(Text.FromBinary(Web.Contents(URL)), #shared) otherwise null,


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
