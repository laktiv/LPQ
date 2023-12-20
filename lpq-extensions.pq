/**
*  Inspired by Power-BI-Desktop-Query-Extensions - https://github.com/tnclark8012/Power-BI-Desktop-Query-Extensions
*  Expression.Evaluate(Text.FromBinary(Web.Contents("https://raw.githubusercontent.com/laktiv/LPQ_Extensions/main/l-pq-extensions.pq?token=GHSAT0AAAAAACI72RZ2G7TJGDOWJV7W4XA4ZLY4GTA")),#shared)
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
// Number              //
/////////////////////////

Number.ProperNumber = Document(
    "Number.ProperNumber",
    "Concerts text as number to proper number type. Also will erase inconvinient thouthands and decimal symbols to regular format.
    In case of error will return input text",
    {[ Description = "Allow commas", Code="LPQ[Number.ProperNumber](""It's over 9,000!"", null, {"",""}, 0)", Result = "9000" ]},
(text as text, optional digitsAfter as number, optional allowCharacters as list, optional startIndex as number) => 
    let
        after = if digits_after is null then 0 else digits_after,
        consider = if startIndex is null then text else Text.Range(text,startIndex), 
        _allowCharacters = if allowCharacters is null then {} else allowCharacters,
        numberSeries = List.FirstN(List.Skip(Text.ToList(consider), each not Text.IsNumber(_)), each Text.IsNumber(_) or List.Contains(_allowCharacters, _))
    in 
        if text is null then null else Number.Round(Text.FromList(numberSeries), after)
),



///////////////////////// 
// Text                //
/////////////////////////

Text.PowerTrim = ((text as text, optional char_to_trim as text) =>
    let 
        char = if char_to_trim = null then " " else char_to_trim,
		split = Text.Split(text, char),
		removeblanks = List.Select(split, each _ <> "")
	in  
        Text.Combine(removeblanks, char)
),

Text.ReplaceByList = (try (txt as text, replacementLists as list) as text => List.Accumulate(replacementList, txt, (s,x) => Text.Replace(s,x{0},x{1})) otherwise null),

Text.RegexMatch = ((txt as text, regex as text, delim as text) => try 
    Web.Page("<script>var x='"&txt&"';var delim='"&delim&"';var regex=/"&regex&"/gi;var result=x.match(regex).join(delim);document.write(result);</script>")[Data]{0}[Children]{0}[Children]{1}[Text]{0}
    otherwise null
),

// Code                //
Text.RemoveHtmlTags = (let func = (HTML) =>
    let
        Check = if Value.Is(Value.FromText(HTML), type text) then HTML else "",
        Source = Text.From(Check),
        SplitAny = Text.SplitAny(Source,"<>"),
        ListAlternate = List.Alternate(SplitAny,1,1,1),
        ListSelect = List.Select(ListAlternate, each _<>""),
        TextCombine = Text.Combine(ListSelect, "")
    in
        TextCombine, documentation = [
            Documentation.Name =  " Text.RemoveHtmlTags"
            , Documentation.Description = "Remove Html Tags"
            , Documentation.LongDescription = " Removes all Html tags from a text"
            , Documentation.Category = " Text.Modification"
            , Documentation.Source = " Inspired by a solution from Bill Szysz"
            , Documentation.Author = " Imke Feldmann: www.TheBIccountant.com & Mike Carlo: PowerBI.Tips"
            , Documentation.Examples = {[
                Description = "Function that enables one to pass in a column that has HTML tags."
                , Code = "<div>my bit of text</div>"
                , Result = "my bit of text "]}]
    in
        Value.ReplaceType(func, Value.ReplaceMetadata(Value.Type(func), documentation))
)




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
