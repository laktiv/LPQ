/**
*  Inspired by Power-BI-Desktop-Query-Extensions - https://github.com/tnclark8012/Power-BI-Desktop-Query-Extensions
*  Expression.Evaluate(Text.FromBinary(Web.Contents("https://gist.github.com/laktiv/700a6d65e098189881ecd77e585b233a")),#shared)
**/

let

MissingTests = List.Select(Record.FieldNames(_extensionLibrary), each not
	List.MatchesAny(Tests, (testRecord as record) => Text.Contains(testRecord[CaseName], _))),
MissingDocumentation = List.Select(Record.FieldNames(_extensionLibrary), each not
	Record.HasFields(Value.Metadata(Value.Type(Record.Field(_extensionLibrary, _))), "Documentation.Examples")),
Tests = {
	[ CaseName = "List.From",
		Test = (library as record) =>
			TestUtils[AssertEqual](
				{"A", "B", "C"},
				library[List.From]("{A, B, C}"),
				"Text should be split on ,") ],
	[ CaseName = "List.From - no braces",
		Test = (library as record) =>
			TestUtils[AssertEqual](
				{"A", "B", "C"},
				library[List.From]("A, B, C"),
				"Text should be split on ,") ],
	[ CaseName = "List.Flatten",
		Test = (library as record) =>
			TestUtils[AssertEqual](
				{ 1, 2, Table.FromRecords({[x=1]}), 3, 4, 5},
				library[List.Flatten]({ 1, 2, Table.FromRecords({[x=1]}), {3, 4, 5} }),
				"Flattened list") ],
	[ CaseName = "Number.ToLetters",
		Test = (library as record) =>
			TestUtils[AssertEqual](
				{ "A", "BC", "BAD" },
				{ library[Number.ToLetters](1),
					library[Number.ToLetters](55),
					library[Number.ToLetters](1382) },
				"1, 2, and 3 letter translations") ],
	[ CaseName = "Table.RenameColumn",
		Test = (library as record) =>
			TestUtils[AssertEqual](
				{"NewName", "NumberCol"},
				Table.ColumnNames(library[Table.RenameColumn](TestUtils[SimpleTable], "TextCol", "NewName")),
				"Column should have new name") ],
	[ CaseName = "Text.Until",
		Test = (library as record) =>
			TestUtils[AssertEqual](
				"abc ",
				library[Text.Until]("abc 123", "1"),
				"Proper substring should be found") ],
	[ CaseName = "Text.Substring - infer end",
		Test = (library as record) =>
			TestUtils[AssertEqual](
				"BI Rules!",
				library[Text.Substring]("Power BI Rules!", 6),
				"Text.Substring('Power BI Rules', 6)") ],
	[ CaseName = "Switch case pairs",
		Test = (library as record) =>
			TestUtils[AssertEqual](
				2,
				library[Switch]("B", {{"A", 1}, {"B", 2}, {"C", 3}}),
				"B is second case") ],
	[ CaseName = "Switch separate case and result",
		Test = (library as record) =>
			TestUtils[AssertEqual](
				4,
				library[Switch]("TE", {"QB", "RB", "WR", "TE", "K", "D/ST"}, {1, 2, 3, 4, 5, 6}),
				"TE is 4th case") ] },

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
			if not Value.Is(actualType, Value.Type(expectedType)) then error "Expected type " & typeAsText(expectedType) & " does not match actual type " & typeAsText(actualType)
			else if expected = actual then true
			else if expectedAsText = actualAsText then error "Cannot provide accurate failure message! Text versions of expected and actual are identical, but they failed the equality test!"
			else error "Expected: " & expectedAsText & "; Actual: " & actualAsText & "; Reason: " & description,
	SimpleTable = Table.FromRecords({[TextCol = "A", NumberCol = "1"], [TextCol = "B", NumberCol = 2], [TextCol = "C", NumberCol = 3]}) ],

TestResults =
	let failedTests = List.Select(
			List.Transform(Tests, (suite as record) =>
				let
					testResult = try suite[Test](_extensionLibrary)
				in
					if testResult[HasError] then Error.Record(suite[CaseName], testResult[Error][Message], null) else true), each _ <> true)
	in 	if List.IsEmpty(failedTests) then "All " & Text.From(List.Count(Tests)) & " tests passed! :)" else failedTests,

_extensionLibrary = [

// #	DEPENDENCIES
Document = (name as text, description as text, category as text, valueOrExample as any, optional valueIfExample as any) =>
	let value = if valueIfExample is null then valueOrExample else valueIfExample,
			examples = if valueIfExample is null then {} else valueOrExample,
			category = if category is null then {} else category
	in 	Value.ReplaceType(value, Value.Type(value) meta [
			// [Description="", Code="", Result=""]
			Documentation.Name = name,
			Documentation.Description = description,
			Documentation.Category = category,
			Documentation.Examples = examples
		]),
Switch = Document(
	"Switch",
	"Given a value, find it's paired item <br>"&
		"Switch(value as any, cases as list, results as list, optional default as any) <br>"&
		"Switch(value as any, pairs as list, optional default as any)",
	"Development", {
	[ Description = "Using separate lists",
		Code="LPQ[Switch](1, {1, 2, 3}, {""A"", ""B"", ""C""})",
		Result="A"],
	[ Description = "Using one paired list",
		Code="LPQ[Switch](1, {{1, ""A""}, {2, ""B""}, {3, ""C""}})",
		Result="A"]},
	(value as any, casesOrPairs as list, optional resultsOrDefault as any, optional default as any) =>
		let hasPairs = List.First(casesOrPairs) is list,
				usingPairs =
					let targetPosition = List.PositionOf(casesOrPairs, value, Occurrence.First, (case, theValue) => theValue = case{0})
					in 	if targetPosition = -1 then resultsOrDefault else casesOrPairs{targetPosition}{1},
				usingCases =
					let cases = casesOrPairs,
							results = resultsOrDefault
					in 	if List.IsEmpty(cases) or List.IsEmpty(results) then default else if value = List.First(cases) then List.First(results) else @Switch(value, List.Skip(cases, 1), List.Skip(results, 1), default)
		in 	if hasPairs then usingPairs else usingCases)
,


// ## 	BULK REPLACE & TABLES
MyFindReplace =
	let table = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("JcpBDoIwEIXhq5iuJMGF8QRQSpxYW0LHuCAsEFk0Ykuwxus7Lbs33/xdxxBYztCyPu8YOOdPFd2glKYRDX/WxeIOamv4GaQksduIZr7LtBKZWyPaJNy/l2EMZFxfm4Jj0nIexhdZKQt+SVKv/hMeNkykdasNloAiffbHg5l92ImnDda7jIJIUuNOVICgVcb6/g8=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Find = _t, Replace = _t])
	in  Table.ReplaceValue(table,"",null,Replacer.ReplaceValue,{"Replace"}),
MyFindReplace_Sample =
	let table = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("TY7LDoIwFAV/hXQrxijInrY8GnkFKi4ICyQsGhUIYvh925JCdzOTe5NTVQCXBJgA1GYFQhxvXAauQpwpMg7GFnFuaWxrfNXYUezYkM1KMkSOnpLldA7gLhddLF3sXSjjRInkoPOHqe14CDw/zZEnK+n7wcI8kiRJOcivhfXi70GS9QaFJIp4YSuIVvzGbuKpuGdeLgsaPmPTiu0ojTMXUVnhu2lfvMHIRTdZ/Gn4zk82iyV+nhYUEsq31H8=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Find = _t, Replace = _t])
	in  Table.ReplaceValue(table,"",null,Replacer.ReplaceValue,{"Replace"}),
MyFindReplace_InnoSpecs =
	let table = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("bY/PboJAEIdfZcKpjcQo9gVaOOiBlMSaHoiHLYxKxJ3NsqTi03f/AAumt5nv983Mbp4H8SF5h5gkNkFom+AY5sEHUaMgrqm4wku6fbzqcGBXZ7AGnTDmjrh43wpR4w25YrKDjH5R6iucY6FImlNPyA6leCNt261aGdt+pWK8ZLKEgRM/VWfv7asHTtfsuEJ5YgV65asTM8X08F2V6uIdPzY7mlSNqFk3/8Rnq0SrrLhN0h1Ey7Wmppyy1c8zZFPQb85IKlgv30yWZP8lPtjcQVemhfvGkmhCovFOCM7oTVjA7BkujoaxMT7+AQ==", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Find = _t, Replace = _t])
	in  Table.ReplaceValue(table,"",null,Replacer.ReplaceValue,{"Replace"}),
BulkTableReplace =
	let BulkReplace = (
		DataTable as table, 
		FindReplaceTable as table, 
		DataTableColumn as list, 
		optional ReplaceFullCell as logical
	) =>
		/*
		*   Replace values in column by table list of replaces
		*   Create extra table "MyFindReplace" with [Find] and [Replace] columns
		*/
		let
			FindReplaceList = Table.ToRows(FindReplaceTable), 	//Convert the FindReplaceTable to a list using the Table.ToRows function so we can reference the list with an index number
			Counter = Table.RowCount(FindReplaceTable),			//Count number of rows in the FindReplaceTable to determine how many iterations are needed
			BulkReplaceValues = (DataTableTemp, n) =>			//Define a function to iterate over our list with the Table.ReplaceValue function
			let
				ReplaceType = if ReplaceFullCell=null or ReplaceFullCell=false 	// set Replace type
					then Replacer.ReplaceText 
					else Replacer.ReplaceValue,

				ReplaceTable = Table.ReplaceValue(				//Replace values using nth item in FindReplaceList
					DataTableTemp,
					if FindReplaceList{n}{0} = null then "" else FindReplaceList{n}{0},		//replace null with empty string
					if FindReplaceList{n}{1} = null then "" else FindReplaceList{n}{1},
					ReplaceType,
					DataTableColumn
				)
			in
				if n = Counter - 1			//if we are not at the end of the FindReplaceList then iterate through Table.ReplaceValue again
					then ReplaceTable
					else @BulkReplaceValues(ReplaceTable, n+1),
			Output = BulkReplaceValues(DataTable, 0)		//Evaluate the sub-function at the first row
		in
			Output
	in
		BulkReplace
,


// #	DATE
// ## 	Inno
Date.DateFromISOwn = Document(
	"Date.DateFromISOwn",
	"Compute date of the first day of the week given an ISO weeknumber. If the Year is not entered, it will default to 'this' year.",
	"Dates.Convertasion",
	{[Description="Year must be a full year: eg 2021 vs 21. The latter will not be interpreted as the year 2021.",
		Code="LPQ[Date.DateFromISOwn](""week number"", opt. ""year"")",
		Result="Date on start of week #"   ]},
	(wn as number, optional year as number) as date =>
		let
			yr = if year = null then Date.Year(DateTime.LocalNow()) else year,
			wn1Start = Date.StartOfWeek(#date(yr,1,1), Day.Monday),
			w1 = if Date.AddDays(wn1Start,3) < #date(yr,1,1) then Date.AddDays(wn1Start,7) else wn1Start
			//-2 if not 01.01.01 start of the year(?)
		in Date.AddDays(w1, 7*(wn-1)) ),
DateTime.FromUnix = Document(
	"DateTime.FromUnix",
	"Converts UNIX timestamp to DateTime.",
	"Date.Transformation",
	{[Description="Converts UNIX timestamp to DateTime.",
		Code="LPQ[DateTime.FromUnix](""1710291332078"")",
		Result="2024-03-13T00:55:32.0780000"   ]},
	(ts) => #datetime(1970,1,1,0,0,0) + #duration(0,0,0,Number.FromText(ts)/1000) ),
// ## 	Basic calendar
Date.Calendar = Document(
	"Date.Calendar",
	"Generate a calendar table for a given date span - can be text or proper dates. Current columns are Date, DayOfWeek, Month, MonthNum, WeekStartData, WeekStart, Year, YearMonth",
	"Dates.Date",
	{[Description="2024 calendar",
		Code="LPQ[Date.Calendar](""1/1/2024"", ""12/31/2024""",
		Result="2024 calendar"	 ]},
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
		in 	Result),
Date.DatesBetween = Document(
	"Date.DatesBetween",
	"Returns a list of dates in a given span (inclusive). Start and end parameters can be any order",
	"Dates.Date",
	{[Description="Date range",
		Code="LPQ[Date.DatesBetween](""1/1/2016"", ""1/3/2016"")",
		Result="{""1/1/2016"", ""1/2/2016"", ""1/3/2016""}" ]},
	(start as any, end as any) =>
		let
			StartDate = Date.From(start),
			EndDate = Date.From(end),
			adjustedStart = List.Min({StartDate, EndDate}),
			adjustedEnd = List.Max({StartDate, EndDate}),
			GetDates = (start as date, end as date, dates as list)=> if start > end then dates else @GetDates(Date.AddDays(start, 1), end, List.Combine({dates, {start}})),
			Dates = GetDates(adjustedStart, adjustedEnd, {})
		in
			Dates),
Date.DayName = Document(
	"Date.DayName",
	"Returns the English day of the week name for a date",
	"Dates.Date",
	{[Description="Get the day name",
		Code="LPQ[Date.DayName](""9/9/2016"")",
		Result="Friday" ]},
	(date as any) =>
		let Eng = {"Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"},
			Ukr = {"Неділя","Понеділок","Вівторок","Середа","Четвер","Пятниця","Субота"},
			Lang = Eng
		in  Switch(Date.DayOfWeek(DateTime.From(date)), {0,1,2,3,4,5,6}, Lang, null)),
Date.MonthName = Document(
	"Date.MonthName",
	"Returns the English month name for a date",
	"Dates.Date",
	{[Description="Get the month name",
		Code="LPQ[Date.MonthName](""9/9/2016"")",
		Result="September"	 ]},
	(date as any, optional lang as text) =>
		let eng = {"January","February","March","April","May","June","July","August","September","October","November","December"},
			ukr = {"Січень","Лютий","Березень","Квітень","Травень","Червень","Липень","Серпень","Вересень","Жовтень","Листопад","Грудень"},
			Lang = if lang is null or lang="" then eng else ukr,
			monthNumber = if date is number then date else Date.Month(DateTime.From(date))
		in  Switch(monthNumber, {1,2,3,4,5,6,7,8,9,10,11,12}, Lang, null))
,


// #	LIST
// ## 	Compare
List.CompareListsCheck = (ListSource as list, ListCompare as list, optional CaseIgnore as logical) =>
	//fxList = (l) => try Text.Combine( List.Transform(List.Distinct(List.Transform(l, each Text.Split(_, " "){1})), each Text.From(_)), "," ) otherwise "",
	if (CaseIgnore ?? {CaseIgnore=""}{1}?) <> null
		then List.ContainsAny(ListSource, ListCompare, (x, y) => Text.Contains(x, y, Comparer.OrdinalIgnoreCase))
		else List.ContainsAny(ListSource, ListCompare, (x, y) => x=y),
List.CompareListsMatches = (ListSource as list, ListCompare as list, optional CaseIgnore as logical) =>
	//fxList = (l) => try Text.Combine( List.Transform(List.Distinct(List.Transform(l, each Text.Split(_, " "){1})), each Text.From(_)), "," ) otherwise "",
	if (CaseIgnore ?? {CaseIgnore=""}{1}?) <> null
		then List.Select(ListSource, each List.Contains(ListCompare, _, (x,y) => Text.Contains(x, y, Comparer.OrdinalIgnoreCase)))
		else List.Select(ListSource, each List.Contains(ListCompare, _, (x,y) => Text.Contains(x, y, Comparer.Equals))),
// ## 	Transform
List.Flatten = Document(
	"List.Flatten",
	"Recursively flattens list elements. The end result is a single list",
	"Lists.Transformation",
	{[Description="Flattening nested lists into one",
		Code="LPQ[List.Flatten]({ 1, 2, Table.FromRecords({[x=1]}), {3, 4, 5} })",
		Result="{ 1, 2, Table.FromRecords({[x=1]}), 3, 4, 5}"  ]},
	(list as list) => List.Accumulate(list, {}, (state, current) =>
		let
			currentListContent = if current is list then @List.Flatten(current) else {current}
		in
			List.Combine({state, currentListContent}))),
List.FlattenAdvanced = Document(
	"List.FlattenAdvanced",
	"Recursively flattens list elements. The end result is a single list",
	"Lists.Transformation",
	{[Description="Flatten a nested list with mixed levels into a simple list.",
		Code="LPQ[List.Flatten]({1, {2, 3}, {1, {2, 3}}})",
		Result="{1, 2, 3, 1, 2, 3}"  ]},
	(L as list) as list =>
		let
				IsListType = (x) as logical => Value.Type(x) = List.Type,
				FlattenOneLevel = (L as list) as list =>
						List.Combine(
								List.Transform(L, each if IsListType(_) then _ else {_})
						),
				FlattenAll = (L as list) as list =>
						if List.MatchesAny(L, IsListType) // Check if L is nested
						then @FlattenAll(FlattenOneLevel(L))
						else L
		in
				FlattenAll(L)),
List.From = Document(
	"List.From",
	"Converts a text representation of a list into a list of the elements. Items are considered to be split by ,",
	"Lists.Transformation",
	{[Description="Convert a text list",
		Code="LPQ[List.From](""{A, B, C}"")",
		Result="{ ""A"", ""B"", ""C"" }"   ]},
	(simpleTextList as text) =>
		let
			trimWhitespace = Text.Trim(simpleTextList),
			listToSplit = Text.TrimEnd(Text.TrimStart(trimWhitespace, "{"), "}"),
			Result = List.Transform(Text.Split(listToSplit, ","), each Text.Trim(_))
		in
			Result),
List.ToText = Document(
	"List.ToText",
	"Converts a list to a textual representation. Inverse of List.From",
	"Lists.Transformation",
	{[Description="Conver to text",
		Code="LPQ[List.ToText]({ 1, 2, 3})",
		Result="""{1, 2, 3}"""	 ]},
	(list as list) => List.Accumulate(list, "{", (state, current) => current & Text.From(current)) & "}"),
	List.Swap = (_,_from as number ,_to as number) =>
	let
		from = List.Min({_from,_to}),
		to = List.Max({_from,_to})
	in if from=to then _ else
		List.Range(_,0,from)
		&{_{to}}
		&List.Range(_,from+1,to-from-1)
		&{_{from}}
		&List.Range(_,to+1),
		List.Shuffle = (n) => List.Accumulate({0..(n-2)},{0..n},(_,iterator)=>
								List.Swap(List.Buffer(_),
									Number.Round(Number.RandomBetween(iterator,n)),
									iterator
									)),
List.ToRecord = Document(
	"List.ToRecord",
	"Transform a list of strings to a record using a given lambda (passed values k). The built-in Record.FromList only takes static lists...",
	"List.ToRecord",
	{[Description="Transform a list of strings to a record using a given lambda (passed values k)",
		Code="LPQ[List.ToRecord]({""a"",""b""}, (k) => Text.Upper(k)))",
		Result="[a=""A"", b=""B""]"	 ]},
	(List as list, Lambda as function) as record =>
		let
			Transformed = List.Transform(List, Lambda)	//each Lambda(_)
		in
			Record.FromList(Transformed, List))
,


// #	NUMBERS
Number.Digits = {0,1,2,3,4,5,6,7,8,9},
Number.ParseText = Document(
	"Number.ParseText",
	"Returns the first number (1+ consecutive digits) in the given text. Optionally, provide allowed characters to ignore them",
	"Numbers.Transformation",
	{[Description="Allow commas",
		Code="LPQ[Number.ParseText](""It's over 9,000!"", 0, {"",""})",
		Result="9000"  ]},
	(text as text, optional startIndex as number, optional allowCharacters as list) =>
		let
			consider = if startIndex is null then text else Text.Range(text,startIndex),
			_allowCharacters = if allowCharacters is null then {} else allowCharacters,
			numberSeries = List.FirstN(List.Skip(Text.ToList(consider), each not Text.IsNumber(_)), each Text.IsNumber(_) or List.Contains(_allowCharacters, _))
		in
			if text is null then null else Text.FromList(numberSeries)),
Number.ToLetters = Document(
	"Number.ToLetters",
	"Converts a number (starting at 1) to an alphabet representation. Works like column headers in Excel.",
	"Numbers.Transformation",
	{[Description="Column 27",
		Code="LPQ[Number.ToLetters](27)",
		Result="AB" ]},
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
			if value <= 26 then GetLetter(value) else func(value, 1) & GetLetter(value)),
Number.ProperNumber = Document(
	"Number.ProperNumber",
	"Concerts text as number to proper number type. Also will erase inconvinient thouthands and decimal symbols to regular format. <br>" &
	"In case of error will return input text",
	"Numbers.Transformation",
	{[Description="Allow commas",
		Code="LPQ[Number.ProperNumber](""It's over 9,000!"", null, {"",""}, 0)",
		Result="9000"  ]},
	(text as text, optional digitsAfter as number, optional allowCharacters as list, optional startIndex as number) =>
	let
		after = if digitsAfter is null then 0 else digitsAfter,
		consider = if startIndex is null then text else Text.Range(text,startIndex),
		_allowCharacters = if allowCharacters is null then {} else allowCharacters,
		numberSeries = List.FirstN(List.Skip(Text.ToList(consider), each not Text.IsNumber(_)), each Text.IsNumber(_) or List.Contains(_allowCharacters, _))
	in
		if text is null then null else Number.Round(Text.FromList(numberSeries), after))
,


// #	SPLITTERS
Splitter.SplitTextByNonAlpha = Document(
	"Splitter.SplitTextByNonAlpha",
	"Splits text by characters that aren't [A-Za-zА-Яа-я]",
	"Splitter.Transformation",
	{[Description="Split text",
		Code="LPQ[Splitter.SplitTextByNonAlpha](""A1B,C"")",
		Result="{ ""A"", ""B"", ""C"" }"   ]},
	(line as text) => Splitter.SplitTextByNotIn(Text.Alphabet)),
Splitter.SplitTextByNotIn = Document(
	"Splitter.SplitTextByNotIn",
	"Splits text on any characters that aren't the provided 'safe' characters",
	"Splitter.Transformation",
	{[Description="Split on non-alphanumeric",
		Code="LPQ[Splitter.SplitTextByNotIn](LPQ[Text.AlphaNumeric])(""Power BI is #1"")",
		Result="{""Power BI is "", ""1""}"	 ]},
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


// #	TABLE
Table.AddBlankRow = (table as table) => Table.Combine({table, Table.FromRecords({[]})}),
Table.DrillIntoColumn = (table as table, columnName as text) =>
	let FindValue = (value as any) =>
		if value is list then
			if List.Count(value) = 1 then @FindValue(List.First(value))
			else if List.Count(value) = 0 then null
			else error "Couldn't find single value"
		else if value is table then
			if Table.RowCount(value) = 1 then @FindValue(List.First(Table.ToColumns(value)))
									else if Table.RowCount(value) = 0 then null
			else error "Couldn't find single value"
		else  value,
		Result = Table.TransformColumns(table, {{columnName, FindValue}})
	in  Result,
Table.EnterDataFormula = Document(
	"Table.EnterDataFormula",
	"Convert a table into a formula that's compatible with the Enter Data UI. This function returns an expression you can copy and paste to editthe table using Enter Data. Paste the formula, then click the gear icon next to the step name. This is an easy way to allow people to customize a table you've queried",
	"Table.Convertion",
	{[Description="Simple table",
		Code="Web.Page(Web.Contents(""https://en.wikipedia.org/wiki/List_of_Super_Bowl_champions"")){3}[Data]",
		Result="Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText(""jVVtb9s2EP4rRD4aRqI3S9TH2Gm7ZWkb1MGGom6Aq83IhBkyoOik2a8fT6Ski6Fu+yDdWea9Pffc8du3M342P7uVzrU/jrbZs7UTQgnb+q+lfzL/nFeLxIu0ror55pglWcK8viA6J3o96vVwJksSqvNRT5Ognn2fh1yuQClo2cq8/DCvmMbCPzmmUWaLkEYyn2GkNIhFEFUQvBN1FkQexGI2uP8kXtg73SjQO3YLzkrjMEgRn/NFEmrlpL66JLmnRM+JXhC9IvWR82lxWqvQz8KypTV6a9pYJxZ5nlex1mpMg5dEp98p5OR7TWEmqaYLkgY2eQ2avbegt7LdGlbUof1dApgLz/MASQCcF0EEpH3wDuJiNkfXWUAabT9YITRbwqvHeXsIPovBZ4S5jDWl3kdZjbqHfBbKmXUkGd1iA78ae2AfJGg36ZRH46QzTqKIrtLR1Wc4IA82FzemZZe68cRv2ReQu5Bs3vO/7JOt5kixEt886d45Vp1kg8u/oN1L3Tij2Rexaw9ST3qqAj95FIGmPFC4Jhl+lPAo2ZVRT/vgKeuHoeg9pWT2MqLnhCDkO6ccROSWoJx8NFZsLn7XO48pPBklcQCV6yNmdDJKPrbJj+LIejol9UmUj1Jr0RoH7E/pQWnQc9LPXNJ7ronngEkVyFaVszHh48MDKMOWUqlJN3VCZoDAUxMY6pwkiIASCmwu1u6c3ZijRDY8tpFeHQh5nIUq0r4O3Iotw//WApxTgnm5h5fDhHW3DGdhKDtRBOuMdsNHfhZDy7vCYoEdo8dhzkghePQP0C3uT+le2WovxUOfQfqmieX8Hkeuvh8M/eEtNB5XAXbShoesk3JM93YvFewE0hPYO2g8ejHbjHaExxEkha6kXzdag5M+oG5ATRuS9nF+UukKrGeqBr9ftNuHmT3xgDs6LKZ4BaR0g1wL1xdKAUaK3w+n7+DxCbo1tjxut6CFsFNGuAXehvhslfC9YGuQejIOjslogkv4SgrfgNUebDOWk77hNh1f/OfS+RXmx+o9qK3R0zZ8jHIncA5bIdiddPCL8/VpDCv/Nh5nD/jO460mrPBSD1Hw10qJZ9Fdsf5qexmi4IOuUxzXT5uLy8HgSjhrpGM30vyP09f+PmmNfvYLQLBraI5g/9voN3NscS3fiZ/wbzG+/wM="", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type text) meta [Serialized.Text = true]) in type table [Appearances = _t, Team = _t, Wins = _t, Losses = _t, Winning percentage = _t, Season(s) = _t])" ]},
	(Table as table) =>
		let
			Encoded = Table.JsonEncode(Table),
			ColumnMeta= "[" & Text.Range(List.Accumulate(Table.ColumnNames(Table), "", (state, current) => state & ", " & current & " = _t"), 2) & "]",
			Text =
				"Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText(""" & Encoded & """, BinaryEncoding.Base64), Compression.Deflate))," &
				" let _t = ((type text) meta [Serialized.Text = true]) in type table " & ColumnMeta & ")"
		in
			Text
	),
// if fieldNames aren't specified, use the field names from the first row of the column.
// Perform a cross join of lists. Example usage:
// Table.FromListCrossJoin({ {ColorsTable[ColorName], "Color"}, {NumbersTable[Number], "Number"}})
// Will give me a new table with two columns, "Color" and "Number" which contains one row for each possible
// combination of colors and numbers
// Table.FromListCrossJoin({{"Red", "Blue"}, "Color"}, {{1,2,3}, "Number"}}) =
//	Table.FromRecords({[Color="Red", Number=1],[Color="Red", Number = 2],[Color="Red", Number = 3],[Color="Blue", Number=1],[Color="Blue", Number=2],[Color="Blue", Number=3]})
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
   in  Result,
Table.JsonDecode = (encoded as text) =>
	let Decompressed = Binary.Decompress(Binary.FromText(encoded, BinaryEncoding.Base64), Compression.Deflate),
		Decoded = Table.FromRows(Json.Document(Decompressed), let _t = ((type text) meta [Serialized.Text = true]) in type table [Name = _t, Number = _t])
	in  Decoded,
Table.JsonEncode = (table as table) =>
	let Rows = Table.ToRows(table),
		Json = Json.FromValue(Rows),
		Compressed = Binary.Compress(Json, Compression.Deflate),
		Encoded = Binary.ToText(Compressed, BinaryEncoding.Base64)
	in  Encoded,
// ## 	Replaces a value if it matches a predicate
Table.ReplaceValueIf = (table as table, replaceIf as function, after as any, columnNameOrList as any) =>
	Table.ReplaceValue(table, null,after, (text, old, new)=>if replaceIf(text) then new else text, if columnNameOrList is list then columnNameOrList else {columnNameOrList}),
// ## 	Splits camelCased and PascalCased column names.
Table.SplitColumnNames = (table as table) => Table.RenameColumns(table, List.Transform(Table.ColumnNames(table), each {_, Text.SplitCamelCase(_)})),
// ## 	Splits camelCased and PascalCased text in a column.
Table.SplitColumnText = (table as table, columns as list) => if List.Count(columns) = 0 then table else Table.TransformColumns(@Table.SplitColumnText(table, List.Skip(columns, 1)), {{List.First(columns), Text.SplitCamelCase}}),
Table.TransformColumn = (table as table, column as text, transform as function) => Table.TransformColumns(table, {{column, transform}}),
Table.RenameAndTransformColumn = (table, currentName as text, newName as text, transform as function) => Table.TransformColumn(Table.RenameColumns(table, {currentName, newName}), newName, transform)
,
Table.ToXml = (t) =>
		"<table>"&
		{List.Accumulate(
			Table.ToRecords(t),
			"",
			(state,curr)=>state&
			"<row>"&
				{List.Accumulate(
						Record.FieldNames(curr),
				"",
				(s,c)=>s&Text.Format("<#{0}>#{1}</#{0}>",
									{c,
								Record.Field(curr,c)})
				)}
			&"</row>"
		)}
	&"</table>"
,
Table.NumberColumns = (table as table, columns as list) as table =>
	let	NumberTypes = List.Repeat({type number}, List.Count(columns)),
		DataWithTypes = Table.TransformColumnTypes( table, List.Zip({columns, NumberTypes}) ),
		NumberNotNull = (number) => if number is null then 0 else number,
		DataNoNulls = Table.TransformColumns( DataWithTypes, List.Zip({ columns, List.Repeat({NumberNotNull}, List.Count(columns)), NumberTypes }) )
	in	Table.ReplaceErrorValues(DataNoNulls, List.Transform(columns, each {_,0})),
Table.TrimHeader = (Table as table) =>
	let	hRaw = Table.ColumnNames(Table),
		hTrim = List.Transform(hRaw, each Text.Trim(_))
	in	if List.Count(List.Difference(hRaw, hTrim)) > 0 then Table.RenameColumns(Table, List.Zip({hRaw, hTrim})) else Table
,
// ## 	Renaming columns
Table.RenameColumn = Document(
	"Table.RenameColumn",
	"Simple rename 1 column in the table.",
	"Table.Transformation",
	{[Description="Rename a column in table.",
		Code="LPQ[Table.RenameColumn](table, ""old column name"", ""new column name"")",
		Result="| new column |  column   |#(lf)| ---------- | ---------- |"	 ]},
	(Table as table, column as text, newName as text) => Table.RenameColumns(Table, {{column, newName}})
	),
Table.RenameColumnsAdv = Document(
	"Table.RenameColumnsAdv",
	"Advanced rename columns in table by Lists and with optional Filter.",
	"Table.Transformation",
	{[Description="Rename mane columns in table with possible selection.",
		Code="LPQ[Table.RenameColumnsAdv](table, ""{""col1"",""col2""}"", ""{""new1"",""new2""}"", True)",
		Result="Table with new column's names and Selected with only renamed columns.#(lf)" &
					"|	new1   |	new2	|"	 ]},
	(Table as table, OldColumns as list, NewColumns as list, optional SelectHeaders as logical) =>
		let	selectColumns = if (SelectHeaders ?? {SelectHeaders=""}{1}?) = null then false else SelectHeaders,
			// Merge Old names and New names
			columnsZip = List.Zip({ OldColumns, NewColumns }),
			// Get table headers and Select list ny Old names
			headers = List.Transform(Table.ColumnNames(Table), each _),
			selectHeaders = List.Select(headers, (x) => List.AnyTrue(List.Transform(OldColumns, each x = _))),
			// Get only picked columns
			SelectTable = if selectColumns then Table.SelectColumns(Table, selectHeaders) else Table.SelectColumns(Table, headers),
			// Rename columns
			RenameTable = Table.RenameColumns(SelectTable, columnsZip)
		in	try RenameTable otherwise null),
Table.RenameColumnsByIndices = (Source as table, Indices as list, ColumnNamesNew as list) =>
	let	ColumnNames = Table.ColumnNames(Source),
		ColumnNamesList = List.Transform(Indices, each ColumnNames{_} ),
		ZippedList = List.Zip({ ColumnNamesList, ColumnNamesNew } ),
		#"Renamed Columns" = Table.RenameColumns(Source, ZippedList )
	in 	#"Renamed Columns",
Table.RenameColumnsByOrder = (InputTable as table, ColumnNumbers as list, NewColumnNames as list) =>
	let
		OldColumnNames = Table.ColumnNames(InputTable),
		Indexed = List.Zip({OldColumnNames, {0..-1+List.Count(OldColumnNames)}}),
		Filtered = List.Select(Indexed, each List.Contains(ColumnNumbers,_{1})),
		IndexRemoved = List.Transform(Filtered, each _{0}),
		RenameList = List.Zip({IndexRemoved,NewColumnNames}),
		RenamedColumns = Table.RenameColumns(InputTable, RenameList)
	in
		RenamedColumns,
Table.TrimHeaders = (Table as table) =>
	let hRaw = Table.ColumnNames(Table),
		hTrim = List.Transform(hRaw, each Text.Trim(_))
	in  if List.Count(List.Difference(hRaw, hTrim)) > 0 then Table.RenameColumns(Table, List.Zip({hRaw, hTrim})) else Table,
Table.PromoteHeadersNonEmpty = (table as table) as table =>
	/** Promote the first table row as column names, but keep the current header if the value in the first row is empty. For simpler use cases check built-in function `Table.PromoteHeaders` **/
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
	in  RenamedColumns,
Table.SelectHeaders = (List as list, ColumnStart as text, ColumnEnd as text, optional Index as text) =>
	let pos1 = List.PositionOf(List, ColumnStart),
		pos2 = List.PositionOf(List, ColumnEnd),
		range = List.Range(List, pos1 + 1, pos2 - pos1 - 1)
	in  if Index=null then range else List.Union({{Index}, range}),
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
		Table.RemoveColumns(Filter_out_empty_rows, {"Custom"}),
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
		Table.SelectColumns(tbl, non_empty_columns(tbl_no_erors)),
// ## 	return a table with all errors replaced with specified value or null
Table.ReplaceAllErrorsInTable = (tbl as table, optional replacer_value) => Table.ReplaceErrorValues(tbl, List.Transform(Table.ColumnNames(tbl), each {_, replacer_value}))
,
// ## 	Sort table
Table.SortColumnByList = Document(
	"Table.SortColumnByList",
	"Sort table's column text by List.",
	"Table.Sorting",
	{[Description="Sort table's column text by List.",
		Code="LPQ[Table.SortColumnByList](Table, ""Category"", ""{""File"",""Edit""}"", True)",
		Result="Sorted table Table in column 'Category' by List sort order Ascending." ]},
	(Table as table, ColumnName as text, ListSort as list, Order as logical) =>
		Table.RemoveColumns(
			Table.Sort(
				Table.AddIndexColumn(Table,"i",0,1,Int64.Type), {
					{each List.PositionOf(ListSort, Record.Field(_,ColumnName)), if Order then Order.Ascending else Order.Descending },
					{"i",Order.Ascending}   }
				),
			{"i"})),
// ## 	Running Totals
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
		Consolidate,
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


// #	ARRAYS
// ## 	VLOOKUP in other table
fnVlookUp = Document(
	"fnVlookUp",
	"LookUp or CheckUp value in other table (return value from other column or just search)",
	"Tables.Search",
	{[Description="Value, Table, Search at Column name, Return Column name.",
		Code="LPQ[fnVlookUp]([1], Products, ""ProductKey"", ""Price"")",
		Result="9000 OR 1" ]},
	(lookupValue as any, lookupTable as table, lookupColumnName as text, optional returnColumnValue as text) =>
		let colLookup = Table.Column(lookupTable, lookupColumnName),
			lookup = List.PositionOf(colLookup, lookupValue, 0),
			count = lookup >= 0
		in  if count then
				if returnColumnValue is null then 1
					else Table.Column(lookupTable, returnColumnValue){lookup}
				else null)
,
// ## 	Search Name in List
fnCheckInList = Document(
	"fnCheckInList",
	"Search word in the list. True - case not sesetive, False - case sensetive.",
	"Lists.Match",
	{[Description="Using separate lists. Params: Word, List, Case sensetive.",
		Code="LPQ[fnCheckInList](""Word"", {""Text"", ""Word"", ""Letter""}, true)",
		Result="TRUE"	 ]},
	(Name as text, List as list, optional Match as logical) =>
		let
			match = if (Match ?? {Match=""}{1}?) = null then true else Match,
			p = try Text.Trim(Text.BeforeDelimiter(Name,"(")) otherwise Name,
			pList = Text.Split(p," "),
			result = if match
				then List.ContainsAny(pList, List, (x,y) => Text.Contains(x,y,Comparer.OrdinalIgnoreCase))
				else List.ContainsAny(pList, List, (x,y) => x=y)
		in
			try result otherwise null)
,
Text.RegExp = (txt as text, regex as text, delim as text) =>
	let RegExp =
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
Text.RegexMatch = (txt as text, regex as text, delim as text) => try
	Web.Page("<script>var x='"&txt&"';var delim='"&delim&"';var regex=/"&regex&"/gi;var result=x.match(regex).join(delim);document.write(result);</script>")[Data]{0}[Children]{0}[Children]{1}[Text]{0}
	otherwise null
,
Text.GetMatch = (Text as text, List as list) =>
	try List.Select(Text.Split(Text," "),
			each List.ContainsAny(List.LastN(Text.Split(_," "),1), List, (x,y) => Text.Contains(x,y,Comparer.OrdinalIgnoreCase))
		){0}
	otherwise null
,


// #	TEXT
Text.Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz" & "АБВГДЕЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдежзийклмнопрстуфхцчшщъыьэюя",
Text.AlphaNumeric = Text.Alphabet & "0123456789",
Text.FromList = (list as list) => List.Accumulate(list, "", (state, current) => state & Text.From(current)),
Text.IsNumber = (text as text) => try Number.FromText(text) is number otherwise false
,
// ## 	Is text all uppercase? returns false if any non-alpha characters are present
Text.IsUpperCase = (text as text) => List.AllTrue(List.Transform(Text.ToList(text), (letter)=>Text.Contains(Text.Alphabet, letter) and letter = Text.Upper(letter))),
Text.IsAlpha = (text as text) => List.MatchesAll(Text.ToList(text), each Text.Contains(Text.Alphabet, _)),
Text.RemoveExtraWhitespace = (text as text) => Text.Combine(Splitter.SplitTextByWhitespace()(text)," ")
,
// ## 	Splits camelCased and PascalCased text and separates by a space. Ex: "thisIsAColumn" -> "this Is A Column"
Text.SplitCamelCase = (text as nullable text) => if text is null then null else List.Accumulate(Text.ToList(text),"", (state, current) =>
	let
		PreviousLetter = Text.End(state, 1),
		Ignore = (text as text) => text = " " or text = "."
	in state & (if not Text.IsUpperCase(PreviousLetter)
					and not Ignore(PreviousLetter)
					and not Ignore(current)
					and Text.IsUpperCase(current)
				then " " else "" ) &
		current),
Text.SplitOnNotIn = (line as nullable text, validCharacters as text) => Splitter.SplitTextByNotIn(validCharacters)(line),
Text.SplitOnNonAlpha = (line as nullable text) =>
	if line is null then null else List.Accumulate(Text.ToList(line), {null} , (state, current) =>
		let
			doSkip = not Text.Contains(Text.Alphabet, current),
			lastItem = List.Last(state),
			appendLast = lastItem<>null
		in
			if doSkip then
				if lastItem is null then state
				else List.Combine({state, {null}})
			else if appendLast
				then List.Combine({List.RemoveLastN(state, 1), {lastItem & current}})
				else List.Combine({List.RemoveLastN(state, 1), {current}})),
Text.Substring = (text as text, start as number, optional count as number) =>
   let
	  start = if start >= 0 then start else error "start index should be >= 0",
	  end = if count = null then Text.Length(text) else if count <= Text.Length(text) then count else error "count should be <= text length",
	  textList = Text.ToList(text),
	  substr = Text.FromList(List.FirstN(List.Skip(textList, start), end - start))
   in substr,
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
Text.AsciiOnly = (s as text) as text =>
	let
		Listified = Text.ToList(s),
		Numbered = List.Transform(Listified, each Character.ToNumber(_)),
		Filtered = List.Select(Numbered, each _ <= 255),
		Stringified = List.Transform(Filtered, each Character.FromNumber(_)),
		Joined = Text.Combine(Stringified, ""),
		Return = Joined
	in
		Return,

Text.Between = (Haystack as text, After as text, Before as text) as text =>
	let
		CutAfter = Text.Split(Haystack, After),
		CutBefore = Text.Split(CutAfter{1}, Before),
		Needle = if List.Count(CutAfter) > 1
		then (if List.Count(CutBefore) > 1 then CutBefore{0} else Error.Record("FindTextFailed","The text did not contain the keyword " & Before, Haystack))
		else error Error.Record("FindTextFailed","The text did not contain the keyword " & After, Haystack)
	in Needle,

Text.ContainsAny = (str as text, needles as list) as logical =>
	let count = List.Count(needles)
	in  List.AnyTrue(
		List.Generate(
			()=>[i=0],
			each [i] < count,
			each [i=[i]+1],
			each Text.Contains(str,needles{[i]}) )),

Text.Count = (Haystack as text, Needle as text) as number =>
	List.Count(Text.Split(Haystack, Needle)) - 1,

Text.EachBetween = (Haystack as text, After as text, Before as text) as list =>
	let
		CutAfter = Text.Split(Haystack, After),
		SkipFirst = List.Skip(CutAfter),
		CutEach = List.Transform(SkipFirst, each Text.Split(_, Before){0})
	in
		CutEach,

Text.EachFromTo = (Haystack as text, After as text, Before as text) as text =>
	let CutAfter = Text.Split(Haystack, After),
		SkipFirst = List.Skip(CutAfter),
		CutEach = List.Transform(SkipFirst, each After & Text.Split(_, Before){0} & Before)
	in  CutEach,

Text.FromTo = (Haystack as text, From as text, UpTo as text) as text =>
	let
		CutAfter = Text.Split(Haystack, From),
		CutBefore = Text.Split(CutAfter{1}, UpTo),
		Needle = if List.Count(CutAfter) > 1
		then (if List.Count(CutBefore) > 1 then From & CutBefore{0} & UpTo else Error.Record("FindTextFailed","The text did not contain the keyword " & UpTo, Haystack))
		else error Error.Record("FindTextFailed","The text did not contain the keyword " & From, Haystack)
	in Needle,

Text.RemoveSymbols = (inputtext as text) as text =>
	let
		//get a list of lists containing the numbers of Unicode punctuation characters
		numberlists = {{0..31},{33..47},{58..64},{91..96},{123..191}},
		//turn this into a single list
		combinedlist = List.Combine(numberlists),
		//get a list of all the punctuation characters that these numbers represent
		punctuationlist = List.Transform(combinedlist, each Character.FromNumber(_)),
		//some text to test this on
		//inputtext = "Hello! My name is Chris, and I'm hoping that this *cool* post will help you!",
		//the text with punctuation removed
		outputtext = Text.Remove(inputtext, punctuationlist)
	in
		outputtext,

Text.ReplaceAll = (str as text, Replacements as list) as text =>
	let count = List.Count(Replacements)
	in List.Last(
		List.Generate(
			()=>[i=0, s=str],
			each [i] <= count,
			each [
				s=Text.Replace([s],Replacements{[i]}{0},Replacements{[i]}{1}),
				i=[i]+1
			],
			each [s] )),
Text.PowerTrim = Document(
	"Text.PowerTrim",
	"Power Trim all text.",
	"Text.Transform",
	{[Description="Power Trim all text.",
		Code="LPQ[Text.PowerTrim](""	  A   B  C "")",
		Result="A B C" ]},
	(text as text, optional char_to_trim as text) =>
		let char = if char_to_trim = null then " " else char_to_trim,
			split = Text.Split(text, char),
			removeblanks = List.Select(split, each _ <> "")
		in  Text.Combine(removeblanks, char)),
Text.ReplaceByList = try (txt as text, replacementLists as list) as text => List.Accumulate(replacementLists, txt, (s,x) => Text.Replace(s,x{0},x{1})) otherwise null
,
//[ ]: add description
Text.ContainsList = 	
	(Text as text, List as list, optional ContainsMatch as logical) =>
	let
			match = ContainsMatch ?? {ContainsMatch=""}{1}? ?? false,
			product = try Text.Trim(Text.BeforeDelimiter(Text.Lower(Text),"(")) otherwise Text.Lower(Text),
			pList = Text.Split(product, " "),       
			compareList = List.Transform(List, Text.Lower),
			result =
					if match then
							List.ContainsAny(pList, compareList, (x,y) => Text.Contains(x, y, Comparer.OrdinalIgnoreCase))
					else
							List.ContainsAny(pList, compareList, (x,y) => x=y)
	in
			try result otherwise null
,

// #	CODE AND WEB
Web.StripHtmlList = Document(
	"Web.StripHtmlList",
	"Cleans all html tags in the column of the table.",
	"Web.Modification",
	{[Description="Cleans all html tags in the column of the table.",
		Code="LPQ[Web.StripHtmlList](table, ""column name with html"")",
		Result="https://inno3d.com/en/PRODUCT_INNO3D_GEFORCE_RTX_3070_ICHILL_X3_LHR" ]},
	(tbl as table, columnName as text) =>
		let
			t1 = Table.Column(tbl, columnName),
			l1 = List.Select(List.Transform(t1, each Text.Clean(Text.PowerTrim(_))), each _ <> ""),
			lBodyStart = List.PositionOf( l1, List.Select(l1, each Text.Contains(_, "<body", Comparer.OrdinalIgnoreCase)){0} ),
			lBodyEnd = List.PositionOf(l1, List.Select(l1, each Text.Contains(_, "</body", Comparer.OrdinalIgnoreCase)){0} ),
			L = List.Range(l1, lBodyStart, lBodyEnd - lBodyStart - 1)
		in Table.FromList(L, Splitter.SplitByNothing(), null, null, ExtraValues.Error)),
Web.RemoveHtmlTags = Document(
	"Web.RemoveHtmlTags",
	"Cleans all html tags in the text. With trim and clean of whitespace.",
	"Web.Modification", {[
		Description = "Cleans all html tags in the text. ALso cleans and trims the text.",
		Code="LPQ[Web.RemoveHtmlTags](""<div>my bit of text</div>"")",
		Result="my bit of text" ]},
	let func = (HTML) =>
		let Check = if Value.Is(Value.FromText(HTML), type text) then HTML else "",
				Source = Text.From(Check),
				SplitAny = Text.SplitAny(Source,"<>"),
				ListAlternate = List.Alternate(SplitAny,1,1,1),
				ListSelect = List.Select(ListAlternate, each _<>"")
		in 	Text.RemoveExtraWhitespace(Text.Combine(ListSelect, "")),
				documentation = [
					Documentation.Name= "Web.RemoveHtmlTags",
					Documentation.Description= "Remove Html Tags",
					Documentation.LongDescription= "Removes all Html tags from a text",
					Documentation.Category= "Web.Modification",
					Documentation.Source= "Inspired by a solution from Bill Szysz",
					Documentation.Author= "Imke Feldmann: www.TheBIccountant.com & Mike Carlo: PowerBI.Tips",
					Documentation.Examples= {[
						Description="Function that enables one to pass in a column that has HTML tags.",
						Code			 ="<div>my bit of text</div>",
						Result		 ="my bit of text"	]}]
	in Value.ReplaceType(func, Value.ReplaceMetadata(Value.Type(func), documentation))),
Web.TableToMarkdown = Document(
	"Web.TableToMarkdown",
	"Converts table to text table with Markdown syntax.",
	"Web.Modification",
	{[Description="Creates Markdown table.",
		Code="LPQ[Web.TableToMarkdown](table)",
		Result="| Опис та коментарь | TODO	|#(lf)| ----------------- | ------- |" ]},
	(Table as table) =>
		let
			TableValues = Table.AddColumn(Table, "Custom", each Text.Combine(Record.FieldValues(_), " | ")),
			HyphenLine = Text.Combine(List.Transform(Table.ColumnNames(Table), each Text.Repeat("-",Text.Length(_))), " | "),
			CombineList = List.Combine({{Text.Combine(Table.ColumnNames(Table), " | ")}, {HyphenLine}, TableValues[Custom]}),
			TransferToMarkdown = Text.Combine(CombineList, "#(lf)")
		in
			TransferToMarkdown)
,
Web.GetAllLinks = (SourceURL as text, optional AttributeDelimiter as text) =>
	let AttributeDelimiter = if AttributeDelimiter is null then """" else AttributeDelimiter,
	    //Get HTML source
	    Source = Text.FromBinary(Web.Contents(SourceURL)),
	    //Function to find each link
	    GetLink = (Counter as number) =>
	        let  CurrentLink = Text.BetweenDelimiters(
	                Source,
	                "href=" & AttributeDelimiter,
	                AttributeDelimiter,
	                Counter
	            )
	        in  if CurrentLink = "" then {}
	            else List.Combine({{CurrentLink}, @GetLink(Counter + 1)}),
	    //Call function
	    Output = GetLink(0)
	in  Output
,
Api.MakeApiCallTemplate = Document(
	"Api.MakeApiCallTemplate",
	"Converts table to text table with Markdown syntax.",
	"Web.Api",
	{[Description="See code and mofify for running"]},
	/*	// Define the main function and parameters
			myFunction = () =>
			let
					// Make the API call
					baseUrl = "https://api.worldbank.org/v2/",
					response = Json.Document(
							Web.Contents(baseUrl, [
									RelativePath = "country",
									Query = [
											format="json",
											per_page = "1000"
									],
									Headers = []
							])
					),
					response1 = response{1},
					#"Converted to Table" = Table.FromList(response1, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
					#"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", 
			{"id", "iso2Code", "name", "region", "adminregion", "incomeLevel", "lendingType", "capitalCity", "longitude", "latitude"}, 
			{"ISO 3 code", "ISO 2 code", "Country name", "Region", "Admin region", "Income level", "Lending type", "Capital city", "Longitude", "Latitude"})
			in
					#"Expanded Column1")
    */ (text)=>text)
,

// #	IMPORT PQ
// ## 	Minimal script loader
ImportPq = Document(
	"ImportPq",
	"Most minimal external PQ script file loader.",
	"Files.Import",
	{[Description="Load PQ file to Excel",
		Code="LPQ[ImportPq](""C:\File.pq"")",
		Result="code from .pq file" ]},
	(FilePath as text, optional Encoding as nullable number) =>
		try Expression.Evaluate(Text.FromBinary(File.Contents(FilePath),(Encoding ?? TextEncoding.Utf8)),#shared) otherwise null),
ImportPq.GitHubRaw = Document(
	"ImportPq.GitHubRaw",
	"Import file from GitHub.",
	"Web.Import",
	{[Description="Load file from GitHub",
		Code="LPQ[ImportPq.GitHubRaw](""https://github.com/laktiv/LPQ/blob/40a01a1998f1b88a03502e6f994625b108c8722b/LPQ.pq"")",
		Result="File from GitHub" ]},
	(URL as text) =>
		try Expression.Evaluate(Text.FromBinary(Web.Contents(URL)),#shared) otherwise null),
ImportPq.GitHubRepo = Document(
	"ImportPq.GitHubRepo",
	"Get list of files from GitHub repository.",
	"Web.Import",
	{[Description="Load file from GitHub",
		Code="LPQ[ImportPq.GitHubRepo](""https://github.com/laktiv/LPQ/blob/40a01a1998f1b88a03502e6f994625b108c8722b/LPQ.pq"")",
		Result="File list from GitHub repository." ]},
	(GitHubUser as text, GitHubRepo as text) =>
		let BaseURL = "https://api.github.com",
				params = [
						RelativePath = Text.Format(
								"/repos/#[user]/#[repo]/git/trees/main", [
										user = GitHubUser,
										repo = GitHubRepo
								]
						),
						ManualStatusHandling = {400, 404, 429, 500 }
				],
				rawText  = Text.FromBinary( response, TextEncoding.Utf8 ),
				response = Web.Contents( BaseURL, params ),
				Json     = Json.Document( response, TextEncoding.Utf8 ),
				TreeTble = Json[tree],
				Converted= Table.FromList(TreeTble, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
				Expanded = Table.ExpandRecordColumn(Converted, "Column1", {"path","mode","type","sha","url","size"})
		in  Expanded)
,


// #	VALUE
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
	else error "unknown -- not a primitive type!",

Value.ToText = (Val as any, optional RecursTypes as logical) as text =>
	let RecursTypes = if (RecursTypes<>null) then RecursTypes else false,
		Tried = (try Val),
		Value = if Tried[HasError] then Tried[Error] else Tried[Value],
		CaseValues = {
			{ (x)=> Value.Is(x, type type), Type.ToText(Value, RecursTypes) },
			{ (x)=> Value.Is(x, type function),
				let Type = Value.Type(Value),
					Params = Type.FunctionParameters(Type),
					Reqd = Type.FunctionRequiredParameters(Type),
					Ret = Type.FunctionReturn(Type)
				in 	"function (" &
					Record.TransformJoin(Params, (k,v) =>
						(if List.PositionOf(Record.FieldNames(Params), k) >= Reqd then "optional " else "") &
						k & " as " & @Value.ToText(v, RecursTypes)
					)
					& ") as " & @Value.ToText(Ret, RecursTypes)
			},
			{ (x)=> Value.Is(x, type table), "#table(" & @Value.ToText(Table.ColumnNames(Value), RecursTypes) & ", " & @Value.ToText(Table.ToRows(Value), RecursTypes) & ")"},
			{ (x)=> Value.Is(x, type record), "[" &
				Record.TransformJoin(Value, (k,v) => k & "=" & @Value.ToText(v, RecursTypes))
			& "]" },
			{ (x)=> Value.Is(x, type list), "{" & Text.Combine(List.Transform(Value, each @Value.ToText(_, RecursTypes)), ", ") & "}" },
			{ (x)=> x = null, "null" },
			{ (x)=> true, Expression.Constant(Value) }
		},
		Return = List.First(List.Select(CaseValues, each _{0}(Value))){1}
	in  Return,

Value.TypeToText = (Value as any, optional Recurs as logical) as text =>
	let
		Recurs = if (Recurs<>null) then Recurs else false,
		Type = Value.Type(Value),
		ToText = if Value.Is(Value, type type) and Recurs then
			"type " & Type.ToText(Value, Recurs)
		else
			Type.ToText(Type, Recurs),
		Return = ToText
	in Return,

Value.WaitFor = (producer as function, interval as function, optional count as number) as any =>
	let list = List.Generate(
			//start: first try, no result
			() => {0, null},
			//condition: stop if we have the result (try count null'd) or we've exceeded the max tries
			(state) => state{0} <> null and (count = null or state{0} < count),
			//next: stop try tally if we have our result, otherwise check again and tally a try
			(state) => if state{1} <> null
				then {null, state{1}}
				else {1 + state{0}, Function.InvokeAfter(() => producer(state{0}), interval(state{0}))},
			//transformer: only return the result, not try tally
			(state) => state{1})
	in  List.Last(list)
,


// #	OTHER
Type.ToText = (Type as any, optional Recurs as logical) as text =>
	let Recurs = if (Recurs<>null) then Recurs else false,
			CaseValues = {
				{ (x)=> (try x)[HasError], "error" },
				{ (x)=> Type.Is(x, type type), "type"},	//if Recurs then  else
				{ (x)=> Type.Is(x, type function), "function"},
				{ (x)=> Type.Is(x, type table), if Recurs then "table " & @Type.ToText(Type.TableRow(NonNull), Recurs) else "table"},
				{ (x)=> Type.Is(x, type record), if Recurs then
					let Record = Type.RecordFields(NonNull)
					in 	"[" & Record.TransformJoin(Record, (k,v) =>
						(if v[Optional] then "optional " else "") & Expression.Identifier(k) & " = " & @Type.ToText(v[Type], Recurs)
					) & "]"
				else "record"},
				{ (x)=> Type.Is(x, type list), if Recurs then "{" & @Type.ToText(Type.ListItem(NonNull), Recurs) & "}" else "list"},
				{ (x)=> Type.Is(x, type binary), "binary"},
				{ (x)=> Type.Is(x, type logical), "logical"},
				{ (x)=> Type.Is(x, type number), "number"},
				{ (x)=> Type.Is(x, type text), "text"},
				{ (x)=> Type.Is(x, type date), "date"},
				{ (x)=> Type.Is(x, type time), "time"},
				{ (x)=> Type.Is(x, type datetime), "datetime"},
				{ (x)=> Type.Is(x, type datetimezone), "datetimezone"},
				{ (x)=> Type.Is(x, type duration), "duration"},
				{ (x)=> Type.Is(type anynonnull, x), "anynonnull"},
				{ (x)=> Type.Is(type null, x), "null"},
				{ (x)=> Type.Is(None.Type, x), "none"},
				{ (x)=> true, "?"}
			},
			NonNull = Type.NonNullable(Type),
			Return = if Type.Is(type any, Type) then "any"
				else (if Type.IsNullable(Type) then "nullable " else "")
					& List.First(List.Select(CaseValues, each _{0}(NonNull))){1}
	in  Return,

Record.TransformJoin = (Rec as record, Lambda as function, optional Delimiter as text) as text =>
	let Delimiter = if (Delimiter<>null) then Delimiter else ", ",
		Keys = Record.FieldNames(Rec),
		Transformed = List.Transform(Keys, each Lambda(_, Record.Field(Rec,_))),
		Combined = Text.Combine(Transformed, Delimiter),
		Return = Combined
	in  Return


],

Result = _extensionLibrary

in Result

