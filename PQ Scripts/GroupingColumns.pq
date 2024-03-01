/*
GroupingColumns list of text

Options record of any:

    HashFunctions Record of Functions
        - Value should be a function that returns a number from 0 to n - 1 where n is the number of unique values in GroupingColumn
        - Key should be GroupingColumn name to apply the function to
        - Default is grouping column values sorted a-z as text
        - Note that a custom function may be required for a grouping column with numerical values to sort properly since the values are sorted as text

    SummaryFunctions Record of Functions
        - Optional: will add a row under each section displaying a summary value based on the function specified for a particular column.
        - Example: show subtotals for Column A
        SummaryFunctions = [#"Column A" = List.Sum]
        - Example: show average for Column B and count of Column C
        SummaryFunctions = [#"Column B" = List.Average, #"Column C" = List.Count]

    HeadingStyle Text
        "Outline" (Default): 
            - Places group headings in additional columns to the left. For example, with two grouping columns, Outline level 1 headings will be in column A, Outline level 2 headings will be in column B, and the source data will be in columns C+.
        "Inline": 
            - Places group headings in the first column of the source data.

    MaxLevelBlankRows Number
        - Maximum outline level for adding blank rows
        - Default is all outline levels

    BlankGroupName Text
        - Title to use for category rows that are blank (null)
        - Default is "Blank"

    ShowStyleIndex Logical
        - The Style Index is an additional leftmost column with values that can be used for conditional formatting. This column can be hidden or omitted from the print area.
        true (Default): StyleIndex column shown
        false: StyleIndex column not shown

    SummaryFunctionLabels Record or Logical
        - Adds a description row for summary functions (if specified)
        true (Default): uses "Name" from Value.Metadata, or internal name for pre-defined functions such as List.Sum
        false: no description row is shown
        Record: custom labels, e.g. [#"Column A" = "Total"]

*/
(InputTable as table, GroupingColumns as list, optional Options as record, optional OutlineLevel as number) as table =>
    let
        // Set up key variables; create and count a Sorted Set of Categories
        OutlineLevel = OutlineLevel ?? 1,
        InputTable = if (OutlineLevel = 1) then Table.AddIndexColumn(InputTable, "_ID", 1) else InputTable,
        NumCategoryCols = List.Count(GroupingColumns),
        GroupingColumn = GroupingColumns{OutlineLevel - 1},
        CategorySet = List.Sort(List.Distinct(Table.Column(InputTable, GroupingColumn))),
        NumGroups = List.Count(CategorySet),
        // Options
        Options = Options ?? [],
        HashFunctions = Record.FieldOrDefault(Options, "HashFunctions", null),
        SummaryFunctions = Record.FieldOrDefault(Options, "SummaryFunctions", []),
        HeadingStyle = Record.FieldOrDefault(Options, "HeadingStyle", "Outline"),
        MaxLevelBlankRows = Record.FieldOrDefault(Options, "MaxLevelBlankRows", NumCategoryCols),
        BlankGroupName = Record.FieldOrDefault(Options, "BlankGroupName", "Blank"),
        ShowStyleIndex = Record.FieldOrDefault(Options, "ShowStyleIndex", true),
        SummaryFunctionLabels = Record.FieldOrDefault(
            Options, "SummaryFunctionLabels", if (SummaryFunctions <> []) then true else false
        ),
        // Hash Function (Default is sorted a-z ascending)
        CategorySetText = List.Transform(CategorySet, each Text.From(_)),
        MatchTable = Record.FromList(
            List.Numbers(0, NumGroups),
            List.ReplaceValue(CategorySetText, null, BlankGroupName, Replacer.ReplaceValue)
        ),
        DefaultHashFunction = (key) => Record.Field(MatchTable, Text.From(key ?? BlankGroupName)),
        HashFunction = Record.FieldOrDefault(HashFunctions, GroupingColumn, DefaultHashFunction),
        // Heading title function
        FormatHeading = (xText) => Text.Upper(Text.From(xText ?? BlankGroupName)),
        // Create empty record ingredients
        FirstRecord = InputTable{0},
        EmptyValues = List.Repeat({null}, Record.FieldCount(FirstRecord)),
        EmptyKeys = Record.FieldNames(FirstRecord),
        EmptyRecord = Record.FromList(EmptyValues, EmptyKeys),
        // Set up heading row function for groups
        HeadingCols = Record.FromList(
            List.Repeat({null}, NumCategoryCols), List.Transform(GroupingColumns, each "H_" & _)
        ),
        HeadingName = "H_" & GroupingColumn,
        GroupHeaders = (xTable, FirstID) =>
            Table.Combine(
                {
                    // add a blank row
                    let
                        AddHeadingCols = if (HeadingStyle = "Outline") then HeadingCols else [],
                        AddBlank =
                            (OutlineLevel <= MaxLevelBlankRows)
                            and (Record.Field(xTable{0}, "_ID") <> FirstID)
                    in
                        if (AddBlank) then
                            Table.FromRecords(
                                {Record.Combine({[StyleIndex = "blank"], AddHeadingCols, EmptyRecord})}
                            )
                        else
                            #table({}, {}),
                    // add heading row
                    let
                        // Second record replaces value from <<HeadingName>> field of the first record
                        HeadingRowOutline = Table.FromRecords(
                            {
                                Record.Combine(
                                    {
                                        [
                                            StyleIndex = "Heading" & Text.From(OutlineLevel)
                                        ],
                                        HeadingCols,
                                        Record.FromList(
                                            {FormatHeading(Record.Field(xTable{0}, GroupingColumn))}, {HeadingName}
                                        ),
                                        EmptyRecord
                                    }
                                )
                            }
                        ),
                        // Second record replaces value from first field of the first record
                        HeadingRowInline = Table.FromRecords(
                            {
                                Record.Combine(
                                    {
                                        [
                                            StyleIndex = "Heading" & Text.From(OutlineLevel)
                                        ],
                                        EmptyRecord,
                                        Record.FromList(
                                            {FormatHeading(Record.Field(xTable{0}, GroupingColumn))}, {EmptyKeys{0}}
                                        )
                                    }
                                )
                            }
                        )
                    in
                        if (HeadingStyle = "Outline") then
                            HeadingRowOutline
                        else
                            HeadingRowInline
                }
            ),
        // Set up summary row function for groups
        // For the Outline heading style, heading columns will be automatically added from the rows appended above
        SummaryRow = (xTable) =>
            let
                SummaryRowVals = Table.AggregateTableColumn(
                    #table({"t"}, {{xTable}}),
                    "t",
                    List.Transform(
                        Table.ColumnNames(xTable),
                        each {_, Record.FieldOrDefault(SummaryFunctions, _, (x) => null), _}
                    )
                )
            in
                Table.AddColumn(SummaryRowVals, "StyleIndex", each "Summary" & Text.From(OutlineLevel)),
        // Summary Function Labels
        // Get ResourceExpression Name or Metadata Name
        GetFunctionLabel = (xFunction as function) =>
            let
                InternalName = Record.FieldOrDefault(Value.ResourceExpression(xFunction), "Name", null),
                MeaningfulInternalName =
                    if (Text.StartsWith(InternalName, "List.")) then
                        Text.AfterDelimiter(InternalName, ".")
                    else
                        InternalName,
                NameSelection = List.RemoveNulls(
                    {
                        Record.FieldOrDefault(Value.Metadata(xFunction), "Name", null),
                        MeaningfulInternalName,
                        "Custom"
                        // failsafe
                    }
                )
            in
                NameSelection{0},
        LabeledColumns =
            if (SummaryFunctionLabels = true) then
                Record.FromList(
                    List.Transform(Record.ToList(SummaryFunctions), each GetFunctionLabel(_)),
                    Record.FieldNames(SummaryFunctions)
                )
            else if (SummaryFunctionLabels is record) then
                SummaryFunctionLabels
            else
                [],
        SummaryFunctionRow = Record.Combine(
            {
                [
                    StyleIndex = "SummaryFunctionLabels"
                ],
                if (HeadingStyle = "Outline") then
                    HeadingCols
                else
                    [],
                EmptyRecord,
                LabeledColumns
            }
        ),
        // Check if last group
        LastGroup = OutlineLevel = NumCategoryCols,
        // Partition and combine with headers
        #"Partitioned Table" = Table.Partition(InputTable, GroupingColumn, NumGroups, HashFunction),
        FirstID = #"Partitioned Table"{0}{0}[_ID],
        #"Added Headers" = List.Transform(
            #"Partitioned Table",
            each
                Table.Combine(
                    {
                        GroupHeaders(_, FirstID),
                        if (LastGroup) then
                            _
                        else
                            @GroupByCategory(_, GroupingColumns, Options, OutlineLevel + 1),
                        if (SummaryFunctions <> []) then
                            SummaryRow(_)
                        else
                            #table({}, {})
                    }
                )
        ),
        #"Recombined Tables" = Table.Combine(#"Added Headers"),
        #"Removed Internal Columns" =
            if (ShowStyleIndex = false) then
                Table.RemoveColumns(#"Recombined Tables", {"StyleIndex", "_ID"})
            else
                Table.RemoveColumns(#"Recombined Tables", "_ID"),
        #"Added Summary Function Label Row" =
            if (SummaryFunctionLabels <> false and OutlineLevel = 1) then
                Table.InsertRows(#"Removed Internal Columns", 0, {SummaryFunctionRow})
            else
                #"Removed Internal Columns"
    in
        #"Added Summary Function Label Row"
