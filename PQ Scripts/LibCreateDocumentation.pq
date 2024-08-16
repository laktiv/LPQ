let
    Table = Table.SelectRows(
        Table.ReplaceErrorValues(Record.ToTable(LPQ), {{"Value",null}}),
        each [Name] <> "Document" and [Value] <> null
    ),
    Category = Table.AddColumn(Table, "Domain", each Text.Split(Text.Replace([Name], "_", "."), "."){0}),
    Sort = Table.Sort(
        Category,
        {
            {"Domain", Order.Ascending},
            {each LPQ[Value.TypeText]([Value]), Order.Ascending},
            {"Name", Order.Ascending}
        }
    ),
    #"meta" = Table.AddColumn(Sort, "meta", each if Value.Type([Value]) is text then null else Value.Metadata(Value.Type([Value]))),
    Unfold = Table.ExpandRecordColumn(#"meta", "meta",
        {
            "Documentation.Name",
            "Documentation.Description",
            "Documentation.LongDescription",
            "Documentation.Examples",
            "Documentation.Category"
        },
        {"meta.Name", "meta.Description", "meta.LongDescription", "meta.Examples", "meta.Category"}
    ),
    SelectColumns = Table.SelectColumns(Unfold,{"meta.Category", "Domain", "Name", "meta.Name", "meta.Description", "meta.LongDescription", "meta.Examples"})
in
    SelectColumns