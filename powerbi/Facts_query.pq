// Power Query (M) for the "Facts" table: combines tblBIFacts from every monthly result file in the SharePoint folder.
// In Power BI Desktop: Home > Transform data > New source > Blank query > Advanced editor, paste this, then change
// SiteUrl and FolderName. Name the query "Facts".
// If a month was saved twice, only the newest file of that month is kept, so no month is ever counted twice.
let
    SiteUrl = "https://YOURTENANT.sharepoint.com/sites/YOURSITE",
    FolderName = "02 Monthly Results",
    Files = SharePoint.Files(SiteUrl, [ApiVersion = 15]),
    InFolder = Table.SelectRows(Files, each Text.Contains([Folder Path], "/" & FolderName & "/")
        and Text.Lower([Extension]) = ".xlsx" and not Text.StartsWith([Name], "~")),
    // tblBIFacts of each file (a file without it is skipped)
    WithFacts = Table.AddColumn(InFolder, "Facts", each try Excel.Workbook([Content], true){[Item = "tblBIFacts", Kind = "Table"]}[Data] otherwise null),
    Usable = Table.SelectRows(WithFacts, each [Facts] <> null),
    Slim = Table.SelectColumns(Usable, {"Facts", "Date modified"}),
    Expanded = Table.ExpandTableColumn(Slim, "Facts", {"Month", "Month Date", "Section", "Category", "Item", "Metric", "Value"}),
    Typed = Table.TransformColumnTypes(Expanded, {{"Month", type text}, {"Month Date", type date}, {"Section", type text},
        {"Category", type text}, {"Item", type text}, {"Metric", type text}, {"Value", type number}, {"Date modified", type datetime}}),
    // Newest file per month only
    Newest = Table.Group(Typed, {"Month"}, {{"Newest", each List.Max([Date modified]), type datetime}}),
    Joined = Table.NestedJoin(Typed, {"Month"}, Newest, {"Month"}, "N", JoinKind.Inner),
    WithNewest = Table.ExpandTableColumn(Joined, "N", {"Newest"}),
    Latest = Table.SelectRows(WithNewest, each [Date modified] = [Newest]),
    Result = Table.RemoveColumns(Latest, {"Date modified", "Newest"})
in
    Result
