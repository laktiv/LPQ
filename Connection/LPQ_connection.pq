let
    g_id = "700a6d65e098189881ecd77e585b233a",
    Gist = Json.Document(Web.Contents("https://api.github.com/gists/"& g_id ))[files][LPQ.pq][raw_url],

    SourceList = {
        Gist,
        "https://raw.githubusercontent.com/laktiv/LPQ/main/LPQ.pq",
        "C:\Users\Laktionov.I\OneDrive\INNO\PowerQuery\LPQ.pq" },
    LPQ = Expression.Evaluate(Text.FromBinary(Web.Contents(SourceList{ 0 })), #shared)
in
    LPQ
