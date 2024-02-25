let
    fGist = (gist_id as text, optional load_text as logical) =>
        Json.Document(Web.Contents("https://api.github.com/gists/" & gist_id))[files][LPQ.pq][raw_url],
    gist_url = fGist("700a6d65e098189881ecd77e585b233a", false),

    SourceList = {
        gist_url,
        "https://raw.githubusercontent.com/laktiv/LPQ/main/LPQ.pq",
        "C:\Users\Laktionov.I\OneDrive\INNO\PowerQuery\LPQ.pq" },
    i = 0,

    LPQ = Expression.Evaluate(Text.FromBinary(Web.Contents(SourceList{i})), #shared)
in
    LPQ
