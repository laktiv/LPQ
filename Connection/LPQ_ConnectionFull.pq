let
	g_id = "700a6d65e098189881ecd77e585b233a",
	Gist = Json.Document(Web.Contents("https://api.github.com/gists/"& g_id))[files][LPQ.pq][raw_url],
	SourceList = { Gist,
        "https://raw.githubusercontent.com/laktiv/LPQ/main/LPQ.pq",
        "C:\Users\Laktionov.I\OneDrive\INNO\PowerQuery\LPQ.pq" }
in
	Expression.Evaluate(Text.FromBinary(Web.Contents(SourceList{ 0 })), #shared)

/*
let lpq = Text.FromBinary(Web.Contents(luri))
in Expression.Evaluate(lpq, #shared)
*/

/* lpq_uri
"https://gist.githubusercontent.com/laktiv/700a6d65e098189881ecd77e585b233a/raw/d6dbd2dcececdfec69e0a1514b7da0a84233efdf/LPQ.pq" meta [
    IsParameterQuery=true, List={
        "https://gist.githubusercontent.com/laktiv/700a6d65e098189881ecd77e585b233a/raw/d6dbd2dcececdfec69e0a1514b7da0a84233efdf/LPQ.pq",
        "https://raw.githubusercontent.com/laktiv/LPQ/main/LPQ.pq",
        "C:\Users\Laktionov.I\OneDrive\INNO\PowerQuery\LPQ.pq"
    },
    DefaultValue="https://gist.githubusercontent.com/laktiv/700a6d65e098189881ecd77e585b233a/raw/d6dbd2dcececdfec69e0a1514b7da0a84233efdf/LPQ.pq",
    Type="Any",
    IsParameterQueryRequired=false ]
*/
//Expression.Evaluate(Text.FromBinary(Web.Contents(lpq_uri)),#shared)
