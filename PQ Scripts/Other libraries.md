## Other libraries

My favorites:

PBI**
```
let
    tnclark8012link = "https://raw.githubusercontent.com/tnclark8012/Power-BI-Desktop-Query-Extensions/master/power-query-extensions.pq",
    tnclark8012 = Expression.Evaluate(Text.FromBinary(Web.Contents(tnclark8012link)),#shared)
in
    tnclark8012
```