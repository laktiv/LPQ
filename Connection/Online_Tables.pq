[
MARG=#table({"Marg","k"},{{10,0.9091},{9,0.9175},{8,0.92601},{7,0.9346},{6,0.9435},{5,0.95245},{4,0.9616},{3,0.9709},{2,0.9805}}),
VAT =#table({"Country","Code","VAT","k"},{{"Bulgaria","BG","20","0,83334"},{"Cyprus","CY","19","0,84035"},{"Czech","CZ","21","0,826510201"},{"Estonia","EE","20","0,83334"},{"Hungary","HR","27","0,7874016"},{"Lithuania","LV","21","0,826510201"},{"Poland","PL","23","0,81301"},{"Romania","RO","19","0,84035"},{"Serbia","RS","20","0,83334"},{"Slovakia","SK","20","0,83334"},{"Turkey","TR","18","0,847458"},{"Belarus","BY","20","0,83334"},{"Georgia","GE","18","0,847458"},{"Kazakhstan","KZ","12","0,928575"},{"Ukraine","UA","20","0,83334"},{"Uzbekistan","UZ","12","0,928575"}}),
CUR = 
let fpln =(t)=> Text.BetweenDelimiters(Web.BrowserContents("https://nbp.pl/statystyka-i-sprawozdawczosc/kursy/tabela-a/"),t&"</td><td>","</td>",0),
rec=[czk =Number.FromText(Text.BetweenDelimiters(Web.BrowserContents("https://www.cnb.cz/en/financial-markets/foreign-exchange-market/central-bank-exchange-rate-fixing/central-bank-exchange-rate-fixing/daily.txt"),"USD|","<",0),"en-US"),
		 eur =Number.From(Text.BeforeDelimiter(Xml.Tables(Web.Contents("https://www.ecb.europa.eu/rss/fxref-usd.html")){0}[Table]{1}[Table]{0}[title]{0}[#"Element:Text"]," "),"en-US"),
		 ron =Number.FromText(Table.SelectRows(Xml.Tables(Web.Contents("https://www.bnr.ro/nbrfxrates.xml")){1}[Table]{0}[Cube]{0}[Rate],(t)=>t[#"Attribute:currency"]="USD"){0}[#"Element:Text"],"en-US"),
		 pln =fpln("1 USD"), pln_eur =fpln("1 EUR")]
	in Table.RenameColumns(Table.TransformColumnTypes(Record.ToTable(rec),{{"Value",type number}}),{{"Name","Curr"},{"Value","Rate"}})
]