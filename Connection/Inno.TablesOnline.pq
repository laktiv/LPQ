[
	VAT  = #table({"VAT","k"},{{12,0.928575},{18,0.847458},{19,0.84035},{20,0.83334},{21,0.826510201},{23,0.81301},{27,0.7874016}}),
	CODE = #table({"Code","VAT"},{{"BY",20},{"GE",18},{"HK",0},{"KZ",12},{"RU",20},{"UA",20},{"UZ",12},{"BG",20},{"CY",19},{"CZ",21},{"EE",20},{"HR",27},{"LV",21},{"PL",23},{"RO",19},{"RS",20},{"SK",20},{"TR",18},{"US",0}}),
	MARG = #table({"Marg","k"},{{10,0.9091},{9,0.9175},{8,0.92601},{7,0.9346},{6,0.9435},{5,0.95245},{4,0.9616},{3,0.9709},{2,0.9805}}),
	CUR =
		let
			fpln = (t) => Text.BetweenDelimiters(Web.BrowserContents("https://nbp.pl/statystyka-i-sprawozdawczosc/kursy/tabela-a/"), t&"</td><td>","</td>", 0),
			rec = [
				czk = Number.FromText(Text.BetweenDelimiters(Web.BrowserContents("https://www.cnb.cz/en/financial-markets/foreign-exchange-market/central-bank-exchange-rate-fixing/central-bank-exchange-rate-fixing/daily.txt"),"USD|","<",0),"en-US"),
				eur = Number.From(Text.BeforeDelimiter(Xml.Tables(Web.Contents("https://www.ecb.europa.eu/rss/fxref-usd.html")){0}[Table]{1}[Table]{0}[title]{0}[#"Element:Text"]," "),"en-US"),
				ron = Number.FromText(Table.SelectRows(Xml.Tables(Web.Contents("https://www.bnr.ro/nbrfxrates.xml")){1}[Table]{0}[Cube]{0}[Rate],(t)=>t[#"Attribute:currency"]="USD"){0}[#"Element:Text"],"en-US"),
				pln = fpln("1 USD"),
				pln_eur = fpln("1 EUR") ]
		in
			Table.RenameColumns(Table.TransformColumnTypes(Record.ToTable(rec),{{"Value", type number}}),{{"Name","Curr"},{"Value","Rate"}})
]