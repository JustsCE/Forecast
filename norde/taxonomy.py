"""Taxonomy — the fixed vocabulary of the NORDE forecast and the one DAX that feeds it.

line     P&L leaf line, by code; statement order = code order; totals are derived
entity   dim Strukturvienibas[ACCOUNT_NR]; each in one division; active = carries P&L postings
month    first day of the month (the Sheets write tools turn "2026-09" into that date)
amount   EUR, ledger sign, the [amount] measure with the year-end close (Kontu reformācija) excluded
"""

LINES = [  # (code, name, kind)
    ("1.1", "Auto pārdošanas ieņēmumi", "revenue"),
    ("1.2", "Lietotu auto pārdošanas ieņēmumi", "revenue"),
    ("1.3", "Rezerves daļu un aksesuāru pārdošanas ieņēmumi", "revenue"),
    ("1.4", "Ieņēmumi no garantijas remonta, detaļas", "revenue"),
    ("1.5", "Ieņēmumi no garantijas remonta, darbs", "revenue"),
    ("1.6", "Ieņēmumi no servisa pakalpojuma", "revenue"),
    ("1.7", "Ieņēmumi no automašīnu nomas", "revenue"),
    ("1.8", "Citi ieņēmumi", "revenue"),
    ("2.1", "Auto iegādes izdevumi", "expense"),
    ("2.2", "Materiālu izdevumi", "expense"),
    ("2.3", "Algu izdevumi un tiem pielīdzināmie", "expense"),
    ("2.4", "Ēkas izdevumi, komunālie maksājumi u.tml.", "expense"),
    ("2.5", "Transporta izdevumi", "expense"),
    ("2.6", "Pagarinātās garantijas", "expense"),
    ("2.7", "Reklāmas izdevumi", "expense"),
    ("2.8", "Citi izdevumi bez finanšu izdevumiem", "expense"),
    ("2.9", "Bankas komisijas", "expense"),
    ("2.91", "Līzinga procenti, operatīvais līzings", "expense"),
    ("2.92", "Krājuma finansējums, procenti", "expense"),
    ("2.93", "Pamatlīdzekļu nolietojums, norakstījumi", "expense"),
]
TOTALS = [("1.99999", "KOPĀ IEŅĒMUMI"), ("2.99999", "KOPĀ IZDEVUMI"), ("3.99999", "PEĻŅA")]

DIVISIONS = ["Pārdošana", "Serviss", "Virsbūves", "Auto noma", "Administrācija", "Kopējie"]

ENTITIES = [  # (code, name, division, active)
    ("02", "Pārdošana", "Pārdošana", True),
    ("02.1", "Pārdošanas jaunas auto", "Pārdošana", True),
    ("02.2", "Pārdošanas lietoti auto", "Pārdošana", True),
    ("02.3", "Pirmspārdošanas serviss", "Pārdošana", True),
    ("02.4", "Lietotie auto", "Pārdošana", False),
    ("AS", "Autosalons", "Pārdošana", False),
    ("AS2", "Autosalons2", "Pārdošana", False),
    ("LA", "Lietotie auto", "Pārdošana", False),
    ("NCS_R - Jaunu a/m tirdzniecība", "Pārdošanas jaunas auto", "Pārdošana", False),
    ("03", "Pēcpārdošanas serviss", "Serviss", True),
    ("03.2", "Serviss", "Serviss", True),
    ("03.2.1", "Rezerves daļas", "Serviss", True),
    ("03.2.2", "Garantijas apkalpošana", "Serviss", True),
    ("PRT_R - Rezerves daļas", "Serviss", "Serviss", False),
    ("S", "Serviss", "Serviss", False),
    ("SRV_R - Serviss", "Serviss", "Serviss", False),
    ("03.1", "Virsbūves serviss", "Virsbūves", True),
    ("BOD_R - Virsbūvju serviss", "Virsbūves serviss", "Virsbūves", False),
    ("KR", "Auto krāsošana", "Virsbūves", False),
    ("06", "Auto noma", "Auto noma", True),
    ("01", "Administrācija", "Administrācija", True),
    ("A", "Administrācija", "Administrācija", False),
    ("04", "Saimniecība", "Kopējie", True),
    ("05", "Kopējie", "Kopējie", True),
    ("07", "Auto mazgātuve", "Kopējie", True),
    ("V", "Neattiecināts", "Kopējie", True),
]

DAX = """EVALUATE
SELECTCOLUMNS(
    SUMMARIZECOLUMNS(
        'HOR Financials'[PNL_LINE],
        'dim Strukturvienibas'[ACCOUNT_NR],
        'dim Datumi'[yyyy-mm],
        FILTER ( ALL ( 'dim Datumi'[yyyy-mm] ), 'dim Datumi'[yyyy-mm] >= "@since" ),
        FILTER ( ALL ( 'HOR Financials'[PNL_LINE_Name] ), 'HOR Financials'[PNL_LINE_Name] <> "na" ),
        FILTER ( ALL ( 'HOR Financials'[PNL_LINE] ), NOT CONTAINSSTRING ( 'HOR Financials'[PNL_LINE], "99999" ) ),
        FILTER ( ALL ( 'HOR Financials'[transaction_type] ), 'HOR Financials'[transaction_type] <> "Kontu reformācija" ),
        "a", [amount]
    ),
    "l", LEFT ( 'HOR Financials'[PNL_LINE], FIND ( " ", 'HOR Financials'[PNL_LINE] ) - 1 ),
    "e", 'dim Strukturvienibas'[ACCOUNT_NR],
    "m", 'dim Datumi'[yyyy-mm],
    "a", ROUND ( [a], 0 )
)"""


def dax(since):
    """The query to run with powerbi_run_dax; `since` = first month to pull, 'YYYY-MM'. Paste its rows as [l, e, m, a]."""
    return DAX.replace("@since", since)
