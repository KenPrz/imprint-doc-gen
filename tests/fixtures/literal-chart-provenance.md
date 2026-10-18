# Literal category chart fixture

Authored for this repository on 2026-10-18 using openpyxl (no third-party document content).
The workbook has one empty sheet and a bar chart with literal category and numeric series data.
It intentionally has no chart formula references or embedded workbook relationship.
Independent LibreOffice rendering verifies that the chart remains visible after updating its data.

Construction: Workbook(); BarChart(); Series(idx=0, order=0, tx=SeriesLabel(v='Original'),
val=NumDataSource(numLit=NumData(ptCount=2, pt=[NumVal(idx=0,v=1), NumVal(idx=1,v=2)])),
cat=AxDataSource(strLit=StrData(ptCount=2, pt=[StrVal(idx=0,v='Old A'), StrVal(idx=1,v='Old B')]))).
Chart title: Pilot revenue. Axis titles: Amount / Month. Anchor A1 on sheet Chart.
