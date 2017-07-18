dbMemo "SQL" ="PARAMETERS qid Long, sid Long, pct IEEESingle;\015\012UPDATE SurfaceCover SET Pe"
    "rcentCover = [pct]\015\012WHERE Quadrat_ID = [qid] AND Surface_ID = [sid];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbText "Filter" ="[Unit_Code]='CEBR' AND [Plot_ID]=133"
Begin
End
