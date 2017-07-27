dbMemo "SQL" ="PARAMETERS scid Long, qid Long, plant Text ( 50 ), dead Long, pct IEEEDouble;\015"
    "\012UPDATE SpeciesCover SET PlantCode = [plant], IsDead = [dead], PercentCover ="
    " [pct]\015\012WHERE ID = [scid]\015\012AND\015\012Quadrat_ID = [qid];\015\012"
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
