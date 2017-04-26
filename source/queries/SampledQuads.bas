dbMemo "SQL" ="SELECT Transect_ID, SUM(IsSampled) AS NumSampledQuads\015\012FROM Quadrat\015\012"
    "GROUP BY Transect_ID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="NumSampledQuads"
        dbInteger "ColumnWidth" ="2340"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_ID"
        dbLong "AggregateType" ="-1"
    End
End
