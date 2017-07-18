dbMemo "SQL" ="TRANSFORM Min(qs.SpeciesCover_ID) AS MinOfSpeciesCover_ID\015\012SELECT qs.Trans"
    "ect_ID, qs.PlantCode, qs.IsDead\015\012FROM Quadrat_Species_Position AS qs\015\012"
    "GROUP BY qs.Transect_ID, qs.PlantCode, qs.IsDead\015\012PIVOT qs.Quad_CoverID;\015"
    "\012"
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
        dbText "Name" ="qs.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qs.SpeciesCover_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qs.PlantCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qs.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qs.NumSampledQuads"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPercentCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SpeciesCoverID_Q1"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SpeciesCoverID_Q2"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SpeciesCoverID_Q3"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
