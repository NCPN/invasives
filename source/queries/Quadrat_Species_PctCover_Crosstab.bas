dbMemo "SQL" ="TRANSFORM Min(qs.PercentCover) AS MinOfPercentCover\015\012SELECT qs.Transect_ID"
    ", qs.PlantCode, qs.IsDead, qs.NumSampledQuads, Sum(qs.PercentCover) AS SumOfPerc"
    "entCover, SumOfPercentCover/(IIf(NumSampledQuads=0,1,NumSampledQuads)) AS AvgCov"
    "er\015\012FROM Quadrat_Species_Position AS qs\015\012GROUP BY qs.Transect_ID, qs"
    ".PlantCode, qs.IsDead, qs.NumSampledQuads\015\012PIVOT qs.Quad_Pos;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "OrderByOn" ="0"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="[Transect_ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[PlantCode]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[IsDead]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="<>"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species.[Transect_ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species.Quadrat"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species.[PlantCode]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species.[IsDead]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PIVOT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species.[IsSampled]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species.NoExotics"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species.Position_m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quad_Pos"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="1_"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="1_0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="1_3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2_"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2_5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2_8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="3_"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="3_10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="3_13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species.IsSampled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species.PlantCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q2_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q1_0m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q2_8m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species.Transect_ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3420"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Quadrat_Species.Quad_Pos"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q1_3m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q3_10m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q3_13m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species.NumSampledQuads"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgCover"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species.PercentCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPercentCover"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qs.SpeciesCover_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q3_9_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qs.Transect_ID"
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
        dbText "Name" ="Q2_4_5m"
        dbLong "AggregateType" ="-1"
    End
End
