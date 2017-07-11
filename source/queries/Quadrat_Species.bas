dbMemo "SQL" ="SELECT q.ID, q.Transect_ID, q.Quadrat, q.IsSampled, q.NoExotics, sc.*, \"Q\" & q"
    ".Quadrat & IIF(IsNull([Position_m]),\"\",\"_\")  & sc.Position_m & IIF(IsNull([P"
    "osition_m]),\"\",\"m\") AS Quad_Pos, NumSampledQuads\015\012FROM (Quadrat AS q I"
    "NNER JOIN SpeciesCover AS sc ON sc.Quadrat_ID = q.ID) INNER JOIN SampledQuads AS"
    " sq ON sq.Transect_ID = q.Transect_ID;\015\012"
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
        dbText "Name" ="q.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q.Quadrat"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q.IsSampled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q.NoExotics"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Quadrat_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.PlantCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.PercentCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quad_Pos"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NumSampledQuads"
        dbLong "AggregateType" ="-1"
    End
End
