dbMemo "SQL" ="SELECT DISTINCT q.ID AS Quadrat_ID, q.Transect_ID, q.Quadrat, q.IsSampled, q.NoE"
    "xotics, sc.ID AS SpeciesCover_ID, sc.PlantCode, sc.IsDead, sc.PercentCover, \"Q\""
    " & q.Quadrat & \015\012IIF(IsNull([Position_m]),\"\",\"_\")  & esp.Position_m & "
    "\015\012IIF(IsNull([Position_m]),\"\",\"m\") AS Quad_Pos, esp.Position_m, NumSam"
    "pledQuads, \"SpeciesCoverID_Q\" & q.Quadrat AS Quad_CoverID\015\012FROM (((Speci"
    "esCover AS sc INNER JOIN Quadrat AS q ON q.ID = sc.Quadrat_ID) INNER JOIN Transe"
    "ct AS t ON t.Transect_ID = q.Transect_ID) INNER JOIN SampledQuads AS sq ON sq.Tr"
    "ansect_ID = q.Transect_ID) INNER JOIN EventSamplePosition AS esp ON (esp.Quadrat"
    " = q.Quadrat) AND (esp.Event_ID = t.Event_ID);\015\012"
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
    Begin
        dbText "Name" ="SpeciesCover_ID"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="esp.Position_m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quad_CoverID"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
