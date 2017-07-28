dbMemo "SQL" ="SELECT sc.*\015\012FROM (SurfaceCover AS sc INNER JOIN Quadrat AS q ON q.ID = sc"
    ".Quadrat_ID) INNER JOIN Transect AS t ON t.Transect_ID = q.Transect_ID\015\012WH"
    "ERE sc.PercentCover = 0\015\012AND\015\012q.NoExotics =1;\015\012"
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
        dbText "Name" ="NumRecords"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PlantCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1005"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SpeciesCoverID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SpeciesCover_ID"
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
        dbText "Name" ="sc.Surface_ID"
        dbLong "AggregateType" ="-1"
    End
End
