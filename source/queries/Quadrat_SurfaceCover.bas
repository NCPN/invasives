dbMemo "SQL" ="SELECT sc.ID AS SurfaceCover_ID, sc.Quadrat_ID, sc.Surface_ID, sc.PercentCover, "
    "s.Surface, s.Description, t.Transect_ID, q.Quadrat, s.ColName & '_Q' & q.Quadrat"
    " AS QSfcColName, s.ColName & '_CoverID_Q' & q.Quadrat AS Quad_CoverID\015\012FRO"
    "M ((SurfaceCover AS sc INNER JOIN Surface AS s ON s.ID = sc.Surface_ID) INNER JO"
    "IN Quadrat AS q ON q.ID = sc.Quadrat_ID) LEFT JOIN Transect AS t ON t.Transect_I"
    "D = q.Transect_ID;\015\012"
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
        dbText "Name" ="q.Quadrat"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Surface_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Quadrat_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Surface"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.PercentCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QSfcColName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2625"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SurfaceCover_ID"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quad_CoverID"
        dbInteger "ColumnWidth" ="3990"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QSfcCoverID"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
