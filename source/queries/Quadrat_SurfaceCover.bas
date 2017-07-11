dbMemo "SQL" ="SELECT sc.ID, sc.Quadrat_ID, sc.Surface_ID, sc.PercentCover, s.Surface, s.Descri"
    "ption, t.Transect_ID, q.Quadrat, s.ColName & '_Q' & q.Quadrat AS QSfcColName\015"
    "\012FROM ((SurfaceCover AS sc INNER JOIN Surface AS s ON s.ID = sc.Surface_ID) I"
    "NNER JOIN Quadrat AS q ON q.ID = sc.Quadrat_ID) LEFT JOIN Transect AS t ON t.Tra"
    "nsect_ID = q.Transect_ID;\015\012"
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
        dbText "Name" ="sc.ID"
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
    End
End
