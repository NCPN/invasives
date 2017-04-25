dbMemo "SQL" ="PARAMETERS tid Text ( 50 );\015\012SELECT t.Transect_ID, q.Quadrat, s.Surface, s"
    ".ColName, s.ColName & '_Q' & q.Quadrat AS ControlName, sc.ID, sc.Quadrat_ID, sc."
    "Surface_ID, sc.PercentCover\015\012FROM (((SurfaceCover AS sc INNER JOIN Surface"
    " AS s ON s.ID = sc.Surface_ID) INNER JOIN Quadrat AS q ON q.ID = sc.Quadrat_ID) "
    "INNER JOIN Transect AS t ON t.Transect_ID = q.Transect_ID) INNER JOIN tbl_Events"
    " AS e ON e.Event_ID = t.Event_ID\015\012WHERE t.Transect_ID = [tid];\015\012"
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
        dbText "Name" ="t.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Surface"
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
        dbText "Name" ="sc.PercentCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q.Quadrat"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.ColName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ControlName"
        dbLong "AggregateType" ="-1"
    End
End
