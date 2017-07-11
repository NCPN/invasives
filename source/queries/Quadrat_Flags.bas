dbMemo "SQL" ="SELECT q.*, 'IsSampled_Q' & q.Quadrat AS QIsSampledColName, 'NoExotics_Q' & q.Qu"
    "adrat AS QNoExoticsColName\015\012FROM Quadrat AS q;\015\012"
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
        dbText "Name" ="q.IsSampled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q.NoExotics"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QIsSampledColName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QNoExoticsColName"
        dbLong "AggregateType" ="-1"
    End
End
