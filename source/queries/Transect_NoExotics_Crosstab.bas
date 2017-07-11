dbMemo "SQL" ="TRANSFORM Min(qf.NoExotics) AS MinOfNoExotics\015\012SELECT qf.Transect_ID\015\012"
    "FROM Quadrat_Flags AS qf\015\012GROUP BY qf.Transect_ID\015\012PIVOT qf.QNoExoti"
    "csColName;\015\012"
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
        dbText "Name" ="qf.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NoExotics_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NoExotics_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NoExotics_Q3"
        dbLong "AggregateType" ="-1"
    End
End
