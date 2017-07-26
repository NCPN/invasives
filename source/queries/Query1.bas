dbMemo "SQL" ="PARAMETERS eid Text ( 50 );\015\012SELECT Transect_ID, Transect\015\012FROM Tran"
    "sect\015\012WHERE Event_ID = [eid];\015\012"
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
        dbText "Name" ="Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect"
        dbLong "AggregateType" ="-1"
    End
End
