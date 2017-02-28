dbMemo "SQL" ="SELECT 0 AS ID, 'Alive' AS Flag \015\012FROM (SELECT COUNT(*) FROM MSysResources"
    ") AS DUAL\015\012UNION\015\012SELECT 1 AS ID, 'Dead' AS Flag \015\012FROM (SELEC"
    "T COUNT(*) FROM MSysResources) AS DUAL\015\012UNION SELECT ID, Flag FROM Flag;\015"
    "\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbText "Description" ="List of IsDead & Flag options"
Begin
    Begin
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Flag"
        dbLong "AggregateType" ="-1"
    End
End
