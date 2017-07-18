dbMemo "SQL" ="PARAMETERS qid Long, [is] Long, ne Long;\015\012UPDATE Quadrat SET IsSampled = ["
    "is], NoExotics = [ne]\015\012WHERE ID = [qid];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
