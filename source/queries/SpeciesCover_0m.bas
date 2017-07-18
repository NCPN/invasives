dbMemo "SQL" ="SELECT SpeciesCover_ID AS ID_hm, Quadrat_ID AS QuadID_hm, NULL AS ID_5m, NULL AS"
    " QuadID_5m, NULL AS ID_10m, NULL AS QuadID_10m, PlantCode, IsDead, PercentCover "
    "AS Q1_hm\015\012FROM Quadrat_Species_Position\015\012WHERE Position_m = 0;\015\012"
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
        dbText "Name" ="ID_hm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuadID_hm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuadID_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID_10m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuadID_10m"
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
        dbText "Name" ="Q1_hm"
        dbLong "AggregateType" ="-1"
    End
End
