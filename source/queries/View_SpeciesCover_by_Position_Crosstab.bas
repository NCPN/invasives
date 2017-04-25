dbMemo "SQL" ="SELECT vw.Transect_ID, vw.PlantCode, vw.ID_hm, vw.QuadID_hm, vw.ID_5m, vw.QuadID"
    "_5m, vw.ID_10m, vw.QuadID_10m, vw.PlantCode, vw.IsDead, vw.Q1_hm, vw.Q2_5m, vw.Q"
    "3_10m\015\012FROM View_SpeciesCover_by_Position AS vw\015\012GROUP BY vw.Transec"
    "t_ID, vw.PlantCode, vw.ID_hm, vw.QuadID_hm, vw.ID_5m, vw.QuadID_5m, vw.ID_10m, v"
    "w.QuadID_10m, vw.PlantCode, vw.IsDead, vw.Q1_hm, vw.Q2_5m, vw.Q3_10m;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "OrderByOn" ="0"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="vw.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw.Q1_hm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw.Q2_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw.Q3_10m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw.PlantCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1001"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw.ID_hm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw.QuadID_hm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw.ID_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw.QuadID_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw.ID_10m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw.QuadID_10m"
        dbLong "AggregateType" ="-1"
    End
End
