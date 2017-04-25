dbMemo "SQL" ="SELECT q.Transect_ID, sc.ID AS ID_hm, Quadrat_ID AS QuadID_hm, IIf(False, 0, NUL"
    "L) AS ID_5m, IIf(False, 0, NULL) AS QuadID_5m, IIf(False, 0, NULL) AS ID_10m, II"
    "f(False, 0, NULL) AS QuadID_10m, PlantCode, IsDead, PercentCover AS Q1_hm, IIf(F"
    "alse, 0, NULL) AS Q2_5m, IIf(False, 0, NULL) AS Q3_10m\015\012FROM SpeciesCover "
    "AS sc INNER JOIN Quadrat AS q ON q.ID = sc.Quadrat_ID\015\012WHERE Position_m = "
    "0;\015\012"
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
        dbText "Name" ="v.ID_hm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.QuadID_hm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.ID_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.QuadID_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.ID_10m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.QuadID_10m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.PlantCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.Q1_hm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.Q2_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.Q3_10m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q.Transect_ID"
        dbInteger "ColumnWidth" ="3405"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
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
    Begin
        dbText "Name" ="Q2_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q3_10m"
        dbLong "AggregateType" ="-1"
    End
End
