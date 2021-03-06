﻿dbMemo "SQL" ="(SELECT \015\012q.Transect_ID, \015\012sc.SpeciesCover_ID AS ID_hm, \015\012Quad"
    "rat_ID AS QuadID_hm, \015\012IIf(False, 0, NULL) AS ID_5m, IIf(False, 0, NULL) A"
    "S QuadID_5m,\015\012IIf(False, 0, NULL) AS ID_10m, IIf(False, 0, NULL) AS QuadID"
    "_10m,\015\012PlantCode, IsDead, \015\012PercentCover AS Q1_hm, IIf(False, 0, NUL"
    "L) AS Q2_5m, IIf(False, 0, NULL) AS Q3_10m\015\012FROM Quadrat_Species_Position "
    "sc\015\012WHERE Position_m = 0)\015\012UNION ALL\015\012(SELECT\015\012q.Transec"
    "t_ID, \015\012IIf(False, 0, NULL) AS ID_hm, IIf(False, 0, NULL) AS QuadID_hm, \015"
    "\012sc.SpeciesCover_ID AS ID_5m, Quadrat_ID AS QuadID_5m,\015\012IIf(False, 0, N"
    "ULL) AS ID_10m, IIf(False, 0, NULL) AS QuadID_10m,\015\012PlantCode, IsDead, \015"
    "\012IIf(False, 0, NULL) AS Q1_hm, PercentCover AS Q2_5m, IIf(False, 0, NULL) AS "
    "Q3_10m\015\012FROM Quadrat_Species_Position sc\015\012WHERE Position_m = 5)\015\012"
    "UNION ALL (SELECT\015\012q.Transect_ID, \015\012IIf(False, 0, NULL) AS ID_hm, II"
    "f(False, 0, NULL) AS QuadID_hm, \015\012IIf(False, 0, NULL) AS ID_5m, IIf(False,"
    " 0, NULL) AS QuadID_5m,\015\012sc.SpeciesCover_ID AS ID_10m, Quadrat_ID AS QuadI"
    "D_10m,\015\012PlantCode, IsDead, \015\012IIf(False, 0, NULL) AS Q1_hm, IIf(False"
    ", 0, NULL) AS Q2_5m, PercentCover AS Q3_10m\015\012FROM Quadrat_Species_Position"
    " sc\015\012WHERE Position_m = 10);\015\012"
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
        dbText "Name" ="q.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
End
