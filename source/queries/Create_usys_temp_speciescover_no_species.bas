﻿dbMemo "SQL" ="SELECT t.Transect_ID, u.PlantCode, u.IsDead, u.NumSampledQuads, u.SumOfPercentCo"
    "ver, u.AvgCover, u.Q1_0m, u.Q2_4_5m, u.Q3_9_5m, u.SpeciesCoverID_Q1, u.SpeciesCo"
    "verID_Q2, u.SpeciesCoverID_Q3 INTO usys_temp_speciescover\015\012FROM Transect A"
    "S t LEFT JOIN usys_temp_speciescover_empty AS u ON t.Transect_ID = u.Transect_ID"
    ";\015\012"
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