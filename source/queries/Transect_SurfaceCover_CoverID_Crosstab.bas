dbMemo "SQL" ="TRANSFORM Min(sc.SurfaceCover_ID) AS MinOfSurfaceCover_ID\015\012SELECT sc.Trans"
    "ect_ID\015\012FROM Quadrat_SurfaceCover AS sc\015\012GROUP BY sc.Transect_ID\015"
    "\012PIVOT sc.Quad_CoverID;\015\012"
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
        dbText "Name" ="sc.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Surface"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SurfaceCoverID_Q1"
        dbInteger "ColumnWidth" ="2055"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SurfaceCoverID_Q2"
        dbInteger "ColumnWidth" ="2505"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SurfaceCoverID_Q3"
        dbInteger "ColumnWidth" ="2730"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mineral_Soil_Sediment_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cryptogram_CoverID_Q1"
        dbInteger "ColumnWidth" ="2490"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dead_Wood_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dung_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fungus_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fungus_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lichen_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Litter_Duff_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Litter_Duff_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Litter_Duff_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mineral_Soil_Sediment_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Moss_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Road_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Road_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rock_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Root_Bole_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Trash_Junk_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cryptogram_CoverID_Q2"
        dbInteger "ColumnWidth" ="2490"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dead_Wood_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Root_Bole_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Standing_Water_Flooded_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stream_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stream_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Trash_Junk_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Moss_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cryptogram_CoverID_Q3"
        dbInteger "ColumnWidth" ="2490"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dung_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dung_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fungus_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lichen_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lichen_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mineral_Soil_Sediment_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Moss_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Road_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rock_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rock_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Root_Bole_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dead_Root_Bole_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dead_Root_Bole_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dead_Root_Bole_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dead_Wood_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Standing_Water_Flooded_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Standing_Water_Flooded_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stream_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Trash_Junk_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
End
