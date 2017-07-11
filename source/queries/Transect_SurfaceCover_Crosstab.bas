dbMemo "SQL" ="TRANSFORM Min(qsc.PercentCover) AS MinOfPercentCover\015\012SELECT qsc.Transect_"
    "ID\015\012FROM Quadrat_SurfaceCover AS qsc\015\012GROUP BY qsc.Transect_ID\015\012"
    "PIVOT qsc.QSfcColName;\015\012"
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
        dbText "Name" ="Dead_Root_Bole_Q1"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dead_Root_Bole_Q3"
        dbInteger "ColumnWidth" ="2850"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dead_Wood_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Root_Bole_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Trash_Junk_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qsc.Quadrat"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cryptogram_Q3"
        dbInteger "ColumnWidth" ="1965"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dead_Wood_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Litter_Duff_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mineral_Soil_Sediment_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mineral_Soil_Sediment_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mineral_Soil_Sediment_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Moss_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Moss_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Moss_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Road_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Road_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Road_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rock_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rock_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rock_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Root_Bole_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Trash_Junk_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qsc.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qsc.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cryptogram_Q2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Dead_Root_Bole_Q2"
        dbInteger "ColumnWidth" ="2445"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dead_Wood_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Standing_Water_Flooded_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Standing_Water_Flooded_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Standing_Water_Flooded_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stream_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stream_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stream_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Trash_Junk_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qsc.Quadrat_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cryptogram_Q1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2010"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Dung_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dung_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dung_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fungus_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fungus_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fungus_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lichen_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lichen_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lichen_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Litter_Duff_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Litter_Duff_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Root_Bole_Q3"
        dbLong "AggregateType" ="-1"
    End
End
