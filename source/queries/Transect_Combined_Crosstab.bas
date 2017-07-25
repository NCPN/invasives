dbMemo "SQL" ="SELECT t.*, s.IsSampled_Q1, s.IsSampled_Q2, s.IsSampled_Q3, ne.NoExotics_Q1, ne."
    "NoExotics_Q2, ne.NoExotics_Q3, sc.*\015\012FROM ((Transect AS t LEFT JOIN Transe"
    "ct_IsSampled_Crosstab AS s ON s.Transect_ID = t.Transect_ID) LEFT JOIN Transect_"
    "NoExotics_Crosstab AS ne ON ne.Transect_ID = t.Transect_ID) LEFT JOIN Transect_S"
    "urfaceCover_Crosstab AS sc ON sc.Transect_ID = t.Transect_ID\015\012ORDER BY t.E"
    "vent_ID, t.Transect;\015\012"
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
        dbText "Name" ="ne.NoExotics_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Dead_Root_Bole_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Dead_Root_Bole_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Vert_Prec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Horz_Prec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Std_Dev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Comments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Stop_Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Observer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Recorder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.IsSampled_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Dead_Wood_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Root_Bole_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Trash_Junk_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.IsSampled_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Cryptogram_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Rcvr_Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Elevation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.UTM_Zone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Datum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Feat_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Unfilt_Pos"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Dead_Wood_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Litter_Duff_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Mineral_Soil_Sediment_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Mineral_Soil_Sediment_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Mineral_Soil_Sediment_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Moss_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Moss_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Moss_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Road_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Road_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Road_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Rock_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Rock_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Rock_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Root_Bole_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Trash_Junk_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.E_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.N_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ne.NoExotics_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Cryptogram_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Dead_Root_Bole_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.GPS_File_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Max_PDOP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Corr_Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.GPS_Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Update_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Dead_Wood_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Standing_Water_Flooded_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Standing_Water_Flooded_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Standing_Water_Flooded_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Stream_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Stream_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Stream_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Trash_Junk_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ne.NoExotics_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Cryptogram_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Event_ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3330"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="t.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Start_Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.IsSampled_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Dung_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Dung_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Dung_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Fungus_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Fungus_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Fungus_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Lichen_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Lichen_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Lichen_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Litter_Duff_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Litter_Duff_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Root_Bole_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_ID"
        dbLong "AggregateType" ="-1"
    End
End
