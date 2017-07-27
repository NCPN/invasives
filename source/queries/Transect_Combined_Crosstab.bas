dbMemo "SQL" ="SELECT t.*, s.IsSampled_Q1, s.IsSampled_Q2, s.IsSampled_Q3, ne.NoExotics_Q1, ne."
    "NoExotics_Q2, ne.NoExotics_Q3, sc.*, ci.*\015\012FROM (((Transect AS t LEFT JOIN"
    " Transect_IsSampled_Crosstab AS s ON s.Transect_ID = t.Transect_ID) LEFT JOIN Tr"
    "ansect_NoExotics_Crosstab AS ne ON ne.Transect_ID = t.Transect_ID) LEFT JOIN Tra"
    "nsect_SurfaceCover_Crosstab AS sc ON sc.Transect_ID = t.Transect_ID) LEFT JOIN T"
    "ransect_SurfaceCover_CoverID_Crosstab AS ci ON ci.Transect_ID = t.Transect_ID\015"
    "\012ORDER BY t.Event_ID, t.Transect;\015\012"
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
    Begin
        dbText "Name" ="ci.SurfaceCoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Surface"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.SurfaceCoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.SurfaceCoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Transect_ID"
        dbInteger "ColumnWidth" ="3330"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Rock_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Cryptogram_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Dung_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Dung_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Fungus_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Lichen_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Lichen_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Mineral_Soil_Sediment_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Moss_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Moss_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Road_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Rock_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Root_Bole_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Dead_Root_Bole_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Dead_Root_Bole_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Dead_Root_Bole_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Dead_Wood_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Standing_Water_Flooded_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Standing_Water_Flooded_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Stream_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Trash_Junk_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Rock_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Cryptogram_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Dead_Wood_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Dung_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Fungus_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Fungus_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Lichen_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Litter_Duff_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Litter_Duff_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Litter_Duff_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Mineral_Soil_Sediment_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Mineral_Soil_Sediment_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Moss_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Road_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Road_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Root_Bole_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Trash_Junk_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Cryptogram_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Dead_Wood_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Root_Bole_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Standing_Water_Flooded_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Stream_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Stream_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ci.Trash_Junk_CoverID_Q3"
        dbInteger "ColumnWidth" ="2805"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
