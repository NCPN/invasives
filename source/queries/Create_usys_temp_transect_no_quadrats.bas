dbMemo "SQL" ="SELECT t.*, u.IsSampled_Q1, u.IsSampled_Q2, u.IsSampled_Q3, u.NoExotics_Q1, u.No"
    "Exotics_Q2, u.NoExotics_Q3, u.sc_Transect_ID, u.Cryptogram_Q1, u.Cryptogram_Q2, "
    "u.Cryptogram_Q3, u.Dead_Root_Bole_Q1, u.Dead_Root_Bole_Q2, u.Dead_Root_Bole_Q3, "
    "u.Dead_Wood_Q1, u.Dead_Wood_Q2, u.Dead_Wood_Q3, u.Dung_Q1, u.Dung_Q2, u.Dung_Q3,"
    " u.Fungus_Q1, u.Fungus_Q2, u.Fungus_Q3, u.Lichen_Q1, u.Lichen_Q2, u.Lichen_Q3, u"
    ".Litter_Duff_Q1, u.Litter_Duff_Q2, u.Litter_Duff_Q3, u.Mineral_Soil_Sediment_Q1,"
    " u.Mineral_Soil_Sediment_Q2, u.Mineral_Soil_Sediment_Q3, u.Moss_Q1, u.Moss_Q2, u"
    ".Moss_Q3, u.Road_Q1, u.Road_Q2, u.Road_Q3, u.Rock_Q1, u.Rock_Q2, u.Rock_Q3, u.Ro"
    "ot_Bole_Q1, u.Root_Bole_Q2, u.Root_Bole_Q3, u.Standing_Water_Flooded_Q1, u.Stand"
    "ing_Water_Flooded_Q2, u.Standing_Water_Flooded_Q3, u.Stream_Q1, u.Stream_Q2, u.S"
    "tream_Q3, u.Trash_Junk_Q1, u.Trash_Junk_Q2, u.Trash_Junk_Q3, u.ci_Transect_ID, u"
    ".Cryptogram_CoverID_Q1, u.Cryptogram_CoverID_Q2, u.Cryptogram_CoverID_Q3, u.Dead"
    "_Root_Bole_CoverID_Q1, u.Dead_Root_Bole_CoverID_Q2, u.Dead_Root_Bole_CoverID_Q3,"
    " u.Dead_Wood_CoverID_Q1, u.Dead_Wood_CoverID_Q2, u.Dead_Wood_CoverID_Q3, u.Dung_"
    "CoverID_Q1, u.Dung_CoverID_Q2, u.Dung_CoverID_Q3, u.Fungus_CoverID_Q1, u.Fungus_"
    "CoverID_Q2, u.Fungus_CoverID_Q3, u.Lichen_CoverID_Q1, u.Lichen_CoverID_Q2, u.Lic"
    "hen_CoverID_Q3, u.Litter_Duff_CoverID_Q1, u.Litter_Duff_CoverID_Q2, u.Litter_Duf"
    "f_CoverID_Q3, u.Mineral_Soil_Sediment_CoverID_Q1, u.Mineral_Soil_Sediment_CoverI"
    "D_Q2, u.Mineral_Soil_Sediment_CoverID_Q3, u.Moss_CoverID_Q1, u.Moss_CoverID_Q2, "
    "u.Moss_CoverID_Q3, u.Road_CoverID_Q1, u.Road_CoverID_Q2, u.Road_CoverID_Q3, u.Ro"
    "ck_CoverID_Q1, u.Rock_CoverID_Q2, u.Rock_CoverID_Q3, u.Root_Bole_CoverID_Q1, u.R"
    "oot_Bole_CoverID_Q2, u.Root_Bole_CoverID_Q3, u.Standing_Water_Flooded_CoverID_Q1"
    ", u.Standing_Water_Flooded_CoverID_Q2, u.Standing_Water_Flooded_CoverID_Q3, u.St"
    "ream_CoverID_Q1, u.Stream_CoverID_Q2, u.Stream_CoverID_Q3, u.Trash_Junk_CoverID_"
    "Q1, u.Trash_Junk_CoverID_Q2, u.Trash_Junk_CoverID_Q3 INTO usys_temp_transect\015"
    "\012FROM Transect AS t LEFT JOIN usys_temp_transect_quadrat_empty AS u ON t.Tran"
    "sect_ID = u.Transect_ID;\015\012"
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
        dbText "Name" ="usys_temp_quadrat_empty.Standing_Water_Flooded_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Standing_Water_Flooded_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Standing_Water_Flooded_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Stream_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Stream_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Stream_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Trash_Junk_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Cryptogram_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Event_ID"
        dbLong "AggregateType" ="-1"
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
        dbText "Name" ="usys_temp_quadrat_empty.IsSampled_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.NoExotics_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Cryptogram_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Dead_Root_Bole_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Dead_Wood_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Dead_Wood_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Root_Bole_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Standing_Water_Flooded_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Stream_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Stream_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Trash_Junk_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.ci_Transect_ID"
        dbInteger "ColumnWidth" ="1665"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Cryptogram_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Dung_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Dung_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Fungus_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Lichen_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Lichen_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Mineral_Soil_Sediment_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Moss_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Moss_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Road_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Rock_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Rock_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Root_Bole_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Root_Bole_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.ci_Transect_ID"
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
        dbText "Name" ="usys_temp_quadrat_empty.IsSampled_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.NoExotics_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Cryptogram_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Dung_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Dung_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Dung_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Fungus_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Fungus_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Fungus_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Lichen_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Lichen_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Lichen_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Litter_Duff_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Litter_Duff_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Mineral_Soil_Sediment_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Mineral_Soil_Sediment_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Mineral_Soil_Sediment_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Moss_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Moss_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Moss_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Road_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Road_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Road_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Rock_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Rock_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Rock_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Root_Bole_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Root_Bole_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Root_Bole_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Standing_Water_Flooded_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Standing_Water_Flooded_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Standing_Water_Flooded_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Stream_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Dung_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Dung_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Fungus_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Lichen_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Lichen_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Mineral_Soil_Sediment_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Moss_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Moss_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Road_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Rock_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Rock_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Root_Bole_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Lichen_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Lichen_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Lichen_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Litter_Duff_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Litter_Duff_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Litter_Duff_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Trash_Junk_Q3"
        dbInteger "ColumnWidth" ="1740"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Dead_Root_Bole_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Dead_Root_Bole_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Dead_Root_Bole_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Dead_Wood_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Standing_Water_Flooded_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Standing_Water_Flooded_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Stream_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Trash_Junk_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Cryptogram_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Root_Bole_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Trash_Junk_Q3"
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
        dbText "Name" ="t.E_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.N_Coord"
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
        dbText "Name" ="usys_temp_quadrat_empty.NoExotics_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Dead_Root_Bole_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Dead_Root_Bole_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Dead_Wood_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Trash_Junk_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Standing_Water_Flooded_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Standing_Water_Flooded_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Stream_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Trash_Junk_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Cryptogram_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Dead_Wood_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Dung_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Fungus_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Fungus_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Lichen_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Litter_Duff_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Litter_Duff_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Litter_Duff_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Mineral_Soil_Sediment_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Mineral_Soil_Sediment_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Moss_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Road_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Road_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Rock_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Root_Bole_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Trash_Junk_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Dead_Root_Bole_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Dead_Root_Bole_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Dead_Root_Bole_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Dead_Wood_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Dead_Root_Bole_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Dead_Root_Bole_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Dead_Wood_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Dead_Wood_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Dead_Wood_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Dung_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.IsSampled_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.IsSampled_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.IsSampled_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.NoExotics_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.NoExotics_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.NoExotics_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.sc_Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Cryptogram_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Cryptogram_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Cryptogram_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Dead_Root_Bole_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Dung_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Dung_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Fungus_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Fungus_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Fungus_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Rock_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Root_Bole_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Cryptogram_CoverID_Q1"
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
        dbText "Name" ="usys_temp_quadrat_empty.IsSampled_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.sc_Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Cryptogram_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Dead_Wood_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Litter_Duff_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Mineral_Soil_Sediment_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Mineral_Soil_Sediment_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Mineral_Soil_Sediment_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Moss_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Moss_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Moss_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Road_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Road_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Road_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Rock_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Rock_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Trash_Junk_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Stream_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Stream_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Trash_Junk_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Dung_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Fungus_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Fungus_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Lichen_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Litter_Duff_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Litter_Duff_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Litter_Duff_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Mineral_Soil_Sediment_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Mineral_Soil_Sediment_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Moss_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Road_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Road_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Rock_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Root_Bole_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Trash_Junk_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Cryptogram_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Dead_Wood_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Root_Bole_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Standing_Water_Flooded_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Stream_CoverID_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Stream_CoverID_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="u.Trash_Junk_CoverID_Q3"
        dbInteger "ColumnWidth" ="2505"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="usys_temp_quadrat_empty.Dead_Wood_CoverID_Q2"
        dbLong "AggregateType" ="-1"
    End
End
