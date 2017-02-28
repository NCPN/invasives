dbMemo "SQL" ="SELECT *\015\012FROM (tbl_Quadrat_Species AS s INNER JOIN tbl_Quadrat_Transect A"
    "S t ON t.Transect_ID = s.Transect_ID) INNER JOIN tbl_Events AS e ON e.Event_ID ="
    " t.Event_ID\015\012WHERE e.Start_Date > [What start date after which you want th"
    "e values set (YYYY/MM/DD)?];\015\012"
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
        dbText "Name" ="t.Dead_Wood_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Root_Bole_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Dead_Root_Bole_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Mineral_Soil_Sediment_Q1"
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
        dbText "Name" ="s.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Q3_10m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Average_Cover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Q2_8m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Transect_ID"
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
        dbText "Name" ="t.Road_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Road_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Road_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Rock_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Rock_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Standing_Water_Flooded_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Standing_Water_Flooded_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Standing_Water_Flooded_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Stream_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Stream_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Stream_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Trash_Junk_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Observer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Dead_Wood_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Litter_Duff_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Root_Bole_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Rock_Q1"
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
        dbText "Name" ="s.Species_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Q3_13m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Avg_Cover_2009"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.GPS_File_Name"
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
        dbText "Name" ="t.Recorder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Cryptogram_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Comments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Dead_Wood_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Dead_Root_Bole_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Dead_Root_Bole_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Q1_hm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Avg_Cover_2008"
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
        dbText "Name" ="t.Stop_Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Observer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Cryptogram_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Trash_Junk_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.version_key_number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Start_Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Dung_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Dung_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Dung_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Fungus_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Fungus_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Fungus_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Lichen_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Lichen_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Lichen_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Litter_Duff_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Litter_Duff_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Root_Bole_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Plant_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Q2_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Q1_3m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Event_ID"
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
        dbText "Name" ="t.Mineral_Soil_Sediment_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Mineral_Soil_Sediment_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Moss_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Moss_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Moss_Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Cryptogram_Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Trash_Junk_Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Protocol_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Start_Date"
        dbLong "AggregateType" ="-1"
    End
End
