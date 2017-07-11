dbMemo "SQL" ="TRANSFORM Min(Transect_Quadrat.[IsSampled]) AS MinOfIsSampled\015\012SELECT Tran"
    "sect_Quadrat.[t].[Transect_ID], Transect_Quadrat.[Event_ID], Transect_Quadrat.[T"
    "ransect], Transect_Quadrat.Start_Time, Transect_Quadrat.GPS_File_Name, Transect_"
    "Quadrat.Rcvr_Type, Transect_Quadrat.Elevation, Transect_Quadrat.N_Coord, Transec"
    "t_Quadrat.UTM_Zone, Transect_Quadrat.Datum, Transect_Quadrat.Max_PDOP, Transect_"
    "Quadrat.Corr_Type, Transect_Quadrat.GPS_Time, Transect_Quadrat.Update_Status, Tr"
    "ansect_Quadrat.Feat_Name, Transect_Quadrat.Vert_Prec, Transect_Quadrat.Std_Dev, "
    "Transect_Quadrat.Recorder\015\012FROM Transect_Quadrat\015\012GROUP BY Transect_"
    "Quadrat.[t].[Transect_ID], Transect_Quadrat.[Event_ID], Transect_Quadrat.[Transe"
    "ct], Transect_Quadrat.Start_Time, Transect_Quadrat.GPS_File_Name, Transect_Quadr"
    "at.Rcvr_Type, Transect_Quadrat.Elevation, Transect_Quadrat.N_Coord, Transect_Qua"
    "drat.UTM_Zone, Transect_Quadrat.Datum, Transect_Quadrat.Max_PDOP, Transect_Quadr"
    "at.Corr_Type, Transect_Quadrat.GPS_Time, Transect_Quadrat.Update_Status, Transec"
    "t_Quadrat.Feat_Name, Transect_Quadrat.Vert_Prec, Transect_Quadrat.Std_Dev, Trans"
    "ect_Quadrat.Recorder\015\012PIVOT \"Q\" & [Quadrat];\015\012"
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
        dbText "Name" ="[t].[Transect_ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Event_ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Transect]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Quadrat.[t].[Transect_ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Quadrat.Start_Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Quadrat.GPS_File_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Quadrat.Rcvr_Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Quadrat.Elevation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Quadrat.N_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Quadrat.UTM_Zone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Quadrat.Datum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Quadrat.Max_PDOP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Quadrat.Corr_Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Quadrat.GPS_Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Quadrat.Update_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Quadrat.Feat_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Quadrat.Vert_Prec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Quadrat.Std_Dev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Quadrat.Comments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Quadrat.Recorder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Quadrat.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PIVOT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Quadrat.[Event_ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Quadrat.[Transect]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q3"
        dbLong "AggregateType" ="-1"
    End
End
