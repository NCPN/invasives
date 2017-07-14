dbMemo "SQL" ="PARAMETERS tid Text ( 50 ), oid Long, start DateTime, cmt Text ( 255 );\015\012U"
    "PDATE Transect SET Start_Time = [start], Observer = [oid], Comments = [cmt]\015\012"
    "WHERE Transect_ID = [tid];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbInteger "RowHeight" ="2130"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Transect.Elevation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.E_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.N_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.UTM_Zone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.Datum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.Max_PDOP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.Corr_Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.GPS_Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.Update_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.Feat_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.Unfilt_Pos"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.Start_Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.GPS_File_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.Rcvr_Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.Vert_Prec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.Horz_Prec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.Std_Dev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.Comments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.Stop_Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.Observer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.Recorder"
        dbLong "AggregateType" ="-1"
    End
End
