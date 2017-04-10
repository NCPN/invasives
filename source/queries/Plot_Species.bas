dbMemo "SQL" ="SELECT tbl_Locations.Unit_Code, tbl_Locations.Plot_ID, SpeciesCover.PlantCode\015"
    "\012FROM (((tbl_Locations INNER JOIN tbl_Events ON tbl_Locations.Location_ID=tbl"
    "_Events.Location_ID) LEFT JOIN Transect ON tbl_Events.Event_ID=Transect.Event_ID"
    ") LEFT JOIN Quadrat ON Quadrat.Transect_ID=Transect.Transect_ID) LEFT JOIN Speci"
    "esCover ON SpeciesCover.Quadrat_ID = Quadrat.ID\015\012GROUP BY tbl_Locations.Un"
    "it_Code, tbl_Locations.Plot_ID, SpeciesCover.PlantCode\015\012HAVING (((SpeciesC"
    "over.PlantCode) Is Not Null));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbText "Description" ="Query all species by plot by park"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SpeciesCover.PlantCode"
        dbLong "AggregateType" ="-1"
    End
End
