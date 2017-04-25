dbMemo "SQL" ="SELECT tbl_Locations.Unit_Code, SpeciesCover.PlantCode\015\012FROM (((tbl_Locati"
    "ons INNER JOIN tbl_Events ON tbl_Locations.Location_ID=tbl_Events.Location_ID) L"
    "EFT JOIN Transect ON tbl_Events.Event_ID=Transect.Event_ID) LEFT JOIN Quadrat ON"
    " Quadrat.Transect_ID=Transect.Transect_ID) LEFT JOIN SpeciesCover ON SpeciesCove"
    "r.Quadrat_ID = Quadrat.ID\015\012GROUP BY tbl_Locations.Unit_Code, SpeciesCover."
    "PlantCode\015\012HAVING (((SpeciesCover.PlantCode) Is Not Null));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbText "Description" ="Query all species by plot by park"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SpeciesCover.PlantCode"
        dbLong "AggregateType" ="-1"
    End
End
