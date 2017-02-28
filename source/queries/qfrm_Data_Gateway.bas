Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Location_ID"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tbl_Locations.Updated_Date"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbText "Description" ="List of sample locations and associated sampling events"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Location_ID"
        dbInteger "ColumnWidth" ="3900"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =101
    Top =82
    Right =878
    Bottom =402
    Left =-1
    Top =-1
    Right =770
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =87
        Top =2
        Right =205
        Bottom =101
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
End
