Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
End
Begin OutputColumns
    Expression ="tbl_Events.*"
    Expression ="tbl_Locations.Plot_ID"
    Expression ="tbl_Locations.Unit_Code"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID=tbl_Events.Location_ID"
    Flag =3
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbText "Description" ="Data entry form record source"
Begin
End
Begin
    State =0
    Left =-33
    Top =62
    Right =1003
    Bottom =382
    Left =-1
    Top =-1
    Right =1029
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =306
        Top =6
        Right =431
        Bottom =120
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =2
        Name ="tbl_Events"
        Name =""
    End
End
