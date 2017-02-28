Operation =1
Option =0
Having ="(((Year([Start_Date])) Is Not Null))"
Begin InputTables
    Name ="tbl_Events"
End
Begin OutputColumns
    Alias ="Visit_Year"
    Expression ="Year([Start_Date])"
End
Begin OrderBy
    Expression ="Year([Start_Date])"
    Flag =0
End
Begin Groups
    Expression ="Year([Start_Date])"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="Visit_Year"
    End
End
Begin
    State =0
    Left =18
    Top =40
    Right =985
    Bottom =353
    Left =-1
    Top =-1
    Right =956
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =543
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
