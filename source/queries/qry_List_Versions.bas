Operation =1
Option =0
Begin InputTables
    Name ="tbl_master_version"
End
Begin OutputColumns
    Expression ="tbl_master_version.project_ID"
    Expression ="tbl_master_version.version_key_number"
    Expression ="tbl_master_version.version_key_date"
    Expression ="tbl_master_version.narrative_version"
    Expression ="tbl_master_version.version_comments"
End
Begin OrderBy
    Expression ="tbl_master_version.version_key_number"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
End
Begin
    State =0
    Left =18
    Top =40
    Right =1002
    Bottom =353
    Left =-1
    Top =-1
    Right =973
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="tbl_master_version"
        Name =""
    End
End
