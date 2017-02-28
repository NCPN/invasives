Operation =1
Option =0
Begin InputTables
    Name ="tbl_master_version"
    Name ="tbl_SOP_version"
End
Begin OutputColumns
    Expression ="tbl_master_version.project_ID"
    Expression ="tbl_master_version.version_key_number"
    Expression ="tbl_master_version.version_key_date"
    Expression ="tbl_master_version.narrative_version"
    Expression ="tbl_master_version.version_comments"
    Expression ="tbl_SOP_version.SOP_number"
    Expression ="tbl_SOP_version.SOP_version_number"
    Expression ="tbl_SOP_version.active_flag"
End
Begin Joins
    LeftTable ="tbl_master_version"
    RightTable ="tbl_SOP_version"
    Expression ="tbl_master_version.version_key_number=tbl_SOP_version.version_key_number"
    Flag =2
End
Begin OrderBy
    Expression ="tbl_master_version.version_key_number"
    Flag =0
    Expression ="tbl_SOP_version.SOP_number"
    Flag =0
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
    Right =823
    Bottom =353
    Left =-1
    Top =-1
    Right =798
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =45
        Top =6
        Right =187
        Bottom =120
        Top =1
        Name ="tbl_master_version"
        Name =""
    End
    Begin
        Left =236
        Top =7
        Right =380
        Bottom =121
        Top =0
        Name ="tbl_SOP_version"
        Name =""
    End
End
