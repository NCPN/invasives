Operation =1
Option =0
Begin InputTables
    Name ="tbl_wrk_Species_List"
    Name ="tlu_NCPN_Plants"
    Name ="tlu_Parks"
End
Begin OutputColumns
    Expression ="tbl_wrk_Species_List.Plot_ID"
    Expression ="tbl_wrk_Species_List.P1"
    Expression ="tbl_wrk_Species_List.P2"
    Expression ="tbl_wrk_Species_List.P3"
    Expression ="tbl_wrk_Species_List.P4"
    Expression ="tbl_wrk_Species_List.P5"
    Expression ="tbl_wrk_Species_List.P6"
    Expression ="tbl_wrk_Species_List.P7"
    Expression ="tbl_wrk_Species_List.P8"
    Expression ="tbl_wrk_Species_List.P9"
    Expression ="tbl_wrk_Species_List.P10"
    Expression ="tlu_NCPN_Plants.Master_Family"
    Expression ="tlu_NCPN_Plants.Utah_Species"
    Expression ="tlu_Parks.ParkName"
End
Begin Joins
    LeftTable ="tbl_wrk_Species_List"
    RightTable ="tlu_NCPN_Plants"
    Expression ="tbl_wrk_Species_List.Plant_Code=tlu_NCPN_Plants.Master_PLANT_Code"
    Flag =2
    LeftTable ="tbl_wrk_Species_List"
    RightTable ="tlu_Parks"
    Expression ="tbl_wrk_Species_List.Park_Code=tlu_Parks.ParkCode"
    Flag =2
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
        Top =10
        Name ="tbl_wrk_Species_List"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =284
        Bottom =120
        Top =5
        Name ="tlu_NCPN_Plants"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =105
        Top =0
        Name ="tlu_Parks"
        Name =""
    End
End
