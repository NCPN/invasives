Operation =1
Option =1
Begin InputTables
    Name ="Quadrat_Species_Crosstab"
End
Begin OutputColumns
    Alias ="SumCover"
    Expression ="Q1+Q1_3m+Q1_0m+Q2+Q2_8m+Q2_5m+Q3+Q3_13m+Q3_10m"
    Alias ="AvgCover"
    Expression ="SumCover/NumSampledQuads"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ID_hm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuadID_hm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuadID_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID_10m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuadID_10m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PlantCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q2_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Q3_10m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Q3_13m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species_Crosstab.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species_Crosstab.PlantCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species_Crosstab.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species_Crosstab.NumSampledQuads"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species_Crosstab.Q3_10m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species_Crosstab.Q3_13m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Q2_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species_Crosstab.Q2_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Q1_0m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Q2_8m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species_Crosstab.Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species_Crosstab.Q1_0m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species_Crosstab.Q1_3m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species_Crosstab.Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species_Crosstab.Q2_8m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.PlantCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.NumSampledQuads"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Q1_3m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Species_Crosstab.Q3"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =55
    Top =15
    Right =867
    Bottom =803
    Left =-1
    Top =-1
    Right =780
    Bottom =341
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =461
        Bottom =304
        Top =0
        Name ="Quadrat_Species_Crosstab"
        Name =""
    End
End
