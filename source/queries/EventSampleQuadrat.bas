dbMemo "SQL" ="SELECT DISTINCT e.Start_Date, IIF(Year(e.Start_Date) = SamplingYear, SamplingYea"
    "r, \015\012         (SELECT MAX (SamplingYear) FROM QuadratPosition) ) AS Sampli"
    "ngYr, Quadrat\015\012FROM tbl_Events AS e, QuadratPosition AS qp\015\012WHERE Ye"
    "ar(e.Start_Date) = SamplingYear\015\012OR\015\012Year(e.Start_Date) > (SELECT MA"
    "X(SamplingYear) FROM QuadratPosition);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="e.Start_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SamplingYr"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1665"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Quadrat"
        dbLong "AggregateType" ="-1"
    End
End
