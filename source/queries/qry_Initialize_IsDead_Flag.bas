﻿dbMemo "SQL" ="UPDATE (tbl_Quadrat_Species AS s INNER JOIN tbl_Quadrat_Transect AS t ON t.Trans"
    "ect_ID = s.Transect_ID) INNER JOIN tbl_Events AS e ON e.Event_ID = t.Event_ID SE"
    "T IsDead = [What IsDead value do you want to set? (NULL, 0-alive, 1-dead, ?)]\015"
    "\012WHERE e.Start_Date > [What start date after which you want the values set (Y"
    "YYY/MM/DD)?];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End