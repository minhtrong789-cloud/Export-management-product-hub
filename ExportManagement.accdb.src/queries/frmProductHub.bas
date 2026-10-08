Operation =1
Option =0
Begin InputTables
    Name ="MarketingCategories"
End
Begin OutputColumns
    Expression ="MarketingCategories.MarketingCategory"
    Expression ="MarketingCategories.MarketingCategoryID"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="MarketingCategories.MarketingCategoryID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MarketingCategories.MarketingCategory"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1498
    Bottom =672
    Left =-1
    Top =-1
    Right =1473
    Bottom =406
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =72
        Top =18
        Right =288
        Bottom =234
        Top =0
        Name ="MarketingCategories"
        Name =""
    End
End
