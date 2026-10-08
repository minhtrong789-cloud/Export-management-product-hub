CREATE TABLE [OrderDetails] (
  [OrderDetailID] AUTOINCREMENT CONSTRAINT [Index_46B42D85_6755_44CB] PRIMARY KEY UNIQUE NOT NULL,
  [OrderID] LONG CONSTRAINT [FK_OrderDetails_Orders] REFERENCES [Orders] ([OrderID]) ON DELETE CASCADE ,
  [ProductID] LONG,
  [Quantity] DOUBLE,
  [UnitPrice] CURRENCY,
  [UnitNetWeight] DOUBLE,
  [UnitGrossWeight] DOUBLE,
  [UnitCBM] DOUBLE,
  [IsActive] BIT,
  [InsertDate] DATETIME,
  [ChangeDate] DATETIME
)
