CREATE TABLE [CustomerAddresses] (
  [AddressID] AUTOINCREMENT CONSTRAINT [Index_6339310D_2535_4F0A] PRIMARY KEY UNIQUE NOT NULL,
  [CustomerID] LONG CONSTRAINT [FK_CustomerAddresses_Customers] REFERENCES [Customers] ([CustomerID]) ON DELETE CASCADE ,
  [AddressLabel] VARCHAR (100),
  [RecipientName] VARCHAR (190),
  [AddressLine1] LONGTEXT,
  [AddressLine2] LONGTEXT,
  [Country] VARCHAR (50)
)
