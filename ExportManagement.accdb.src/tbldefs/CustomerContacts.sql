CREATE TABLE [CustomerContacts] (
  [ContactID] AUTOINCREMENT CONSTRAINT [Index_377A19BF_43ED_4460] PRIMARY KEY UNIQUE NOT NULL,
  [CustomerID] LONG CONSTRAINT [FK_CustomerContacts_Customers] REFERENCES [Customers] ([CustomerID]) ON DELETE CASCADE ,
  [ContactPerson] VARCHAR (100),
  [JobTitle] VARCHAR (80),
  [ContactPersonEmail] VARCHAR (100),
  [ContactPersonPhone] VARCHAR (50),
  [IsPrimary] BIT
)
