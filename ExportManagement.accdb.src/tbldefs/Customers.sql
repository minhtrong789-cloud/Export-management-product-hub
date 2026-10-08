CREATE TABLE [Customers] (
  [CustomerID] AUTOINCREMENT CONSTRAINT [Index_167C762D_C156_4BF7] PRIMARY KEY UNIQUE NOT NULL,
  [CompanyName] VARCHAR (190),
  [Country] VARCHAR (50),
  [Continent] VARCHAR (50),
  [TaxID] VARCHAR (50),
  [CompanyPhone] VARCHAR (50),
  [CompanyEmail] VARCHAR (150),
  [CompanyWebsite] VARCHAR (150),
  [CompanyAddress1] LONGTEXT,
  [CompanyAddress2] LONGTEXT,
  [Notes] LONGTEXT,
  [IsActive] BIT,
  [InsertDate] DATETIME,
  [ChangeDate] DATETIME
)
