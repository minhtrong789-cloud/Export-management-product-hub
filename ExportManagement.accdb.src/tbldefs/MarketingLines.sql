CREATE TABLE [MarketingLines] (
  [MarketingLineID] AUTOINCREMENT CONSTRAINT [Index_995CDE23_0BE1_4F4A] PRIMARY KEY UNIQUE NOT NULL,
  [MarketingLine] VARCHAR (50) CONSTRAINT [Index_11DFCCB8_A7EF_4CD2] UNIQUE,
  [MarketingCategoryID] LONG CONSTRAINT [FK_MarketingLines_MarketingCategories] REFERENCES [MarketingCategories] ([MarketingCategoryID]) ON UPDATE CASCADE 
)
