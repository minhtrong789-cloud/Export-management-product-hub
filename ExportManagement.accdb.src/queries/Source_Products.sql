SELECT
  p.ProductID,
  p.MarketingLineID,
  mktLines.MarketingLine,
  mktLines.MarketingCategoryID,
  mktCats.MarketingCategory,
  p.Barcode,
  p.HSCode,
  p.ProductName_en,
  p.ProductName_vn,
  p.PackageTypeID,
  pt.PackageType,
  p.Capacity,
  p.CapacityUnitID,
  cunits.CapacityUnit,
  p.Capacity_oz,
  p.Price_vnd,
  p.Price_usd,
  p.Length_cm,
  p.Width_cm,
  p.Height_cm,
  p.Weight_gram,
  p.Ingredients,
  p.USP_en,
  p.USP_vn,
  p.HowToUse_en,
  p.HowToUse_vn,
  p.ImagePath,
  p.OuterPaperBox,
  p.ShelfLife_month,
  p.ChangeDate,
  p.InsertDate,
  p.IsActive
FROM
  (
    (
      Products AS p
      LEFT JOIN PackageTypes AS pt ON p.PackageTypeID = pt.PackageTypeID
    )
    LEFT JOIN (
      MarketingLines AS mktLines
      LEFT JOIN MarketingCategories AS mktCats ON mktLines.MarketingCategoryID = mktCats.MarketingCategoryID
    ) ON p.MarketingLineID = mktLines.MarketingLineID
  )
  LEFT JOIN CapacityUnits AS cunits ON p.CapacityUnitID = cunits.CapacityUnitID
ORDER BY
  p.ProductID;
