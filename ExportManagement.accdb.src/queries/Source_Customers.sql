SELECT
  c.CustomerID,
  c.CompanyName,
  cc.ContactPerson,
  cc.ContactPersonEmail,
  cc.ContactPersonPhone,
  c.CompanyEmail,
  c.CompanyPhone,
  c.CompanyAddress1 AS ConsigneeAddress,
  c.CompanyAddress2 AS BillingAddress,
  c.Country,
  c.TaxID,
  NULL AS DefaultCurrency,
  NULL AS DefaultPaymentTerm,
  NULL AS DefaultIncoterm,
  c.Notes,
  c.IsActive,
  c.InsertDate,
  c.ChangeDate
FROM
  Customers AS c
  LEFT JOIN (
    SELECT
      *
    FROM
      CustomerContacts
    WHERE
      IsPrimary = True
  ) AS cc ON c.CustomerID = cc.CustomerID;
