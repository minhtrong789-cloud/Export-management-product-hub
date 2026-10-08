dbMemo "SQL" ="SELECT c.CustomerID, c.CompanyName, cc.ContactPerson, cc.ContactPersonEmail, cc."
    "ContactPersonPhone, c.CompanyEmail, c.CompanyPhone, c.CompanyAddress1 AS Consign"
    "eeAddress, c.CompanyAddress2 AS BillingAddress, c.Country, c.TaxID, NULL AS Defa"
    "ultCurrency, NULL AS DefaultPaymentTerm, NULL AS DefaultIncoterm, c.Notes, c.IsA"
    "ctive, c.InsertDate, c.ChangeDate\015\012FROM Customers AS c LEFT JOIN (SELECT *"
    " FROM CustomerContacts WHERE IsPrimary = True)  AS cc ON c.CustomerID = cc.Custo"
    "merID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
