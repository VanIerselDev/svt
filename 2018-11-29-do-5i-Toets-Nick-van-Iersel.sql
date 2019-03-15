-- Pagina 24
-- Oef 2
SELECT Naam FROM tblKlanten WHERE Klantnummer = 1
-- Oef 4
SELECT Gemeente FROM tblKlanten WHERE Klantnummer = 268

-- Pagina 48 (Bovenaan)
-- Oef 2
SELECT NederlandseNaam FROM tblProducten WHERE (Bestelpunt > Voorraad)

-- Pagina 48 (Onderaan)
-- Oef 2
SELECT * FROM tblKlanten WHERE Gemeente = 'Leuven'
-- Oef 4
SELECT * FROM tblKlanten WHERE Saldo > 70 ORDER BY Naam
-- Oef 6
SELECT Saldo FROM tblKlanten WHERE BTWnr IS NOT NULL

-- Pagina 59
-- Oef 2
SELECT NederlandseNaam, PrijsPerEenheid, ((PrijsPerEenheid * (1 + 0.33)) * (1 + 0.21)) As Verkoopprijs FROM tblProducten
-- Oef 4
SELECT Functie FROM tblWerknemers GROUP BY Functie
-- Oef 6
SELECT Productnummer, Productnaam, NederlandseNaam FROM tblProducten WHERE Leveranciersnummer = :Leveranciersnummer

-- Pagina 77
-- Oef 2
SELECT * FROM tblOrderinformatie WHERE "Order-ID" = 10001
-- Oef 4
SELECT tblKlanten.Straat, tblKlanten.Postcode ,tblKlanten.Gemeente FROM tblOrders INNER JOIN tblKlanten ON tblOrders.Klantnummer = tblKlanten.Klantnummer WHERE "Order-ID" = 10740
-- Oef 6
SELECT DISTINCT Klantnummer FROM tblOrders RIGHT JOIN tblKlanten ON tblOrders.Klantnummer = tblKlanten.Klantnummer LEFT JOIN  tblOrderinformatie ON tblOrders."Order-ID" = tblOrderinformatie."Order-ID" WHERE "Order-ID" IS NULL

-- Pagina 122-123
-- Oef 2
SELECT AVG(PrijsPerEenheid) FROM tblProducten
-- Oef 4
SELECT COUNT(*) FROM tblProducten WHERE Categorienummer = :Categorienummer
-- Oef 6
SELECT * FROM tblCategorieen INNER JOIN tblProducten ON tblCategorieen.Categorienummer = tblProducten.Categorienummer WHERE tblCategorieen.Categorienummer = 6 OR tblCategorieen.Categorienummer = 8
