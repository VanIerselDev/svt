SELECT "Familienaam","Voornaam","Werknemer-ID","Geboortedatum" FROM tblWerknemers WHERE MONTH(Geboortedatum) = 1 OR MONTH(Geboortedatum) = 12

SELECT Productnummer,NederlandseNaam,Voorraad,Bestelpunt,InBestelling FROM tblProducten WHERE InBestelling = False AND Voorraad < Bestelpunt

SELECT Productnummer,NederlandseNaam,HoeveelheidPerEenheid FROM tblProducten WHERE HoeveelheidPerEenheid LIKE '*lik*'

SELECT Klantnummer, Naam, Straat, Postcode, Gemeente FROM tblKlanten

SELECT Naam, BTWnr FROM tblKlanten WHERE BTWnr IS NULL ORDER BY Naam

SELECT * FROM tblKlanten WHERE ((Type = 'W') OR (Type = 'R')) And ((Gemeente = 'Aarschot') OR (Gemeente = 'Leuven'))

SELECT * FROM tblWerknemers WHERE (Geslacht = '2') AND (Gemeente = 'Leuven')

SELECT * FROM tblKlanten WHERE Naam LIKE 'Van*'

SELECT Voornaam, Familienaam, Postcode, Gemeente FROM tblWerknemers

SELECT NederlandseNaam, PrijsPerEenheid, Categorienummer FROM tblProducten ORDER BY NederlandseNaam

SELECT Productnaam, Voorraad AS Spoedig FROM tblProducten

SELECT * FROM tblProducten WHERE Categorienummer =:Geeft_Catergorie
