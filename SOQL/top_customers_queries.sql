-- Query 1
SELECT Id, Name
FROM Account

-- Query 2
SELECT Id, Name, Amount
FROM Opportunity
ORDER BY Amount DESC

-- Query 3
SELECT Name, StageName, Amount
FROM Opportunity
WHERE IsClosed = false

-- Query 4
SELECT Name, CloseDate
FROM Opportunity
ORDER BY CloseDate DESC

-- Query 5
SELECT Name, Industry
FROM Account
ORDER BY Name
