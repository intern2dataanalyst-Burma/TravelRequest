-- Travel Accommodation Portal - migration 005
-- Gives the portal its OWN employee -> manager routing table, maintained by HR in User Management.
-- dbo.checklist_empmanager_mapping is shared with the checklist apps, so the portal no longer writes to it.
--
-- The new table is seeded once by READING the shared table (only when the new table is empty), so approval
-- routing keeps working from day one. After that the two tables are independent: HR edits here; changes made
-- by the checklist apps to their table do not flow into this one.
-- Safe to re-run.

USE MingalabaDB;
GO

IF OBJECT_ID('dbo.TravelApproverMapping', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.TravelApproverMapping (
        id           int IDENTITY(1, 1) NOT NULL CONSTRAINT PK_TravelApproverMapping PRIMARY KEY,
        EmpCode      nvarchar(50)  NOT NULL,
        EmpEmail     nvarchar(256) NULL,
        ManagerCode  nvarchar(50)  NULL,     -- NULL = no manager (HOD): requests skip manager approval
        ManagerEmail nvarchar(256) NULL,
        CreatedAt    datetime2     NOT NULL CONSTRAINT DF_TravelApproverMapping_CreatedAt DEFAULT SYSUTCDATETIME(),
        UpdatedAt    datetime2     NOT NULL CONSTRAINT DF_TravelApproverMapping_UpdatedAt DEFAULT SYSUTCDATETIME(),
        CONSTRAINT UX_TravelApproverMapping_EmpCode UNIQUE (EmpCode)
    );
    PRINT 'Created dbo.TravelApproverMapping';
END
GO

-- One-time seed from the shared checklist table (read-only on that table)
IF NOT EXISTS (SELECT 1 FROM dbo.TravelApproverMapping)
BEGIN
    INSERT INTO dbo.TravelApproverMapping (EmpCode, EmpEmail, ManagerCode, ManagerEmail)
    SELECT LTRIM(RTRIM(src.EmpCode)),
           NULLIF(LTRIM(RTRIM(src.EmpEmail)), ''),
           NULLIF(LTRIM(RTRIM(src.ManagerCode)), ''),
           NULLIF(LTRIM(RTRIM(src.ManagerEmail)), '')
    FROM (
        -- latest row per employee, same rule the portal used when reading the shared table
        SELECT *, ROW_NUMBER() OVER (PARTITION BY LTRIM(RTRIM(EmpCode)) ORDER BY id DESC) AS rn
        FROM dbo.checklist_empmanager_mapping
        WHERE NULLIF(LTRIM(RTRIM(EmpCode)), '') IS NOT NULL
    ) src
    WHERE src.rn = 1;
    PRINT CONCAT('Seeded ', @@ROWCOUNT, ' rows from dbo.checklist_empmanager_mapping');
END
GO
