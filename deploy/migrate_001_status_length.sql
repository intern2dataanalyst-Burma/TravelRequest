-- Travel Accommodation Portal - migration 001
-- Widens TravelRequests.Status so it can hold 'Pending Manager Approval' (24 chars; the column was nvarchar(20)).
-- Workflow statuses: Pending Booking -> Pending Manager Approval -> Approved / Rejected (or MailFailed).
-- Safe to re-run. Widening keeps existing data; the DEFAULT constraint does not block a length change.

USE MingalabaDB;
GO

IF COL_LENGTH('dbo.TravelRequests', 'Status') < 60   -- COL_LENGTH is in bytes: nvarchar(30) = 60
BEGIN
    ALTER TABLE dbo.TravelRequests ALTER COLUMN Status nvarchar(30) NOT NULL;
    PRINT 'TravelRequests.Status widened to nvarchar(30)';
END
ELSE
    PRINT 'TravelRequests.Status already nvarchar(30) or wider - nothing to do';
GO
