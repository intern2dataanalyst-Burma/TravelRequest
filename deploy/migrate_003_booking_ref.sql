-- Travel Accommodation Portal - migration 003
-- Adds a short, auto-incrementing Booking Reference (#10001, #10002, ...) shown in all email subjects and bodies.
-- SQL Server assigns the number on INSERT, so two requests can never get the same reference.
-- Existing rows are numbered from 10001 when the column is added.
-- Safe to re-run.

USE MingalabaDB;
GO

IF COL_LENGTH('dbo.TravelRequests', 'BookingRefID') IS NULL
    ALTER TABLE dbo.TravelRequests ADD BookingRefID int IDENTITY(10001, 1) NOT NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes
               WHERE object_id = OBJECT_ID('dbo.TravelRequests') AND name = 'UX_TravelRequests_BookingRefID')
    CREATE UNIQUE INDEX UX_TravelRequests_BookingRefID ON dbo.TravelRequests (BookingRefID);
GO
