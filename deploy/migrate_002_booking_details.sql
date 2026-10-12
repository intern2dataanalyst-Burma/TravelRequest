-- Travel Accommodation Portal - migration 002
-- Adds the booking details the Travel Desk enters on /travel-desk-confirm before the request goes to the manager.
-- Nullable: they stay empty until the Travel Desk submits the booking form.
-- Safe to re-run: each column is added only if it is missing.

USE MingalabaDB;
GO

IF COL_LENGTH('dbo.TravelRequests', 'BookingArea') IS NULL
    ALTER TABLE dbo.TravelRequests ADD BookingArea nvarchar(150) NULL;      -- Place/Area, e.g. Indiranagar
GO
IF COL_LENGTH('dbo.TravelRequests', 'PropertyName') IS NULL
    ALTER TABLE dbo.TravelRequests ADD PropertyName nvarchar(200) NULL;     -- e.g. The Leela Palace
GO
IF COL_LENGTH('dbo.TravelRequests', 'Cost') IS NULL
    ALTER TABLE dbo.TravelRequests ADD Cost decimal(12, 2) NULL;            -- final booked price, INR
GO
