-- Travel Accommodation Portal - migration 004
-- Lets the Travel Desk offer several accommodation options; the manager approves one of them.
--   BookingOptions : JSON array of every option, e.g. [{"area":"Indiranagar","property":"The Leela Palace","cost":"12500.00"}]
--   SelectedOption : 1-based number of the option the manager approved
-- BookingArea / PropertyName / Cost keep holding the final (approved) booking.
-- Safe to re-run; existing rows are untouched (older single bookings are read as "Option 1" by the app).

USE MingalabaDB;
GO

IF COL_LENGTH('dbo.TravelRequests', 'BookingOptions') IS NULL
    ALTER TABLE dbo.TravelRequests ADD BookingOptions nvarchar(max) NULL;
GO
IF COL_LENGTH('dbo.TravelRequests', 'SelectedOption') IS NULL
    ALTER TABLE dbo.TravelRequests ADD SelectedOption tinyint NULL;
GO

-- Reject anything that is not valid JSON
IF NOT EXISTS (SELECT 1 FROM sys.check_constraints WHERE name = 'CK_TravelRequests_BookingOptions_IsJson')
    ALTER TABLE dbo.TravelRequests ADD CONSTRAINT CK_TravelRequests_BookingOptions_IsJson
        CHECK (BookingOptions IS NULL OR ISJSON(BookingOptions) = 1);
GO
