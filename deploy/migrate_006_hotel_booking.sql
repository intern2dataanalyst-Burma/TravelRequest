-- Travel Accommodation Portal - migration 006 (hotel booking workflow)
--   TravelRequests.PreferredHotel : optional hotel the employee would like (request form, section 3 "Occupancy")
--   dbo.TravelBookingVouchers     : the final booking voucher (PDF) the Admin Team uploads after the manager approves.
--                                   Kept out of TravelRequests so normal page loads never read the file.
-- The Admin-only fields (type of occupancy, refund status, remarks) live inside the existing BookingOptions JSON,
-- so they need no new columns. New status values ("Pending Final Booking", "Booked") fit the nvarchar(30) Status.
-- Safe to re-run; existing rows are untouched.

USE MingalabaDB;
GO

IF COL_LENGTH('dbo.TravelRequests', 'PreferredHotel') IS NULL
    ALTER TABLE dbo.TravelRequests ADD PreferredHotel nvarchar(200) NULL;
GO

IF OBJECT_ID('dbo.TravelBookingVouchers', 'U') IS NULL
    CREATE TABLE dbo.TravelBookingVouchers (
        RequestId  uniqueidentifier NOT NULL
            CONSTRAINT PK_TravelBookingVouchers PRIMARY KEY
            CONSTRAINT FK_TravelBookingVouchers_Request REFERENCES dbo.TravelRequests (RequestId),
        FileName   nvarchar(255)    NOT NULL,
        Content    varbinary(max)   NOT NULL,
        UploadedAt datetime2        NOT NULL CONSTRAINT DF_TravelBookingVouchers_UploadedAt DEFAULT SYSUTCDATETIME()
    );
GO
