create database Contacts
Go
use Contacts
Create Table dbo.Customers(
                CustomerID int not null identity primary key,
                CustomerNo varchar(20),
                CustomerName varchar(100) not null
)

Create table dbo.Items(
                ItemID int not null identity primary key,
                ItemNo varchar(20),
                ItemName varchar(100) not null
)

Create table dbo.CustomerCodes(
                CustomerCodeID int not null identity primary key,
                CustomerID int,
                ItemID int,
                CustomerCode varchar(50),
                [Description] varchar(8000)


)

Create table dbo.CustomerCodeDetails(
                CustomerCodeDetailsID int not null identity primary key,
                CustomerCodeID int not null,
                QtyPriced float,
                UnitPrice float,
                StartDate datetime,
                FinishDate datetime,
                DiscountPct float
)
Go

INSERT INTO dbo.customers
values('C001', 'Bugs Bunny'),('C002', 'Daffy Duck'),('C003', 'Wiley Coyote')

INSERT INTO dbo.Items
                values('I001', 'Giant Kite'), ('I002','Hi-Speed Tonic'), ('I003', 'Rocket Powered Unicycle')
Go
INSERT INTO dbo.CustomerCodes(customerid, itemid, customercode)
VALUES(1,1,'Albuquerque Flyer')
Go
