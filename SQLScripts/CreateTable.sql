CREATE TABLE [dbo].[StudentScore]
(
	[Id] INT NOT NULL PRIMARY KEY IDENTITY,
	[Name] NVARCHAR(50) NOT NULL,
	[Class] INT NOT NULL,
	[Score] INT NULL, 
    [Sex] NCHAR(10) NOT NULL
)

insert into studentscore ([Name],[Class],[Score],[Sex])
values ('Nancy',2,89,'Female')
insert into studentscore ([Name],[Class],[Score],[Sex])
values ('Leo',2,88,'Male')



