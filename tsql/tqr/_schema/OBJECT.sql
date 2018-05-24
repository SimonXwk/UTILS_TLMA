-- auto-generated definition
create table OBJECT
(
  ADMITNAME varchar(50) not null
    constraint PK_Object
    primary key,
  STAGEID   float,
  OWNER     varchar(50),
  ADMITTYPE varchar(20),
  FILENAME  varchar(50),
  FILEPATH  varchar(255),
  TITLE     varchar(80),
  SUBJECT   varchar(50),
  AUTHOR    varchar(50),
  KEYWORDS  varchar(50),
  COMMENTS  varchar(255),
  CREATED   datetime,
  REMOVED   datetime,
  MODIFIED  datetime
)
go

create index I_Object_Modified
  on OBJECT (MODIFIED)
go

create index I_Object_Removed
  on OBJECT (REMOVED)
go

create index I_Object_StageID
  on OBJECT (STAGEID)
go

create index I_Object_AdmitType
  on OBJECT (ADMITTYPE)
go

