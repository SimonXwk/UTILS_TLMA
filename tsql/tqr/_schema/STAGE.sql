-- auto-generated definition
create table STAGE
(
  STAGEID     float         not null
    constraint PK_Stage
    primary key,
  STAGE       varchar(50),
  STAGEALIAS  varchar(50),
  DESCRIPTION varchar(50),
  HISTORY     bit default 0 not null
)
go

create index IX_STAGE_STAGE
  on STAGE (STAGE)
go
