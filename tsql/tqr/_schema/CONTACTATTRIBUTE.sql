-- auto-generated definition
CREATE TABLE CONTACTATTRIBUTE
(
  SERIALNUMBER              VARCHAR(50) NOT NULL
    CONSTRAINT PK_CONTACTATTRIBUTE
    PRIMARY KEY,
  HTMLNOTES                 VARCHAR(MAX),
  COUNTAMOUNT               FLOAT,
  SUMAMOUNT                 FLOAT,
  MAXAMOUNT                 FLOAT,
  AVGAMOUNT                 FLOAT,
  MAXDATE                   DATETIME,
  MINDATE                   DATETIME,
  LEADCONTACT               SMALLINT DEFAULT 0,
  JOINTSALUTATION           VARCHAR(255),
  JOINTLETTERSALUTATION     VARCHAR(255),
  DEFAULTTITLE              VARCHAR(30),
  DEFAULTFIRSTNAME          VARCHAR(50),
  DEFAULTKEYNAME            VARCHAR(100),
  DEFAULTADDRESSLINE1       VARCHAR(255),
  DEFAULTADDRESSLINE2       VARCHAR(100),
  DEFAULTADDRESSLINE3       VARCHAR(100),
  DEFAULTADDRESSLINE4       VARCHAR(100),
  DEFAULTPOSTCODE           VARCHAR(10),
  DEFAULTCOUNTRY            VARCHAR(50),
  DEFAULTADDRESSCODE1       VARCHAR(50),
  DEFAULTADDRESSCODE2       VARCHAR(50),
  EXTERNALREF               VARCHAR(50),
  EXTERNALREFTYPE           VARCHAR(50),
  EXTERNALREFDATE           DATETIME,
  SUBADDRESS                VARCHAR(100),
  PLAINTEXTEMAIL            SMALLINT DEFAULT 0,
  PICTUREFILE2              VARCHAR(50),
  DONOTCALL                 SMALLINT DEFAULT 0,
  SECONDCONTACT             VARCHAR(50),
  CANVASSEE                 SMALLINT DEFAULT 0,
  DONOTPUBLISH              SMALLINT DEFAULT 0,
  VOLUNTEER                 SMALLINT DEFAULT 0,
  VOLSOURCE                 VARCHAR(30),
  VOLSPECIALREQ             VARCHAR(255),
  EMERGENCYNAME             VARCHAR(200),
  EMERGENCYNUMBER           VARCHAR(30),
  VOLUNTEERMANAGER          VARCHAR(200),
  DONOTVOLUNTEER            SMALLINT DEFAULT 0,
  CRMPRIMARYMANAGER         VARCHAR(50),
  CRMSECONDARYMANAGER       VARCHAR(50),
  CRMPLAN                   VARCHAR(20),
  CRMREFERMANAGER           SMALLINT DEFAULT 0,
  CRMNOTE                   VARCHAR(1500),
  CCEMAILADDRESS            VARCHAR(255),
  CCNAME                    VARCHAR(255),
  SOURCE2                   VARCHAR(50),
  DEFAULTSOURCE2            VARCHAR(50),
  MEDICALREF1               VARCHAR(50),
  MEDICALREF2               VARCHAR(50),
  MEDICALREF2DATE           DATETIME,
  DEFAULTRECEIPTREQUIRED    VARCHAR(3),
  DEFAULTRECEIPTSUMMARY     VARCHAR(3),
  SOURCEDATE                DATETIME,
  TEMPMAINTITLE             VARCHAR(30),
  TEMPMAINFIRSTNAME         VARCHAR(50),
  TEMPMAINKEYNAME           VARCHAR(100),
  TEMPSECONDARYTITLE        VARCHAR(30),
  TEMPSECONDARYFIRSTNAME    VARCHAR(50),
  TEMPSECONDARYKEYNAME      VARCHAR(100),
  DEFAULTPRODUCTLEVEL       VARCHAR(255),
  NICKNAME                  VARCHAR(255),
  CONFIGTEXT1               VARCHAR(255),
  CONFIGTEXT2               VARCHAR(255),
  CONFIGTEXT3               VARCHAR(255),
  CONFIGTEXT4               VARCHAR(255),
  CONFIGTEXT5               VARCHAR(255),
  IC                        VARCHAR(50),
  ICTYPE                    VARCHAR(255),
  ALTTITLE                  NVARCHAR(500),
  ALTFIRSTNAME              NVARCHAR(500),
  ALTKEYNAME                NVARCHAR(500),
  ALTADDRESS                NVARCHAR(4000),
  DONOTEMAIL                SMALLINT DEFAULT 0,
  SUMAMOUNTSOFT             FLOAT,
  COUNTAMOUNTSOFT           INT,
  SUMAMOUNTHARDNSOFT        FLOAT,
  COUNTAMOUNTHARDNSOFT      INT,
  REFERRALSOURCE            VARCHAR(50),
  QLDHEALTHREGION           VARCHAR(50),
  QLDREGION                 VARCHAR(100),
  ISLANDER                  SMALLINT,
  WRITTENPRIVACYCONSENTHELD SMALLINT,
  VERBALPRIVACYCONSENTGIVEN SMALLINT,
  NOTES                     VARCHAR(MAX)
)
GO
CREATE UNIQUE INDEX UI_CONTACTATTRIBUTE
  ON CONTACTATTRIBUTE (SERIALNUMBER)
GO
CREATE INDEX IX_CONTACTATTRIBUTE_NICKNAME
  ON CONTACTATTRIBUTE (NICKNAME)
GO
CREATE INDEX IX_CONTACTATTRIBUTE_COUNTAMOUNT
  ON CONTACTATTRIBUTE (COUNTAMOUNT)
GO
CREATE INDEX IX_CONTACTATTRIBUTE_SUMAMOUNT
  ON CONTACTATTRIBUTE (SUMAMOUNT)
GO
CREATE INDEX IX_CONTACTATTRIBUTE_MAXAMOUNT
  ON CONTACTATTRIBUTE (MAXAMOUNT)
GO
CREATE INDEX IX_CONTACTATTRIBUTE_MAXDATE
  ON CONTACTATTRIBUTE (MAXDATE)
GO
CREATE INDEX IX_CONTACTATTRIBUTE_MINDATE
  ON CONTACTATTRIBUTE (MINDATE)
GO
CREATE INDEX IX_CONTACTATTRIBUTE_JOINTSALUTATION
  ON CONTACTATTRIBUTE (JOINTSALUTATION)
GO
CREATE INDEX IX_CONTACTATTRIBUTE_LETTERSALUTATION
  ON CONTACTATTRIBUTE (JOINTLETTERSALUTATION)
GO
CREATE INDEX IX_CONTACTATTRIBUTE_DONOTPUBLISH
  ON CONTACTATTRIBUTE (DONOTPUBLISH)
GO
CREATE INDEX IX_CONTACTATT_DEFFIRSTNAME
  ON CONTACTATTRIBUTE (DEFAULTFIRSTNAME)
GO
CREATE INDEX IX_CONTACTATT_DEFKEYNAME
  ON CONTACTATTRIBUTE (DEFAULTKEYNAME)
GO
CREATE INDEX IX_CONTACTATT_DEFADDRESS1
  ON CONTACTATTRIBUTE (DEFAULTADDRESSLINE1)
GO
CREATE INDEX IX_CONTACTATT_DEFADDRESS3
  ON CONTACTATTRIBUTE (DEFAULTADDRESSLINE3)
GO
CREATE INDEX IX_CONTACTATT_DEFADDRESS4
  ON CONTACTATTRIBUTE (DEFAULTADDRESSLINE4)
GO
CREATE INDEX IX_CONTACTATT_DEFPOSTCODE
  ON CONTACTATTRIBUTE (DEFAULTPOSTCODE)
GO
CREATE INDEX IX_CONTACTATTRIBUTE_EXTERNALREF
  ON CONTACTATTRIBUTE (EXTERNALREF)
GO
CREATE INDEX IX_CONTACTATTRIBUTE_EXTERNALREFTYPE
  ON CONTACTATTRIBUTE (EXTERNALREFTYPE)
GO
CREATE INDEX IX_CONTACTATTRIBUTE_EXTERNALREFDATE
  ON CONTACTATTRIBUTE (EXTERNALREFDATE)
GO
CREATE INDEX I_CONTACTATTR_PLAINTEXTEMAIL
  ON CONTACTATTRIBUTE (PLAINTEXTEMAIL)
GO
CREATE INDEX IX_CONTACTATTRIBUTE_DONOTCALL
  ON CONTACTATTRIBUTE (DONOTCALL)
GO
CREATE INDEX IX_CONTACTATTRIBUTE_SECONDCONTACT
  ON CONTACTATTRIBUTE (SECONDCONTACT)
GO
CREATE INDEX IX_CONTACTATTRIBUTE_CANVASSEE
  ON CONTACTATTRIBUTE (CANVASSEE)
GO
CREATE INDEX IX_CONTACTATTRIBUTE_VOLUNTEER
  ON CONTACTATTRIBUTE (VOLUNTEER)
GO
CREATE INDEX IX_CONTACTATTRIBUTE_VOLSOURCE
  ON CONTACTATTRIBUTE (VOLSOURCE)
GO
CREATE INDEX I_CONTACTATTR_CRMPRIMARY
  ON CONTACTATTRIBUTE (CRMPRIMARYMANAGER)
GO
CREATE INDEX I_CONTACTATTR_CRMSECOND
  ON CONTACTATTRIBUTE (CRMSECONDARYMANAGER)
GO
CREATE INDEX I_CONTACTATTR_CRMREFER
  ON CONTACTATTRIBUTE (CRMREFERMANAGER)
GO
CREATE INDEX IX_CONTACTATTRIBUTE_DEFRCPTSUMMARY
  ON CONTACTATTRIBUTE (DEFAULTRECEIPTSUMMARY)
GO
CREATE INDEX IX_CONTACTATTRIBUTE_DEFRCPTREQUIRED
  ON CONTACTATTRIBUTE (DEFAULTRECEIPTREQUIRED)
GO
CREATE INDEX IX_CONTACTATTRIBUTE_IC
  ON CONTACTATTRIBUTE (IC)
GO
