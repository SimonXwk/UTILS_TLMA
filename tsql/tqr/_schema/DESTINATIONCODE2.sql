-- auto-generated definition
CREATE TABLE DESTINATIONCODE2
(
  DESTINATIONCODE2       VARCHAR(25)   NOT NULL,
  DESTINATIONTYPE        VARCHAR(50),
  DESTINATIONDESCRIPTION VARCHAR(255),
  DESTINATIONNOTES       VARCHAR(1000),
  LEDGERCODE             VARCHAR(50),
  RECORDOWNER            VARCHAR(25),
  MODIFIED               DATETIME,
  MODIFIEDBY             VARCHAR(50),
  CREATED                DATETIME,
  CREATEDBY              VARCHAR(50),
  EXCLUDEFROMDROPDOWN    SMALLINT DEFAULT 0,
  ARCHIVE                FLOAT    DEFAULT 0,
  OWNERCODE              INT DEFAULT 0 NOT NULL,
  CONSTRAINT PK_DESTINATIONCODE2
  PRIMARY KEY (DESTINATIONCODE2, OWNERCODE)
)
GO
CREATE INDEX IX_DESTINATION_DESTINATIONTYPE
  ON DESTINATIONCODE2 (DESTINATIONTYPE)
GO
CREATE INDEX IX_DESTINATION_DESTINATIONDESC
  ON DESTINATIONCODE2 (DESTINATIONDESCRIPTION)
GO
CREATE INDEX IX_DESTINATION_LEDGERCODE
  ON DESTINATIONCODE2 (LEDGERCODE)
GO
CREATE INDEX IX_DESTINATION_RECORDOWNER
  ON DESTINATIONCODE2 (RECORDOWNER)
GO
CREATE INDEX IX_DESTINATION_MODIFIED
  ON DESTINATIONCODE2 (MODIFIED)
GO
CREATE INDEX IX_DESTINATION_MODIFIEDBY
  ON DESTINATIONCODE2 (MODIFIEDBY)
GO
CREATE INDEX IX_DESTINATION_CREATED
  ON DESTINATIONCODE2 (CREATED)
GO
CREATE INDEX IX_DESTINATION_CREATEDBY
  ON DESTINATIONCODE2 (CREATEDBY)
GO
CREATE INDEX IX_DESTINATION_EXCLUDEFROM
  ON DESTINATIONCODE2 (EXCLUDEFROMDROPDOWN)
GO
CREATE INDEX IX_DESTINATION_ARCHIVE
  ON DESTINATIONCODE2 (ARCHIVE)
GO
