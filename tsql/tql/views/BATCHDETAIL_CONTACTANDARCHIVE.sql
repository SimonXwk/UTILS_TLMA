CREATE VIEW [dbo].[BATCHDETAIL_CONTACTANDARCHIVE] AS
SELECT dbo.BATCHITEM.OWNERCODE,
       dbo.BATCHITEM.ADMITNAME,
       dbo.BATCHITEM.RECEIPTNO AS AUTOID,
       dbo.BATCHITEM.SERIALNUMBER,
       dbo.BATCHITEMPLEDGE.PLEDGEID,
       dbo.BATCHITEMPLEDGE.PLEDGELINENO,
       dbo.BATCHITEM.DATEOFPAYMENT,
       dbo.BATCHITEM.DATEPOSTED,
       dbo.BATCHITEM.PAYMENTTYPE,
       dbo.BATCHITEMSPLIT.PAYMENTAMOUNT,
       dbo.BATCHITEMSPLIT.TAXCLAIMABLE,
       dbo.BATCHITEMPLEDGE.RECEIPTREQUIRED,
       dbo.BATCHITEM.MANUALRECEIPTNO,
       dbo.BATCHITEM.TREATASANON,
       dbo.BATCHITEMSPLIT.DESTINATIONCODE,
       dbo.BATCHITEMSPLIT.DESTINATIONCODE2,
       dbo.BATCHITEMSPLIT.SOURCECODE,
       dbo.BATCHITEMSPLIT.SOURCECODE2,
       dbo.BATCHITEMSPLIT.CAMPAIGNYEAR,
       dbo.BATCHITEM.NARRATIVE,
       dbo.BATCHITEM.INCLUDEINTHANKYOU,
       dbo.BATCHITEM.REVERSED,
       dbo.BATCHITEMSPLIT.EXTERNALREF,
       dbo.BATCHITEMSPLIT.EXTERNALREFTYPE,
       dbo.BATCHITEMPLEDGE.RECEIPTLETTERID,
       dbo.BATCHITEM.ACCOUNTNAME,
       dbo.BATCHITEM.SORTCODE,
       dbo.BATCHITEM.ACCOUNTNUMBER,
       dbo.BATCHITEM.ROLLNUMBER,
       dbo.BATCHITEM.TEXT1,
       dbo.BATCHITEM.TEXT2,
       dbo.BATCHITEM.BANKNAME,
       dbo.BATCHITEM.BANKADDRESS,
       dbo.BATCHITEM.BANKPOSTCODE,
       dbo.BATCHITEM.CREDITCARDTYPE,
       dbo.BATCHITEM.CREDITCARDNUMBER,
       dbo.BATCHITEM.CREDITCARDMASKED,
       dbo.BATCHITEM.CREDITCARDHOLDER,
       dbo.BATCHITEM.CREDITCARDEXPIRY,
       dbo.BATCHITEM.CREDITCARDISSUENUMBER,
       dbo.BATCHITEM.CREDITCARDAUTHCODE,
       dbo.BATCHITEM.CREDITCARDMETHOD,
       dbo.BATCHITEM.CHEQUENUMBER,
       dbo.BATCHITEM.CREATED,
       dbo.BATCHITEM.CREATEDBY,
       dbo.BATCHITEM.MODIFIED,
       dbo.BATCHITEM.MODIFIEDBY,
       dbo.BATCHITEMSPLIT.NEW,
       dbo.BATCHITEMSPLIT.EXISTING,
       dbo.BATCHITEMSPLIT.RECOVERED,
       dbo.BATCHITEM.NOTES,
       dbo.BATCHITEMSPLIT.PAYMENTAMOUNTNETT,
       dbo.BATCHITEMSPLIT.GSTAMOUNT,
       dbo.BATCHITEM.RESULT,
       dbo.BATCHITEM.SUCCESSFAILURE,
       dbo.BATCHITEMPLEDGE.RECEIPTSUMMARY,
       dbo.BATCHITEMPLEDGE.RCPTSUMMARYLETTERID,
       dbo.BATCHITEM.APPLYJOINTSALUTATION,
       CASE
           WHEN dbo.CONTACT.SERIALNUMBER IS NOT NULL THEN dbo.CONTACT.CONTACTTYPE
           ELSE dbo.CONTACTSARCHIVE.CONTACTTYPE
       END AS TheContactType,
       CASE
           WHEN dbo.CONTACT.SERIALNUMBER IS NOT NULL THEN dbo.CONTACT.PRIMARYCATEGORY
           ELSE dbo.CONTACTSARCHIVE.PRIMARYCATEGORY
       END AS PrimaryCategory,
       COALESCE(BATCHITEM.TITLE,dbo.CONTACT.TITLE,dbo.CONTACTSARCHIVE.TITLE) AS TheTitle,
       COALESCE(BATCHITEM.FIRSTNAME,dbo.CONTACT.FIRSTNAME,dbo.CONTACTSARCHIVE.FIRSTNAME) AS TheFirstName,
       COALESCE(BATCHITEM.KEYNAME,dbo.CONTACT.KEYNAME,dbo.CONTACTSARCHIVE.KEYNAME + ' ( Archived )') AS TheKeyName,
       COALESCE(dbo.CONTACT.TITLE,dbo.CONTACTSARCHIVE.TITLE) AS TheContactTitle,
       COALESCE(dbo.CONTACT.FIRSTNAME,dbo.CONTACTSARCHIVE.FIRSTNAME) AS TheContactFirstName,
       COALESCE(dbo.CONTACT.KEYNAME,dbo.CONTACTSARCHIVE.KEYNAME + ' ( Archived )') AS TheContactKeyName,
       COALESCE(BATCHITEM.ADDRESSLINE1,CASE WHEN dbo.CONTACT.SERIALNUMBER IS NOT NULL THEN CASE WHEN ISNULL(dbo.CONTACT.DEFAULTADDRESS, '') = '' THEN dbo.CONTACT.ADDRESSLINE1 ELSE dbo.CONTACTATTRIBUTE.DEFAULTADDRESSLINE1 END ELSE CASE WHEN ISNULL(dbo.CONTACTSARCHIVE.DEFAULTADDRESS, '') = '' THEN dbo.CONTACTSARCHIVE.ADDRESSLINE1 ELSE dbo.CONTACTATTRIBUTESARCHIVE.DEFAULTADDRESSLINE1 END END) AS TheAddress1,
       CASE
           WHEN ISNULL(BATCHITEM.ADDRESSLINE1,'')<>'' THEN BATCHITEM.ADDRESSLINE2
           ELSE CASE
                    WHEN dbo.CONTACT.SERIALNUMBER IS NOT NULL THEN CASE
                                                                       WHEN ISNULL(dbo.CONTACT.DEFAULTADDRESS, '') = '' THEN dbo.CONTACT.ADDRESSLINE2
                                                                       ELSE dbo.CONTACTATTRIBUTE.DEFAULTADDRESSLINE2
                                                                   END
                    ELSE CASE
                             WHEN ISNULL(dbo.CONTACTSARCHIVE.DEFAULTADDRESS, '') = '' THEN dbo.CONTACTSARCHIVE.ADDRESSLINE2
                             ELSE dbo.CONTACTATTRIBUTESARCHIVE.DEFAULTADDRESSLINE2
                         END
                END
       END AS TheAddress2,
       CASE
           WHEN ISNULL(BATCHITEM.ADDRESSLINE1,'')<>'' THEN BATCHITEM.ADDRESSLINE3
           ELSE CASE
                    WHEN dbo.CONTACT.SERIALNUMBER IS NOT NULL THEN CASE
                                                                       WHEN ISNULL(dbo.CONTACT.DEFAULTADDRESS, '') = '' THEN dbo.CONTACT.ADDRESSLINE3
                                                                       ELSE dbo.CONTACTATTRIBUTE.DEFAULTADDRESSLINE3
                                                                   END
                    ELSE CASE
                             WHEN ISNULL(dbo.CONTACTSARCHIVE.DEFAULTADDRESS, '') = '' THEN dbo.CONTACTSARCHIVE.ADDRESSLINE3
                             ELSE dbo.CONTACTATTRIBUTESARCHIVE.DEFAULTADDRESSLINE3
                         END
                END
       END AS TheSuburb,
       CASE
           WHEN ISNULL(BATCHITEM.ADDRESSLINE1,'')<>'' THEN BATCHITEM.ADDRESSLINE4
           ELSE CASE
                    WHEN dbo.CONTACT.SERIALNUMBER IS NOT NULL THEN CASE
                                                                       WHEN ISNULL(dbo.CONTACT.DEFAULTADDRESS, '') = '' THEN dbo.CONTACT.ADDRESSLINE4
                                                                       ELSE dbo.CONTACTATTRIBUTE.DEFAULTADDRESSLINE4
                                                                   END
                    ELSE CASE
                             WHEN ISNULL(dbo.CONTACTSARCHIVE.DEFAULTADDRESS, '') = '' THEN dbo.CONTACTSARCHIVE.ADDRESSLINE4
                             ELSE dbo.CONTACTATTRIBUTESARCHIVE.DEFAULTADDRESSLINE4
                         END
                END
       END AS TheState,
       CASE
           WHEN ISNULL(BATCHITEM.ADDRESSLINE1,'')<>'' THEN BATCHITEM.POSTCODE
           ELSE CASE
                    WHEN dbo.CONTACT.SERIALNUMBER IS NOT NULL THEN CASE
                                                                       WHEN ISNULL(dbo.CONTACT.DEFAULTADDRESS, '') = '' THEN dbo.CONTACT.POSTCODE
                                                                       ELSE dbo.CONTACTATTRIBUTE.DEFAULTPOSTCODE
                                                                   END
                    ELSE CASE
                             WHEN ISNULL(dbo.CONTACTSARCHIVE.DEFAULTADDRESS, '') = '' THEN dbo.CONTACTSARCHIVE.POSTCODE
                             ELSE dbo.CONTACTATTRIBUTESARCHIVE.DEFAULTPOSTCODE
                         END
                END
       END AS ThePostcode,
       COALESCE(BATCHITEM.DAYTELEPHONE,dbo.CONTACT.DAYTELEPHONE,dbo.CONTACTSARCHIVE.DAYTELEPHONE) AS TheTelephone,
       COALESCE(BATCHITEM.FAXNUMBER,dbo.CONTACT.FAXNUMBER,dbo.CONTACTSARCHIVE.FAXNUMBER) AS TheFaxNumber,
       COALESCE(BATCHITEM.EMAILADDRESS,dbo.CONTACT.EMAILADDRESS,dbo.CONTACTSARCHIVE.EMAILADDRESS) AS TheEmail,
       COALESCE(BATCHITEM.EVENINGTELEPHONE,dbo.CONTACT.EVENINGTELEPHONE,dbo.CONTACTSARCHIVE.EVENINGTELEPHONE) AS TheEveningTelephone,
       COALESCE(BATCHITEM.MOBILENUMBER,dbo.CONTACT.MOBILENUMBER,dbo.CONTACTSARCHIVE.MOBILENUMBER) AS TheMobile,
       CASE
           WHEN CONTACT.ANONYMOUS=-1
                OR dbo.BATCHITEM.TREATASANON = - 1 THEN 'Yes'
           ELSE ''
       END AS Anonymous,
       CASE
           WHEN dbo.CONTACT.SERIALNUMBER IS NOT NULL THEN dbo.CONTACT.SALESREGION
           ELSE dbo.CONTACTSARCHIVE.SALESREGION
       END AS SALESREGION,
       CASE
           WHEN dbo.CONTACT.SERIALNUMBER IS NOT NULL THEN dbo.CONTACT.SOURCE
           ELSE dbo.CONTACTSARCHIVE.SOURCE
       END AS CONTACTSOURCE,
       CASE
           WHEN dbo.CONTACT.SERIALNUMBER IS NOT NULL THEN dbo.CONTACT.MAINCONTACT
           ELSE dbo.CONTACTSARCHIVE.MAINCONTACT
       END AS MAINCONTACT,
       dbo.OBJECT.STAGEID,
       dbo.BATCHHEADER.ACCOUNTREFERENCE,
       dbo.BATCHITEMSPLIT.LINEID,
       dbo.BATCHITEMSPLIT.SPLITID,
       dbo.BATCHITEMSPLIT.REMAININGAMOUNT,
       dbo.BATCHITEMSPLIT.REMAININGAMOUNTNETT,
       dbo.BATCHITEMSPLIT.REMAININGGSTAMOUNT,
       dbo.BATCHITEMSPLIT.RANK
FROM
	dbo.CONTACT 
	LEFT OUTER JOIN dbo.CONTACTATTRIBUTE 
		ON dbo.CONTACT.SERIALNUMBER = dbo.CONTACTATTRIBUTE.SERIALNUMBER
	RIGHT OUTER JOIN (dbo.BATCHITEM
		INNER JOIN dbo.BATCHITEMPLEDGE 
		ON dbo.BATCHITEM.ADMITNAME = dbo.BATCHITEMPLEDGE.ADMITNAME
			AND dbo.BATCHITEM.RECEIPTNO = dbo.BATCHITEMPLEDGE.RECEIPTNO
			AND dbo.BATCHITEM.SERIALNUMBER = dbo.BATCHITEMPLEDGE.SERIALNUMBER
		INNER JOIN dbo.BATCHITEMSPLIT 
		ON dbo.BATCHITEMPLEDGE.ADMITNAME = dbo.BATCHITEMSPLIT.ADMITNAME
			AND dbo.BATCHITEMPLEDGE.RECEIPTNO = dbo.BATCHITEMSPLIT.RECEIPTNO
			AND dbo.BATCHITEMPLEDGE.SERIALNUMBER = dbo.BATCHITEMSPLIT.SERIALNUMBER
			AND dbo.BATCHITEMPLEDGE.LINEID = dbo.BATCHITEMSPLIT.LINEID
		INNER JOIN 
		dbo.BATCHHEADER INNER JOIN dbo.OBJECT ON dbo.BATCHHEADER.ADMITNAME = dbo.OBJECT.ADMITNAME 
	ON dbo.BATCHITEM.ADMITNAME = dbo.BATCHHEADER.ADMITNAME
	LEFT OUTER JOIN 
	dbo.CONTACTATTRIBUTESARCHIVE RIGHT OUTER JOIN dbo.CONTACTSARCHIVE 
	ON dbo.CONTACTATTRIBUTESARCHIVE.SERIALNUMBER = dbo.CONTACTSARCHIVE.SERIALNUMBER 
	ON dbo.BATCHITEM.SERIALNUMBER = dbo.CONTACTSARCHIVE.SERIALNUMBER 
	)ON dbo.CONTACT.SERIALNUMBER = dbo.BATCHITEM.SERIALNUMBER
GO