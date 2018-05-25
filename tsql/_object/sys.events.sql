-- Contains a row for each event for which a trigger or event notification fires.
-- These events represent the event types that are specified when the trigger or event notification is created by using
-- CREATE TRIGGER 
-- or
-- CREATE EVENT NOTIFICATION

SELECT
  *
FROM sys.events