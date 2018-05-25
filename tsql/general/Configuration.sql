SELECT
  -- Returns system and build information for the current installation of SQL Server.
  @@VERSION AS [Version]
  -- Returns the name of the registry key under which SQL Server is running.
  ,@@SERVICENAME AS [Service Name]
  -- Returns the name of the local server that is running SQL Server.
  ,@@SERVERNAME AS [Server Name]
  -- Returns the local language identifier (ID) of the language that is currently being used.
  ,@@LANGID AS [Language ID]
  -- Returns the name of the language currently being used.
  ,@@LANGUAGE AS [Language]
  -- returns the last-used timestamp value of the current database.
  ,@@DBTS AS [Last Used Timestamp]
  -- Returns the session ID of the current user process.
  ,@@SPID AS [Session ID]
  -- Returns the current lock time-out setting in milliseconds for the current session.
  ,@@LOCK_TIMEOUT AS [Lock Time-out MS]
   -- Returns the current value of the TEXTSIZE option.
  ,@@TEXTSIZE AS [Text Size]
  -- Returns the maximum number of simultaneous user connections allowed on an instance of SQL Server. The number returned is not necessarily the number currently configured.
  ,@@MAX_CONNECTIONS AS [Max Connections]
  -- Returns the precision level used by decimal and numeric data types as currently set in the server.
  ,@@MAX_PRECISION AS [Max Precision]
  -- Returns the nesting level of the current stored procedure execution (initially 0) on the local server.
  ,@@NESTLEVEL AS [Nest Level]
  -- Returns information about the current SET options.
  ,@@OPTIONS  AS [Options]