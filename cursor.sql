CREATE TABLE #tmp (banco VARCHAR(200),  total_usuarios int);

DECLARE DBCursor CURSOR LOCAL STATIC FORWARD_ONLY READ_ONLY
FOR
    SELECT name 
	FROM Sys.Databases
OPEN DBCursor;

DECLARE @DBName VARCHAR(200) = '';
FETCH NEXT FROM DBCursor INTO @DBName;

WHILE @@FETCH_STATUS = 0
    BEGIN
        DECLARE @SQL NVARCHAR(MAX) = N'USE ' + QUOTENAME(@DBName) + '
                                        IF EXISTS 
                                            (   SELECT  1
                                                FROM    sys.tables
                                                WHERE   [Object_ID] = OBJECT_ID(N''dbo.Usuarios'')
                                            )
                                            BEGIN
                                                INSERT #tmp (banco, total_usuarios) 
                                                SELECT @DB, SUM(quantidade) as total
                                                FROM    dbo.Usuarios						
                                            END';
        EXECUTE SP_EXECUTESQL @SQL, N'@DB VARCHAR(200)', @DBName;
        FETCH NEXT FROM DBCursor INTO @DBName;
    END

CLOSE DBCursor;
DEALLOCATE DBCursor;

SELECT  *
FROM    #tmp;

--DROP TABLE #tmp;