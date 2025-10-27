/* =========================================================
   SQL Performance Observation Tool (SPOT) - FULL DEPLOY
   Database : utility

   ========================================================= */
USE [utility];
GO

/*------------------------- Ensure schema -------------------------*/
IF NOT EXISTS (
    SELECT 1
    FROM sys.schemas
    WHERE name = N'SPOT'
)
    EXEC ('CREATE SCHEMA SPOT');
GO

/*------------------------- Tables -------------------------*/

-- Run header
IF OBJECT_ID('SPOT.Snapshots', 'U') IS NULL
BEGIN
    CREATE TABLE SPOT.Snapshots (
        SnapshotId        DATETIME2(3) NOT NULL PRIMARY KEY, -- UTC run id
        StartedAtUtc      DATETIME2(3) NOT NULL,
        ServerName        SYSNAME NOT NULL,
        SampleCount       INT NOT NULL,
        SampleTimeSeconds INT NOT NULL,
        GetPlans          BIT NOT NULL,
        CreatedAtUtc      DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME()
    );
END
GO

-- Per-sample header
IF OBJECT_ID('SPOT.Samples', 'U') IS NULL
BEGIN
    CREATE TABLE SPOT.Samples (
        SampleId        BIGINT IDENTITY(1,1) PRIMARY KEY,
        SnapshotId      DATETIME2(3) NOT NULL,
        CaptureTimestamp DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME(),
        ServerName      SYSNAME NOT NULL DEFAULT @@SERVERNAME,
        SampleNumber    INT NOT NULL,
        CONSTRAINT FK_Samples_Snapshots FOREIGN KEY (SnapshotId)
            REFERENCES SPOT.Snapshots(SnapshotId)
    );
END
ELSE
BEGIN
    IF COL_LENGTH('SPOT.Samples', 'SnapshotId') IS NULL
        ALTER TABLE SPOT.Samples ADD SnapshotId DATETIME2(3) NULL;
END
GO

-- WhoIsActive rows
IF OBJECT_ID('SPOT.WhoIsActive', 'U') IS NULL
BEGIN
    CREATE TABLE SPOT.WhoIsActive (
        SampleId             BIGINT NOT NULL,
        SnapshotId           DATETIME2(3) NOT NULL,
        [dd hh:mm:ss.mss]    NVARCHAR(MAX) NULL,
        [session_id]         NVARCHAR(MAX) NULL,
        [sql_text]           XML NULL,
        [login_name]         NVARCHAR(MAX) NULL,
        [wait_info]          NVARCHAR(MAX) NULL,
        [CPU]                NVARCHAR(MAX) NULL,
        [tempdb_allocations] NVARCHAR(MAX) NULL,
        [tempdb_current]     NVARCHAR(MAX) NULL,
        [blocking_session_id] NVARCHAR(MAX) NULL,
        [reads]              NVARCHAR(MAX) NULL,
        [writes]             NVARCHAR(MAX) NULL,
        [physical_reads]     NVARCHAR(MAX) NULL,
        [used_memory]        NVARCHAR(MAX) NULL,
        [status]             NVARCHAR(MAX) NULL,
        [open_tran_count]    NVARCHAR(MAX) NULL,
        [percent_complete]   NVARCHAR(MAX) NULL,
        [host_name]          NVARCHAR(MAX) NULL,
        [database_name]      NVARCHAR(MAX) NULL,
        [program_name]       NVARCHAR(MAX) NULL,
        [start_time]         NVARCHAR(MAX) NULL,
        [login_time]         NVARCHAR(MAX) NULL,
        [request_id]         NVARCHAR(MAX) NULL,
        [query_plan]         XML NULL,
        [collection_time]    NVARCHAR(MAX) NULL,
        CONSTRAINT FK_WIA_Samples FOREIGN KEY (SampleId)
            REFERENCES SPOT.Samples(SampleId),
        CONSTRAINT FK_WIA_Snapshots FOREIGN KEY (SnapshotId)
            REFERENCES SPOT.Snapshots(SnapshotId)
    );
END
ELSE
BEGIN
    IF COL_LENGTH('SPOT.WhoIsActive', 'SnapshotId') IS NULL
        ALTER TABLE SPOT.WhoIsActive ADD SnapshotId DATETIME2(3) NULL;

    IF COL_LENGTH('SPOT.WhoIsActive', 'sql_text') IS NULL
        ALTER TABLE SPOT.WhoIsActive ADD [sql_text] XML NULL;
    ELSE
    BEGIN
        -- ensure sql_text is XML
        IF EXISTS (
            SELECT 1
            FROM sys.columns c
            JOIN sys.types t ON t.user_type_id = c.user_type_id
            WHERE c.object_id = OBJECT_ID('SPOT.WhoIsActive')
              AND c.name = 'sql_text'
              AND t.name <> 'xml'
        )
            ALTER TABLE SPOT.WhoIsActive ALTER COLUMN [sql_text] XML NULL;
    END

    IF COL_LENGTH('SPOT.WhoIsActive', 'query_plan') IS NULL
        ALTER TABLE SPOT.WhoIsActive ADD [query_plan] XML NULL;
END
GO

-- Health checks per sample
IF OBJECT_ID('SPOT.HealthChecks', 'U') IS NULL
BEGIN
    CREATE TABLE SPOT.HealthChecks (
        SampleId          BIGINT NOT NULL,
        SnapshotId        DATETIME2(3) NOT NULL,
        [Check Description] NVARCHAR(255) NOT NULL,
        [Purpose]           NVARCHAR(1000) NOT NULL,
        [Current Value]     NVARCHAR(MAX) NULL,
        CONSTRAINT FK_HC_Samples FOREIGN KEY (SampleId)
            REFERENCES SPOT.Samples(SampleId),
        CONSTRAINT FK_HC_Snapshots FOREIGN KEY (SnapshotId)
            REFERENCES SPOT.Snapshots(SnapshotId)
    );
END
ELSE
BEGIN
    IF COL_LENGTH('SPOT.HealthChecks', 'SnapshotId') IS NULL
        ALTER TABLE SPOT.HealthChecks ADD SnapshotId DATETIME2(3) NULL;
END
GO

-- AG health per sample
IF OBJECT_ID('SPOT.AGHealth', 'U') IS NULL
BEGIN
    CREATE TABLE SPOT.AGHealth (
        SampleId                BIGINT NOT NULL,
        SnapshotId              DATETIME2(3) NOT NULL,
        replica_server_name     NVARCHAR(256) NULL,
        ag_name                 NVARCHAR(256) NULL,
        database_name           NVARCHAR(256) NULL,
        synchronization_state_desc NVARCHAR(60) NULL,
        is_suspended            NVARCHAR(10) NULL,
        log_send_queue_size     NVARCHAR(MAX) NULL,
        redo_queue_size         NVARCHAR(MAX) NULL,
        CONSTRAINT FK_AG_Samples FOREIGN KEY (SampleId)
            REFERENCES SPOT.Samples(SampleId),
        CONSTRAINT FK_AG_Snapshots FOREIGN KEY (SnapshotId)
            REFERENCES SPOT.Snapshots(SnapshotId)
    );
END
ELSE
BEGIN
    IF COL_LENGTH('SPOT.AGHealth', 'SnapshotId') IS NULL
        ALTER TABLE SPOT.AGHealth ADD SnapshotId DATETIME2(3) NULL;
END
GO

-- Wait stats per sample
IF OBJECT_ID('SPOT.WaitStats', 'U') IS NULL
BEGIN
    CREATE TABLE SPOT.WaitStats (
        SampleId            BIGINT NOT NULL,
        SnapshotId          DATETIME2(3) NOT NULL,
        wait_type           NVARCHAR(120) NOT NULL,
        waiting_tasks_count NVARCHAR(MAX) NOT NULL,
        wait_time_ms        NVARCHAR(MAX) NOT NULL,
        max_wait_time_ms    NVARCHAR(MAX) NOT NULL,
        signal_wait_time_ms NVARCHAR(MAX) NOT NULL,
        CONSTRAINT FK_WS_Samples FOREIGN KEY (SampleId)
            REFERENCES SPOT.Samples(SampleId),
        CONSTRAINT FK_WS_Snapshots FOREIGN KEY (SnapshotId)
            REFERENCES SPOT.Snapshots(SnapshotId)
    );
END
ELSE
BEGIN
    IF COL_LENGTH('SPOT.WaitStats', 'SnapshotId') IS NULL
        ALTER TABLE SPOT.WaitStats ADD SnapshotId DATETIME2(3) NULL;
END
GO

-- Blocking per sample
IF OBJECT_ID('SPOT.Blocking', 'U') IS NULL
BEGIN
    CREATE TABLE SPOT.Blocking (
        SampleId            BIGINT NOT NULL,
        SnapshotId          DATETIME2(3) NOT NULL,
        waiting_session_id  NVARCHAR(MAX) NOT NULL,
        blocking_session_id NVARCHAR(MAX) NOT NULL,
        wait_duration_ms    NVARCHAR(MAX) NULL,
        waiting_query       NVARCHAR(MAX) NULL,
        blocking_query      NVARCHAR(MAX) NULL,
        CONSTRAINT FK_BL_Samples FOREIGN KEY (SampleId)
            REFERENCES SPOT.Samples(SampleId),
        CONSTRAINT FK_BL_Snapshots FOREIGN KEY (SnapshotId)
            REFERENCES SPOT.Snapshots(SnapshotId)
    );
END
ELSE
BEGIN
    IF COL_LENGTH('SPOT.Blocking', 'SnapshotId') IS NULL
        ALTER TABLE SPOT.Blocking ADD SnapshotId DATETIME2(3) NULL;
END
GO

/*------------------------- Indexes -------------------------*/
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID('SPOT.Samples') AND name = 'IX_Samples_SnapshotId')
    CREATE INDEX IX_Samples_SnapshotId ON SPOT.Samples (SnapshotId, SampleNumber);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID('SPOT.WhoIsActive') AND name = 'IX_WIA_SnapshotId')
    CREATE INDEX IX_WIA_SnapshotId ON SPOT.WhoIsActive (SnapshotId);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID('SPOT.HealthChecks') AND name = 'IX_HC_SnapshotId')
    CREATE INDEX IX_HC_SnapshotId ON SPOT.HealthChecks (SnapshotId);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID('SPOT.AGHealth') AND name = 'IX_AG_SnapshotId')
    CREATE INDEX IX_AG_SnapshotId ON SPOT.AGHealth (SnapshotId);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID('SPOT.WaitStats') AND name = 'IX_WS_SnapshotId')
    CREATE INDEX IX_WS_SnapshotId ON SPOT.WaitStats (SnapshotId);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID('SPOT.Blocking') AND name = 'IX_BL_SnapshotId')
    CREATE INDEX IX_BL_SnapshotId ON SPOT.Blocking (SnapshotId);
GO

/*------------------------- Stored procedure -------------------------*/

IF OBJECT_ID('SPOT.CaptureSnapshot', 'P') IS NOT NULL
    DROP PROCEDURE SPOT.CaptureSnapshot;
GO

CREATE PROCEDURE SPOT.CaptureSnapshot
    @SampleCount INT = 3,
    @SampleTimeSeconds INT = 5,
    @GetPlans BIT = 0  -- 1 = include query plans
AS
BEGIN
    SET NOCOUNT ON;

    IF @SampleCount IS NULL OR @SampleCount < 1 SET @SampleCount = 1;
    IF @SampleTimeSeconds IS NULL OR @SampleTimeSeconds < 1 SET @SampleTimeSeconds = 1;

    -- Run-level SnapshotId
    DECLARE @SnapshotId DATETIME2(3) = CAST(SYSUTCDATETIME() AS DATETIME2(3));
    DECLARE @ServerName SYSNAME = @@SERVERNAME;

    INSERT INTO SPOT.Snapshots (SnapshotId, StartedAtUtc, ServerName, SampleCount, SampleTimeSeconds, GetPlans)
    VALUES (@SnapshotId, @SnapshotId, @ServerName, @SampleCount, @SampleTimeSeconds, @GetPlans);

    DECLARE @i INT = 1;

    WHILE @i <= @SampleCount
    BEGIN
        DECLARE @SampleId BIGINT;
        INSERT INTO SPOT.Samples (SnapshotId, CaptureTimestamp, ServerName, SampleNumber)
        VALUES (@SnapshotId, SYSUTCDATETIME(), @ServerName, @i);
        SET @SampleId = SCOPE_IDENTITY();

        /* -------- WhoIsActive: pre-create #w with sql_text XML + query_plan XML -------- */
        IF OBJECT_ID('tempdb..#w') IS NOT NULL DROP TABLE #w;
        CREATE TABLE #w (
            [dd hh:mm:ss.mss] NVARCHAR(MAX) NULL,
            [session_id] NVARCHAR(MAX) NULL,
            [sql_text] XML NULL,
            [login_name] NVARCHAR(MAX) NULL,
            [wait_info] NVARCHAR(MAX) NULL,
            [CPU] NVARCHAR(MAX) NULL,
            [tempdb_allocations] NVARCHAR(MAX) NULL,
            [tempdb_current] NVARCHAR(MAX) NULL,
            [blocking_session_id] NVARCHAR(MAX) NULL,
            [reads] NVARCHAR(MAX) NULL,
            [writes] NVARCHAR(MAX) NULL,
            [physical_reads] NVARCHAR(MAX) NULL,
            [used_memory] NVARCHAR(MAX) NULL,
            [status] NVARCHAR(MAX) NULL,
            [open_tran_count] NVARCHAR(MAX) NULL,
            [percent_complete] NVARCHAR(MAX) NULL,
            [host_name] NVARCHAR(MAX) NULL,
            [database_name] NVARCHAR(MAX) NULL,
            [program_name] NVARCHAR(MAX) NULL,
            [start_time] NVARCHAR(MAX) NULL,
            [login_time] NVARCHAR(MAX) NULL,
            [request_id] NVARCHAR(MAX) NULL,
            [query_plan] XML NULL,
            [collection_time] NVARCHAR(MAX) NULL
        );

        DECLARE @OutCols NVARCHAR(MAX) =
            '[dd hh:mm:ss.mss][session_id][sql_text][login_name][wait_info][CPU][tempdb_allocations][tempdb_current][blocking_session_id]' +
            '[reads][writes][physical_reads][used_memory][status][open_tran_count][percent_complete][host_name][database_name][program_name]' +
            '[start_time][login_time][request_id]' +
            CASE WHEN @GetPlans = 1 THEN '[query_plan]' ELSE '' END +
            '[collection_time]';

        BEGIN TRY
            EXEC sp_WhoIsActive
                @get_plans = @GetPlans,
                @output_column_list = @OutCols,
                @destination_table = '#w';
        END TRY
        BEGIN CATCH
            PRINT CONCAT('WhoIsActive failed: ', ERROR_MESSAGE());
        END CATCH;

        -- Persist
        INSERT INTO SPOT.WhoIsActive (
            SampleId, SnapshotId, [dd hh:mm:ss.mss], [session_id], [sql_text], [login_name],
            [wait_info], [CPU], [tempdb_allocations], [tempdb_current], [blocking_session_id],
            [reads], [writes], [physical_reads], [used_memory], [status], [open_tran_count],
            [percent_complete], [host_name], [database_name], [program_name], [start_time],
            [login_time], [request_id], [query_plan], [collection_time]
        )
        SELECT
            @SampleId, @SnapshotId, [dd hh:mm:ss.mss], [session_id], [sql_text], [login_name],
            [wait_info], [CPU], [tempdb_allocations], [tempdb_current], [blocking_session_id],
            [reads], [writes], [physical_reads], [used_memory], [status], [open_tran_count],
            [percent_complete], [host_name], [database_name], [program_name], [start_time],
            [login_time], [request_id], [query_plan], [collection_time]
        FROM #w;

        /* ---------------- Health Checks ---------------- */
        DECLARE @HealthChecks TABLE (
            [Check Description] NVARCHAR(255),
            [Purpose] NVARCHAR(1000),
            [Current Value] NVARCHAR(MAX)
        );

        DECLARE @AvgSqlCpuUtilization INT, @AvgTotalCpuUtilization INT;

        ;WITH RingRaw AS (
            SELECT CONVERT(XML, rb.record) AS x
            FROM sys.dm_os_ring_buffers rb
            WHERE rb.ring_buffer_type = N'RING_BUFFER_SCHEDULER_MONITOR'
        ),
        Parsed AS (
            SELECT
                x.value('(./Record/SchedulerMonitorEvent/SystemHealth/SystemIdle)[1]', 'int') AS SystemIdle,
                x.value('(./Record/SchedulerMonitorEvent/SystemHealth/ProcessUtilization)[1]', 'int') AS ProcessUtilization
            FROM RingRaw
        )
        SELECT
            @AvgSqlCpuUtilization = CASE WHEN COUNT(*) > 0 THEN CONVERT(INT, ROUND(AVG(CAST(ProcessUtilization AS FLOAT)), 0)) ELSE NULL END,
            @AvgTotalCpuUtilization = CASE WHEN COUNT(*) > 0 THEN CONVERT(INT, ROUND(100 - AVG(CAST(SystemIdle AS FLOAT)), 0)) ELSE NULL END
        FROM Parsed
        WHERE ProcessUtilization IS NOT NULL OR SystemIdle IS NOT NULL;

        INSERT INTO @HealthChecks VALUES
            (N'CPU Utilization (SQL Server Process)', N'Shows the percentage of total CPU time being used by the SQL Server process itself. High values indicate CPU pressure from queries.', ISNULL(CAST(@AvgSqlCpuUtilization AS NVARCHAR(10)) + N'%', N'Data not available')),
            (N'CPU Utilization (Total - Including OS)', N'Shows the total CPU usage on the server. If this is high but SQL CPU is low, another process is consuming CPU resources.', ISNULL(CAST(@AvgTotalCpuUtilization AS NVARCHAR(10)) + N'%', N'Data not available'));

        INSERT INTO @HealthChecks
        SELECT
            N'Runnable Schedulers (Signal Waits)',
            N'A count of tasks ready to run but waiting for CPU. Consistently high values (e.g., > 5–10) indicate significant CPU pressure.',
            CAST(SUM(runnable_tasks_count) AS NVARCHAR(10))
        FROM sys.dm_os_schedulers
        WHERE scheduler_id < 255;

        INSERT INTO @HealthChecks
        SELECT
            N'Memory Grants Pending',
            N'Number of queries waiting for a memory grant to execute. Any value > 0 is a clear sign of memory pressure.',
            CAST(COUNT(*) AS NVARCHAR(10))
        FROM sys.dm_exec_query_memory_grants
        WHERE grant_time IS NULL;

        DECLARE @FinalPLEValue NVARCHAR(MAX), @FinalBCHRValue NVARCHAR(MAX);
        BEGIN TRY
            DECLARE @PLEValue BIGINT;
            SELECT @PLEValue = cntr_value
            FROM sys.dm_os_performance_counters
            WHERE [object_name] LIKE '%Buffer Manager%' AND counter_name = 'Page life expectancy';
            SET @FinalPLEValue = ISNULL(CAST(@PLEValue AS NVARCHAR(20)) + N' seconds', N'N/A');

            DECLARE @BCHRValue DECIMAL(10,2);
            SELECT @BCHRValue = CAST((a.cntr_value * 1.0 / NULLIF(b.cntr_value,0)) * 100 AS DECIMAL(10,2))
            FROM sys.dm_os_performance_counters AS a
            JOIN sys.dm_os_performance_counters AS b ON a.object_name = b.object_name
            WHERE a.counter_name = 'Buffer cache hit ratio'
              AND b.counter_name = 'Buffer cache hit ratio base'
              AND a.object_name LIKE '%Buffer Manager%';
            SET @FinalBCHRValue = ISNULL(CAST(@BCHRValue AS NVARCHAR(20)) + N'%', N'N/A');
        END TRY
        BEGIN CATCH
            IF @FinalPLEValue IS NULL SET @FinalPLEValue = N'Calculation Error';
            IF @FinalBCHRValue IS NULL SET @FinalBCHRValue = N'Calculation Error';
        END CATCH;

        INSERT INTO @HealthChecks VALUES
            (N'Page Life Expectancy (PLE)', N'Seconds a data page will stay in cache. A low value indicates memory pressure.', @FinalPLEValue),
            (N'Buffer Cache Hit Ratio', N'Percentage of pages found in memory without having to be read from disk. Should be > 95% for OLTP systems. Low values indicate insufficient memory.', @FinalBCHRValue);

        IF OBJECT_ID('tempdb..#LogSpace') IS NOT NULL DROP TABLE #LogSpace;
        CREATE TABLE #LogSpace (
            [Database Name] NVARCHAR(128),
            [Log Size (MB)] DECIMAL(18,5),
            [Log Space Used (%)] DECIMAL(18,5),
            [Status] INT
        );
        INSERT INTO #LogSpace EXEC ('DBCC SQLPERF(LOGSPACE);');

        INSERT INTO SPOT.HealthChecks (SampleId, SnapshotId, [Check Description], [Purpose], [Current Value])
        SELECT @SampleId, @SnapshotId, [Check Description], [Purpose], [Current Value]
        FROM @HealthChecks;

        DROP TABLE #LogSpace;

        /* AG Health */
        BEGIN TRY
            INSERT INTO SPOT.AGHealth (
                SampleId, SnapshotId, replica_server_name, ag_name, database_name,
                synchronization_state_desc, is_suspended, log_send_queue_size, redo_queue_size
            )
            SELECT
                @SampleId, @SnapshotId, ar.replica_server_name, ag.name AS ag_name,
                DB_NAME(drs.database_id) AS database_name, drs.synchronization_state_desc,
                CAST(drs.is_suspended AS NVARCHAR(10)), CAST(drs.log_send_queue_size AS NVARCHAR(MAX)),
                CAST(drs.redo_queue_size AS NVARCHAR(MAX))
            FROM sys.dm_hadr_database_replica_states drs
            JOIN sys.availability_replicas ar ON drs.replica_id = ar.replica_id
            JOIN sys.availability_groups ag ON drs.group_id = ag.group_id;
        END TRY
        BEGIN CATCH
            PRINT CONCAT('AGHealth skip/info: ', ERROR_MESSAGE());
        END CATCH;

        /* Wait Stats */
        INSERT INTO SPOT.WaitStats (
            SampleId, SnapshotId, wait_type, waiting_tasks_count,
            wait_time_ms, max_wait_time_ms, signal_wait_time_ms
        )
        SELECT
            @SampleId, @SnapshotId, wait_type,
            CAST(waiting_tasks_count AS NVARCHAR(MAX)),
            CAST(wait_time_ms AS NVARCHAR(MAX)),
            CAST(max_wait_time_ms AS NVARCHAR(MAX)),
            CAST(signal_wait_time_ms AS NVARCHAR(MAX))
        FROM sys.dm_os_wait_stats
        WHERE wait_type NOT IN (
            'BROKER_TASK_STOP','BROKER_EVENTHANDLER','BROKER_RECEIVE_WAITFOR','CHECKPOINT_QUEUE',
            'CLR_AUTO_EVENT','DBMIRROR_DBM_EVENT','FT_IFTS_SCHEDULER_IDLE_WAIT','HADR_CLUSAPI_CALL',
            'HADR_TIMER_TASK','LAZYWRITER_SLEEP','LOGMGR_QUEUE','REQUEST_FOR_DEADLOCK_SEARCH',
            'SLEEP_TASK','SQLTRACE_BUFFER_FLUSH','WAITFOR','XE_DISPATCHER_WAIT','XE_TIMER_EVENT'
        );

        /* Blocking */
        INSERT INTO SPOT.Blocking (
            SampleId, SnapshotId, waiting_session_id, blocking_session_id,
            wait_duration_ms, waiting_query, blocking_query
        )
        SELECT
            @SampleId, @SnapshotId,
            CAST(wt.session_id AS NVARCHAR(MAX)),
            CAST(wt.blocking_session_id AS NVARCHAR(MAX)),
            CAST(wt.wait_duration_ms AS NVARCHAR(MAX)),
            (SELECT [text] FROM sys.dm_exec_requests r CROSS APPLY sys.dm_exec_sql_text(r.sql_handle)
             WHERE r.session_id = wt.session_id),
            (SELECT [text] FROM sys.dm_exec_requests r CROSS APPLY sys.dm_exec_sql_text(r.sql_handle)
             WHERE r.session_id = wt.blocking_session_id)
        FROM sys.dm_os_waiting_tasks wt
        WHERE wt.blocking_session_id IS NOT NULL AND wt.blocking_session_id <> 0;

        /* Sleep between samples */
        IF @i < @SampleCount
        BEGIN
            DECLARE @delay CHAR(12) =
                RIGHT('00' + CAST(@SampleTimeSeconds / 3600 AS VARCHAR(2)), 2) + ':' +
                RIGHT('00' + CAST((@SampleTimeSeconds % 3600) / 60 AS VARCHAR(2)), 2) + ':' +
                RIGHT('00' + CAST(@SampleTimeSeconds % 60 AS VARCHAR(2)), 2) + '.000';
            WAITFOR DELAY @delay;
        END

        SET @i += 1;
    END

    /* Summary for this run */
    SELECT
        s.SnapshotId,
        s.StartedAtUtc,
        s.ServerName,
        s.SampleCount,
        s.SampleTimeSeconds,
        s.GetPlans,
        WIA = (SELECT COUNT(*) FROM SPOT.WhoIsActive WHERE SnapshotId = s.SnapshotId),
        HC  = (SELECT COUNT(*) FROM SPOT.HealthChecks WHERE SnapshotId = s.SnapshotId),
        AG  = (SELECT COUNT(*) FROM SPOT.AGHealth WHERE SnapshotId = s.SnapshotId),
        WS  = (SELECT COUNT(*) FROM SPOT.WaitStats WHERE SnapshotId = s.SnapshotId),
        BL  = (SELECT COUNT(*) FROM SPOT.Blocking WHERE SnapshotId = s.SnapshotId)
    FROM SPOT.Snapshots s
    WHERE s.SnapshotId = @SnapshotId;
END
GO
-- Adding purge routine
IF OBJECT_ID('SPOT.PurgeOldData', 'P') IS NOT NULL
    DROP PROCEDURE SPOT.PurgeOldData;
GO

CREATE PROCEDURE SPOT.PurgeOldData
(
      @RetainDays INT  -- keep this many days of history, purge anything older
)
AS
BEGIN
    SET NOCOUNT ON;

    ----------------------------------------------------------------
    -- 1. Validate / normalise input
    ----------------------------------------------------------------
    IF @RetainDays IS NULL OR @RetainDays < 0
    BEGIN
        RAISERROR('RetainDays must be 0 or greater.', 16, 1);
        RETURN;
    END

    /*
        Define cutoff point.
        Anything with StartedAtUtc < @CutoffUtc will be purged.
        Example: @RetainDays = 7 means keep last 7 days, purge older than 7 days.
    */
    DECLARE @CutoffUtc DATETIME2(3) = DATEADD(DAY, -@RetainDays, SYSUTCDATETIME());

    ----------------------------------------------------------------
    -- 2. Identify old SnapshotIds first (drives all child rows)
    ----------------------------------------------------------------
    ;WITH OldSnaps AS
    (
        SELECT SnapshotId
        FROM SPOT.Snapshots
        WHERE StartedAtUtc < @CutoffUtc
    )
    ----------------------------------------------------------------
    -- 3. Purge children in FK-safe order
    ----------------------------------------------------------------
    -- WhoIsActive
    DELETE WIA
    FROM SPOT.WhoIsActive AS WIA
    INNER JOIN OldSnaps OS ON WIA.SnapshotId = OS.SnapshotId;

    -- HealthChecks
    DELETE HC
    FROM SPOT.HealthChecks AS HC
    INNER JOIN OldSnaps OS ON HC.SnapshotId = OS.SnapshotId;

    -- AGHealth
    DELETE AG
    FROM SPOT.AGHealth AS AG
    INNER JOIN OldSnaps OS ON AG.SnapshotId = OS.SnapshotId;

    -- WaitStats
    DELETE WS
    FROM SPOT.WaitStats AS WS
    INNER JOIN OldSnaps OS ON WS.SnapshotId = OS.SnapshotId;

    -- Blocking
    DELETE BL
    FROM SPOT.Blocking AS BL
    INNER JOIN OldSnaps OS ON BL.SnapshotId = OS.SnapshotId;

    -- Samples (depends on Snapshots)
    DELETE S
    FROM SPOT.Samples AS S
    INNER JOIN OldSnaps OS ON S.SnapshotId = OS.SnapshotId;

    ----------------------------------------------------------------
    -- 4. Finally purge Snapshots themselves
    ----------------------------------------------------------------
    DELETE SN
    FROM SPOT.Snapshots AS SN
    INNER JOIN OldSnaps OS ON SN.SnapshotId = OS.SnapshotId;

    ----------------------------------------------------------------
    -- 5. Return info about what is now left (optional summary)
    ----------------------------------------------------------------
    SELECT
        RemainingSnapshots = COUNT(*) ,
        OldestSnapshotUtc  = MIN(StartedAtUtc),
        NewestSnapshotUtc  = MAX(StartedAtUtc)
    FROM SPOT.Snapshots;

END
GO
