/* 
Script to generate empty CVIP database for assortis clients
Rename inrae -> real name of a client
*/

BEGIN TRANSACTION            
  DECLARE @JobID BINARY(16)  
  DECLARE @ReturnCode INT    
  SELECT @ReturnCode = 0     
IF (SELECT COUNT(*) FROM msdb.dbo.syscategories WHERE name = N'[Uncategorized (Local)]') < 1 
  EXECUTE msdb.dbo.sp_add_category @name = N'[Uncategorized (Local)]'

  -- Delete the job with the same name (if it exists)
  SELECT @JobID = job_id     
  FROM   msdb.dbo.sysjobs    
  WHERE (name = N'CVIP inrae. Experts. Ranking update')       
  IF (@JobID IS NOT NULL)    
  BEGIN  
  -- Check if the job is a multi-server job  
  IF (EXISTS (SELECT  * 
              FROM    msdb.dbo.sysjobservers 
              WHERE   (job_id = @JobID) AND (server_id <> 0))) 
  BEGIN 
    -- There is, so abort the script 
    RAISERROR (N'Unable to import job ''CVIP inrae. Experts. Ranking update'' since there is already a multi-server job with this name.', 16, 1) 
    GOTO QuitWithRollback  
  END 
  ELSE 
    -- Delete the [local] job 
    EXECUTE msdb.dbo.sp_delete_job @job_name = N'CVIP inrae. Experts. Ranking update' 
    SELECT @JobID = NULL
  END 

BEGIN 

  -- Add the job
  EXECUTE @ReturnCode = msdb.dbo.sp_add_job @job_id = @JobID OUTPUT , @job_name = N'CVIP inrae. Experts. Ranking update', @owner_login_name = N'sa', @description = N'No description available.', @category_name = N'[Uncategorized (Local)]', @enabled = 1, @notify_level_email = 0, @notify_level_page = 0, @notify_level_netsend = 0, @notify_level_eventlog = 2, @delete_level= 0
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 

  -- Add the job steps
  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 1, @step_name = N'SctRanking', @command = N'DELETE FROM lnkExp_RankSct

INSERT INTO lnkExp_RankSct
SELECT T1.id_Expert, T1.id_Sector, 
--CONVERT(decimal(8,4), CONVERT(real, T1.PointsPerSector)*100/(T2.PointsSeniority+1))+T1.PointsPerSector AS rnkSctValue
CONVERT(real, T1.PointsPerSector)*TotalRecords/T2.TotalMainSectors AS rnkSctValue
FROM
(
SELECT id_Expert, id_Sector, SUM(DATEDIFF(m, wkeStartDate, wkeEndDate)) AS PointsPerSector
FROM lnkExp_Wke EW INNER JOIN lnkWke_Sct WS ON EW.id_ExpWke=WS.id_ExpWke
WHERE wkeEndDate>=wkeStartDate
GROUP BY id_Expert, id_Sector
) AS T1
LEFT OUTER JOIN
(
SELECT id_Expert, COUNT(DISTINCT WS.id_Sector) AS TotalSectors, COUNT(DISTINCT S.id_MainSector) AS TotalMainSectors
FROM lnkExp_Wke EW INNER JOIN lnkWke_Sct WS ON EW.id_ExpWke=WS.id_ExpWke
INNER JOIN tbl_Sectors S ON WS.id_Sector=S.id_Sector
WHERE wkeEndDate>=wkeStartDate
GROUP BY id_Expert
) AS T2
ON T1.id_Expert=T2.id_Expert
LEFT OUTER JOIN
(
SELECT id_Expert, id_Sector, COUNT(DISTINCT EW.id_ExpWke) AS TotalRecords
FROM lnkExp_Wke EW INNER JOIN lnkWke_Sct WS ON EW.id_ExpWke=WS.id_ExpWke
WHERE wkeEndDate>=wkeStartDate 
AND DATEDIFF(m, wkeStartDate, wkeEndDate)>12
GROUP BY id_Expert, id_Sector
) AS T3
ON T1.id_Expert=T3.id_Expert AND T1.id_Sector=T3.id_Sector
ORDER BY T1.id_Expert
', @database_name = N'rms_inrae', @server = N'', @database_user_name = N'', @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 1, @output_file_name = N'', @on_success_step_id = 0, @on_success_action = 3, @on_fail_step_id = 0, @on_fail_action = 3
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 
  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 2, @step_name = N'CouRanking', @command = N'DELETE FROM lnkExp_RankCou

INSERT INTO lnkExp_RankCou
SELECT T1.id_Expert, T1.id_Country, CONVERT(decimal(8,4), CONVERT(real, T1.PointsPerCountry)*100/(T2.PointsTotal+1))+T1.PointsPerCountry AS rnkCouValue
FROM
(
SELECT id_Expert, id_Country, SUM(DATEDIFF(m, wkeStartDate, wkeEndDate)) AS PointsPerCountry
FROM lnkExp_Wke EW INNER JOIN lnkWke_Cou WS ON EW.id_ExpWke=WS.id_ExpWke
WHERE wkeEndDate>=wkeStartDate
GROUP BY id_Expert, id_Country
) AS T1
LEFT OUTER JOIN
(
SELECT id_Expert, SUM(DATEDIFF(m, ISNULL(wkeStartDate, DATEADD(y, -1, wkeEndDate)), ISNULL(wkeEndDate, DATEADD(y, 1, wkeStartDate)))) AS PointsTotal
FROM lnkExp_Wke EW 
WHERE wkeEndDate>=wkeStartDate
GROUP BY id_Expert
) AS T2
ON T1.id_Expert=T2.id_Expert
ORDER BY T1.id_Expert
', @database_name = N'rms_inrae', @server = N'', @database_user_name = N'', @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 1, @output_file_name = N'', @on_success_step_id = 0, @on_success_action = 1, @on_fail_step_id = 0, @on_fail_action = 2
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 
  EXECUTE @ReturnCode = msdb.dbo.sp_update_job @job_id = @JobID, @start_step_id = 1 

  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 

  -- Add the job schedules
  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id = @JobID, @name = N'Daily at 11-15', @enabled = 1, @freq_type = 4, @active_start_date = 20040422, @active_start_time = 111500, @freq_interval = 1, @freq_subday_type = 1, @freq_subday_interval = 0, @freq_relative_interval = 0, @freq_recurrence_factor = 0, @active_end_date = 99991231, @active_end_time = 235959
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 

  -- Add the Target Servers
  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @JobID, @server_name = N'(local)' 
  IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback 

END
COMMIT TRANSACTION          
GOTO   EndSave              
QuitWithRollback:
  IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION 
EndSave: 


