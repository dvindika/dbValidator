$currentDir = Convert-Path .
$a = “<style>”

$a = $a + “BODY{font-family: Verdana, Arial, Helvetica, sans-serif;font-size:10;font-color: #000000}”

$a = $a + “TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}”

$a = $a + “TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color: #E8E8E8}”

$a = $a + “TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black}”

$a = $a + “</style>”
$ConnString = Read-Host -Prompt "Enter Connection String for database (HOST_NAME\INSTANCE_NAME)"
$Result = Invoke-Sqlcmd -Query ";WITH CTE_Backup AS
(
SELECT  database_name,backup_start_date,type,physical_device_name
       ,Row_Number() OVER(PARTITION BY database_name,BS.type 
        ORDER BY backup_start_date DESC) AS RowNum
FROM    msdb..backupset BS
JOIN    msdb.dbo.backupmediafamily BMF
ON      BS.media_set_id=BMF.media_set_id
)
SELECT      D.name
           ,ISNULL(CONVERT(VARCHAR,backup_start_date),'No backups') AS last_backup_time
           ,D.recovery_model_desc
           ,state_desc,
            CASE WHEN type ='D' THEN 'Full database'
            WHEN type ='I' THEN 'Differential database' 
            WHEN type ='L' THEN 'Log' 
            WHEN type ='F' THEN 'File or filegroup' 
            WHEN type ='G' THEN 'Differential file' 
            WHEN type ='P' THEN 'Partial' 
            WHEN type ='Q' THEN 'Differential partial' 
            ELSE 'Unknown' END AS backup_type 
           ,physical_device_name
FROM        sys.databases D
LEFT JOIN   CTE_Backup CTE
ON          D.name = CTE.database_name
AND         RowNum = 1
ORDER BY    D.name,type;" -ServerInstance $ConnString

$Result|ConvertTo-Html -head $a|Out-File "C:\temp\Temp.html"
Invoke-Item "C:\temp\Temp.html"
cd $currentDir