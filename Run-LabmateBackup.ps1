[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'labmate-config.json'),
    [switch]$Force,
    [switch]$CompareOnly,
    [string[]]$OnlyTable
)

$ErrorActionPreference = 'Stop'
$script:ReportWriter = $null
$script:StagedBackupPath = $null
$script:RestoreSucceeded = $false
$script:ProcessedBackupStatePath = Join-Path $PSScriptRoot 'processed-backup-state.json'

function Write-RunLog {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR','OK')][string]$Level = 'INFO')
    $line = '[{0:yyyy-MM-dd HH:mm:ss}] [{1}] {2}' -f (Get-Date), $Level, $Message
    Write-Host $line
    if ($null -ne $script:ReportWriter) {
        $script:ReportWriter.WriteLine($line)
        $script:ReportWriter.Flush()
    }
}

function Write-ReportLine {
    param([string]$Text = '')
    if ($null -ne $script:ReportWriter) {
        $script:ReportWriter.WriteLine($Text)
        $script:ReportWriter.Flush()
    }
}

function Quote-SqlIdentifier {
    param([string]$Name)
    return '[' + $Name.Replace(']', ']]') + ']'
}

function Quote-MySqlIdentifier {
    param([string]$Name)
    return '`' + $Name.Replace('`', '``') + '`'
}

function ConvertTo-SqlLiteral {
    param([string]$Value)
    return "N'" + $Value.Replace("'", "''") + "'"
}

function Get-SqlConnectionString {
    param($Config, [string]$Database = 'master')
    if ([bool]$Config.SqlServer.UseWindowsAuthentication) {
        return "Server=$($Config.SqlServer.Server);Database=$Database;Integrated Security=True;TrustServerCertificate=True;"
    }
    $passwordName = [string]$Config.SqlServer.PasswordEnvironmentVariable
    $password = [Environment]::GetEnvironmentVariable($passwordName)
    if ([string]::IsNullOrEmpty($password)) {
        throw "SQL password environment variable '$passwordName' is empty."
    }
    return "Server=$($Config.SqlServer.Server);Database=$Database;User ID=$($Config.SqlServer.Username);Password=$password;TrustServerCertificate=True;"
}

function Open-SqlConnection {
    param($Config, [string]$Database = 'master')
    $connection = New-Object System.Data.SqlClient.SqlConnection (Get-SqlConnectionString $Config $Database)
    $connection.Open()
    return $connection
}

function Invoke-SqlScalar {
    param($Connection, [string]$Sql, [int]$Timeout = 120)
    $command = $Connection.CreateCommand()
    try {
        $command.CommandTimeout = $Timeout
        $command.CommandText = $Sql
        return $command.ExecuteScalar()
    }
    finally { $command.Dispose() }
}

function Invoke-SqlNonQuery {
    param($Connection, [string]$Sql, [int]$Timeout = 0)
    $command = $Connection.CreateCommand()
    try {
        $command.CommandTimeout = $Timeout
        $command.CommandText = $Sql
        $null = $command.ExecuteNonQuery()
    }
    finally { $command.Dispose() }
}

function Start-MySqlQuery {
    param($Config, [string]$Sql)
    $exe = [string]$Config.MySql.Exe
    if (-not (Test-Path -LiteralPath $exe)) { throw "MySQL client not found: $exe" }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $exe
    $psi.Arguments = '-h "{0}" -P {1} -u "{2}" --default-character-set=utf8mb4 --batch --skip-column-names "{3}"' -f `
        $Config.MySql.Host, $Config.MySql.Port, $Config.MySql.Username, $Config.MySql.Database
    $psi.UseShellExecute = $false
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $passwordName = [string]$Config.MySql.PasswordEnvironmentVariable
    $password = [Environment]::GetEnvironmentVariable($passwordName)
    if ([string]::IsNullOrEmpty($password) -and -not [string]::IsNullOrEmpty([string]$Config.MySql.PasswordFile)) {
        $passwordFile = [string]$Config.MySql.PasswordFile
        if (-not (Test-Path -LiteralPath $passwordFile -PathType Leaf)) {
            throw "MySQL encrypted password file not found: $passwordFile"
        }
        try {
            $securePassword = Get-Content -LiteralPath $passwordFile -Raw | ConvertTo-SecureString
            $credential = New-Object System.Management.Automation.PSCredential('mysql-user',$securePassword)
            $password = $credential.GetNetworkCredential().Password
        }
        catch { throw "Could not decrypt MySQL password file: $($_.Exception.Message)" }
    }
    if ($null -ne $password) { $psi.EnvironmentVariables['MYSQL_PWD'] = $password }

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi
    $null = $process.Start()
    $process.StandardInput.WriteLine($Sql)
    $process.StandardInput.Close()
    return $process
}

function Complete-MySqlQuery {
    param($Process)
    $Process.WaitForExit()
    $errorText = $Process.StandardError.ReadToEnd()
    if ($Process.ExitCode -ne 0) {
        throw "MySQL query failed (exit $($Process.ExitCode)): $errorText"
    }
    $Process.Dispose()
}

function Invoke-MySqlLines {
    param($Config, [string]$Sql)
    $process = Start-MySqlQuery $Config $Sql
    $lines = New-Object System.Collections.Generic.List[string]
    while (($line = $process.StandardOutput.ReadLine()) -ne $null) { $lines.Add($line) }
    Complete-MySqlQuery $process
    return $lines.ToArray()
}

function ConvertFrom-MySqlBatchField {
    param([string]$Field)
    if ($Field -eq '\N' -or $Field -eq 'NULL') { return $null }
    $builder = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $Field.Length; $i++) {
        $ch = $Field[$i]
        if ($ch -ne '\' -or $i + 1 -ge $Field.Length) {
            $null = $builder.Append($ch)
            continue
        }
        $i++
        switch ($Field[$i]) {
            't' { $null = $builder.Append("`t") }
            'n' { $null = $builder.Append("`n") }
            'r' { $null = $builder.Append("`r") }
            '0' { $null = $builder.Append([char]0) }
            'Z' { $null = $builder.Append([char]26) }
            '\' { $null = $builder.Append('\') }
            default { $null = $builder.Append($Field[$i]) }
        }
    }
    return $builder.ToString()
}

function Normalize-KeyValue {
    param($Value)
    if ($null -eq $Value -or $Value -is [DBNull]) { return '<NULL>' }
    if ($Value -is [DateTime]) { return $Value.ToString('yyyy-MM-dd HH:mm:ss.fffffff').TrimEnd('0').TrimEnd('.') }
    if ($Value -is [bool]) { if ($Value) { return '1' } else { return '0' } }
    return ([string]$Value).Trim().ToUpperInvariant()
}

function New-CanonicalKey {
    param([object[]]$Values)
    $parts = foreach ($value in $Values) {
        $normalized = Normalize-KeyValue $value
        '{0}:{1}' -f $normalized.Length, $normalized
    }
    return $parts -join '|'
}

function Format-ValueForReport {
    param($Value, [int]$MaxChars = 1000)
    if ($null -eq $Value -or $Value -is [DBNull]) { return 'NULL' }
    if ($Value -is [bool]) { if ($Value) { return '1' } else { return '0' } }
    if ($Value -is [DateTime]) { return $Value.ToString('yyyy-MM-dd HH:mm:ss.fffffff').TrimEnd('0').TrimEnd('.') }
    $text = ([string]$Value).Replace("`r", '\r').Replace("`n", '\n').Replace("`t", '\t')
    if ($text.Length -gt $MaxChars) { return $text.Substring(0, $MaxChars) + '...[TRUNCATED]' }
    return $text
}

function Assert-Administrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'Restore mode requires Administrator privileges. Right-click RUN_AUTOMATION.bat and choose Run as administrator.'
    }
}

function Get-NewestStableBackup {
    param($Config)
    $folder = [string]$Config.BackupFolder
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) { throw "Backup folder does not exist: $folder" }
    $minimumAge = [int]$Config.MinimumBackupAgeSeconds
    $cutoff = (Get-Date).AddSeconds(-$minimumAge)
    $files = @(Get-ChildItem -LiteralPath $folder -File -Filter ([string]$Config.BackupFilePattern) |
        Where-Object { $_.LastWriteTime -le $cutoff } |
        Sort-Object LastWriteTime -Descending)
    if ($files.Count -eq 0) {
        throw "No stable backup matching '$($Config.BackupFilePattern)' found in $folder. Wait for the download to finish."
    }
    return $files[0]
}

function Get-BackupProcessingInfo {
    param([System.IO.FileInfo]$Backup)
    return [pscustomobject]@{
        FileName = $Backup.Name
        FullPath = $Backup.FullName
        SizeBytes = [long]$Backup.Length
        LastWriteTimeUtc = $Backup.LastWriteTimeUtc.ToString('o')
    }
}

function Test-BackupAlreadyProcessed {
    param($Info, $Config)
    if (-not (Test-Path -LiteralPath $script:ProcessedBackupStatePath -PathType Leaf)) { return $false }
    try { $state = Get-Content -LiteralPath $script:ProcessedBackupStatePath -Raw | ConvertFrom-Json }
    catch {
        Write-RunLog "Processed-backup state could not be read and will be ignored: $($_.Exception.Message)" 'WARN'
        return $false
    }
    return (
        [string]$state.FileName -ceq [string]$Info.FileName -and
        [string]$state.SqlServer -ieq [string]$Config.SqlServer.Server -and
        [string]$state.SqlDatabase -ieq [string]$Config.SqlServer.Database
    )
}

function Save-ProcessedBackupState {
    param($Info, $Config)
    $state = [ordered]@{
        FileName = [string]$Info.FileName
        FullPath = [string]$Info.FullPath
        SizeBytes = [long]$Info.SizeBytes
        LastWriteTimeUtc = [string]$Info.LastWriteTimeUtc
        SqlServer = [string]$Config.SqlServer.Server
        SqlDatabase = [string]$Config.SqlServer.Database
        ProcessedSuccessfullyAt = (Get-Date).ToString('o')
    }
    $json = $state | ConvertTo-Json
    [System.IO.File]::WriteAllText($script:ProcessedBackupStatePath,$json,(New-Object System.Text.UTF8Encoding($false)))
    Write-RunLog "Saved processed-backup state: $script:ProcessedBackupStatePath" 'OK'
}

function Stage-And-RestoreBackup {
    param($Config, [switch]$SkipConfirmation)
    Assert-Administrator
    $backup = Get-NewestStableBackup $Config
    Write-RunLog "Selected backup: $($backup.FullName) ($([math]::Round($backup.Length / 1GB, 3)) GB)"
    $processingInfo = Get-BackupProcessingInfo $backup
    if (Test-BackupAlreadyProcessed $processingInfo $Config) {
        Write-RunLog 'This backup has already been processed. Restore skipped.' 'WARN'
        Write-ReportLine 'Duplicate backup result: already processed; restore skipped, comparison will run.'
        return [pscustomobject]@{ AlreadyProcessed=$true; Info=$processingInfo }
    }

    if (-not $SkipConfirmation) {
        Write-Host ''
        Write-Host "WARNING: This will overwrite SQL database '$($Config.SqlServer.Database)' on '$($Config.SqlServer.Server)'." -ForegroundColor Yellow
        $answer = Read-Host "Type RESTORE to continue"
        if ($answer.Trim() -ine 'RESTORE') { throw 'Restore cancelled by user.' }
    }

    $workingFolder = [string]$Config.SqlServer.BackupWorkingFolder
    if (-not (Test-Path -LiteralPath $workingFolder)) { $null = New-Item -ItemType Directory -Path $workingFolder -Force }
    Get-ChildItem -LiteralPath $workingFolder -File -Filter 'Labmate_Automation_*.bak' -ErrorAction SilentlyContinue |
        Remove-Item -Force
    $stagedName = 'Labmate_Automation_{0}' -f $backup.Name
    $script:StagedBackupPath = Join-Path $workingFolder $stagedName
    Write-RunLog "Copying backup to SQL Server working folder..."
    Copy-Item -LiteralPath $backup.FullName -Destination $script:StagedBackupPath -Force
    $staged = Get-Item -LiteralPath $script:StagedBackupPath
    if ($staged.Length -ne $backup.Length) { throw 'Backup copy size mismatch.' }
    Write-RunLog "Backup copy verified byte-for-byte." 'OK'

    $master = Open-SqlConnection $Config 'master'
    try {
        $backupLiteral = ConvertTo-SqlLiteral $script:StagedBackupPath
        Write-RunLog 'Running RESTORE VERIFYONLY. Current SQL database is still untouched.'
        Invoke-SqlNonQuery $master "RESTORE VERIFYONLY FROM DISK=$backupLiteral;" 0
        Write-RunLog 'Backup validation passed.' 'OK'

        $fileCommand = $master.CreateCommand()
        $fileCommand.CommandTimeout = 0
        $fileCommand.CommandText = "RESTORE FILELISTONLY FROM DISK=$backupLiteral;"
        $reader = $fileCommand.ExecuteReader()
        $backupFiles = New-Object System.Collections.Generic.List[object]
        while ($reader.Read()) {
            $backupFiles.Add([pscustomobject]@{
                LogicalName = [string]$reader['LogicalName']
                Type = [string]$reader['Type']
                Size = [long]$reader['Size']
            })
        }
        $reader.Close()
        $fileCommand.Dispose()

        $dataFiles = @($backupFiles | Where-Object Type -eq 'D')
        $logFiles = @($backupFiles | Where-Object Type -eq 'L')
        if ($dataFiles.Count -ne 1 -or $logFiles.Count -ne 1) {
            throw "Automation currently requires exactly one data and one log file. Backup has D=$($dataFiles.Count), L=$($logFiles.Count)."
        }

        $databaseLiteral = ConvertTo-SqlLiteral ([string]$Config.SqlServer.Database)
        $pathCommand = $master.CreateCommand()
        $pathCommand.CommandText = "SELECT type,physical_name FROM sys.master_files WHERE database_id=DB_ID($databaseLiteral) ORDER BY file_id;"
        $pathReader = $pathCommand.ExecuteReader()
        $currentDataPath = $null
        $currentLogPath = $null
        while ($pathReader.Read()) {
            if ([int]$pathReader['type'] -eq 0 -and $null -eq $currentDataPath) { $currentDataPath = [string]$pathReader['physical_name'] }
            if ([int]$pathReader['type'] -eq 1 -and $null -eq $currentLogPath) { $currentLogPath = [string]$pathReader['physical_name'] }
        }
        $pathReader.Close()
        $pathCommand.Dispose()
        if ([string]::IsNullOrEmpty($currentDataPath) -or [string]::IsNullOrEmpty($currentLogPath)) {
            throw 'Could not determine current SQL data/log paths.'
        }

        $dbName = Quote-SqlIdentifier ([string]$Config.SqlServer.Database)
        $restoreSql = @"
BEGIN TRY
    ALTER DATABASE $dbName SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    RESTORE DATABASE $dbName FROM DISK=$backupLiteral WITH REPLACE,
        MOVE $(ConvertTo-SqlLiteral $dataFiles[0].LogicalName) TO $(ConvertTo-SqlLiteral $currentDataPath),
        MOVE $(ConvertTo-SqlLiteral $logFiles[0].LogicalName) TO $(ConvertTo-SqlLiteral $currentLogPath),
        RECOVERY, STATS=5;
    ALTER DATABASE $dbName SET MULTI_USER;
END TRY
BEGIN CATCH
    IF DB_ID($databaseLiteral) IS NOT NULL AND DATABASEPROPERTYEX($databaseLiteral,'Status')='ONLINE'
        ALTER DATABASE $dbName SET MULTI_USER;
    THROW;
END CATCH;
"@
        Write-RunLog "Restoring database $($Config.SqlServer.Database). Active users will be disconnected."
        $master.FireInfoMessageEventOnUserErrors = $true
        $master.add_InfoMessage({ param($sender, $eventArgs) Write-Host $eventArgs.Message })
        Invoke-SqlNonQuery $master $restoreSql 0
        $state = [string](Invoke-SqlScalar $master "SELECT state_desc + '|' + user_access_desc FROM sys.databases WHERE name=$databaseLiteral;")
        if ($state -ne 'ONLINE|MULTI_USER') { throw "Restore completed but database state is $state" }
        $script:RestoreSucceeded = $true
        Write-RunLog "Restore successful. Database state: $state" 'OK'
    }
    finally { $master.Close(); $master.Dispose() }
    return [pscustomobject]@{ AlreadyProcessed=$false; Info=$processingInfo }
}

function Get-SqlColumns {
    param($Connection, [string]$TableName)
    $command = $Connection.CreateCommand()
    $command.CommandText = "SELECT COLUMN_NAME,DATA_TYPE,ORDINAL_POSITION FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA='dbo' AND TABLE_NAME=@name ORDER BY ORDINAL_POSITION;"
    $null = $command.Parameters.Add('@name',[System.Data.SqlDbType]::NVarChar,128)
    $command.Parameters['@name'].Value = $TableName
    $reader = $command.ExecuteReader()
    $result = New-Object System.Collections.Generic.List[object]
    while ($reader.Read()) {
        $result.Add([pscustomobject]@{ Name=[string]$reader[0]; Type=[string]$reader[1]; Ordinal=[int]$reader[2] })
    }
    $reader.Close(); $command.Dispose()
    return $result.ToArray()
}

function Get-MySqlColumns {
    param($Config, [string]$TableName)
    $safe = $TableName.Replace("'", "''")
    $lines = Invoke-MySqlLines $Config "SELECT COLUMN_NAME,DATA_TYPE,ORDINAL_POSITION FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='$safe' ORDER BY ORDINAL_POSITION;"
    $result = New-Object System.Collections.Generic.List[object]
    foreach ($line in $lines) {
        $parts = $line -split "`t", 3
        if ($parts.Count -eq 3) { $result.Add([pscustomobject]@{Name=$parts[0];Type=$parts[1];Ordinal=[int]$parts[2]}) }
    }
    return $result.ToArray()
}

function Compare-OneTable {
    param($Config, $SqlConnection, $TableConfig)
    $localName = [string]$TableConfig.LocalName
    $sqlName = [string]$TableConfig.SqlName
    Write-RunLog "Comparing $localName <-> dbo.$sqlName"
    Write-ReportLine ''
    Write-ReportLine ('=' * 100)
    Write-ReportLine "TABLE: $localName  <->  dbo.$sqlName"
    Write-ReportLine ('=' * 100)

    $localColumns = @(Get-MySqlColumns $Config $localName)
    $sqlColumns = @(Get-SqlColumns $SqlConnection $sqlName)
    if ($localColumns.Count -eq 0) { Write-ReportLine 'ERROR: Local table missing.'; return }
    if ($sqlColumns.Count -eq 0) { Write-ReportLine 'ERROR: SQL table missing.'; return }

    $localLookup = @{}
    foreach ($column in $localColumns) { $localLookup[$column.Name.ToLowerInvariant()] = $column }
    $sqlLookup = @{}
    foreach ($column in $sqlColumns) { $sqlLookup[$column.Name.ToLowerInvariant()] = $column }
    $commonColumns = @($localColumns | Where-Object { $sqlLookup.ContainsKey($_.Name.ToLowerInvariant()) } | ForEach-Object Name)
    $localOnly = @($localColumns | Where-Object { -not $sqlLookup.ContainsKey($_.Name.ToLowerInvariant()) } | ForEach-Object Name)
    $sqlOnly = @($sqlColumns | Where-Object { -not $localLookup.ContainsKey($_.Name.ToLowerInvariant()) } | ForEach-Object Name)
    $expectedLocalOnly = @($TableConfig.ExpectedLocalOnlyColumns | ForEach-Object { [string]$_ })
    $unexpectedLocalOnly = @($localOnly | Where-Object { $expectedLocalOnly -inotcontains $_ })
    Write-ReportLine "Local columns: $($localColumns.Count); SQL columns: $($sqlColumns.Count); Common: $($commonColumns.Count)"
    Write-ReportLine "Local-only columns: $(if($localOnly.Count){$localOnly -join ', '}else{'None'})"
    Write-ReportLine "Unexpected local-only columns: $(if($unexpectedLocalOnly.Count){$unexpectedLocalOnly -join ', '}else{'None'})"
    Write-ReportLine "SQL-only columns: $(if($sqlOnly.Count){$sqlOnly -join ', '}else{'None'})"

    $keys = @($TableConfig.Keys | ForEach-Object { [string]$_ })
    $missingKeys = @($keys | Where-Object { -not $localLookup.ContainsKey($_.ToLowerInvariant()) -or -not $sqlLookup.ContainsKey($_.ToLowerInvariant()) })
    if ($missingKeys.Count -gt 0) {
        Write-ReportLine "ERROR: Configured key columns missing: $($missingKeys -join ', '). Row-level comparison skipped."
        return
    }
    Write-ReportLine "Comparison key: $($keys -join ' + ')"

    $localWhereClause = ''
    $sqlWhereClause = ''
    if ($localName -ieq 'address') {
        $localWhereClause = " WHERE UPPER(TRIM($(Quote-MySqlIdentifier 'Atype'))) IN ('C','D')"
        $sqlWhereClause = " WHERE UPPER(LTRIM(RTRIM($(Quote-SqlIdentifier 'Atype')))) IN ('C','D')"
        Write-ReportLine 'Row filter: Atype IN (C, D) only'
    }

    $localKeySet = New-Object 'System.Collections.Generic.HashSet[string]'
    [long]$localRows = 0
    [long]$localDuplicateKeys = 0
    # New-row-only mode: read only the configured keys from local MySQL.
    # Existing rows and their non-key values are never compared.
    $localSelect = ($keys | ForEach-Object { Quote-MySqlIdentifier $_ }) -join ','
    $localProcess = Start-MySqlQuery $Config "SELECT $localSelect FROM $(Quote-MySqlIdentifier $localName)$localWhereClause;"
    while (($line = $localProcess.StandardOutput.ReadLine()) -ne $null) {
        $localRows++
        $fields = @($line -split "`t", -1 | ForEach-Object { ConvertFrom-MySqlBatchField $_ })
        $canonical = New-CanonicalKey @($fields)
        if (-not $localKeySet.Add($canonical)) { $localDuplicateKeys++ }
    }
    Complete-MySqlQuery $localProcess

    $selectSql = ($commonColumns | ForEach-Object { Quote-SqlIdentifier $_ }) -join ','
    $command = $SqlConnection.CreateCommand()
    $command.CommandTimeout = 0
    $command.CommandText = "SELECT $selectSql FROM [dbo].$(Quote-SqlIdentifier $sqlName)$sqlWhereClause;"
    $reader = $command.ExecuteReader()
    $ordinals = @{}
    for ($i=0; $i -lt $reader.FieldCount; $i++) { $ordinals[$reader.GetName($i).ToLowerInvariant()] = $i }
    [long]$sqlRows = 0
    [long]$sqlDuplicateKeys = 0
    [long]$sqlOnlyRows = 0
    [long]$writtenSqlOnly = 0
    $sqlSeenKeys = New-Object 'System.Collections.Generic.HashSet[string]'
    $maxDetails = [long]$Config.MaxDetailedRowsPerTable
    $sqlOnlyDetails = New-Object System.Collections.Generic.List[string]
    $panelRateCounts = @{}
    while ($reader.Read()) {
        $sqlRows++
        $keyValues = foreach ($key in $keys) { $reader.GetValue($ordinals[$key.ToLowerInvariant()]) }
        $canonical = New-CanonicalKey @($keyValues)
        if (-not $sqlSeenKeys.Add($canonical)) { $sqlDuplicateKeys++ }
        if (-not $localKeySet.Contains($canonical)) {
            $sqlOnlyRows++
            if ($localName -ieq 'panelrates' -and $ordinals.ContainsKey('compcatid')) {
                $groupValue = Format-ValueForReport $reader.GetValue($ordinals['compcatid']) 300
                if (-not $panelRateCounts.ContainsKey($groupValue)) { $panelRateCounts[$groupValue] = 0L }
                $panelRateCounts[$groupValue]++
            }
            if ($maxDetails -eq 0 -or $writtenSqlOnly -lt $maxDetails) {
                $rowNumber = $sqlOnlyRows
                $lines = New-Object System.Collections.Generic.List[string]
                $lines.Add(('NEW ROW #{0}' -f $rowNumber))
                foreach ($column in $commonColumns) {
                    $lines.Add(('  {0,-32}: {1}' -f $column,(Format-ValueForReport $reader.GetValue($ordinals[$column.ToLowerInvariant()]))))
                }
                $sqlOnlyDetails.Add($lines -join [Environment]::NewLine)
                $writtenSqlOnly++
            }
        }
    }
    $reader.Close(); $command.Dispose()

    Write-ReportLine "Local rows: $localRows; SQL rows: $sqlRows"
    Write-ReportLine "Local distinct keys: $($localKeySet.Count); SQL distinct keys: $($sqlSeenKeys.Count)"
    Write-ReportLine "Duplicate-key rows beyond first: Local=$localDuplicateKeys; SQL=$sqlDuplicateKeys"
    Write-ReportLine "NEW ROWS (present in SQL Server, absent in target MySQL): $sqlOnlyRows"
    Write-ReportLine 'Existing/matched rows: ignored (values not compared)'
    Write-ReportLine 'Local-only rows: ignored'
    if ($localDuplicateKeys -gt 0 -or $sqlDuplicateKeys -gt 0) {
        Write-ReportLine 'WARNING: Duplicate keys make row-level interpretation ambiguous. Review the configured key.'
    }

    Write-ReportLine ''
    if ($localName -ieq 'panelrates') {
        Write-ReportLine 'PANELRATES NEW ROW COUNT BY CompCatID:'
        if ($panelRateCounts.Count -eq 0) { Write-ReportLine 'None' }
        else {
            foreach ($group in @($panelRateCounts.GetEnumerator() | Sort-Object { if ($_.Key -match '^-?\d+$') { [long]$_.Key } else { [long]::MaxValue } }, Key)) {
                Write-ReportLine ('  CompCatID={0,-12} NewRows={1}' -f $group.Key,$group.Value)
            }
        }
        Write-ReportLine ''
    }
    Write-ReportLine 'NEW ROW DETAILS:'
    if ($sqlOnlyDetails.Count -eq 0) { Write-ReportLine 'None' }
    else { foreach ($detail in $sqlOnlyDetails) { Write-ReportLine $detail; Write-ReportLine '' } }
    if ($writtenSqlOnly -lt $sqlOnlyRows) { Write-ReportLine "... $($sqlOnlyRows-$writtenSqlOnly) additional new rows omitted by MaxDetailedRowsPerTable." }

    Write-RunLog "Completed ${localName}: Local=$localRows SQL=$sqlRows New=$sqlOnlyRows" 'OK'
}

if (-not (Test-Path -LiteralPath $ConfigPath)) { throw "Config file not found: $ConfigPath" }
$config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$reportFolder = [string]$config.ReportFolder
if (-not (Test-Path -LiteralPath $reportFolder)) { $null = New-Item -ItemType Directory -Path $reportFolder -Force }
$reportPath = Join-Path $reportFolder ('Labmate_Comparison_{0:yyyyMMdd_HHmmss}.txt' -f (Get-Date))
$utf8 = New-Object System.Text.UTF8Encoding($false)
$script:ReportWriter = New-Object System.IO.StreamWriter($reportPath,$false,$utf8)

try {
    Write-ReportLine 'LABMATE NEW SQL ROWS REPORT'
    Write-ReportLine "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')"
    Write-ReportLine "SQL Server: $($config.SqlServer.Server) / $($config.SqlServer.Database)"
    Write-ReportLine "Target MySQL: $($config.MySql.Host):$($config.MySql.Port) / $($config.MySql.Database)"
    Write-ReportLine 'Mode: Restore SQL backup, then report only rows present in SQL Server but absent in target MySQL.'
    Write-ReportLine 'Existing matched rows and target-only rows are ignored. Target MySQL is never modified.'
    Write-ReportLine ''

    $backupResult = $null
    if (-not $CompareOnly) { $backupResult = Stage-And-RestoreBackup $config -SkipConfirmation:$Force }
    else { Write-RunLog 'CompareOnly mode: SQL restore skipped.' 'WARN' }

    $tables = @($config.Tables)
    if ($OnlyTable.Count -gt 0) {
        $tables = @($tables | Where-Object { $OnlyTable -icontains $_.LocalName -or $OnlyTable -icontains $_.SqlName })
        if ($tables.Count -eq 0) { throw "No configured tables matched -OnlyTable: $($OnlyTable -join ', ')" }
    }

    $sql = Open-SqlConnection $config ([string]$config.SqlServer.Database)
    try {
        $sqlServerName = [string](Invoke-SqlScalar $sql 'SELECT @@SERVERNAME;')
        $mysqlIdentity = @(Invoke-MySqlLines $config 'SELECT CONCAT(@@hostname,''|'',DATABASE());')[0]
        Write-RunLog "Connections verified: SQL=$sqlServerName; MySQL=$mysqlIdentity" 'OK'
        foreach ($table in $tables) { Compare-OneTable $config $sql $table }
    }
    finally { $sql.Close(); $sql.Dispose() }

    if ($null -ne $backupResult -and -not [bool]$backupResult.AlreadyProcessed) {
        Save-ProcessedBackupState $backupResult.Info $config
    }

    Write-ReportLine ''
    Write-ReportLine 'FINAL STATUS: SUCCESS'
    Write-RunLog "Report created: $reportPath" 'OK'
}
catch {
    Write-ReportLine ''
    Write-ReportLine 'FINAL STATUS: FAILED'
    Write-ReportLine $_.Exception.ToString()
    Write-RunLog $_.Exception.Message 'ERROR'
    throw
}
finally {
    if ($script:RestoreSucceeded -and [bool]$config.SqlServer.RemoveWorkingBackupAfterSuccess -and
        -not [string]::IsNullOrEmpty($script:StagedBackupPath) -and (Test-Path -LiteralPath $script:StagedBackupPath)) {
        try {
            Remove-Item -LiteralPath $script:StagedBackupPath -Force
            Write-RunLog 'Removed temporary SQL Backup-folder copy.' 'OK'
        }
        catch { Write-RunLog "Could not remove temporary SQL backup copy: $($_.Exception.Message)" 'WARN' }
    }
    if ($null -ne $script:ReportWriter) { $script:ReportWriter.Dispose() }
}

Write-Host ''
Write-Host "Done. Report: $reportPath" -ForegroundColor Green
