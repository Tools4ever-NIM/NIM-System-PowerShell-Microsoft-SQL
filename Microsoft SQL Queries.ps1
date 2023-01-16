#
# Microsoft SQL.ps1 - IDM System PowerShell Script for Microsoft SQL Server.
#
# Any IDM System PowerShell Script is dot-sourced in a separate PowerShell context, after
# dot-sourcing the IDM Generic PowerShell Script '../Generic.ps1'.
#


$Log_MaskableKeys = @(
    'password'
)


#
# System functions
#

function Idm-SystemInfo {
    param (
        # Operations
        [switch] $Connection,
        [switch] $TestConnection,
        [switch] $Configuration,
        # Parameters
        [string] $ConnectionParams
    )

    Log info "-Connection=$Connection -TestConnection=$TestConnection -Configuration=$Configuration -ConnectionParams='$ConnectionParams'"
    
    if ($Connection) {
        @(
            @{
                name = 'server'
                type = 'textbox'
                label = 'Server'
                description = 'Name of Microsoft SQL server'
                value = ''
            }
            @{
                name = 'database'
                type = 'textbox'
                label = 'Database'
                description = 'Name of Microsoft SQL database'
                value = ''
            }
            @{
                name = 'use_svc_account_creds'
                type = 'checkbox'
                label = 'Use credentials of service account'
                value = $true
            }
            @{
                name = 'username'
                type = 'textbox'
                label = 'Username'
                label_indent = $true
                description = 'User account name to access Microsoft SQL server'
                value = ''
                hidden = 'use_svc_account_creds'
            }
            @{
                name = 'password'
                type = 'textbox'
                password = $true
                label = 'Password'
                label_indent = $true
                description = 'User account password to access Microsoft SQL server'
                value = ''
                hidden = 'use_svc_account_creds'
            }
            @{
                name = 'nr_of_sessions'
                type = 'textbox'
                label = 'Max. number of simultaneous sessions'
                description = ''
                value = 5
            }
            @{
                name = 'sessions_idle_timeout'
                type = 'textbox'
                label = 'Session cleanup idle time (minutes)'
                description = ''
                value = 30
            }
            @{
                name = 'table_1_name'
                type = 'textbox'
                label = 'Query 1 - Name of Table'
                description = ''
            }
            @{
                name = 'table_1_query'
                type = 'textbox'
                label = 'Query 1 - SQL Statement'
                description = ''
            }
            @{
                name = 'table_2_name'
                type = 'textbox'
                label = 'Query 2 - Name of Table'
                description = ''
            }
            @{
                name = 'table_2_query'
                type = 'textbox'
                label = 'Query 2 - SQL Statement'
                description = ''
            }
            @{
                name = 'table_3_name'
                type = 'textbox'
                label = 'Query 3 - Name of Table'
                description = ''
            }
            @{
                name = 'table_3_query'
                type = 'textbox'
                label = 'Query 3 - SQL Statement'
                description = ''
            }
            @{
                name = 'table_4_name'
                type = 'textbox'
                label = 'Query 4 - Name of Table'
                description = ''
            }
            @{
                name = 'table_4_query'
                type = 'textbox'
                label = 'Query 4 - SQL Statement'
                description = ''
            }
            @{
                name = 'table_5_name'
                type = 'textbox'
                label = 'Query 5 - Name of Table'
                description = ''
            }
            @{
                name = 'table_5_query'
                type = 'textbox'
                label = 'Query 5 - SQL Statement'
                description = ''
            }
            @{
                name = 'table_6_name'
                type = 'textbox'
                label = 'Query 6 - Name of Table'
                description = ''
            }
            @{
                name = 'table_6_query'
                type = 'textbox'
                label = 'Query 6 - SQL Statement'
                description = ''
            }
            @{
                name = 'table_7_name'
                type = 'textbox'
                label = 'Query 7 - Name of Table'
                description = ''
            }
            @{
                name = 'table_7_query'
                type = 'textbox'
                label = 'Query 7 - SQL Statement'
                description = ''
            }
            @{
                name = 'table_8_name'
                type = 'textbox'
                label = 'Query 8 - Name of Table'
                description = ''
            }
            @{
                name = 'table_8_query'
                type = 'textbox'
                label = 'Query 8 - SQL Statement'
                description = ''
            }
            @{
                name = 'table_9_name'
                type = 'textbox'
                label = 'Query 9 - Name of Table'
                description = ''
            }
            @{
                name = 'table_9_query'
                type = 'textbox'
                label = 'Query 9 - SQL Statement'
                description = ''
            }
            @{
                name = 'table_10_name'
                type = 'textbox'
                label = 'Query 10 - Name of Table'
                description = ''
            }
            @{
                name = 'table_10_query'
                type = 'textbox'
                label = 'Query 10 - SQL Statement'
                description = ''
            }
            @{
                name = 'table_11_name'
                type = 'textbox'
                label = 'Query 11 - Name of Table'
                description = ''
            }
            @{
                name = 'table_11_query'
                type = 'textbox'
                label = 'Query 11 - SQL Statement'
                description = ''
            }
            @{
                name = 'table_12_name'
                type = 'textbox'
                label = 'Query 12 - Name of Table'
                description = ''
            }
            @{
                name = 'table_12_query'
                type = 'textbox'
                label = 'Query 12 - SQL Statement'
                description = ''
            }
            @{
                name = 'table_13_name'
                type = 'textbox'
                label = 'Query 13 - Name of Table'
                description = ''
            }
            @{
                name = 'table_13_query'
                type = 'textbox'
                label = 'Query 13 - SQL Statement'
                description = ''
            }
            @{
                name = 'table_14_name'
                type = 'textbox'
                label = 'Query 14 - Name of Table'
                description = ''
            }
            @{
                name = 'table_14_query'
                type = 'textbox'
                label = 'Query 14 - SQL Statement'
                description = ''
            }
            @{
                name = 'table_15_name'
                type = 'textbox'
                label = 'Query 15 - Name of Table'
                description = ''
            }
            @{
                name = 'table_15_query'
                type = 'textbox'
                label = 'Query 15 - SQL Statement'
                description = ''
            }
            @{
                name = 'table_16_name'
                type = 'textbox'
                label = 'Query 16 - Name of Table'
                description = ''
            }
            @{
                name = 'table_16_query'
                type = 'textbox'
                label = 'Query 16 - SQL Statement'
                description = ''
            }
            @{
                name = 'table_17_name'
                type = 'textbox'
                label = 'Query 17 - Name of Table'
                description = ''
            }
            @{
                name = 'table_17_query'
                type = 'textbox'
                label = 'Query 17 - SQL Statement'
                description = ''
            }
            @{
                name = 'table_18_name'
                type = 'textbox'
                label = 'Query 18 - Name of Table'
                description = ''
            }
            @{
                name = 'table_18_query'
                type = 'textbox'
                label = 'Query 18 - SQL Statement'
                description = ''
            }
            @{
                name = 'table_19_name'
                type = 'textbox'
                label = 'Query 19 - Name of Table'
                description = ''
            }
            @{
                name = 'table_19_query'
                type = 'textbox'
                label = 'Query 19 - SQL Statement'
                description = ''
            }
            @{
                name = 'table_20_name'
                type = 'textbox'
                label = 'Query 20 - Name of Table'
                description = ''
            }
            @{
                name = 'table_20_query'
                type = 'textbox'
                label = 'Query 20 - SQL Statement'
                description = ''
            }

        )
    }

    if ($TestConnection) {
        Open-MsSqlConnection $ConnectionParams
    }

    if ($Configuration) {
        @()
    }

    Log info "Done"
}


function Idm-OnUnload {
    Close-MsSqlConnection
}

#
# CRUD functions
#

$ColumnsInfoCache = @{}

$SqlInfoCache = @{}

function Fill-SqlInfoCache {
    param (
        [switch] $Force,
        [string] $Query,
        [string] $Class
    )

    # Refresh cache
    $sql_command = New-MsSqlCommand $Query

    $result = (Invoke-MsSqlCommand $sql_command) | Get-Member -MemberType Properties | Select-Object Name

    Dispose-MsSqlCommand $sql_command

    $objects = New-Object System.Collections.ArrayList
    $object = @{}
    # Process in one pass
    foreach ($row in $result) {
            $object = @{
                full_name = $Class
                type      = 'Query'
                columns   = New-Object System.Collections.ArrayList
            }

        $object.columns.Add(@{
            name           = $row.Name
            is_primary_key = $false
            is_identity    = $false
            is_computed    = $false
            is_nullable    = $true
        }) | Out-Null
    }

    if ($object.full_name -ne $null) {
        $objects.Add($object) | Out-Null
    }
    @($objects)
}


function Idm-Dispatcher {
    param (
        # Optional Class/Operation
        [string] $Class,
        [string] $Operation,
        # Mode
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-Class='$Class' -Operation='$Operation' -GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $connection_params = ConvertFrom-Json2 $SystemParams

    if ($Class -eq '') {

        if ($GetMeta) {
            #
            # Get all tables and views in database
            #

            Open-MsSqlConnection $SystemParams
            
            #
            # Output list of supported operations per table/view (named Class)
            #
            for (($i = 0), ($j = 0); $i -lt 21; $i++)
            {
                if($connection_params."table_$($i)_name".length -gt 0)
                {
                    @(
                        [ordered]@{
                            Class = $connection_params."table_$($i)_name"
                            Operation = 'Read'
                            'Source type' = 'Query'
                            'Primary key' = ''
                            'Supported operations' = 'R'
                            'Query' = $connection_params."table_$($i)_query"
                        }
                    )
                    }
            }
        }
        else {
            # Purposely no-operation.
        }

    }
    else {

        if ($GetMeta) {
            #
            # Get meta data
            #

            Open-MsSqlConnection $SystemParams

            @()

        }
        else {
            #
            # Execute function
            #

            Open-MsSqlConnection $SystemParams

            Log debug "MY CLASS NAME IS $($class)"

            for (($i = 0), ($j = 0); $i -lt 21; $i++)
            {
                if($connection_params."table_$($i)_name" -eq $class)
                {
                    $class_query = $connection_params."table_$($i)_query"
                    break
                }
            }
            
            $columns = Fill-SqlInfoCache -Query $class_query
            $Global:ColumnsInfoCache[$Class] = @{
                primary_keys = @($columns | Where-Object { $_.is_primary_key } | ForEach-Object { $_.name })
                identity_col = @($columns | Where-Object { $_.is_identity    } | ForEach-Object { $_.name })[0]
            }

            $primary_keys = $Global:ColumnsInfoCache[$Class].primary_keys
            $identity_col = $Global:ColumnsInfoCache[$Class].identity_col

            $function_params = ConvertFrom-Json2 $FunctionParams

            # Replace $null by [System.DBNull]::Value
            $keys_with_null_value = @()
            foreach ($key in $function_params.Keys) { if ($function_params[$key] -eq $null) { $keys_with_null_value += $key } }
            foreach ($key in $keys_with_null_value) { $function_params[$key] = [System.DBNull]::Value }

            $sql_command = New-MsSqlCommand

            $projection = if ($function_params['selected_columns'].count -eq 0) { '*' } else { @($function_params['selected_columns'] | ForEach-Object { "[$_]" }) -join ', ' }

            switch ($Operation) {
                'Read' {
                    $sql_command.CommandText = $class_query
                    break
                }
            }

            if ($sql_command.CommandText) {
                $deparam_command = DeParam-MsSqlCommand $sql_command

                LogIO info ($deparam_command -split ' ')[0] -In -Command $deparam_command

                if ($Operation -eq 'Read') {
                    # Streamed output
                    Invoke-MsSqlCommand $sql_command $deparam_command
                }
                else {
                    # Log output
                    $rv = Invoke-MsSqlCommand $sql_command $deparam_command
                    LogIO info ($deparam_command -split ' ')[0] -Out $rv

                    $rv
                }
            }

            Dispose-MsSqlCommand $sql_command

        }

    }

    Log info "Done"
}


#
# Helper functions
#

function New-MsSqlCommand {
    param (
        [string] $CommandText
    )

    New-Object System.Data.SqlClient.SqlCommand($CommandText, $Global:MsSqlConnection)
}


function Dispose-MsSqlCommand {
    param (
        [System.Data.SqlClient.SqlCommand] $SqlCommand
    )

    $SqlCommand.Dispose()
}


function AddParam-MsSqlCommand {
    param (
        [System.Data.SqlClient.SqlCommand] $SqlCommand,
        $Param
    )

    $param_name = "@param$($SqlCommand.Parameters.Count)_"
    $SqlCommand.Parameters.AddWithValue($param_name, $Param) | Out-Null

    return $param_name
}


function DeParam-MsSqlCommand {
    param (
        [System.Data.SqlClient.SqlCommand] $SqlCommand
    )

    $deparam_command = $SqlCommand.CommandText

    foreach ($p in $SqlCommand.Parameters) {
        $value_txt = 
            if ($p.Value -eq [System.DBNull]::Value) {
                'NULL'
            }
            else {
                switch ($p.SqlDbType) {
                    { $_ -in @(
                        [System.Data.SqlDbType]::Char
                        [System.Data.SqlDbType]::Date
                        [System.Data.SqlDbType]::DateTime
                        [System.Data.SqlDbType]::DateTime2
                        [System.Data.SqlDbType]::DateTimeOffset
                        [System.Data.SqlDbType]::NChar
                        [System.Data.SqlDbType]::NText
                        [System.Data.SqlDbType]::NVarChar
                        [System.Data.SqlDbType]::Text
                        [System.Data.SqlDbType]::Time
                        [System.Data.SqlDbType]::VarChar
                        [System.Data.SqlDbType]::Xml
                    )} {
                        "'" + $p.Value.ToString().Replace("'", "''") + "'"
                        break
                    }
        
                    default {
                        $p.Value.ToString().Replace("'", "''")
                        break
                    }
                }
            }

        $deparam_command = $deparam_command.Replace($p.ParameterName, $value_txt)
    }

    # Make one single line
    @($deparam_command -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) -join ' '
}


function Invoke-MsSqlCommand {
    param (
        [System.Data.SqlClient.SqlCommand] $SqlCommand,
        [string] $DeParamCommand
    )

    # Streaming
    # ERAM dbo.Files (426.977 rows) execution time: ?
    function Invoke-MsSqlCommand-ExecuteReader {
        param (
            [System.Data.SqlClient.SqlCommand] $SqlCommand
        )
        Log debug "Start ExecuteReader"
        $data_reader = $SqlCommand.ExecuteReader()
        Log debug "Done ExecuteReader"
        Log debug "Start GetSchemaTable"
        $column_names = @($data_reader.GetSchemaTable().ColumnName)
        Log debug "Done GetSchemaTable"
        if ($column_names) {

            # Read data
            Log debug "Start While"
            while ($data_reader.Read()) {
                $hash_table = [ordered]@{}

                foreach ($column_name in $column_names) {
                    $hash_table[$column_name] = if ($data_reader[$column_name] -is [System.DBNull]) { $null } else { $data_reader[$column_name] }
                }

                # Output data
                New-Object -TypeName PSObject -Property $hash_table
            }
            Log debug "End While"
        }

        $data_reader.Close()
        Log debug "Close DataReader"
    }

    # Streaming
    # ERAM dbo.Files (426.977 rows) execution time: 16.7 s
    function Invoke-MsSqlCommand-ExecuteReader00 {
        param (
            [System.Data.SqlClient.SqlCommand] $SqlCommand
        )

        $data_reader = $SqlCommand.ExecuteReader()
        $column_names = @($data_reader.GetSchemaTable().ColumnName)

        if ($column_names) {

            # Initialize result
            $hash_table = [ordered]@{}

            for ($i = 0; $i -lt $column_names.Count; $i++) {
                $hash_table[$column_names[$i]] = ''
            }

            $result = New-Object -TypeName PSObject -Property $hash_table

            # Read data
            while ($data_reader.Read()) {
                foreach ($column_name in $column_names) {
                    $result.$column_name = $data_reader[$column_name]
                }

                # Output data
                $result
            }

        }

        $data_reader.Close()
    }

    # Streaming
    # ERAM dbo.Files (426.977 rows) execution time: 01:11.9 s
    function Invoke-MsSqlCommand-ExecuteReader01 {
        param (
            [System.Data.SqlClient.SqlCommand] $SqlCommand
        )

        $data_reader = $SqlCommand.ExecuteReader()
        $field_count = $data_reader.FieldCount

        while ($data_reader.Read()) {
            $hash_table = [ordered]@{}
        
            for ($i = 0; $i -lt $field_count; $i++) {
                $hash_table[$data_reader.GetName($i)] = $data_reader.GetValue($i)
            }

            # Output data
            New-Object -TypeName PSObject -Property $hash_table
        }

        $data_reader.Close()
    }

    # Non-streaming (data stored in $data_table)
    # ERAM dbo.Files (426.977 rows) execution time: 15.5 s
    function Invoke-MsSqlCommand-DataAdapter-DataTable {
        param (
            [System.Data.SqlClient.SqlCommand] $SqlCommand
        )

        $data_adapter = New-Object System.Data.SqlClient.SqlDataAdapter($SqlCommand)
        $data_table   = New-Object System.Data.DataTable
        $data_adapter.Fill($data_table) | Out-Null

        # Output data
        $data_table.Rows

        $data_table.Dispose()
        $data_adapter.Dispose()
    }

    # Non-streaming (data stored in $data_set)
    # ERAM dbo.Files (426.977 rows) execution time: 14.8 s
    function Invoke-MsSqlCommand-DataAdapter-DataSet {
        param (
            [System.Data.SqlClient.SqlCommand] $SqlCommand
        )

        $data_adapter = New-Object System.Data.SqlClient.SqlDataAdapter($SqlCommand)
        $data_set     = New-Object System.Data.DataSet
        $data_adapter.Fill($data_set) | Out-Null

        # Output data
        $data_set.Tables[0]

        $data_set.Dispose()
        $data_adapter.Dispose()
    }

    if (! $DeParamCommand) {
        $DeParamCommand = DeParam-MsSqlCommand $SqlCommand
    }

    Log debug $DeParamCommand

    try {
        Invoke-MsSqlCommand-ExecuteReader $SqlCommand
    }
    catch {
        Log error "Failed: $_"
        Write-Error $_
    }

    Log debug "Done"
}


function Open-MsSqlConnection {
    param (
        [string] $ConnectionParams
    )

    $connection_params = ConvertFrom-Json2 $ConnectionParams

    $cs_builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder

    # Use connection related parameters only
    $cs_builder['Data Source']     = $connection_params.server
    $cs_builder['Initial Catalog'] = $connection_params.database

    if ($connection_params.use_svc_account_creds) {
        $cs_builder['Integrated Security'] = 'SSPI'
    }
    else {
        $cs_builder['User ID']  = $connection_params.username
        $cs_builder['Password'] = $connection_params.password
    }

    $connection_string = $cs_builder.ConnectionString

    if ($Global:MsSqlConnection -and $connection_string -ne $Global:MsSqlConnectionString) {
        Log info "MsSqlConnection connection parameters changed"
        Close-MsSqlConnection
    }

    if ($Global:MsSqlConnection -and $Global:MsSqlConnection.State -ne 'Open') {
        Log warn "MsSqlConnection State is '$($Global:MsSqlConnection.State)'"
        Close-MsSqlConnection
    }

    if ($Global:MsSqlConnection) {
        #Log debug "Reusing MsSqlConnection"
    }
    else {
        Log info "Opening MsSqlConnection '$connection_string'"

        try {
            $connection = New-Object System.Data.SqlClient.SqlConnection($connection_string)
            $connection.Open()

            $Global:MsSqlConnection       = $connection
            $Global:MsSqlConnectionString = $connection_string

            $Global:ColumnsInfoCache = @{}
            $Global:SqlInfoCache = @{}
        }
        catch {
            Log error "Failed: $_"
            Write-Error $_
        }

        Log info "Done"
    }
}


function Close-MsSqlConnection {
    if ($Global:MsSqlConnection) {
        Log info "Closing MsSqlConnection"

        try {
            $Global:MsSqlConnection.Close()
            $Global:MsSqlConnection = $null
        }
        catch {
            # Purposely ignoring errors
        }

        Log info "Done"
    }
}
