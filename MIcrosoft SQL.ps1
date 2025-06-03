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

    Log verbose "-Connection=$Connection -TestConnection=$TestConnection -Configuration=$Configuration -ConnectionParams='$ConnectionParams'"
    
    if ($Connection) {
        @(
            @{
                name = 'server'
                type = 'textbox'
                label = 'Server'
                tooltip = 'Name of Microsoft SQL server'
                value = ''
            }
            @{
                name = 'database'
                type = 'textbox'
                label = 'Database'
                tooltip = 'Name of Microsoft SQL database'
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
                tooltip = 'User account name to access Microsoft SQL server'
                value = ''
                hidden = 'use_svc_account_creds'
            }
            @{
                name = 'password'
                type = 'textbox'
                password = $true
                label = 'Password'
                label_indent = $true
                tooltip = 'User account password to access Microsoft SQL server'
                value = ''
                hidden = 'use_svc_account_creds'
            }
            @{
                name = 'nr_of_sessions'
                type = 'textbox'
                label = 'Max. number of simultaneous sessions'
                tooltip = ''
                value = 5
            }
            @{
                name = 'sessions_idle_timeout'
                type = 'textbox'
                label = 'Session cleanup idle time (minutes)'
                tooltip = ''
                value = 30
            }
        )
    }

    if ($TestConnection) {
        Open-MsSqlConnection $ConnectionParams
    }

    if ($Configuration) {
        @()
    }

    Log verbose "Done"
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
        [switch] $Force
    )

    if (!$Force -and $Global:SqlInfoCache.Ts -and ((Get-Date) - $Global:SqlInfoCache.Ts).TotalMilliseconds -le [Int32]600000) {
        return
    }

    # Refresh cache
    $sql_command = New-MsSqlCommand "
        SELECT 
        ss.NAME + '.[' + st.NAME + ']' AS full_object_name, 
        (
            CASE WHEN st.type = 'U' THEN 'Table' WHEN st.type = 'V' THEN 'View' ELSE 'Other' END
        ) AS object_type, 
        sc.NAME AS column_name, 
        Cast(
            CASE WHEN pk.column_name IS NULL THEN 0 ELSE 1 END AS BIT
        ) AS is_primary_key, 
        sc.is_identity, 
        sc.is_computed, 
        Cast(
            CASE WHEN pk.column_name IS NULL 
            AND (
            sc.is_nullable = 1 
            OR sc.default_object_id <> 0
            ) THEN 1 ELSE 0 END AS BIT
        ) AS is_nullable 
        FROM 
        sys.schemas AS ss 
        INNER JOIN (
            SELECT 
            NAME, 
            object_id, 
            schema_id, 
            type 
            FROM 
            sys.tables 
            UNION ALL 
            SELECT 
            NAME, 
            object_id, 
            schema_id, 
            type 
            FROM 
            sys.views
        ) AS st ON ss.schema_id = st.schema_id 
        INNER JOIN sys.columns AS sc ON st.object_id = sc.object_id 
        LEFT JOIN (
            SELECT 
            CCU.* 
            FROM 
            information_schema.constraint_column_usage AS CCU 
            INNER JOIN information_schema.table_constraints AS TC ON CCU.constraint_name = TC.constraint_name 
            WHERE 
            TC.constraint_type = 'PRIMARY KEY'
        ) AS pk ON ss.NAME = pk.table_schema 
        AND st.NAME = pk.table_name 
        AND sc.NAME = pk.column_name 
        UNION ALL 
        SELECT 
        ss.NAME + '.[' + sp.NAME + ']' AS full_object_name, 
        'Stored Procedure' AS object_type, 
        p.NAME AS column_name, 
        '0' AS is_primary_key, 
        '0' AS is_identity, 
        '0' AS is_computed, 
        p.is_nullable AS is_nullable 
        FROM 
        sys.procedures sp 
        INNER JOIN sys.schemas ss ON sp.schema_id = ss.schema_id 
        INNER JOIN sys.parameters p ON sp.object_id = p.object_id 
        ORDER BY 
        full_object_name, 
        column_name"

    $objects = New-Object System.Collections.ArrayList
    $object = @{}

    # Process in one pass
    Invoke-MsSqlCommand $sql_command | ForEach-Object {
        if ($_.full_object_name -ne $object.full_name) {
            if ($object.full_name -ne $null) {
                $objects.Add($object) | Out-Null
            }

            $object = @{
                full_name = $_.full_object_name
                type      = $_.object_type
                columns   = New-Object System.Collections.ArrayList
            }
        }

        $object.columns.Add(@{
            name           = $_.column_name
            is_primary_key = $_.is_primary_key
            is_identity    = $_.is_identity
            is_computed    = $_.is_computed
            is_nullable    = $_.is_nullable
        }) | Out-Null
    }

    if ($object.full_name -ne $null) {
        $objects.Add($object) | Out-Null
    }

    Dispose-MsSqlCommand $sql_command

    $Global:SqlInfoCache.Objects = $objects
    $Global:SqlInfoCache.Ts = Get-Date
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

    Log verbose "-Class='$Class' -Operation='$Operation' -GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($Class -eq '') {

        if ($GetMeta) {
            #
            # Get all tables and views in database
            #

            Open-MsSqlConnection $SystemParams

            Fill-SqlInfoCache -Force

            #
            # Output list of supported operations per table/view (named Class)
            #

            @(
                foreach ($object in $Global:SqlInfoCache.Objects) {
                    $primary_keys = $object.columns | Where-Object { $_.is_primary_key } | ForEach-Object { $_.name }

                    if ($object.type -eq 'Stored Procedure') {
                        [ordered]@{
                            Class = $object.full_name
                            Operation = 'Execute'
                        }

                        [ordered]@{
                            Class = $object.full_name
                            Operation = 'Read'
                            'Source type' = $object.type
                            'Primary key' = ""
                            'Supported operations' = "CR"
                        }
                    }
                    elseif ($object.type -eq 'View') {
                        # Non-tables only support 'Read'
                        [ordered]@{
                            Class = $object.full_name
                            Operation = 'Read'
                            'Source type' = $object.type
                            'Primary key' = $primary_keys -join ', '
                            'Supported operations' = 'R'
                        }
                    }
                    else {
                        [ordered]@{
                            Class = $object.full_name
                            Operation = 'Create'
                        }

                        [ordered]@{
                            Class = $object.full_name
                            Operation = 'Read'
                            'Source type' = $object.type
                            'Primary key' = $primary_keys -join ', '
                            'Supported operations' = "CR$(if ($primary_keys) { 'UD' } else { '' })"
                        }

                        if ($primary_keys) {
                            # Only supported if primary keys are present
                            [ordered]@{
                                Class = $object.full_name
                                Operation = 'Update'
                            }

                            [ordered]@{
                                Class = $object.full_name
                                Operation = 'Delete'
                            }
                        }
                    }
                }
            )

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

            Fill-SqlInfoCache

            $columns = ($Global:SqlInfoCache.Objects | Where-Object { $_.full_name -eq $Class }).columns

            switch ($Operation) {
                'Execute' {
                    @{
                        semantics = ''
                        parameters = @(
                            $columns | ForEach-Object {
                                @{
                                    name = $_.name;
                                    allowance = if ($_.is_identity -or $_.is_computed) { 'prohibited' } elseif (! $_.is_nullable) { 'mandatory' } else { 'optional' }
                                }
                            }
                        )
                    }
                    break
                }

                'Create' {
                    @{
                        semantics = 'create'
                        parameters = @(
                            $columns | ForEach-Object {
                                @{
                                    name = $_.name;
                                    allowance = if ($_.is_identity -or $_.is_computed) { 'prohibited' } elseif (! $_.is_nullable) { 'mandatory' } else { 'optional' }
                                }
                            }
                        )
                    }
                    break
                }

                'Read' {
                    @(
                        @{
                            name = 'select_distinct'
                            type = 'checkbox'
                            label = 'Distinct Rows'
                            tooltip = 'Apply Distinct to select'
                            value = $false
                        }
                        @{
                            name = 'where_clause'
                            type = 'textbox'
                            label = 'Filter (SQL where-clause)'
                            tooltip = 'Applied SQL where-clause'
                            value = ''
                        }
                        @{
                            name = 'selected_columns'
                            type = 'grid'
                            label = 'Include columns'
                            tooltip = 'Selected columns'
                            table = @{
                                rows = @($columns | ForEach-Object {
                                    @{
                                        name = $_.name
                                        config = @(
                                            if ($_.is_primary_key) { 'Primary key' }
                                            if ($_.is_identity)    { 'Generated' }
                                            if ($_.is_computed)    { 'Computed' }
                                            if ($_.is_nullable)    { 'Nullable' }
                                        ) -join ' | '
                                    }
                                })
                                settings_grid = @{
                                    selection = 'multiple'
                                    key_column = 'name'
                                    checkbox = $true
                                    filter = $true
                                    columns = @(
                                        @{
                                            name = 'name'
                                            display_name = 'Name'
                                        }
                                        @{
                                            name = 'config'
                                            display_name = 'Configuration'
                                        }
                                    )
                                }
                            }
                            value = @($columns | ForEach-Object { $_.name })
                        }
                    )
                    break
                }

                'Update' {
                    @{
                        semantics = 'update'
                        parameters = @(
                            $columns | ForEach-Object {
                                @{
                                    name = $_.name;
                                    allowance = if ($_.is_primary_key) { 'mandatory' } else { 'optional' }
                                }
                            }
                            @{
                                name = '*'
                                allowance = 'prohibited'
                            }
                        )
                    }
                    break
                }

                'Delete' {
                    @{
                        semantics = 'delete'
                        parameters = @(
                            $columns | ForEach-Object {
                                if ($_.is_primary_key) {
                                    @{
                                        name = $_.name
                                        allowance = 'mandatory'
                                    }
                                }
                            }
                            @{
                                name = '*'
                                allowance = 'prohibited'
                            }
                        )
                    }
                    break
                }
            }

        }
        else {
            #
            # Execute function
            #

            Open-MsSqlConnection $SystemParams

            if (! $Global:ColumnsInfoCache[$Class]) {
                Fill-SqlInfoCache
                $columns = ($Global:SqlInfoCache.Objects | Where-Object { $_.full_name -eq $Class }).columns

                $Global:ColumnsInfoCache[$Class] = @{
                    columns = $columns
                    primary_keys = @($columns | Where-Object { $_.is_primary_key } | ForEach-Object { $_.name })
                    identity_col = @($columns | Where-Object { $_.is_identity    } | ForEach-Object { $_.name })[0]
                    type = ($Global:SqlInfoCache.Objects | Where-Object { $_.full_name -eq $Class }).type
                }
            }

            $primary_keys = $Global:ColumnsInfoCache[$Class].primary_keys
            $identity_col = $Global:ColumnsInfoCache[$Class].identity_col
            $type = $Global:ColumnsInfoCache[$Class].type

            $function_params = ConvertFrom-Json2 $FunctionParams

            # Replace $null by [System.DBNull]::Value
            $keys_with_null_value = @()
            foreach ($key in $function_params.Keys) { if ($function_params[$key] -eq $null) { $keys_with_null_value += $key } }
            foreach ($key in $keys_with_null_value) { $function_params[$key] = [System.DBNull]::Value }

            $sql_command = New-MsSqlCommand

            $projection = if ($function_params['selected_columns'].count -eq 0) { '*' } else { @($function_params['selected_columns'] | ForEach-Object { "[$_]" }) -join ', ' }

            if($function_params['select_distinct']) { $projection = "DISTINCT $($projection)" }

            switch ($Operation) {
                'Execute' {
                    $sql_command.CommandText = "
                        EXECUTE $Class 
                            $(@($function_params.Keys | ForEach-Object { AddParam-MsSqlCommand $sql_command $function_params[$_] }) -join ', ')
                    "
                    break
                }

                'Create' {
                    $filter = if ($identity_col) {
                                  "[$identity_col] = SCOPE_IDENTITY()"
                              }
                              elseif ($primary_keys) {
                                  @($primary_keys | ForEach-Object { "[$_] = $(AddParam-MsSqlCommand $sql_command $function_params[$_])" }) -join ' AND '
                              }
                              else {
                                  @($function_params.Keys | ForEach-Object { "[$_] = $(AddParam-MsSqlCommand $sql_command $function_params[$_])" }) -join ' AND '
                              }

                    $sql_command.CommandText = "
                        INSERT INTO $Class (
                            $(@($function_params.Keys | ForEach-Object { "[$_]" }) -join ', ')
                        )
                        VALUES (
                            $(@($function_params.Keys | ForEach-Object { AddParam-MsSqlCommand $sql_command $function_params[$_] }) -join ', ')
                        );
                        SELECT TOP(1)
                            $projection
                        FROM
                            $Class
                        WHERE
                            $filter
                    "
                    break
                }

                'Read' {
                    if($type -eq 'Stored Procedure') {
                        $projection = @($Global:ColumnsInfoCache[$Class].columns | ForEach-Object { "'' [$($_.name)]";}) -join ', '
                        $sql_command.CommandText = "SELECT $projection WHERE 1=2"
                    } else {
                        $filter = if ($function_params['where_clause'].length -eq 0) { '' } else { " WHERE $($function_params['where_clause'])" }

                        $sql_command.CommandText = "
                            SELECT
                                $projection
                            FROM
                                $Class$filter
                        "
                    }
                    
                    break
                }

                'Update' {
                    $filter = @($primary_keys | ForEach-Object { "[$_] = $(AddParam-MsSqlCommand $sql_command $function_params[$_])" }) -join ' AND '

                    $sql_command.CommandText = "
                        UPDATE TOP(1)
                            $Class
                        SET
                            $(@($function_params.Keys | ForEach-Object { if ($_ -notin $primary_keys) { "[$_] = $(AddParam-MsSqlCommand $sql_command $function_params[$_])" } }) -join ', ')
                        WHERE
                            $filter;
                        SELECT TOP(1)
                            $(@($function_params.Keys | ForEach-Object { "[$_]" }) -join ', ')
                        FROM
                            $Class
                        WHERE
                            $filter
                    "
                    break
                }

                'Delete' {
                    $filter = @($primary_keys | ForEach-Object { "[$_] = $(AddParam-MsSqlCommand $sql_command $function_params[$_])" }) -join ' AND '

                    $sql_command.CommandText = "
                        DELETE TOP(1)
                            $Class
                        WHERE
                            $filter
                    "
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
                    $rv = Invoke-MsSqlCommand $sql_command $deparam_command | ForEach-Object { $_ }
                    LogIO info ($deparam_command -split ' ')[0] -Out $rv

                    $rv
                }
            }

            Dispose-MsSqlCommand $sql_command

        }

    }

    Log verbose "Done"
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
    $param_value = if ($Param -isnot [system.array]) { $Param } else { $Param | ConvertTo-Json -Compress -Depth 32 }

    $SqlCommand.Parameters.AddWithValue($param_name, $param_value) | Out-Null

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
    function Invoke-MsSqlCommand-ExecuteReader {
        param (
            [System.Data.SqlClient.SqlCommand] $SqlCommand
        )

        $data_reader = $SqlCommand.ExecuteReader()
        $column_names = @($data_reader.GetSchemaTable().ColumnName)

        if ($column_names) {

            $hash_table = [ordered]@{}

            foreach ($column_name in $column_names) {
                $hash_table[$column_name] = ""
            }

            $obj = New-Object -TypeName PSObject -Property $hash_table

            # Read data
			if($data_reader.HasRows) {
				while ($data_reader.Read()) {
					foreach ($column_name in $column_names) {
						$obj.$column_name = if ($data_reader[$column_name] -is [System.DBNull]) { $null } else { $data_reader[$column_name] }
					}

					# Output data
					$obj
				}
			} else { [PSCustomObject]$hash_table }
        }

        $data_reader.Close()
    }

    # Streaming
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
        Log verbose "MsSqlConnection connection parameters changed"
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
        Log verbose "Opening MsSqlConnection '$connection_string'"

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

        Log verbose "Done"
    }
}


function Close-MsSqlConnection {
    if ($Global:MsSqlConnection) {
        Log verbose "Closing MsSqlConnection"

        try {
            $Global:MsSqlConnection.Close()
            $Global:MsSqlConnection = $null
        }
        catch {
            # Purposely ignoring errors
        }

        Log verbose "Done"
    }
}
