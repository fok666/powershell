#
# TO DO
#
# [ok] inventário de tabelas, campos, measures, sources, visuals
# [ok] Listas de colunas por tipo (calculated / normal, type, size)
# [ok] estatisticas de colunas: count, distinct, avg size/sum
#
# [  ] relação measures vs colunas vs visuals
# [  ] detectar slices > 10.000
# [ok] detectar count  > 1.000.000
# [ok] detectar measures sem uso
# [ok] detectar colunas sem uso
# [ok] medir % calculated
# [ok] medir # visuals
#
# https://www.kasperonbi.com/dump-the-results-of-a-dax-query-to-csv-using-powershell/ 
# https://www.biinsight.com/connect-to-power-bi-desktop-model-from-excel-and-ssms/
#

# Carrega Assembly do SSAS (AMO)
#
# Analysis Services Management Objects (AMO) driver
# https://docs.microsoft.com/en-us/azure/analysis-services/analysis-services-data-providers#amo-and-adomd-nuget-packages
#
$loadInfo = [Reflection.Assembly]::LoadWithPartialName("Microsoft.AnalysisServices")

# Busca o processo atual:
$process = Get-Process -Id $PID
# Diminui prioridade do processo:
$process.PriorityClass = 'BelowNormal'
# sorteia uma CPU para executar o processo:
$cpu = Get-Random -Minimum 0 -Maximum $env:NUMBER_OF_PROCESSORS
# define a afinidade correspondente:
$process.ProcessorAffinity = [int]([math]::pow(2,$cpu))

# load ZIP methods
Add-Type -AssemblyName System.IO.Compression.FileSystem

# DAX Queries
$query_sum_col = @"
EVALUATE ROW("count", SUM('{0}'[{1}]))
"@

$query_row_col = @"
EVALUATE ROW("count", COUNTROWS('{0}'))
"@

$query_dcount_col = @"
EVALUATE ROW("count", DISTINCTCOUNT('{0}'[{1}]))
"@

# Functions que funcionam linha a linha (RBAR)
$RBAR_FUNCTIONS = @('AVERAGEX', 'CONCATENATEX', 'COUNTX', 'COUNTAX', 'GEOMEANX', 'MAXX', 'MEDIANX', 'MINX', 'PERCENTILEX.EXC', 'PERCENTILEX.INC', 'PRODUCTX', 'RANKX', 'SUMX')
$connectionString = ""
function get-pbixLayout($pbix)
{
    # open ZIP archive for reading
    $zip = [System.IO.Compression.ZipFile]::OpenRead($pbix)
    $outfile = $pbix.Replace(".pbix", ".layout.json")

    # Extract layout file
    $zip.Entries | Where-Object {$_.FullName -eq 'Report/Layout'} | % {
        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, $outfile, $true)
    }
    $zip.Dispose()

    $layout = Get-Content $outfile -Encoding Unicode | ConvertFrom-Json
    $layout | ConvertTo-Json -Depth 100 | Out-File $outfile -Encoding utf8
    return $layout
}

function get-powerBiProcess()
{
    # Processo do Power BI Desktop
    $p = ( get-process -Name PBIDesktop | select-object -First 1 )

    if( $null -eq $p ){
        Write-Host "Nenhum Power BI aberto"
        return $false
    }

    $cmd = ( Get-CimInstance Win32_Process -Filter "name = 'PBIDesktop.exe'" | select-object -First 1 ).CommandLine
    $pbix = $cmd.Remove(0, ($cmd.IndexOf('"', 1)) + 2 ).replace('"', '')

    if( -not ( Test-Path $pbix ) )
    {
        Write-Host "Arquivo não encontrado"
        return $false
    }
    $jsonOut = $pbix.Replace(".pbix", ".json")

    # Processo do Power BI SSAS
    $AsID = ( get-process -Name msmdsrv | select-object -First 1 ).id

    # porta TCP do processo
    $AsPort = ( Get-NetTCPConnection -OwningProcess $AsID -State Listen -LocalAddress 127.0.0.1 | Select-Object -First 1 ).LocalPort

    # SSAS URL
    $ServerName = "127.0.0.1:$AsPort"
    
    return @{pbix=$pbix; jsonOut=$jsonOut; serverName=$ServerName}
}

function connect-PowerBiAs($ServerName)
{
    # Cria instância do server
    $server = New-Object Microsoft.AnalysisServices.Server

    # Conecta no server
    $server.connect($ServerName)

    # Nome do database
    $Database = $server.Databases[0].Name

    # String de conexão
    $connectionString = "Provider=MSOLAP;Data Source=$ServerName;Initial Catalog=$Database;Timeout=0;"

    return @{server=$server; connectionString=$connectionString}
}

$pbiInfo = get-powerBiProcess

if( $False -ne $pbiInfo )
{
    $pbix = $pbiInfo.pbix
    $jsonOut = $pbiInfo.jsonOut
    $ServerName = $pbiInfo.serverName

    $layout = get-pbixLayout($pbix)
} else 
{
    return $false    
}

$connectionInfo = connect-PowerBiAs($ServerName)

$server = $connectionInfo.server
$connectionString = $connectionInfo.connectionString

Function dax($dax)
{
    # write-host $dax
    $result = $False

    $connection = New-Object -TypeName System.Data.OleDb.OleDbConnection
    $connection.ConnectionString = $connectionString

    $command = $connection.CreateCommand()
    $command.CommandText = $dax
    $command.CommandTimeout = 0

    # Write-Host "connected..."

    $adapter = New-Object -TypeName System.Data.OleDb.OleDbDataAdapter $command
    $dataset = New-Object -TypeName System.Data.DataSet

    # Write-Host "fill..."

    try
    {
        $adapter.Fill($dataset)

        if( $dataset.Tables[0].Rows.Count -eq 1 -and $dataset.Tables[0].Columns.Count -eq 1  ){
            $result = $dataset.Tables[0].Rows[0].Item(0)
        } else {
            $result = $dataset.Tables[0].Copy()
        }

        if( "" -eq $result -or $null -eq $result ){
            $result = 0
        }

        $dataset.Tables[0].Clear()
        $adapter.Dispose()
        $dataset.Clear()
        $dataset.Dispose()
    } catch {
        Write-Host "DAX Error: ", $dax -ForegroundColor DarkRed
        Write-Host "Exception: ", $_   -ForegroundColor DarkRed
    } finally {
        $connection.Close()
        $connection.Dispose()

        $dataset = $null
        $adapter = $null
        $command = $null
        $connection = $null
    }
    return $result
}

# https://github.com/CapitalOnTap/marqeta-swagger-json/blob/master/Get-ObjectPropertyValue.ps1
Function Get-ObjPropertyName {
	param(
		[CmdletBinding()]
		[Parameter(Mandatory=$true)]$Object,
		[Parameter(Mandatory=$true)]$Value,
		[Parameter(Mandatory=$false)]$ObjString
	)
	if(!($ObjString)){
		$ObjString = '$Object'
	}
	$ErrorActionPreference = 'SilentlyContinue'
	$return = @()
	(Invoke-Expression -Command $ObjString -ErrorAction SilentlyContinue)  | Get-Member -View All | Where-Object { ($_.MemberType -like "Property") -or ($_.MemberType -like 'NoteProperty')} | ForEach-Object {
		Write-Verbose  $($ObjString + ".$PropertyName")
		$PropertyName = $_.Name
		If((Invoke-Expression -Command $ObjString).$PropertyName -match $value){
			$return += New-Object PSObject -Property @{ 
				Name = $PropertyName
				Location = $ObjString + ".$PropertyName"
			}
		}
        if( $return.Count -gt 0 ){
		    return $return # Get-ObjPropertyName -Object $Object -Value $value -ObjString $($ObjString + ".$PropertyName")
        } else {
    		Get-ObjPropertyName -Object $Object -Value $value -ObjString $($ObjString + ".$PropertyName")
        }
	}
	return $return
}

# https://github.com/RamblingCookieMonster/PowerShell/blob/master/ConvertTo-FlatObject.ps1
Function ConvertTo-FlatObject {
    [cmdletbinding()]
    param(
        [parameter( Mandatory = $True, ValueFromPipeline = $True)]
        [PSObject[]]$InputObject,
        [string[]]$Exclude = "",
        [bool]$ExcludeDefault = $True,
        [string[]]$Include = $null,
        [string[]]$Value = $null,
        [int]$MaxDepth = 10
    )
    Begin
    {
        #region FUNCTIONS

            #Before adding a property, verify that it matches a Like comparison to strings in $Include...
            Function IsIn-Include {
                param($prop)
                if(-not $Include) {$True}
                else {
                    foreach($Inc in $Include)
                    {
                        if($Prop -like $Inc)
                        {
                            $True
                        }
                    }
                }
            }

            #Before adding a value, verify that it matches a Like comparison to strings in $Value...
            Function IsIn-Value {
                param($val)
                if(-not $Value) {$True}
                else {
                    foreach($string in $Value)
                    {
                        if($val -like $string)
                        {
                            $True
                        }
                    }
                }
            }

            Function Get-Exclude {
                [cmdletbinding()]
                param($obj)

                #Exclude default props if specified, and anything the user specified.  Thanks to Jaykul for the hint on [type]!
                    if($ExcludeDefault)
                    {
                        Try
                        {
                            $DefaultTypeProps = @( $obj.gettype().GetProperties() | Select -ExpandProperty Name -ErrorAction Stop )
                            if($DefaultTypeProps.count -gt 0)
                            {
                                Write-Verbose "Excluding default properties for $($obj.gettype().Fullname):`n$($DefaultTypeProps | Out-String)"
                            }
                        }
                        Catch
                        {
                            Write-Verbose "Failed to extract properties from $($obj.gettype().Fullname): $_"
                            $DefaultTypeProps = @()
                        }
                    }
                    
                    @( $Exclude + $DefaultTypeProps ) | Select -Unique
            }

            #Function to recurse the Object, add properties to object
            Function Recurse-Object {
                [cmdletbinding()]
                param(
                    $Object,
                    [string[]]$path = '$Object',
                    [psobject]$Output,
                    $depth = 0
                )

                # Handle initial call
                Write-Verbose "Working in path $Path at depth $depth"
                Write-Debug "Recurse Object called with PSBoundParameters:`n$($PSBoundParameters | Out-String)"
                $Depth++

                #Exclude default props if specified, and anything the user specified.                
                $ExcludeProps = @( Get-Exclude $object )

                #Get the children we care about, and their names
                $Children = $object.psobject.properties | Where {$ExcludeProps -notcontains $_.Name }
                Write-Debug "Working on properties:`n$($Children | select -ExpandProperty Name | Out-String)"

                #Loop through the children properties.
                foreach($Child in @($Children))
                {
                    $ChildName = $Child.Name
                    $ChildValue = $Child.Value

                    Write-Debug "Working on property $ChildName with value $($ChildValue | Out-String)"
                    # Handle special characters...
                        if($ChildName -match '[^a-zA-Z0-9_]')
                        {
                            $FriendlyChildName = "{$ChildName}"
                        }
                        else
                        {
                            $FriendlyChildName = $ChildName
                        }

                    #Handle null...
                        if($ChildValue -eq $null)
                        {
                            Write-Verbose "Skipping NULL $ChildName"
                            continue
                        }

                    #Add the property.
                        if((IsIn-Include $ChildName) -and (IsIn-Value $ChildValue) -and ($Depth -le $MaxDepth) -and ( $ChildName -ne 'Count' ) -and ( $ChildValue.getType().Name -ne 'PSCustomObject' ) -and ( $ChildValue.getType().Name -ne 'Object[]' ) -and ( $ChildValue.getType().Name -ne 'Hashtable' ))
                        {
                            $ThisPath = @( $Path + $FriendlyChildName ) -join "."
                            if( ($Output.psobject.Properties.Match($FriendlyChildName)).Count -gt 0 )
                            {
                                $Output.$FriendlyChildName += $ChildValue
                            } else {
                                $Output | Add-Member -MemberType NoteProperty -Name $FriendlyChildName -Value @($ChildValue)
                            }
                            Write-Verbose "Adding member '$ThisPath'"
                        }

                    #Handle evil looping.  Will likely need to expand this.  Any thoughts on a better approach?
                        if(
                            (
                                $ChildValue.GetType() -eq $Object.GetType() -and
                                $ChildValue -is [datetime]
                            ) -or
                            (
                                $ChildName -eq "SyncRoot" -and
                                -not $ChildValue
                            )
                        )
                        {
                            Write-Verbose "Skipping $ChildName with type $($ChildValue.GetType().fullname)"
                            continue
                        }

                    #Check for arrays
                        $IsArray = @($ChildValue).count -ge 1
                        $count = 0
                        
                    #Set up the path to this node and the data...
                        $CurrentPath = @( $Path + $FriendlyChildName ) -join "."

                    #Exclude default props if specified, and anything the user specified.                
                        $ExcludeProps = @( Get-Exclude $ChildValue )

                    #Get the children's children we care about, and their names.  Also look for signs of a hashtable like type
                        $ChildrensChildren = $ChildValue.psobject.properties | Where {$ExcludeProps -notcontains $_.Name }
                        $HashKeys = if($ChildValue.Keys -notlike $null -and $ChildValue.Values)
                        {
                            $ChildValue.Keys
                        }
                        else
                        {
                            $null
                        }
                        Write-Debug "Found children's children $($ChildrensChildren | select -ExpandProperty Name | Out-String)"

                    #If we aren't at max depth or a leaf...                   
                    if(
                        (@($ChildrensChildren).count -ne 0 -or $HashKeys) -and
                        $Depth -lt $MaxDepth
                    )
                    {
                        #This handles hashtables.  But it won't recurse... 
                            if($HashKeys)
                            {
                                Write-Verbose "Working on hashtable $CurrentPath"
                                foreach($key in $HashKeys)
                                {
                                    Write-Verbose "Adding value from hashtable $CurrentPath['$key']"
                                    $Output | Add-Member -MemberType NoteProperty -name "$CurrentPath['$key']" -value $ChildValue["$key"]
                                    $Output = Recurse-Object -Object $ChildValue["$key"] -Path "$CurrentPath['$key']" -Output $Output -depth $depth 
                                }
                            }
                        #Sub children?  Recurse!
                            else
                            {
                                if($IsArray)
                                {
                                    foreach($item in @($ChildValue))
                                    {  
                                        Write-Verbose "Recursing through array node '$CurrentPath'"
                                        $Output = Recurse-Object -Object $item -Path "$CurrentPath[$count]" -Output $Output -depth $depth
                                        $Count++
                                    }
                                }
                                else
                                {
                                    Write-Verbose "Recursing through node '$CurrentPath'"
                                    $Output = Recurse-Object -Object $ChildValue -Path $CurrentPath -Output $Output -depth $depth
                                }
                            }
                        }
                    }
                
                $Output
            }

        #endregion FUNCTIONS
    }
    Process
    {
        Foreach($Object in $InputObject)
        {
            #Flatten the XML and write it to the pipeline
                Recurse-Object -Object $Object -Output $( New-Object -TypeName PSObject )
        }
    }
}

function get-columnInfo($t, $c)
{
    $rx_tname = [System.Text.RegularExpressions.Regex]::Escape($t.Name)
    $rx_cname = [System.Text.RegularExpressions.Regex]::Escape($c.Name)

    # Busca uso da coluna em Visuals:
    $o1 = "{0}\.{1}" -f $rx_tname, $rx_cname
    $o2 = "{0}(\s?)\[{1}\]" -f $rx_tname, $rx_cname
    $o3 = "'{0}'(\s?)\[{1}\]" -f $rx_tname, $rx_cname
    $o4 = "(^|\W)\[{0}\]"   -f $rx_cname

    $usedInModel = $globalCols.Contains($t.Name + "." + $c.Name)

    $usedCC  = @()
    $usedCC += @( $expr | where-object {$_.Tipo -eq "Calculated"} | where-object { ( $_.Table -ne $t.Name -or ( $_.Table -eq $t.Name -and $_.Name -ne $c.Name ) ) -and ( $_.Expression -match $o1 -or $_.Expression -match $o2 -or $_.Expression -match $o3 ) } )
    $usedCC += @( $expr | where-object {$_.Tipo -eq "Calculated"} | where-object { ( $_.Table -eq $t.Name -and $_.Name -ne $c.Name ) -and ( $_.Expression -match $o4 ) } )

    $usedMS  = @()
    $usedMS += @( $expr | where-object {$_.Tipo -eq "Measure"}    | where-object { ( $_.Table -ne $t.Name -or ( $_.Table -eq $t.Name -and $_.Name -ne $c.Name ) ) -and ( $_.Expression -match $o1 -or $_.Expression -match $o2 -or $_.Expression -match $o3 ) } )
    $usedMS += @( $expr | where-object {$_.Tipo -eq "Measure"}    | where-object { ( $_.Table -eq $t.Name -and $_.Name -ne $c.Name ) -and ( $_.Expression -match $o4 ) } )


    $inRelation = ( $expr | where-object {$_.Tipo -eq "Relation"} | where-object { ( $_.Expression -match $o1 ) } ).Count -gt 0

    $hasRbar = $false
    if( $c.Type -eq "Calculated" )
    {
        $RBAR_FUNCTIONS | ForEach-Object {
            $hasRbar = $c.Expression -contains $_
            if( $hasRbar ){
                break
            }
        }
    }
    # write-host $hasRbar

    $dcount = 0
    $sum = 0

    if( ( $c.Type -eq "Data" -or $c.Type -eq "Calculated" ) -and $c.Type -ne "RowNumber" )
    {
        if( $c.DataType -eq "Double" )
        {
            $dsum = dax( ( $query_sum_col -f $t.Name, $c.Name ) )
            
            if( $null -ne $dsum -and $false -ne $dsum -and [Double]::NaN -ne $dsum )
            {
                $sum = $dsum[1]
            }

        }
    
        $dcount = dax( ( $query_dcount_col -f $t.Name, $c.Name  ) )

        if( $dcount -ne $false )
        {
            $dcount = $dcount[1]
        }

    }

    return @{ Table=$t.Name; Name=$c.Name; DataType=$c.DataType.ToString(); Type=$c.Type; IsKey=( $c.IsKey -or $inRelation ); Distinct=$dcount; Expression=$c.Expression; Sum=$sum; InModel=$usedInModel; InData=( ($usedCC.Count + $usedMS.Count ) -gt 0); InCalculated=$usedCC.Count; InMeasures=$usedMS.Count }
}

function get-measureInfo($t, $m)
{
    $rx_tname = [System.Text.RegularExpressions.Regex]::Escape($t.Name)
    $rx_mname = [System.Text.RegularExpressions.Regex]::Escape($m.Name)

    $o0 = "{0}\.{1}"      -f $rx_tname, $rx_mname
    $o1 = "{0}\[{1}\]"    -f $rx_tname, $rx_mname
    $o2 = "'{0}'\[{1}\]"  -f $rx_tname, $rx_mname
    $o3 = "(^|\W)\[{0}\]" -f $rx_mname
    #$o4 = "`"{0}`""   -f $rx_mname

    $usedInModel = $globalCols.Contains($t.Name + "." + $m.Name)

    $usedCC = ( $expr | where-object {$_.Tipo -eq "Calculated"}  | where-object {( $_.Table -ne $t.Name -or ( $_.Table -eq $t.Name -and $_.Name -ne $m.Name ) ) -and ( $_.Expression -match $o0 -or $_.Expression -match $o1 -or $_.Expression -match $o2 -or $_.Expression -match $o3 ) } )
    $usedMS = ( $expr | where-object {$_.Tipo -eq "Measure"}     | where-object {( $_.Table -ne $t.Name -or ( $_.Table -eq $t.Name -and $_.Name -ne $m.Name ) ) -and ( $_.Expression -match $o0 -or $_.Expression -match $o1 -or $_.Expression -match $o2 -or $_.Expression -match $o3 ) } )
    #write-host $usedCC
    
    $RBAR_FUNCTIONS | ForEach-Object {
        $hasRbar = $m.Expression -contains $_
        if( $hasRbar ){
            break
        }
    }
    # write-host $hasRbar

    return @{ Table=$t.Name; Name=$m.Name; DataType=$m.DataType.ToString(); Type=$null; IsKey=$null; Distinct=$null; Expression=$m.Expression; InModel=$usedInModel; InData=( ($usedCC.Count + $usedMS.Count) -gt 0); InCalculated=$usedCC.Count; InMeasures=$usedMS.Count }
}

function get-visualInfo($s, $v)
{
    $cols = @()

    $filters = @()

    $from   = $null
    $select = $null
    $proj   = $null

    $c = $null
    $f = $null
    $q = $null

    if( $null -ne $v.config )
    {
        $c = $v.config.ToLower() | ConvertFrom-Json

        if( $null -ne $c.singleVisual.prototypeQuery.From ){
            $from   = $c.singleVisual.prototypeQuery.From | ConvertTo-FlatObject 
        }
        if( $null -ne $c.singleVisual.prototypeQuery.Select ){
            $select = $c.singleVisual.prototypeQuery.Select | ConvertTo-FlatObject 

            $select | ForEach-Object {
                $ss = $_
                $cols += $ss.Name
                $ent = ( $from | where-object { $_.Name -eq $ss.Source } )

                if( $null -ne $ent.Entity -and $ent.Entity.count -ge 0 )
                {
                    if( $null -ne $ss.Property )
                    {
                        $cols += ( "{0}.{1}" -f $ent.Entity[0], $ss.Property[0] )
                    }

                    if( $null -ne $ss.Hierarchy )
                    {
                        $cols += ( "{0}.{1}.{2}" -f $ent.Entity[0], $ss.Hierarchy[0], $ss.Level[0] )
                    }

                }
            }
        }

        if( $null -ne $c.singleVisual.projections )
        {
            $proj  = $c.singleVisual.projections | ConvertTo-FlatObject | select-object queryRef
            $cols += $proj.queryRef
        }
    }
    
    if( $null -ne $v.filters )
    {
        $filters = $v.filters | ConvertFrom-Json

        $filters | where-object { $null -ne $_.filter } | ForEach-Object {
            if( $null -ne $_.filter )
            {
                $f = $_.filter.From  | ConvertTo-FlatObject
                $w = $_.filter.Where | ConvertTo-FlatObject
                
                $w | ForEach-Object {
                    $ss = $_
                    $ent = ( $f | where-object { $_.Name -eq $ss.Source } )

                    if( $null -ne $ent.Entity -and $ent.Entity.count -ge 0 )
                    {
                        if( $null -ne $ss.Property )
                        {
                            $cols += ( "{0}.{1}" -f $ent.Entity[0], $ss.Property[0] )
                        }

                        if( $null -ne $ss.Hierarchy )
                        {
                            $cols += ( "{0}.{1}.{2}" -f $ent.Entity[0], $ss.Hierarchy[0], $ss.Level[0] )
                        }
                    }
                }
            }
        }
    }

    if( $null -ne $v.query )
    {
        $q = $v.query   | ConvertFrom-Json

        if( $null -ne $q.Commands -and $q.Commands.Count -gt 0 )
        {
            $f = $q.Commands.SemanticQueryDataShapeCommand.Query.From   | ConvertTo-FlatObject | select-object Name, Entity

            if( $null -ne $q.Commands.SemanticQueryDataShapeCommand.Query.Select )
            {
                $e = $q.Commands.SemanticQueryDataShapeCommand.Query.Select | ConvertTo-FlatObject | select-object Source, Property

                $e | ForEach-Object {
                    $ss = $_
                    $ent = ( $f | where-object { $_.Name -eq $ss.Source[0] } )

                    if( $null -ne $ent.Entity -and $ent.Entity.count -gt 0 )
                    {
                        if( $null -ne $ss.Property )
                        {
                            $cols += ( "{0}.{1}" -f $ent.Entity[0], $ss.Property[0] )
                        }

                        if( $null -ne $ss.Hierarchy )
                        {
                            $cols += ( "{0}.{1}.{2}" -f $ent.Entity[0], $ss.Hierarchy[0], $ss.Level[0] )
                        }
                    }
                }
            }

            if( $null -ne $q.Commands.SemanticQueryDataShapeCommand.Query.Where )
            {
                $w = $q.Commands.SemanticQueryDataShapeCommand.Query.Where  | ConvertTo-FlatObject | select-object Source, Property

                $w | Where-Object { $_.Source  } | ForEach-Object {
                    $ss = $_
                    $ent = ( $f | where-object { $_.Name -eq $ss.Source[0] } )

                    if( $null -ne $ent.Entity -and $ent.Entity.count -gt 0 )
                    {
                        if( $null -ne $ss.Property )
                        {
                            $cols += ( "{0}.{1}" -f $ent.Entity[0], $ss.Property[0] )
                        }

                        if( $null -ne $ss.Hierarchy )
                        {
                            $cols += ( "{0}.{1}.{2}" -f $ent.Entity[0], $ss.Hierarchy[0], $ss.Level[0] )
                        }
                    }
                }
            }
        }

    }
    
    if( $null -ne $v.dataTransforms )
    {
        $dt = $v.dataTransforms | ConvertFrom-Json | ConvertTo-FlatObject | select-object metadata, queryRef, queryName

        $cols += ($dt.metadata)
        $cols += ($dt.queryRef)
        $cols += ($dt.queryName)
    }
    
    $cols = ( $cols | Sort-Object | get-unique )
    # $globalCols += $cols

    return @{
        SectionId = $s.id;
        VisualId = $v.id;
        ConfigName = $c.name;
        Type = $c.singleVisual.visualType;
        Columns = @( if( $cols.count -gt 0 ){ $cols } else { $null } )
    }

}

$info = @()

$info += @{
    Arquivo=$pbix;
    Server=$ServerName;
    TableCount=$server.Databases[0].Model.Tables.Count;
    DatasourcesCount=$server.Databases[0].Model.DataSources.Count;
    ModifiedTime=$server.Databases[0].Model.StructureModifiedTime;
    Paginas=$layout.sections.Count;
    Size=$server.Databases[0].EstimatedSize
}

Write-Host "Arquivo aberto: ", $pbix
Write-Host "Server:         ", $ServerName
Write-host "Tables:         ", $server.Databases[0].Model.Tables.Count
Write-host "Datasources:    ", $server.Databases[0].Model.DataSources.Count
Write-host "Modificado em:  ", $server.Databases[0].Model.ModifiedTime

# Extrai datasources
Write-host "Extraindo datasources"
$datasources = @()

$server.Databases[0].Model.DataSources | ForEach-Object {

    $d = $_
    $ds = ""
    $tmp1 = New-TemporaryFile

    # Isola a definição do datasource (base64)
    $ds = ($d.ConnectionString.Split(";")[2]).Split('"')[1]

    # verifica se não está vazio
    if( -not ( $ds -eq $null -or $ds -eq "" -or $ds.Length -lt 1 ) )
    {

        # converte base64 em bytes
        $bytes = [System.Convert]::FromBase64String($ds)

        # salva os bytes em arquivo temporario (zip)
        [IO.File]::WriteAllBytes($tmp1.FullName, $bytes)
    
        # open ZIP archive for reading
        $zip = [System.IO.Compression.ZipFile]::OpenRead($tmp1.FullName)

        # Extract layout file
        $zip.Entries | Where-Object {$_.FullName -match "^Formulas"} | ForEach-Object {

            # extrai conteudo (rotina M) em arquivo temporario 
            $tmp2 = New-TemporaryFile
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, $tmp2, $true)

            # busca o texto do arquivo temporario
            $def = Get-Content $tmp2 -Raw -Encoding UTF8

            # adiciona na lista de datasources
            $datasources += @{ datasource=$def.ToString().Replace("#(lf)", "`n").Replace("#(tab)", "`t") }

            # remove temporario
            Remove-Item $tmp2
        }

        # encerra abertura do zip
        $zip.Dispose()
    }
    
    # remove temporario
    Remove-Item $tmp1
}

Write-Host "Extraindo Expressions"

$server.Databases[0].Model.Expressions | ForEach-Object {
    $datasources += @{ Name=$_.Name; datasource=$_.Expression.ToString().Replace("#(lf)", "`n").Replace("#(tab)", "`t") }
}

$tables = @()
$cols = @()
$calculated = @()
$measures = @()
$expr = @()

$pages = @()

$relations = @()
$roles = @()

$visuals = @()
$globalCols = @()

$cols = @()

Write-Host "Extraindo Filters"

if( $layout.filters.Length -gt 2 )
{
    $f = $layout.filters | ConvertFrom-Json | ConvertTo-FlatObject | Select-Object Entity, Property
    $f | ForEach-Object {
        $fs = $_
        $n = 0
        $fs.Property | ForEach-Object {
            $cols += ($fs.Entity[$n] + "." + $fs.Property[$n])
            $n += 1
        }
    }
    $globalCols += $cols
}
Write-Host "Extraindo Section Visuals"

Write-Progress -Activity "Visuals" -Id 0 -PercentComplete 0

$cols = @()

$layout.sections | ForEach-Object {
    $s = $_

    $pages += @{ 
        displayName=$_.displayName;
        Visuals=$s.visualContainers.Count;
        SectionId=$s.id
    }

    $vctotal = $s.visualContainers.Count
    $vc = 0

    $s.visualContainers | ForEach-Object {
        $v = $_

        $visuals += ( get-visualInfo -s $s -v $v )

        $vc++
        Write-Progress -Activity "Visuals" -Id 0 -PercentComplete (($vc/$vctotal)*100)

    }

}

$globalCols = ( $visuals.foreach("Columns") | Sort-Object | get-unique )

Write-Progress -Activity "Visuals" -Id 0 -Completed

Write-Host "Extraindo expressions"

$cols = @()

# popula exressions
$server.Databases[0].Model.Tables | ForEach-Object {
    $t = $_
    if ( $t.Name -notmatch "Template" -and $t.Name -notmatch "LocalDateTable" )
    {
        $t.Measures | ForEach-Object {
            $m = $_
            $expr += @{ Table=$t.Name; Tipo="Measure"; Name=$m.Name; Expression=$m.Expression }
        }

        $t.Columns | Where-Object {$_.Type -eq "Calculated"} | ForEach-Object {
            $c = $_
            $expr += @{ Table=$t.Name; Tipo="Calculated"; Name=$c.Name; Expression=$c.Expression; }
        }
	$datasources += @{ Name=$t.Name; datasource=$t.Partitions[0].Source.Expression.ToString().Replace("#(lf)", "`n").Replace("#(tab)", "`t") }
    }
}


Write-Host "Extraindo relations"

# TODO: sentido da relação
$server.Databases[0].Model.Relationships | ForEach-Object {
    $r = $_
    # add to relations list
    $relations += @{ ToTable=$r.ToTable.Name; ToColumn=$r.ToColumn.Name; FromTable=$r.FromTable.Name; FromColumn=$r.FromColumn.Name; Active=$r.IsActive }
    # Add to expressions list
    $expr += @{ Table=$r.FromTable.Name; Tipo="Relation"; Name=$r.FromColumn.Name; Expression=( $r.FromTable.Name + "." + $r.FromColumn.Name ) }
    $expr += @{ Table=$r.ToTable.Name;   Tipo="Relation"; Name=$r.ToColumn.Name;   Expression=( $r.ToTable.Name   + "." + $r.ToColumn.Name   ) }
}

Write-Host "Extraindo tabelas calculadas"

$server.Databases[0].Model.Tables | ForEach-Object {
    $t = $_

    if ( $t.Name -notmatch "Template" -and $t.Name -notmatch "LocalDateTable" )
    {

        $count = dax( ( $query_row_col -f $t.Name ) )

        if( $count -ne $false )
        {
            Write-host "`tContagem: ", $count[1]
        }

        $t.Partitions | ForEach-Object {
            $p = $_

            $srcType  = ""
            $srcQuery = ""

            $t.Partitions | ForEach-Object {
                $p = $_

                $srcType  = $p.SourceType.ToString()
                
                if( $srcType -eq "Calculated" )
                {
                    $srcQuery = $p.Source.Expression
                    $expr += @{ Table=$t.Name; Tipo="CalculatedTable"; Name=""; Expression=$srcQuery }
                } else {
                    $srcQuery = $p.Source.Query
                }
            }

        }

        $tables += @{ Name=$t.Name; Colunas=$t.Columns.Count; Calculadas=( $t.Columns | Where-Object {$_.Type -eq "Calculated"}).Count; Measures=$t.Measures.Count; Categoria=$t.DataCategory; Records=$count[1]; SrourceType=$srcType; SourceQuery=$srcQuery }
    }
}

Write-Host "Extraindo Roles"

$server.Databases[0].Model.Roles | ForEach-Object {
    $r = $_
    $roleName = $r.Name
    $r.TablePermissions | ForEach-Object {
        $p = $_
        $tableName = $p.Name
        $roles += @{ Role=$roleName; Table=$tableName; Filter=$p.FilterExpression }
    }
}

if($roles.Count -eq 0)
{
    $roles += @{ Role=""; Table=""; Filter="" }
}

Write-Host "Percorrendo tabelas"

$server.Databases[0].Model.Tables | ForEach-Object {
    $t = $_

    if ( $t.Name -notmatch "Template" -and $t.Name -notmatch "LocalDateTable" )
    {
        Write-host "`tTabela: ", $t.Name
        #Write-host "`tColunas:    ", $t.Columns.Count
        #Write-host "`tCalculadas: ", ( $t.Columns | Where-Object {$_.Type -eq "Calculated"} ).Count
        #Write-host "`tMedidas:    ", $t.Measures.Count
        #Write-host "`tCategoria:  ", $t.DataCategory

        #
        # MEASURES
        #
         
        if( $t.Measures.Count -gt 0 )
        {
            Write-Progress -Activity "Measures" -Id 0 -PercentComplete 0
            $w = 0
            $total = $t.Measures.Count

            # Write-Host "`tMeasures"

            $t.Measures | ForEach-Object {
                $m = $_
                
                $measures += ( get-measureInfo -t $t -m $m )
                
                $w += 1
                Write-Progress -Activity "Measures" -Id 0 -PercentComplete (($w/$t.Measures.Count)*100)
            }
            Write-Progress -Activity "Measures" -Id 0 -Completed

        }

        Write-Progress -Activity "Calculated" -Id 0 -PercentComplete 0
        $w = 0
        $total = ( $t.Columns | Where-Object {$_.Type -eq "Calculated"} ).Count

        #
        # Colunas Calculadas
        #

        $t.Columns | Where-Object {$_.Type -eq "Calculated"} | ForEach-Object {
            $c = $_

            $calculated += ( get-columnInfo -t $t -c $c )
            
            $w += 1
            Write-Progress -Activity "Calculated" -Id 0 -PercentComplete (($w/$total)*100)
        }

        Write-Progress -Activity "Calculated" -Id 0 -Completed
        Write-Progress -Activity "Columns"    -Id 0 -PercentComplete 0

        $w = 0
        $total = ( $t.Columns | Where-Object {$_.Type -ne "Calculated"} ).Count

        #
        # Colunas de dados
        #

        $t.Columns | Where-Object {$_.Type -ne "Calculated"} | ForEach-Object {
            $c = $_
            
            $cols += ( get-columnInfo -t $t -c $c )
            
            $w += 1
            Write-Progress -Activity "Columns" -Id 0 -PercentComplete (($w/$total)*100)
        }

        Write-Progress -Activity "Columns" -Id 0 -Completed
    }
}

$server.Disconnect()
$server.Dispose()

$json = @{
    Info=$info;
    Pages = @($pages); 
    Tables = @($tables); 
    Columns = @($cols); 
    Calculated = @($calculated); 
    Measures = @($measures); 
    Datasources = @($datasources); 
    Relations = @($relations); 
    Roles = $roles; 
    Visuals = @($visuals)
}
( $json | ConvertTo-Json -Depth 100 ) -replace "\s(NaN)","null$2" -replace "\s(Infinity)","null$2" | Out-File $jsonOut -Encoding utf8

