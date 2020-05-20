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

# load ZIP methods
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Processo do Power BI Desktop
$p = ( get-process -Name PBIDesktop | select-object -First 1 )

if( $null -eq $p ){
    Write-Host "Nenhum Power BI aberto"
    return
}

# Nome do Log
$LogFile = ( $p.MainWindowTitle + ".log" )

$cmd = ( Get-CimInstance Win32_Process -Filter "name = 'PBIDesktop.exe'" | select-object -First 1 ).CommandLine
$pbix = $cmd.Remove(0, ($cmd.IndexOf('"', 1)) + 2 ).replace('"', '')

if( -not ( Test-Path $pbix ) )
{
    Write-Host "Arquivo não encontrado"
    return
}
$jsonOut = $pbix.Replace(".pbix", ".json")

# open ZIP archive for reading
$zip = [System.IO.Compression.ZipFile]::OpenRead($pbix)
$outfile = $pbix.Replace(".pbix", ".layout.json")
# Extract layout file
$zip.Entries | Where-Object {$_.FullName -eq 'Report/Layout'} | % {
    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, $outfile, $true)
}
$zip.Dispose()
# $outfile = & 'C:\Projetos\Gerdau\PBIX\SPGDash - Volume e Desaderencia.layout.json'
$layout = Get-Content $outfile -Encoding Unicode | ConvertFrom-Json
$layout | ConvertTo-Json -Depth 100 | Out-File $outfile -Encoding utf8

# Processo do Power BI SSAS
$AsID = ( get-process -Name msmdsrv | select-object -First 1 ).id

# porta TCP do processo
$AsPort = ( Get-NetTCPConnection -OwningProcess $AsID -State Listen -LocalAddress 127.0.0.1 | Select-Object -First 1 ).LocalPort

# SSAS URL
$ServerName = "127.0.0.1:$AsPort"

# Cria instância do server
$server = New-Object Microsoft.AnalysisServices.Server

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

$query_sum_col = @"
EVALUATE ROW("count", SUM('{0}'[{1}]))
"@

$query_count_col = @"
EVALUATE ROW("count", COUNT('{0}'[{1}]))
"@

$query_dcount_col = @"
EVALUATE ROW("count", DISTINCTCOUNT('{0}'[{1}]))
"@

# -- relations
$query_relations = @"
SELECT [MEASUREGROUP_NAME], [MEASUREGROUP_CARDINALITY], [DIMENSION_UNIQUE_NAME], [DIMENSION_CARDINALITY]
  FROM `$system.MDSCHEMA_MEASUREGROUP_DIMENSIONS;
"@

# -- tables w/ count
$query_tables = @"
SELECT [DIMENSION_UNIQUE_NAME], [DIMENSION_CARDINALITY]
  FROM `$system.MDSCHEMA_DIMENSIONS;
"@

# -- columns
$query_columns = @"
SELECT DISTINCT [HIERARCHY_UNIQUE_NAME]
  FROM `$system.MDSCHEMA_PROPERTIES;
"@

# -- measures
$query_measures = @"
SELECT DISTINCT [MEASURE_UNIQUE_NAME]
  FROM `$system.MDSCHEMA_MEASURES;
"@

# -- cardinality
$query_cols = @"
SELECT [ID], [HierarchyID], [Ordinal], [Name], [Description], [ColumnID], [ModifiedTime] 
  FROM `$system.TMSCHEMA_LEVELS;
"@

# -- tables from TMSCHEMA
$query_tmtables = @"
select [ID], [NAME] from `$SYSTEM.TMSCHEMA_TABLES;
"@

# -- columns from TMSCHEMA
$query_tmcolumns = @"
SELECT [ID], [TableID], [ExplicitName], [Expression]
  FROM  `$system.TMSCHEMA_COLUMNS;
"@

# -- relations from TMSCHEMA
$query_tmrelations = @"
select [ToTableID], [ToColumnID], [FromTableID], [FromColumnID], [IsActive]
  from `$SYSTEM.TMSCHEMA_RELATIONSHIPS 
order by FromTableID;
"@

# Functions que funcionam linha a linha (RBAR)
$RBAR_FUNCTIONS = @('AVERAGEX', 'CONCATENATEX', 'COUNTX', 'COUNTAX', 'GEOMEANX', 'MAXX', 'MEDIANX', 'MINX', 'PERCENTILEX.EXC', 'PERCENTILEX.INC', 'PRODUCTX', 'RANKX', 'SUMX')

# Conecta no server
$server.connect($ServerName)

# Nome do database
$Database = $server.Databases[0].Name
$Tamanho = $server.Databases[0].EstimatedSize

# String de conexão
$connectionString = "Provider=MSOLAP;Data Source=localhost:$AsPort;Initial Catalog=$Database;Timeout=0;"

Write-Host "Arquivo aberto: ", $pbix
Write-Host "Server:         ", $ServerName
Write-host "Tables:         ", $server.Databases[0].Model.Tables.Count
Write-host "Datasources:    ", $server.Databases[0].Model.DataSources.Count
Write-host "Modificado em:  ", $server.Databases[0].Model.ModifiedTime

# Extrai datasources
Write-host "Extraindo datasources"
$datasources = @()

$server.Databases[0].Model.DataSources | % {

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
        $zip.Entries | Where-Object {$_.FullName -match "^Formulas"} | % {

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

$server.Databases[0].Model.Expressions | % {
    $datasources += @{ datasource=$_.Expression.ToString().Replace("#(lf)", "`n").Replace("#(tab)", "`t") }
}

$info = @()
$tables = @()
$cols = @()
$calculated = @()
$measures = @()
$expr = @()

$info += @{
    Arquivo=$pbix;
    Server=$ServerName;
    TableCount=$server.Databases[0].Model.Tables.Count;
    DatasourcesCount=$server.Databases[0].Model.DataSources.Count;
    ModifiedTime=$server.Databases[0].Model.StructureModifiedTime;
    Paginas=$layout.sections.Count;
    Size=$Tamanho
}

$pages = @()

$relations = @()
$roles = @()

$visuals = @()
$globalCols = @()

Write-Progress -Activity "Visuals" -Id 0 -PercentComplete 0

$layout.sections | ForEach-Object {
    $s = $_
    $pages += @{ 
        displayName=$_.displayName;
        Visuals=$s.visualContainers.Count
    }
    $vctotal = $s.visualContainers.Count
    $vc = 0

    $s.visualContainers | ForEach-Object {
        $v = $_

        $cols = @()

        $filters = @()

        $from   = $null
        $select = $null
        $proj   = $null

        $c = $null
        $f = $null
        $q = $null
        $t = $null

        if( $v.config -ne $null )
        {
            $c = $v.config  | ConvertFrom-Json

            if( $c.singleVisual.prototypeQuery.From -ne $null ){
                $from   = $c.singleVisual.prototypeQuery.From | ConvertTo-FlatObject 
            }
            if( $c.singleVisual.prototypeQuery.Select -ne $null ){
                $select = $c.singleVisual.prototypeQuery.Select | ConvertTo-FlatObject 

                $select | ForEach-Object {
                    $ss = $_
                    $cols += $ss.Name
                    $ent = ( $from | where-object { $_.Name -eq $ss.Source } )

                    if( $ent.Entity -ne $null -and $ent.Entity.count -ge 0 )
                    {
                        if( $ss.Property -ne $null )
                        {
                            $cols += ( "{0}.{1}" -f $ent.Entity[0], $ss.Property[0] )
                        }
                        if( $ss.Hierarchy -ne $null )
                        {
                            $cols += ( "{0}.{1}.{2}" -f $ent.Entity[0], $ss.Hierarchy[0], $ss.Level[0] )
                        }

                    }
                }
            }

            if( $c.singleVisual.projections -ne $null )
            {
                $proj   = $c.singleVisual.projections | ConvertTo-FlatObject | select-object queryRef
                $cols += $proj.queryRef
            }
        }
        
        if( $v.filters -ne $null )
        {
            $filters = $v.filters | ConvertFrom-Json

            $filters | where-object { $_.filter -ne $null } | ForEach-Object {
                if( $_.filter -ne $null )
                {
                    $f = $_.filter.From  | ConvertTo-FlatObject
                    $w = $_.filter.Where | ConvertTo-FlatObject
                    
                    $w | ForEach-Object {
                        $ss = $_
                        $ent = ( $f | where-object { $_.Name -eq $ss.Source } )

                        if( $ent.Entity -ne $null -and $ent.Entity.count -ge 0 )
                        {
                            if( $ss.Property -ne $null )
                            {
                                $cols += ( "{0}.{1}" -f $ent.Entity[0], $ss.Property[0] )
                            }
                            if( $ss.Hierarchy -ne $null )
                            {
                                $cols += ( "{0}.{1}.{2}" -f $ent.Entity[0], $ss.Hierarchy[0], $ss.Level[0] )
                            }
                        }
                    }
                }
            }
        }

        if( $v.query -ne $null )
        {
            $q = $v.query   | ConvertFrom-Json

            if( $q.Commands -ne $null -and $q.Commands.Count -gt 0 )
            {
                $f = $q.Commands.SemanticQueryDataShapeCommand.Query.From   | ConvertTo-FlatObject | select-object Name, Entity

                if( $q.Commands.SemanticQueryDataShapeCommand.Query.Select -ne $null )
                {
                    $e = $q.Commands.SemanticQueryDataShapeCommand.Query.Select | ConvertTo-FlatObject | select-object Source, Property

                    $e | ForEach-Object {
                        $ss = $_
                        $ent = ( $f | where-object { $_.Name -eq $ss.Source[0] } )

                        if( $ent.Entity -ne $null -and $ent.Entity.count -gt 0 )
                        {
                            if( $ss.Property -ne $null )
                            {
                                $cols += ( "{0}.{1}" -f $ent.Entity[0], $ss.Property[0] )
                            }
                            if( $ss.Hierarchy -ne $null )
                            {
                                $cols += ( "{0}.{1}.{2}" -f $ent.Entity[0], $ss.Hierarchy[0], $ss.Level[0] )
                            }
                        }
                    }
                }

                if( $q.Commands.SemanticQueryDataShapeCommand.Query.Where -ne $null )
                {
                    $w = $q.Commands.SemanticQueryDataShapeCommand.Query.Where  | ConvertTo-FlatObject | select-object Source, Property

                    $w | ForEach-Object {
                        $ss = $_
                        $ent = ( $f | where-object { $_.Name -eq $ss.Source[0] } )

                        if( $ent.Entity -ne $null -and $ent.Entity.count -gt 0 )
                        {
                            if( $ss.Property -ne $null )
                            {
                                $cols += ( "{0}.{1}" -f $ent.Entity[0], $ss.Property[0] )
                            }
                            if( $ss.Hierarchy -ne $null )
                            {
                                $cols += ( "{0}.{1}.{2}" -f $ent.Entity[0], $ss.Hierarchy[0], $ss.Level[0] )
                            }
                        }
                    }
                }
            }

        }
        
        if( $v.dataTransforms -ne $null )
        {
            $dt = $v.dataTransforms | ConvertFrom-Json | ConvertTo-FlatObject | select-object metadata, queryRef, queryName
            $cols += ($dt.metadata)
            $cols += ($dt.queryRef)
            $cols += ($dt.queryName)
        }

        $visuals += @{
            SectionId = $s.id;
            VisualId = $v.id;
            ConfigName = $c.name;
            Type = $c.singleVisual.visualType;
            Columns = @( if( $cols.count -gt 0 ){ $cols | Sort-Object | get-unique } else { $null } )
        }
        $globalCols += $cols | Sort-Object | get-unique 
        $vc++
        Write-Progress -Activity "Visuals" -Id 0 -PercentComplete (($vc/$vctotal)*100)
    }
}
$globalCols = $globalCols | Sort-Object | get-unique 
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
    }
}


Write-Host "Extraindo relations"

$server.Databases[0].Model.Relationships | % {
    $r = $_
    # add to relations list
    $relations += @{ ToTable=$r.ToTable.Name; ToColumn=$r.ToColumn.Name; FromTable=$r.FromTable.Name; FromColumn=$r.FromColumn.Name; Active=$r.IsActive }
    # Add to expressions list
    $expr += @{ Table=$r.FromTable.Name; Tipo="Relation"; Name=$r.FromColumn.Name; Expression=( $r.FromTable.Name + "." + $r.FromColumn.Name ) }
    $expr += @{ Table=$r.ToTable.Name;   Tipo="Relation"; Name=$r.ToColumn.Name;   Expression=( $r.ToTable.Name   + "." + $r.ToColumn.Name   ) }
}

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
        Write-host "Tabela: ", $t.Name
        Write-host "`tColunas:    ", $t.Columns.Count
        Write-host "`tCalculadas: ", ( $t.Columns | Where-Object {$_.Type -eq "Calculated"} ).Count
        Write-host "`tMedidas:    ", $t.Measures.Count
        Write-host "`tCategoria:  ", $t.DataCategory

        $count = dax( ( $query_count_col -f $t.Name, $t.Columns[1].Name ) )
        if( $count -ne $false )
        {
            Write-host "`tContagem: ", $count[1]
        }

        $tables += @{ Name=$t.Name; Colunas=$t.Columns.Count; Calculadas=( $t.Columns | Where-Object {$_.Type -eq "Calculated"}).Count; Measures=$t.Measures.Count; Categoria=$t.DataCategory; Records=$count[1] }

        #
        # MEASURES
        #
         
        if( $t.Measures.Count -gt 0 )
        {
            Write-Progress -Activity "Measures" -Id 0 -PercentComplete 0
            $w = 0
            $total = $t.Measures.Count

            # Write-Host "`tMeasures"

            $t.Measures | % {
                $m = $_
                $rx_tname = [System.Text.RegularExpressions.Regex]::Escape($t.Name)
                $rx_mname = [System.Text.RegularExpressions.Regex]::Escape($m.Name)

                $o0 = "{0}\.{1}"  -f $rx_tname, $rx_mname
                $o1 = "\W\[{0}\]" -f $rx_mname
                $o2 = "`"{0}`""   -f $rx_mname

                $usedInModel = $globalCols.Contains($t.Name + "." + $m.Name)

                <#
                if( $usedInModel -eq $false )
                {
                    $uso = Get-ObjPropertyName -Object $layout.sections -Value $o0
                    $usedInModel = ( $uso.Count -gt 0 )

                    if( $usedInModel -eq $false )
                    {
                        $uso = Get-ObjPropertyName -Object $layout.sections -Value $o1
                        $usedInModel = ( $uso.Count -gt 0 )

                        if( $usedInModel -eq $false )
                        {
                            $uso = Get-ObjPropertyName -Object $layout.sections -Value $o2
                            $usedInModel = ( $uso.Count -gt 0 )
                        }
                    }
                }
                #>

                $usedCC = ( $expr | where-object {( $_.Table -ne $t.Name -or ( $_.Table -eq $t.Name -and $_.Name -ne $m.Name ) ) -and ( $_.Expression -match $o0 -or $_.Expression -match $o1 -or $_.Expression -match $o2 ) } )
                #write-host $usedCC

                $measures += @{ Table=$t.Name; Name=$m.Name; DataType=$m.DataType.ToString(); Type=$null; IsKey=$null; Distinct=$null; Expression=$m.Expression; InModel=$usedInModel; InData=($usedCC.Count -gt 0) }

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

            $rx_tname = [System.Text.RegularExpressions.Regex]::Escape($t.Name)
            $rx_cname = [System.Text.RegularExpressions.Regex]::Escape($c.Name)

            # Busca uso da coluna em Visuals:
            $o1 = "{0}\.{1}" -f $rx_tname, $rx_cname
            $o2 = "{0}(\s?)\[{1}\]" -f $rx_tname, $rx_cname
            $o3 = "'{0}'(\s?)\[{1}\]" -f $rx_tname, $rx_cname
            $o4 = "`"{0}`""   -f $rx_cname

            $usedInModel = $globalCols.Contains($t.Name + "." + $c.Name)

            <#
            if( $usedInModel -eq $false )
            {
                $uso = Get-ObjPropertyName -Object $layout.sections -Value $o1
                $usedInModel = ( $uso.Count -gt 0 )

                if( $usedInModel -eq $false )
                {
                    $uso = Get-ObjPropertyName -Object $layout.sections -Value $o4
                    $usedInModel = ( $uso.Count -gt 0 )
                }
            }
            #>

            $usedCC = ( $expr | where-object { ( $_.Table -ne $t.Name -or ( $_.Table -eq $t.Name -and $_.Name -ne $c.Name ) ) -and ( $_.Expression -match $o1 -or $_.Expression -match $o2 -or $_.Expression -match $o3 ) } )

            $inRelation = ( $expr | where-object {$_.Tipo -eq "Relation"} | where-object { ( $_.Expression -match $o1 ) } ).Count -gt 0

            if( $c.Type -eq "Calculated" -and $c.Type -ne "RowNumber" )
            {
                if( $c.DataType -eq "Double" )
                {
                    $dsum = dax( ( $query_sum_col -f $t.Name, $c.Name ) )
                    
                    if( $dsum -ne $null -and $dsum -ne $false -and $dsum -ne [Double]::NaN )
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

            $calculated += @{ Table=$t.Name; Name=$c.Name; DataType=$c.DataType.ToString(); Type=$c.Type; IsKey=( $c.IsKey -or $inRelation ); Distinct=$dcount; Expression=$c.Expression; Sum=$sum; InModel=$usedInModel; InData=($usedCC.Count -gt 0) }

            $w += 1
            Write-Progress -Activity "Calculated" -Id 0 -PercentComplete (($w/$total)*100)
        }
        Write-Progress -Activity "Calculated" -Id 0 -Completed

        Write-Progress -Activity "Columns" -Id 0 -PercentComplete 0

        $w = 0
        $total = ( $t.Columns | Where-Object {$_.Type -ne "Calculated"} ).Count

        #
        # Colunas de dados
        #

        $t.Columns | Where-Object {$_.Type -ne "Calculated"} | ForEach-Object {
            $c = $_

            $rx_tname = [System.Text.RegularExpressions.Regex]::Escape($t.Name)
            $rx_cname = [System.Text.RegularExpressions.Regex]::Escape($c.Name)

            $o1 = "{0}\.{1}" -f $rx_tname, $rx_cname
            $o2 = "{0}(\s?)\[{1}\]" -f $rx_tname, $rx_cname
            $o3 = "'{0}'(\s?)\[{1}\]" -f $rx_tname, $rx_cname
            $o4 = "`"{0}`""   -f $rx_cname

            $usedInModel = $globalCols.Contains($t.Name + "." + $c.Name)
            <#
            if( $usedInModel -eq $false )
            {
                $uso = Get-ObjPropertyName -Object $layout.sections -Value $o1
                $usedInModel = ( $uso.Count -gt 0 )

                if( $usedInModel -eq $false )
                {
                    $uso = Get-ObjPropertyName -Object $layout.sections -Value $o4
                    $usedInModel = ( $uso.Count -gt 0 )
                }
            }
            #>

            $usedCC = ( $expr | where-object { ( $_.Table -ne $t.Name -or ( $_.Table -eq $t.Name -and $_.Name -ne $c.Name ) ) -and ( $_.Expression -match $o1 -or $_.Expression -match $o2 -or $_.Expression -match $o3 ) } )

            $inRelation = ( $expr | where-object {$_.Tipo -eq "Relation"} | where-object { ( $_.Expression -match $o1 ) } ).Count -gt 0

            $dcount = $null
            $sum = $null

            if( $c.Type -eq "Data" -and $c.Type -ne "RowNumber" )
            {
                if( $c.DataType -eq "Double" )
                {
                    $dsum = dax( ( $query_sum_col -f $t.Name, $c.Name ) )
                    
                    if( $dsum -ne $null -and $dsum -ne $false -and $dsum -ne [Double]::NaN )
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

            $cols += @{ Table=$t.Name; Name=$c.Name; DataType=$c.DataType.ToString(); Type=$c.Type; IsKey=( $c.IsKey -or $inRelation ); Distinct=$dcount; Expression=$null; Sum=$sum; InModel=$usedInModel; InData=($usedCC.Count -gt 0) }

            $w += 1
            Write-Progress -Activity "Columns" -Id 0 -PercentComplete (($w/$total)*100)
        }
        Write-Progress -Activity "Columns" -Id 0 -Completed
    }
}

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
( $json | ConvertTo-Json -Depth 100 ) -replace "\s(NaN)","null$2" | Out-File $jsonOut -Encoding utf8

