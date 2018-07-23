# Copyright © Microsoft Corporation.  All Rights Reserved.
# This code released under the terms of the 
# Microsoft Public License (MS-PL, http://opensource.org/licenses/ms-pl.html.)
# Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. 
# THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
# INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE. 
# We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code, provided that. 
# You agree: 
# (i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded; 
# (ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; 
# and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code 

Set-StrictMode -Version 4.0

Import-Module -Name SqlServer
Import-Module -Name Microsoft.XRM.Data.PowerShell

function Import-CDSAccessWebApp
{
<#
 .SYNOPSIS
 Imports Access Web App schema and data into a Common Data Service database.
 .DESCRIPTION
 Import-CDSAccessWebApp reads the schema of an Access Web App (AWA) from SQL Azure, 
 creates the same schema in the Common Data Service (CDS), and then imports the data.  
 .PARAMETER XRM
 If you have an existing connection to an XRM/CDS instance by calling Connect-CrmOnlineDiscovery or Connect-CrmOnline, you can pass it in here.  If not, you will be prompted to login and create a connection.
 .PARAMETER SQLServer
 Name of the SQL Server hosting the AWA.  You can obtain this information from the File menu, Connections, in Access when connected to the AWA.
 .PARAMETER SQLDatabase
 Name of teh SQL Database for the AWA.
 .PARAMETER SQLUserName
 Name of the SQL User with at least read permissions to the AWA.
 .PARAMETER SQLPassword
 The password for the SQLUserName.
 .PARAMETER Publisher
 The publisher string to use when creating entities and fields.  By default, the string "AWA" is used.
 .PARAMETER DisplaNamePrefix
 A prefix to include at the beginning of the entity display names.  Paired with the Publisher, use this to test the script and import into a new set of entities within the same databsae.
 .PARAMETER OneEntity
 Used for testing purposes, name of a single entity to load.  No relationships are created.
 .PARAMETER UsePresets
 Read and write configuration settings to a local file.  Use for testing and to import to multiple databases without prompts.
 .PARAMETER NoCreateEntities
 Do not create entities, only load data.
 .PARAMETER NoLoadData
 Do not load data, only create entities.
 .PARAMETER OrganizationOwned
 Use this switch to have the organization own the created  entities.  By default, entities are owned by the user who created them.
 .PARAMETER MaxMemoLength
 Maximum size for memo fields.  Default is 100,000.
 .PARAMETER MaxImageSize
 Maximum size of an image to import.  Default is 12,000,000.
 .PARAMETER Culture
 Default is "en-us".
 .PARAMETER LCID
 Default is 1033.
  .EXAMPLE
 Import-CDSAccessWebApp -SQLServer awa.database.windows.net -SQLDatabase db_xxxx -SQLUserName db_xxxx_ExternalReader -SQLPassword xxxx
 #>
[CmdletBinding()]
param
(
	[system.object] $XRM = $null,
	[parameter(Mandatory=$true)] [string] $SQLServer,
	[parameter(Mandatory=$true)] [string] $SQLDatabase,
	[parameter(Mandatory=$true)] [string] $SQLUsername,
	[parameter(Mandatory=$true)] [string] $SQLPassword,		
	[string]$Publisher = "AWA",
	[string]$DisplayNamePrefix = "",
	[string]$OneEntity = "",
	[switch]$UsePresets,
	[switch]$NoCreateEntities,
	[switch]$NoLoadData,
	[switch]$OrganizationOwned,
	[int]$MaxMemoLength = 100000,
	[int]$MaxImageSize = 12000000,
	[string]$Culture = "en-us",
	[int]$LCID = 1033	
)

function XRMExecute
{
	Param( $cer )

	if( -not $NoCreateEntities )
	{
		for( $retry = 0; $retry -lt 5; $retry++ )
		{
			try
			{
				return( $XRM.Execute( $cer ) )
			}
			catch
			{
				Write-Host -foregroundcolor Red $_.Exception.Message			
				Write-Host -foregroundcolor Red "Retry in 10 seconds..."
			}
			Start-Sleep -s 10
		}
	Write-Host -foregroundcolor Red "Giving up, exiting..."
	break
	}
}

function XRMCreateNewRecord
{
	Param( $ln, $cnr )

	if( -not $NoLoadData )
	{
		for( $retry = 0; $retry -lt 5; $retry++ )
		{
			try
			{
				$x = $XRM.CreateNewRecord( $ln, $cnr )
				if( $x -ne "00000000-0000-0000-0000-000000000000" )
				{
					return( $x );
				}
				else
				{
					Write-Host -foregroundcolor Red ("CreateNewRecord failed: "+$xrm.lastcrmerror)
				}
			}
			catch
			{
				Write-Host -foregroundcolor Red $_.Exception.Message			
			}
			Write-Host -foregroundcolor Red "Retry in 10 seconds..."
			Start-Sleep -s 10
		}
		Write-Host -foregroundcolor Red "Giving up, exiting..."
		break
	}
	else
	{
		return( "1234" )
	}
}

function ReadSQLImage
{
	Param( [string] $table, [string] $id )

	$query = "select datalength([Image]) as binlen from "+$table+" where id="+$id
	$binlen = ReadSQL -query $query
	$len = $binlen["binlen"]

	if( $len -gt $MaxImageSize )
	{
		write-host -foregroundcolor Red ("Image for id="+$id+" with a size of "+$len+" is larger than maximum allowed size of "+$MaxImageSize+" bytes). Use -MAxImageSize argument to change the maximum.")

		return $null
	}
	else
	{
		$query = "select [Image] from "+$table+" where id="+$id
		$image = ReadSQL -query $query -maxbinlen $len

		return $image["Image"]
	}
}

function ReadSQL
{
	Param( [string] $query, [int] $maxbinlen=1024 )

	$r = Invoke-Sqlcmd -Query $query -ServerInstance $SQLServer -Database $SQLDatabase -Username $SQLUsername -Password $SQLPassword -MaxBinaryLength $maxbinlen

	return $r
}

function LogicalName
{
	Param( [string] $name )

	$noprefix = $false;
	if( $name -imatch "EntityImage" )
	{
		$noprefix = $true
	}

	return $PluralService.Singularize( @(if( $noprefix ) { "" } else { $Publisher }) + $name ).ToLower() -replace '[\s\-]',''	
}

function SchemaName
{
	Param( [string] $name )

	$noprefix = $false;
	if( $name -imatch "EntityImage" )
	{
		$noprefix = $true
	}
	
	return $PluralService.Singularize( @(if( $noprefix ) { "" } else { $Publisher }) + (Get-Culture).TextInfo.ToTitleCase( $name ) ) -replace '[\s\-]',''	
}

function SingularDisplayName
{
	Param( [string] $name )

	$s = $DisplayNamePrefix + $PluralService.Singularize( $name )
	$r = new-object Microsoft.Xrm.Sdk.Label( $s, $lcid )
	return $r
}

function PluralDisplayName
{
	Param( [string] $name )

	$s = $DisplayNamePrefix + $PluralService.Pluralize( $name ) 
	$r = new-object Microsoft.Xrm.Sdk.Label( $s, $lcid )
	return $r
}

function DisplayName
{
	Param( [string] $name )

	$s = $DisplayNamePrefix + $name
	$r = new-object Microsoft.Xrm.Sdk.Label( $s, $lcid )
	return $r
}

Add-Type -assembly System.Data.Entity.Design
$PluralService = [System.Data.Entity.Design.PluralizationServices.PluralizationService]::CreateService( $Culture )

if( $XRM -eq $null -and -not $NoCreateEntities -and -not $NoLoadData )
{
	$XRM = Connect-CrmOnlineDiscovery -InteractiveMode
	if( $XRM -eq $null )
	{
		write-host -foregroundcolor Red "No connection to XRM, exiting..."
		break
	}
}

if( -not ($Publisher -match "_$") )
{
	$Publisher = $Publisher+"_"
}

if( $OrganizationOwned )
{
	$Ownership = [microsoft.xrm.sdk.metadata.OwnershipTypes]::OrganizationOwned 
} 
else 
{ 
	$Ownership = [microsoft.xrm.sdk.metadata.OwnershipTypes]::UserOwned
}

#========================================================================================================================
Write-Host "Reading SQL Schema..."

[array]$tables = ReadSQL -query "
	select t.name, t.object_id 
	from sys.tables t join sys.schemas s on t.schema_id = s.schema_id 
	where t.type_desc = 'USER_TABLE' and s.name = 'Access' 
		and t.name <> 'Trace' and t.name <> 'ActionEvents?' and t.name <> 'ActionEventArguments?'
		and right( t.name, 7 ) <>  '?Images' 
"

[array]$columns = ReadSQL -query "
	select c.object_id, c.column_id, c.name, c.max_length, c.precision, c.scale, 
	t.name as 'type', c.is_nullable, c.is_identity,
	(select u.definition from sys.computed_columns u 
		where u.column_id = c.column_id and u.object_id = c.object_id) as 'computed_expr',
	(select kc.referenced_object_id from sys.foreign_keys k join sys.foreign_key_columns kc 
		on k.object_id = kc.constraint_object_id and kc.parent_column_id = c.column_id 
		where k.parent_object_id = c.object_id) as 'lookup_object_id',
		cp.Properties, cp.DefaultValue, cp.ComputedValue
	from sys.columns c 
	join sys.types t on c.user_type_id = t.user_type_id
	left join AccessSystem.ColumnProperties cp on 
		c.name COLLATE DATABASE_DEFAULT = cp.ColumnName COLLATE DATABASE_DEFAULT and cp.ObjectId = 
		(select aso.ID from AccessSystem.Objects aso join Sys.Tables st on 
		aso.ObjectName COLLATE DATABASE_DEFAULT = st.name COLLATE DATABASE_DEFAULT 
		where st.object_id = c.object_id and aso.ObjectTypeNumber = 100)
	order by c.object_id, c.column_id
"

$TableColumns = @{}
$TableIndex = @{}
$identity = @{}
$TableColumnNames = New-Object System.Collections.Generic.List[System.Object]
For( $t = 0; $t -lt $tables.length; $t++ )
{
	$table = $tables[$t]
	$TableColumns[$table["object_id"]] = New-Object System.Collections.Generic.List[System.Object]
	$TableIndex[$table["object_id"]] = $t
}

For( $c = 0; $c -lt $columns.length; $c++ )
{
	$column = $columns[$c]
	if( $TableIndex[$column["object_id"]] -ne $null )
	{
		$TableColumns[$column["object_id"]].Add( $c )
		$a = "{0}:{1}:{2}:{3}" -f $tables[$TableIndex[$column["object_id"]]]["name"], $column["name"], $column["type"], $column["object_id"]
		$TableColumnNames.Add( $a )
		if( $column["is_identity"] )
		{
			$identity[$column["object_id"]] = $column["name"]
		}
	}
}

$LoadOrder = New-Object System.Collections.Generic.List[System.Object]

Do
{
	$m = 0
	ForEach( $table in $tables )
	{
		if( ! $LoadOrder.Contains($table["object_id"]) )
		{
			$l = 1
			ForEach( $ci in $TableColumns[$table["object_id"]] )
			{
				$column = $columns[$ci]
				if( -not [String]::IsNullOrEmpty($column["lookup_object_id"].ToString()) )
				{
					if( ! $LoadOrder.Contains($column["lookup_object_id"]) )
					{
						$l = 0
					}
				}
			}
			if( $l )
			{
				$LoadOrder.Add( $table["object_id"] )
				$m = 1
			}
		}
	}
}
while( $m )

$PrimaryColumn = @{}
$SQLPresets = $SQLDatabase + ".xml"
if( (Test-Path $SQLPresets) -and $UsePresets )
{
	$p = Import-CliXML $SQLPresets

	$match = 1
	if( $p.TableColumnNames.Count -eq $TableColumnNames.Count )
	{
		for( $t = 0; $t -lt $TableColumnNames.Count; $t++ )
		{
			if( $p.TableColumnNames[$t] -ne $TableColumnNames[$t] )
			{
				$match = 0;
			}
		}
	}

	if( $match )
	{
		Write-Host "    Configuration presets read from $SQLPresets..."
		$PrimaryColumn = $p.PrimaryColumn
	}
}

if( $PrimaryColumn.count -eq 0 )
{
	Write-Host "Select Primary Name fields..."

	For( $t = 0; $t -lt $tables.length; $t++ )
	{
		$table = $tables[$t]
		$tc = $TableColumns[$table["object_id"]]

		"    {0}" -f $table["name"]

		$nums = @()
		For( $c = 0; $c -lt $tc.count; $c++ )	
		{
			$col = $columns[$tc[$c]]
			if( ( [String]::IsNullOrEmpty($col["lookup_object_id"].ToString()) -and $col["type"] -eq "nvarchar" -and $col["max_length"] -le 440 -and $col["max_length"] -ge 0 ) -or $col["is_identity"] )
			{
				$nums += $c			
				"        [{0}] {1} {2}" -f ([string]$nums.Count).PadLeft(2), $columns[$tc[$c]]["name"], $(if($col["is_identity"]) { "(autonumber)" } else {""})

			}
		}
		$nums += -1
		"        [{0}] {1}" -f ([string]$nums.Count).PadLeft(2), "New <Primary Name> calculated column, complete later"

		do
		{
		    $p = Read-Host "        Select Primary Name field by index number"
	    	}
		while( [int]$p -lt 1 -or [int]$p -gt $nums.count )
		$PrimaryColumn[$table["object_id"]] = $(if($nums[[int]$p-1] -eq -1) { -1 } else { $tc[$nums[[int]$p-1]] })
	}

	$p = @{ TableColumnNames = $TableColumnNames
		PrimaryColumn = $PrimaryColumn
		}
	if( $UsePresets )
	{
		$p | Export-CliXML $SQLPresets
	}
}

#========================================================================================================================
Write-Host "Creating Entities and Columns..."

$CrmColumnTypes = @{}
$entities = @{}
$logicalEntityName = @{}

ForEach($table in $tables)
{
	if( ($OneEntity -ne "") -and ($table["name"] -ne $OneEntity) )
	{
		continue;
	}
	
	"    {0}" -f $table["name"]

	$em = new-object Microsoft.Xrm.Sdk.Metadata.EntityMetadata

	$em.SchemaName = SchemaName $table["name"] 
	$em.LogicalName = LogicalName $table["name"]
	$logicalEntityName[$table["name"]] = $em.LogicalName
	$em.DisplayName = SingularDisplayName $table["name"] 
	$em.DisplayCollectionName = PluralDisplayName $table["name"]
	$em.OwnershipType = $Ownership
	$em.IsActivity = 0

	$pn = new-object Microsoft.Xrm.Sdk.Metadata.StringAttributeMetadata
	if( $PrimaryColumn[$table["object_id"]] -eq -1 )
	{
		$name = "Primary Name"
		$maxlength = 220
		$CrmColumnTypes[-1] = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::String	
	}
	else
	{
		$p = $columns[$PrimaryColumn[$table["object_id"]]]
		$name = $p["name"]
		$maxlength = $p["max_length"]/2
		if( !$p["is_nullable"] )
		{
			$pn.RequiredLevel = new-object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty( [Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevel]::ApplicationRequired )
		}
		if( $p["is_identity"] )
		{
			$pn.AutoNumberFormat = "{SEQNUM:8}"
			$maxlength = 8
		}
		$CrmColumnTypes[$PrimaryColumn[$table["object_id"]]] = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::String
	}
	$pn.SchemaName = SchemaName $name
	$pn.DisplayName = DisplayName $name
	$pn.LogicalName = LogicalName $name
	$pn.MaxLength = $maxlength
	
	$cer = new-object Microsoft.Xrm.Sdk.Messages.CreateEntityRequest
	$cer.Entity = $em
	$cer.PrimaryAttribute = $pn

	"       {0}" -f $name

	$entities[$table["name"]] = XRMExecute $cer

	ForEach( $ci in $TableColumns[$table["object_id"]] )
	{
		$col = $columns[$ci]
		if( [String]::IsNullOrEmpty($col["lookup_object_id"].ToString()) -and ($ci -ne $PrimaryColumn[$table["object_id"]]) )
		{
			switch( $col["type"] )
			{
				"nvarchar"
				{
					if( $col["Properties"] -imatch 'axl:texttype="hyperlink"' ) 
					{
						$a = New-Object Microsoft.Xrm.Sdk.Metadata.StringAttributeMetadata
						$a.MaxLength = 4000
						$a.Format = [Microsoft.Xrm.Sdk.Metadata.StringFormat]::Url
						$h = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::String					
					}
					elseif( $col["max_length"] -le 440 -and $col["max_length"] -ge 0 )
					{
						$a = New-Object Microsoft.Xrm.Sdk.Metadata.StringAttributeMetadata
						$a.MaxLength = $col["max_length"] / 2
						$h = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::String
					}
					else
					{
						$a = New-Object Microsoft.Xrm.Sdk.Metadata.MemoAttributeMetadata
						$a.MaxLength = $MaxMemoLength
						$h = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::String						
					}
				}
				"decimal"
				{
					$a = New-Object Microsoft.Xrm.Sdk.Metadata.DecimalAttributeMetadata
					$h = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmDecimal
				}
				"int"
				{
					if( $col["Properties"] -match "axl:image" )
					{
						$a = New-Object Microsoft.Xrm.Sdk.Metadata.ImageAttributeMetadata
						$h = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw
					}
					else
					{
						$a = New-Object Microsoft.Xrm.Sdk.Metadata.IntegerAttributeMetadata
						$h = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmNumber
					}
				}
				"date"
				{
					$a = New-Object Microsoft.Xrm.Sdk.Metadata.DateTimeAttributeMetadata
					$a.Format = [Microsoft.Xrm.Sdk.Metadata.DateTimeFormat]::DateOnly
					$h = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmDateTime
				}
				"datetime2"
				{
					$a = New-Object Microsoft.Xrm.Sdk.Metadata.DateTimeAttributeMetadata
					$a.Format = [Microsoft.Xrm.Sdk.Metadata.DateTimeFormat]::DateAndTime
					$h = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmDateTime
				}
				"bit"
				{
					$a = New-Object Microsoft.Xrm.Sdk.Metadata.BooleanAttributeMetadata
					$yes = new-object Microsoft.Xrm.Sdk.Label( "Yes", $lcid )
					$no = new-object Microsoft.Xrm.Sdk.Label( "No", $lcid )
					$yeso = new-object Microsoft.Xrm.Sdk.Metadata.OptionMetadata( $yes, 1 )
					$noo = new-object Microsoft.Xrm.Sdk.Metadata.OptionMetadata( $no, 0 )
					$a.OptionSet = new-object Microsoft.Xrm.Sdk.Metadata.BooleanOptionSetMetadata( $yeso, $noo )
					$h = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmBoolean
				}
				"float"
				{
					$a = New-Object Microsoft.Xrm.Sdk.Metadata.DoubleAttributeMetadata
					$h = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmFloat
				}
				default
				{
					Write-Host "Unknown type: Column " $col["name"] ", " $col["type"]
				}
			}

			$CrmColumnTypes[$ci] = $h

			$a.DisplayName = DisplayName $col["name"]
			
			if( $h -eq [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw )
			{
				$a.SchemaName = SchemaName "EntityImage"
				$a.LogicalName = LogicalName "EntityImage"
			}
			else
			{			
				$a.SchemaName = SchemaName $col["name"]
				$a.LogicalName = LogicalName $col["name"]
			}

			if( !$col["is_nullable"] )
			{
					$pn.RequiredLevel = new-object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty(
													[Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevel]::ApplicationRequired )			
			}

			$car = new-object Microsoft.Xrm.Sdk.Messages.CreateAttributeRequest
			$car.Attribute = $a
			$car.EntityName = $em.LogicalName

			"       {0}" -f $col["name"]

			$na = XRMExecute $car
		}
	}
}

#========================================================================================================================
Write-Host "Creating Relationships..."

$c = 0

ForEach($table in $tables)
{
	if( $OneEntity -ne "" )
	{
		break;
	}

	$t = 0
	ForEach( $ci in $TableColumns[$table["object_id"]] )
	{
		$col = $columns[$ci]
		if( -not [String]::IsNullOrEmpty($col["lookup_object_id"].ToString()) )
		{
			if( -not $t )
			{
				"    {0}" -f $table["name"]
				$t = 1
				$c = 1
		    }
			"       {0}" -f $col["name"]

			ForEach( $r in $tables )
			{
				if( $col["lookup_object_id"] -eq $r["object_id"] )
				{
					$l = $logicalEntityName[$r["name"]]
				}
			}
			
			$la = New-Object Microsoft.Xrm.Sdk.Metadata.LookupAttributeMetadata

			$la.DisplayName = DisplayName $col["name"]
			$la.LogicalName = LogicalName $col["name"]+"_"+$l
			$la.SchemaName = SchemaName $col["name"]+"_"+$l
			if( !$col["is_nullable"] )
			{
				$la.RequiredLevel = new-object Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevelManagedProperty(
							[Microsoft.Xrm.Sdk.Metadata.AttributeRequiredLevel]::ApplicationRequired )
			}

			$one = New-Object Microsoft.Xrm.Sdk.Metadata.OneToManyRelationshipMetadata

			$one.ReferencedEntity = $l
			$one.ReferencingEntity = $logicalEntityName[$table["name"]]
			$one.SchemaName = $logicalEntityName[$table["name"]]+"_"+$l

			$c1m = New-Object Microsoft.Xrm.Sdk.Messages.CreateOneToManyRequest
			$c1m.Lookup = $la
			$c1m.OneToManyRelationship = $one

			$na = XRMExecute $c1m
		}
	}
}

if( ! $c )
{
	"    (no relationships)"
}

#========================================================================================================================
"Loading Data..."

for( $retry = 0; ($retry -lt 5) -and (-not $xrm.isready); $retry++ )
{
	"Waiting for XRM to be ready..."
	Start-Sleep -s 10
}
if( -not $xrm.isready )
{
	Write-Host -foregroundcolor Red "Giving up, exiting..."
	break
}

$IDMap = @{}

for( $t = 0; $t -lt $LoadOrder.count; $t++ )
{
	$table = $tables[$TableIndex[$LoadOrder[$t]]]

	if( ($OneEntity -ne "") -and ($table["name"] -ne $OneEntity) )
	{
		continue;
	}
	
	[array]$data = ReadSQL ("select * from Access.[" + $table["name"] + "]")

	"    {0} ({1} row{2})" -f $table["name"], $data.length, $(if( $data.length -gt 1 ) {"s"} else {""})

	$tc = $TableColumns[$table["object_id"]]

	$IDMap[$table["object_id"]] = @{}

	foreach( $d in $data )
	{
		$cnr = new-object 'system.collections.generic.dictionary[System.String,Microsoft.Xrm.Tooling.Connector.CrmDataTypeWrapper]'

		for( $c = 0; $c -lt $tc.count; $c++ )
		{
			$column = $columns[$tc[$c]]

			# Access Web Apps stored hyperlinks in teh format "DisplayString#URLString#", this extracts the URL
			if( $column["Properties"] -imatch 'axl:texttype="hyperlink"' )
			{
				if( $d[$column["name"]] -match '[^#]*#([^#]+)' )
				{
					$d[$column["name"]] = $matches[1]
				}
			}
			
			elseif( ($d[$column["name"]] -isnot [DBNull]) )
			{
				$ln = LogicalName $column["name"]				
				$lt = new-object Microsoft.Xrm.Tooling.Connector.CrmDataTypeWrapper

				if( $column["lookup_object_id"] -is [DBNull] )
				{
					$lt.Type = $CrmColumnTypes[$tc[$c]]

					if( $lt.Type -eq [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw )
					{
				    	$ln = LogicalName "entityimage"										
						$image = ReadSQLImage -table ("Access.["+$table["name"]+"?images]") -id $d[$column["name"]] 
						$lt.Value = [System.Byte[]]$image
			    	}
					elseif( $lt.Type -eq [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::String )
					{
						$lt.Value = [string]$d[$column["name"]]
					}
					else
					{
						$lt.Value = $d[$column["name"]]
					}

					$cnr.Add( $ln, $lt )						
				}
				else
				{
					$lt.Value = $IDMap[$column["lookup_object_id"]][$d[$column["name"]]]
					$lt.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Lookup

					$lt.ReferencedEntity = LogicalName $tables[$TableIndex[$column["lookup_object_id"]]]["name"]

					$cnr.Add( $ln, $lt )
				}
			}
		}

		$ln = LogicalName $table["name"]

		$na = XRMCreateNewRecord $ln $cnr
				
		$IDMap[$table["object_id"]][$d[$identity[$table["object_id"]]]] = $na
	}	
}
}


