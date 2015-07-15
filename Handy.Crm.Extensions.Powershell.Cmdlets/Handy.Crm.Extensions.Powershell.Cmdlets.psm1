Set-StrictMode -Version Latest


Function Get-CRMOptionSetValue
{
    [CmdletBinding()]
    [OutputType([Microsoft.Xrm.Sdk.OptionSetValue])]
    Param(
        [Parameter(Mandatory=$true)]
        [int]$Value
    )

    Begin {}
    Process
    {
        $optionSetValue = New-Object -TypeName 'Microsoft.Xrm.Sdk.OptionSetValue' -ArgumentList $Value

        $optionSetValue
    }
    End {}
}


Function Get-CRMEntityReference
{
    [CmdletBinding()]
    [OutputType([Microsoft.Xrm.Sdk.EntityReference])]
    Param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$EntityName,

        [Parameter(Mandatory=$true)]
        [guid]$Id
    )

    Begin {}
    Process
    {
        $entityReference = New-Object -TypeName 'Microsoft.Xrm.Sdk.EntityReference' -ArgumentList $EntityName, $Id

        $entityReference
    }
    End {}
}


Function Merge-CRMAttributes
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$From,

        [Parameter(Mandatory=$true)]
        [hashtable]$To
    )

    foreach ($key in $From.Keys)
    {
        $To[$key] = $From[$key]
    }
}


Function Get-CRMSolution
{
    [CmdletBinding()]
    [OutputType([Microsoft.Xrm.Sdk.Entity])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Client.CrmConnection]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )

    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true" count="1">
  <entity name="solution">
    <all-attributes />
    <filter type="and">
      <condition attribute="uniquename" operator="eq" value="{0}" />
    </filter>
  </entity>
</fetch>
"@

    $solution = Get-CRMEntity -Connection $Connection -FetchXml ([string]::Format($fetchXml, $Name)) | Select-Object -Index 0
    $solution
}


Function Set-CRMSDKStepState
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Client.CrmConnection]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Xrm.Sdk.Entity]$Solution,

        [Parameter(Mandatory=$true)]
        [ValidateSet('Enabled', 'Disabled')]
        [string]$State,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$Include = [string]::Empty,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$Exclude = [string]::Empty
    )

    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
  <entity name="sdkmessageprocessingstep">
    <all-attributes />
    <filter type="and">
      <condition attribute="solutionid" operator="eq" value="{0}" />
      {1}
      {2}
    </filter>
  </entity>
</fetch>
"@

    if ($Include -ne [string]::Empty)
    {
        $includeCondition = "<condition attribute=`"name`" operator=`"like`" value=`"$Include`" />"
    }
    else
    {
        $includeCondition = [string]::Empty
    }

    if ($Exclude -ne [string]::Empty)
    {
        $excludeCondition = "<condition attribute=`"name`" operator=`"not-like`" value=`"$Exclude`" />"
    }
    else
    {
        $excludeCondition = [string]::Empty
    }

    Write-Verbose "FetchXML:"
    Write-Verbose "$([string]::Format($fetchXml, $solution.Id, $includeCondition, $excludeCondition))"

    $steps = Get-CRMEntity -Connection $Connection -FetchXml ([string]::Format($fetchXml, $Solution.Id, $includeCondition, $excludeCondition))

    Write-Verbose "Found $($steps.Count) steps"
    Write-Verbose "Enabling them"

    switch ($State)
    {
        'Enabled'
        {
            $response = Set-CRMState -Connection $Connection -Entity $steps -State 0 -Status 1
        }

        'Disabled'
        {
            $response = Set-CRMState -Connection $Connection -Entity $steps -State 1 -Status 2
        }
    }

    $response
}


Function Get-CRMBusinessUnit
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Client.CrmConnection]$Connection,

        [Parameter(Mandatory=$true)]
        [string]$Name
    )

    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true">
  <entity name="businessunit">
    <all-attributes />
    <filter type="and">
      <condition attribute="name" operator="eq" value="{0}"/>
    </filter>
  </entity>
</fetch>
"@

    $bu = Get-CRMEntity -Connection $Connection -FetchXml ([string]::Format($fetchXml, $Name))

    $bu
}


Function New-CRMBusinessUnit
{
    [CmdletBinding()]
    [OutputType([System.Guid])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Client.CrmConnection]$Connection,

        [Parameter(Mandatory=$true)]
        [string]$Name,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [guid]$ParentBusinessUnitId
    )

    if ($ParentBusinessUnitId -eq $null)
    {
        Write-Verbose "Getting root BU"

        $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true" count="1">
  <entity name="businessunit">
    <all-attributes />
    <filter type="and">
      <condition attribute="parentbusinessunitid" operator="null" />
    </filter>
  </entity>
</fetch>
"@

        $bu = Get-CRMEntity -Connection $Connection -FetchXml $fetchXml | Select-Object -Index 0
        $ParentBusinessUnitId = $bu.Id
    }

    $buAttributes = @{}
    $buAttributes['name'] = $Name
    $buAttributes['parentbusinessunitid'] = Get-CRMEntityReference -EntityName 'businessunit' -Id $ParentBusinessUnitId
    
    $resp = New-CRMEntity -Connection $Connection -EntityName 'businessunit' -Attributes $buAttributes -ReturnResponses

    $resp.Responses[0].Response.id
}


Function New-CRMUser
{
    [CmdletBinding()]
    [OutputType([System.Guid])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Client.CrmConnection]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$DomainName,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$FirstName,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$LastName,

        [Parameter(Mandatory=$true)]
        [guid]$BusinessUnitId
    )

    $userAttributes = @{}
    $userAttributes['domainname'] = $DomainName
    $userAttributes['firstname'] = $FirstName
    $userAttributes['lastname'] = $LastName
    $userAttributes['fullname'] = "$($FirstName) $($LastName)"
    $userAttributes['businessunitid'] = Get-CRMEntityReference -EntityName 'businessunit' -Id $BusinessUnitId

    $resp = New-CRMEntity -Connection $Connection -EntityName 'systemuser' -Attributes $userAttributes -ReturnResponses

    $resp.Responses[0].Response.id
}


Function Get-CRMDuplicateRule
{
    [CmdletBinding()]
    [OutputType([Microsoft.Xrm.Sdk.Entity])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Client.CrmConnection]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory=$false)]
        [int]$StateCode = 0,

        [Parameter(Mandatory=$false)]
        [int]$StatusCode = 0
    )
    # statecode	statecodename	statuscode	statuscodename
    # 0         Inactive        0           Unpublished
    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true">
  <entity name="duplicaterule">
    <all-attributes />
    <filter type="and">
      <condition attribute="name" operator="eq" value="{0}" />
      <condition attribute="statecode" operator="eq" value="{1}" />
      <condition attribute="statuscode" operator="eq" value="{2}" />
    </filter>
  </entity>
</fetch>
"@

    $duplicateRule = Get-CRMEntity -Connection $Connection -FetchXml ([string]::Format($fetchXml, $Name, $StateCode, $StatusCode)) | Select-Object -Index 0

    $duplicateRule
}

Function New-CRMQueue
{
    [CmdletBinding()]
    [OutputType([System.Guid])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Client.CrmConnection]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [guid]$OwnerId,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$Email = [string]::Empty
    )

    $queueAttributes = @{}
    $queueAttributes['name'] = $Name
    $queueAttributes['ownerid'] = Get-CRMEntityReference -EntityName 'systemuser' -Id $OwnerId
    $queueAttributes['outgoingemaildeliverymethod'] = Get-CRMOptionSetValue -Value 2 # Email Router

    if ($Email -ne [string]::Empty)
    {
        $queueAttributes['emailaddress'] = $Email
    }

    # TODO Add returnresponse
    $resp = New-CRMEntity -Connection $Connection -EntityName 'queue' -Attributes $queueAttributes -ReturnResponses

    $queueId = $resp.Responses[0].Response.id

    if ($Email -ne [string]::Empty)
    {
        $queueEntity = Get-CRMEntityById -Connection $Connection -EntityName 'queue' -Id $queueId -Colummns 'emailrouteraccessapproval'
        $queueEntity['emailrouteraccessapproval'] = Get-CRMOptionSetValue -Value 1 # Approved
        $resp = Update-CRMEntity -Connection $Connection -Entity $queueEntity
    }

    $queueId
}

Function Set-CRMQueueForUser
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Client.CrmConnection]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [guid]$UserId,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [guid]$QueueId
    )

    $userEntity = Get-CRMEntityById -Connection $Connection -EntityName 'systemuser' -Id $UserId -Columns 'queueid'

    $userEntity['queueid'] = Get-CRMEntityReference -EntityName 'queue' -Id $QueueId

    $resp = Update-CRMEntity -Connection $Connection -Entity $userEntity -ReturnResponses

    $resp
}


Function Add-CRMRoleForUser
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Client.CrmConnection]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [Microsoft.Xrm.Sdk.Entity]$User,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$RoleName
    )

    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true" count="1">
  <entity name="role">
    <all-attributes />
    <filter type="and">
      <condition attribute="name" operator="eq" value="{0}" />
      <condition attribute="businessunitid" operator="eq" value="{1}" />
    </filter>
  </entity>
</fetch>
"@

    $role = Get-CRMEntity -Connection $Connection -FetchXml ([string]::Format($fetchXml, $RoleName, $User['businessunitid'].Id)) | Select-Object -Index 0

    if ($role -eq $null)
    {
        throw "Couldn't find role $($RoleName) or something went wrong."
    }

    $reference = Get-CRMEntityReference -EntityName $role.LogicalName -Id $role.Id

    Add-CRMAssociation -Connection $Connection -EntityName $User.LogicalName -Id $User.Id -Relationship 'systemuserroles_association' -RelatedEntity $reference
}

Function Remove-CRMRoleForUser
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Client.CrmConnection]$Connection,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [Microsoft.Xrm.Sdk.Entity]$User,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$RoleName
    )

    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true" count="1">
  <entity name="role">
    <all-attributes />
    <filter type="and">
      <condition attribute="name" operator="eq" value="{0}" />
      <condition attribute="businessunitid" operator="eq" value="{1}" />
    </filter>
  </entity>
</fetch>
"@

    $role = Get-CRMEntity -Connection $Connection -FetchXml ([string]::Format($fetchXml, $RoleName, $User['businessunitid'].Id)) | Select-Object -Index 0

    if ($role -eq $null)
    {
        throw "Couldn't find role $($RoleName) or something went wrong."
    }

    $reference = Get-CRMEntityReference -EntityName $role.LogicalName -Id $role.Id

    Remove-CRMAssociation -Connection $Connection -EntityName $User.LogicalName -Id $User.Id -Relationship 'systemuserroles_association' -RelatedEntity $reference
}


Add-Type -TypeDefinition @"
    public enum CurrencyCodeEnum
    {
        CHF,
        EUR,
        GBP,
        RUB,
        USD
    }
"@

Function Get-CRMTransactionCurrency
{
    [CmdletBinding()]
    [OutputType([Microsoft.Xrm.Sdk.Entity])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Client.CrmConnection]$Connection,

        [Parameter(Mandatory=$true)]
        [CurrencyCodeEnum]$CurrencyCode
    )

    $fetchXml = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true" count="1">
  <entity name="transactioncurrency">
    <all-attributes />
    <filter type="and">
      <condition attribute="isocurrencycode" operator="eq" value="{0}" />
    </filter>
  </entity>
</fetch>
"@

    $tc = Get-CRMEntity -Connection $Connection -FetchXml ([string]::Format($fetchXml, $CurrencyCode)) | Select-Object -Index 0

    $tc
}

Function New-CRMTransactionCurrency
{
    [CmdletBinding()]
    [OutputType([System.Guid])]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Client.CrmConnection]$Connection,

        [Parameter(Mandatory=$true)]
        [CurrencyCodeEnum]$CurrencyCode,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [hashtable]$AdditionAttributes = @{}
    )

    # Settings
    $currencyAttributes = @{}
    $currencyAttributes['currencyprecision'] = 2
    $currencyAttributes['exchangerate'] = [decimal]1.0
    $currencyAttributes['isocurrencycode'] = [string]$CurrencyCode

    switch ($CurrencyCode)
    {
        ([CurrencyCodeEnum]::CHF -as [string])
        {
            $currencyAttributes['currencyname'] = 'Swiss franc'
            $currencyAttributes['currencysymbol'] = 'Fr.'
            break
        }

        ([CurrencyCodeEnum]::EUR -as [string])
        {
            $currencyAttributes['currencyname'] = 'Euro'
            $currencyAttributes['currencysymbol'] = '€'
            break
        }

        ([CurrencyCodeEnum]::GBP  -as [string])
        {
            $currencyAttributes['currencyname'] = 'Pound Sterling'
            $currencyAttributes['currencysymbol'] = '£'
            break
        }

        ([CurrencyCodeEnum]::RUB -as [string])
        {
            $currencyAttributes['currencyname'] = 'Russian ruble'
            $currencyAttributes['currencysymbol'] = 'р.'
            break
        }

        ([CurrencyCodeEnum]::USD -as [string])
        {
            $currencyAttributes['currencyname'] = 'US Dollar'
            $currencyAttributes['currencysymbol'] = '$'
            break
        }

        defaut
        {
            break
        }
    }

    Merge-CRMAttributes -From $AdditionAttributes -To $currencyAttributes

    $resp = New-CRMEntity -Connection $Connection -EntityName 'transactioncurrency' -Attributes $currencyAttributes -ReturnResponses

    $resp.Responses[0].Response.id
}

Function Set-CRMSDKStepMode
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Xrm.Client.CrmConnection]$Connection,

        [Parameter(Mandatory=$true)]
        [guid]$Id,

        [Parameter(Mandatory=$true)]
        [ValidateSet('Asynchronous', 'Synchronous')]
        [string]$Mode,

        [Parameter(Mandatory=$false)]
        [switch]$SetAutoDelete
    )

    $step = Get-CRMEntityById -Connection $Connection -EntityName 'sdkmessageprocessingstep' -Id $Id -Columns 'mode', 'asyncautodelete'
    # Sync - 0, Async - 1
    # Yes - 1, No - 0

    switch ($Mode)
    {
        'Asynchronous'
        {
            $step['mode'] = Get-CRMOptionSetValue -Value 1
            $step['asyncautodelete'] = $SetAutoDelete.IsPresent
            break
        }
        'Synchronous'
        {
            $step['mode'] = Get-CRMOptionSetValue -Value 0
            $step['asyncautodelete'] = $false
            break
        }
    }

    $resp = Update-CRMEntity -Connection $Connection -Entity $step
    $resp
}
