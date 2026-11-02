# Post-deployment configuration for the sharepointApp.
#
# The project template retains one operation descriptor under infra/operations.
# This script resolves the selected SharePoint site/library parameters,
# authorizes the shared connection, creates one Poll Trigger Config, and
# configures the Function App.

$ErrorActionPreference = 'Stop'

Write-Host "Post-deployment configuration..." -ForegroundColor Yellow

# --- Resolve the retained operation -----------------------------------------
$operationsRoot = Join-Path (Split-Path $PSScriptRoot -Parent) 'operations'
$descriptorFiles = if (Test-Path $operationsRoot) {
    @(Get-ChildItem -Path $operationsRoot -Filter 'trigger-config.json' -File -Recurse)
} else {
    @()
}

if ($descriptorFiles.Count -ne 1) {
    Write-Host "ERROR: expected exactly one retained operation descriptor under '$operationsRoot', found $($descriptorFiles.Count)." -ForegroundColor Red
    exit 1
}

$descriptor = Get-Content -Path $descriptorFiles[0].FullName -Raw | ConvertFrom-Json
foreach ($propertyName in @('functionName', 'triggerConfigName', 'operationName', 'description')) {
    $property = $descriptor.PSObject.Properties[$propertyName]
    if (-not $property -or [string]::IsNullOrWhiteSpace([string]$property.Value)) {
        Write-Host "ERROR: operation descriptor is missing '$propertyName'." -ForegroundColor Red
        exit 1
    }
}
if (-not $descriptor.PSObject.Properties['parameters']) {
    Write-Host "ERROR: operation descriptor must contain a parameters array." -ForegroundColor Red
    exit 1
}

function Resolve-OperationParameters {
    param(
        [Parameter(Mandatory)] [object[]] $Definitions,
        [Parameter(Mandatory)] [hashtable] $EnvironmentValues
    )

    $resolved = [System.Collections.Generic.List[object]]::new()
    foreach ($definition in $Definitions) {
        if ([string]::IsNullOrWhiteSpace([string]$definition.name)) {
            throw "Operation descriptor contains a parameter without a name."
        }

        $valueProperty = $definition.PSObject.Properties['value']
        if ($valueProperty) {
            $resolved.Add(@{ name = $definition.name; value = $valueProperty.Value })
            continue
        }

        $sourceName = [string]$definition.valueFromEnvironment
        if ([string]::IsNullOrWhiteSpace($sourceName)) {
            throw "Parameter '$($definition.name)' must define value or valueFromEnvironment."
        }

        $value = $EnvironmentValues[$sourceName]
        if ([string]::IsNullOrWhiteSpace([string]$value)) {
            if ($definition.required) {
                throw "Required operation parameter '$($definition.name)' has no value in '$sourceName'."
            }
            continue
        }

        $resolved.Add(@{ name = $definition.name; value = $value })
    }

    return $resolved.ToArray()
}

# --- Read azd outputs --------------------------------------------------------
$outputsJson = azd env get-values --output json
$outputs = ($outputsJson -join [Environment]::NewLine) | ConvertFrom-Json

$resourceGroupName = $outputs.resourceGroupName
$connectorNamespaceName = $outputs.connectorNamespaceName
$sharepointConnectionName = $outputs.sharepointConnectionName
$functionAppName = $outputs.functionAppName
$subscriptionId = $outputs.AZURE_SUBSCRIPTION_ID
if (-not $subscriptionId) {
    $subscriptionId = az account show --query id -o tsv
}

if (-not $resourceGroupName -or -not $connectorNamespaceName -or -not $sharepointConnectionName -or -not $functionAppName -or -not $subscriptionId) {
    Write-Host "ERROR: required azd outputs missing. Run 'azd provision' first." -ForegroundColor Red
    exit 1
}

$apiVersion = '2026-05-01-preview'
$namespaceResourceId = "/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.Web/connectorGateways/$connectorNamespaceName"
$triggerConfigResourceId = "$namespaceResourceId/triggerConfigs/$($descriptor.triggerConfigName)"

# --- Resolve canonical SharePoint parameters -------------------------------
$sharepointSiteAddress = [string]$outputs.SHAREPOINT_SITE_ADDRESS
$sharepointTable = [string]$outputs.SHAREPOINT_TABLE
$sharepointFolderPath = [string]$outputs.SHAREPOINT_FOLDER_PATH
$sharepointTableWasConfigured = -not [string]::IsNullOrWhiteSpace($sharepointTable)

function Select-FromList {
    param(
        [Parameter(Mandatory)] [string] $Title,
        [Parameter(Mandatory)] [object[]] $Items,
        [Parameter(Mandatory)] [string] $LabelProperty,
        [string] $SubLabelProperty
    )

    Write-Host ""
    Write-Host $Title -ForegroundColor Cyan
    for ($index = 0; $index -lt $Items.Count; $index++) {
        $item = $Items[$index]
        $label = $item.$LabelProperty
        $subLabel = if ($SubLabelProperty) { $item.$SubLabelProperty } else { $null }
        if ($subLabel) {
            Write-Host ("  [{0}] {1}  ({2})" -f ($index + 1), $label, $subLabel)
        } else {
            Write-Host ("  [{0}] {1}" -f ($index + 1), $label)
        }
    }

    while ($true) {
        $answer = Read-Host "Enter number (1-$($Items.Count))"
        $selection = 0
        if ([int]::TryParse($answer, [ref]$selection) -and $selection -ge 1 -and $selection -le $Items.Count) {
            return $Items[$selection - 1]
        }
        Write-Host "Invalid selection." -ForegroundColor Yellow
    }
}

function Invoke-Graph {
    param(
        [Parameter(Mandatory)] [string] $Url,
        [string[]] $UriParameters,
        [switch] $Quiet
    )

    $arguments = @(
        'rest',
        '--method', 'get',
        '--url', $Url,
        '--resource', 'https://graph.microsoft.com',
        '--output', 'json',
        '--only-show-errors'
    )
    if ($UriParameters) {
        $arguments += '--uri-parameters'
        $arguments += $UriParameters
    }

    $raw = & az @arguments 2>&1
    if ($LASTEXITCODE -ne 0 -or -not $raw) {
        if (-not $Quiet -and $raw) {
            Write-Host "   Microsoft Graph request failed." -ForegroundColor DarkGray
        }
        return $null
    }

    try {
        return (($raw -join [Environment]::NewLine) | ConvertFrom-Json)
    } catch {
        if (-not $Quiet) {
            Write-Host "   Microsoft Graph returned an unexpected response." -ForegroundColor DarkGray
        }
        return $null
    }
}

$selectedSiteId = $null
if (-not $sharepointSiteAddress) {
    Write-Host ""
    Write-Host "Fetching SharePoint sites accessible to your account from Microsoft Graph..." -ForegroundColor Yellow
    $sitesResponse = Invoke-Graph `
        -Url 'https://graph.microsoft.com/v1.0/sites' `
        -UriParameters @('search=*', '$select=id,displayName,webUrl', '$top=100') `
        -Quiet
    $sites = @($sitesResponse.value | Where-Object { $_.displayName } | Sort-Object displayName)

    if ($sites.Count -gt 0) {
        $selectedSite = Select-FromList `
            -Title 'Select a SharePoint site:' `
            -Items $sites `
            -LabelProperty 'displayName' `
            -SubLabelProperty 'webUrl'
        $sharepointSiteAddress = $selectedSite.webUrl
        $selectedSiteId = $selectedSite.id
    } else {
        $sharepointSiteAddress = Read-Host "SharePoint site URL (for example, https://contoso.sharepoint.com/sites/MySite)"
    }

    if (-not $sharepointSiteAddress) {
        Write-Host "ERROR: SHAREPOINT_SITE_ADDRESS is required." -ForegroundColor Red
        exit 1
    }
    azd env set SHAREPOINT_SITE_ADDRESS $sharepointSiteAddress | Out-Null
}

if (-not $selectedSiteId) {
    try {
        $siteUri = [Uri]$sharepointSiteAddress
        $relativePath = $siteUri.AbsolutePath.TrimEnd('/')
        $resolvedSite = Invoke-Graph `
            -Url "https://graph.microsoft.com/v1.0/sites/$($siteUri.Host):$relativePath" `
            -UriParameters @('$select=id') `
            -Quiet
        $selectedSiteId = $resolvedSite.id
    } catch {
        $selectedSiteId = $null
    }
}

$selectedListId = $null
if (-not $sharepointTable) {
    if ($selectedSiteId) {
        $listsResponse = Invoke-Graph `
            -Url "https://graph.microsoft.com/v1.0/sites/$selectedSiteId/lists" `
            -UriParameters @('$select=id,displayName,list,webUrl', '$top=200') `
            -Quiet
        $documentLibraries = @(
            $listsResponse.value |
                Where-Object { $_.list.template -eq 'documentLibrary' } |
                Sort-Object displayName
        )
        if ($documentLibraries.Count -gt 0) {
            $selectedLibrary = Select-FromList `
                -Title 'Select a SharePoint document library:' `
                -Items $documentLibraries `
                -LabelProperty 'displayName' `
                -SubLabelProperty 'webUrl'
            $sharepointTable = $selectedLibrary.id
            $selectedListId = $selectedLibrary.id
        }
    }

    if (-not $sharepointTable) {
        $sharepointTable = Read-Host "SharePoint document library ID or name"
    }
    if (-not $sharepointTable) {
        Write-Host "ERROR: SHAREPOINT_TABLE is required." -ForegroundColor Red
        exit 1
    }
    azd env set SHAREPOINT_TABLE $sharepointTable | Out-Null
}

if (-not $sharepointFolderPath -and -not $sharepointTableWasConfigured) {
    $listIdForFolders = if ($selectedListId) { $selectedListId } else { $sharepointTable }
    $folderChoices = @()
    if ($selectedSiteId -and $listIdForFolders) {
        $childrenResponse = Invoke-Graph `
            -Url "https://graph.microsoft.com/v1.0/sites/$selectedSiteId/lists/$listIdForFolders/drive/root/children" `
            -UriParameters @('$select=name,folder', '$top=200') `
            -Quiet
        $folderChoices = @(
            [pscustomobject]@{ displayName = '(whole library)'; value = '' }
            $childrenResponse.value |
                Where-Object { $_.folder } |
                Sort-Object name |
                ForEach-Object { [pscustomobject]@{ displayName = $_.name; value = "/$($_.name)" } }
        )
    }

    if ($folderChoices.Count -gt 1) {
        $selectedFolder = Select-FromList `
            -Title 'Select a folder, or use the whole library:' `
            -Items $folderChoices `
            -LabelProperty 'displayName'
        $sharepointFolderPath = $selectedFolder.value
    } else {
        $sharepointFolderPath = Read-Host "Optional folder path (press Enter for the whole library)"
    }

    if ($sharepointFolderPath) {
        azd env set SHAREPOINT_FOLDER_PATH $sharepointFolderPath | Out-Null
    }
}

$parameterValues = @{
    SHAREPOINT_SITE_ADDRESS = $sharepointSiteAddress
    SHAREPOINT_TABLE = $sharepointTable
    SHAREPOINT_FOLDER_PATH = $sharepointFolderPath
}

# --- Verify the connector-namespace az CLI extension is installed ----------
$extInstalled = az extension show --name connector-namespace --query name -o tsv 2>$null
if (-not $extInstalled) {
    Write-Host "ERROR: 'connector-namespace' Azure CLI extension is not installed." -ForegroundColor Red
    Write-Host "       Download the latest connector_namespace wheel from:" -ForegroundColor Red
    Write-Host "       https://github.com/Azure/Connectors/releases" -ForegroundColor Red
    exit 2
}

# --- Authorize the SharePoint connection (OAuth consent) --------------------
function Test-ConnectionExists {
    param([Parameter(Mandatory)] [string] $ConnectionName)

    az connector-namespace connection show `
        -g $resourceGroupName `
        --namespace $connectorNamespaceName `
        -n $ConnectionName `
        -o none 2>$null
    return ($LASTEXITCODE -eq 0)
}

function Add-ConnectionAccessPolicy {
    param(
        [Parameter(Mandatory)] [string] $ConnectionName,
        [Parameter(Mandatory)] [string] $PolicyName,
        [Parameter(Mandatory)] [string] $PrincipalObjectId,
        [Parameter(Mandatory)] [string] $TenantId,
        [switch] $Required
    )

    az connector-namespace connection access-policy create `
        -g $resourceGroupName `
        --namespace $connectorNamespaceName `
        --connection-name $ConnectionName `
        -n $PolicyName `
        --principal-type 'ActiveDirectory' `
        --principal-object-id $PrincipalObjectId `
        --principal-tenant-id $TenantId `
        -o none 2>$null

    if ($LASTEXITCODE -ne 0) {
        if ($Required) {
            throw "Failed to create required connection access policy '$PolicyName'."
        }
        Write-Host "   WARNING: could not create optional access policy '$PolicyName'." -ForegroundColor Yellow
    }
}

function New-SharePointConnection {
    param([Parameter(Mandatory)] [string] $ConnectionName)

    Write-Host "-> Creating SharePoint connection '$ConnectionName' on namespace '$connectorNamespaceName'..." -ForegroundColor Yellow
    az connector-namespace connection create `
        -g $resourceGroupName `
        --namespace $connectorNamespaceName `
        -n $ConnectionName `
        --connector-name 'sharepointonline' `
        -o none
    if ($LASTEXITCODE -ne 0) {
        return $false
    }

    $tenantId = az account show --query tenantId -o tsv
    $functionAppPrincipalId = az functionapp identity show `
        -g $resourceGroupName `
        -n $functionAppName `
        --query principalId `
        -o tsv
    $namespacePrincipalId = az resource show `
        --ids $namespaceResourceId `
        --api-version $apiVersion `
        --query identity.principalId `
        -o tsv

    if (-not $functionAppPrincipalId -or -not $namespacePrincipalId) {
        throw "Required managed identity principal IDs were not available for connection access policies."
    }

    Add-ConnectionAccessPolicy -ConnectionName $ConnectionName -PolicyName 'functionapp-msi' -PrincipalObjectId $functionAppPrincipalId -TenantId $tenantId -Required
    Add-ConnectionAccessPolicy -ConnectionName $ConnectionName -PolicyName 'namespace-trigger-runtime' -PrincipalObjectId $namespacePrincipalId -TenantId $tenantId -Required

    $userPrincipalId = az ad signed-in-user show --query id -o tsv 2>$null
    if ($userPrincipalId) {
        Add-ConnectionAccessPolicy -ConnectionName $ConnectionName -PolicyName 'dev-user' -PrincipalObjectId $userPrincipalId -TenantId $tenantId
    }

    return $true
}

function Invoke-AuthorizeConnection {
    param([Parameter(Mandatory)] [string] $ConnectionName)

    Write-Host "-> Authorizing SharePoint Online connection: $ConnectionName" -ForegroundColor Cyan

    if (-not (Test-ConnectionExists -ConnectionName $ConnectionName)) {
        if (-not (New-SharePointConnection -ConnectionName $ConnectionName)) {
            return $false
        }
    }

    $currentStatus = az connector-namespace connection show `
        -g $resourceGroupName `
        --namespace $connectorNamespaceName `
        -n $ConnectionName `
        --query 'properties.overallStatus' `
        -o tsv 2>$null
    if ($currentStatus -and $currentStatus.ToLowerInvariant() -eq 'connected') {
        Write-Host "   already Connected; skipping consent flow" -ForegroundColor Green
        return $true
    }

    $parametersPath = Join-Path $PSScriptRoot (".consent-parameters-{0}.json" -f [guid]::NewGuid())
    [IO.File]::WriteAllText(
        $parametersPath,
        '[{"parameterName":"token","redirectUrl":"https://portal.azure.com"}]',
        [Text.UTF8Encoding]::new($false))
    try {
        $consentJson = $null
        for ($attempt = 1; $attempt -le 5; $attempt++) {
            $consentJson = az connector-namespace connection list-consent-links `
                -g $resourceGroupName `
                --namespace $connectorNamespaceName `
                --connection-name $ConnectionName `
                --parameters "@$parametersPath" `
                -o json
            if ($LASTEXITCODE -eq 0 -and $consentJson) {
                break
            }
            Write-Host "   list-consent-links attempt $attempt failed; retrying in 5s..." -ForegroundColor Yellow
            Start-Sleep -Seconds 5
        }
    } finally {
        Remove-Item $parametersPath -ErrorAction SilentlyContinue
    }

    if (-not $consentJson) {
        return $false
    }

    $parsedConsent = ($consentJson -join [Environment]::NewLine) | ConvertFrom-Json
    $link = if ($parsedConsent.value -and $parsedConsent.value.Count -gt 0) {
        $parsedConsent.value[0].link
    } else {
        $parsedConsent.link
    }
    if (-not $link) {
        return $false
    }

    Write-Host "   opening browser for OAuth consent..." -ForegroundColor Cyan
    try {
        Start-Process $link | Out-Null
    } catch {
        return $false
    }

    $deadline = (Get-Date).AddMinutes(5)
    $lastStatus = ''
    while ((Get-Date) -lt $deadline) {
        $status = az connector-namespace connection show `
            -g $resourceGroupName `
            --namespace $connectorNamespaceName `
            -n $ConnectionName `
            --query 'properties.overallStatus' `
            -o tsv 2>$null
        if ($status -ne $lastStatus) {
            Write-Host "   status: $(if ($status) { $status } else { '?' })" -ForegroundColor Cyan
            $lastStatus = $status
        }
        if ($status -and $status.ToLowerInvariant() -eq 'connected') {
            Write-Host "   connection authenticated" -ForegroundColor Green
            return $true
        }
        Start-Sleep -Seconds 3
    }

    return $false
}

Write-Host ""
Write-Host "Authorizing SharePoint Online connection via Azure CLI..." -ForegroundColor Yellow
if (-not (Invoke-AuthorizeConnection -ConnectionName $sharepointConnectionName)) {
    Write-Host "ERROR: SharePoint connection is not Connected. Complete OAuth consent and rerun the postdeploy hook." -ForegroundColor Red
    exit 1
}

# --- Create the retained Poll Trigger Config --------------------------------
$parameters = @(Resolve-OperationParameters -Definitions @($descriptor.parameters) -EnvironmentValues $parameterValues)
$requestBody = @{
    properties = @{
        connectionDetails = @{
            connectionName = $sharepointConnectionName
            connectorName = 'sharepointonline'
        }
        deliveryMode = 'Poll'
        description = $descriptor.description
        operationName = $descriptor.operationName
        parameters = $parameters
        state = 'Enabled'
        type = 'NotSpecified'
    }
} | ConvertTo-Json -Depth 8 -Compress

$requestBodyPath = Join-Path $PSScriptRoot (".trigger-config-{0}.json" -f [guid]::NewGuid())
[IO.File]::WriteAllText($requestBodyPath, $requestBody, [Text.UTF8Encoding]::new($false))
try {
    $triggerConfigJson = az rest `
        --method put `
        --uri "$triggerConfigResourceId?api-version=$apiVersion" `
        --body "@$requestBodyPath" `
        --only-show-errors `
        -o json
} finally {
    Remove-Item $requestBodyPath -ErrorAction SilentlyContinue
}
if ($LASTEXITCODE -ne 0 -or -not $triggerConfigJson) {
    Write-Host "ERROR: failed to create the Connector Namespace Poll Trigger Config." -ForegroundColor Red
    exit 1
}

$triggerConfig = ($triggerConfigJson -join [Environment]::NewLine) | ConvertFrom-Json
if (-not $triggerConfig.properties.pollingEndpoints) {
    $triggerConfigJson = az rest `
        --method get `
        --uri "$triggerConfigResourceId?api-version=$apiVersion" `
        --only-show-errors `
        -o json
    if ($LASTEXITCODE -ne 0 -or -not $triggerConfigJson) {
        Write-Host "ERROR: failed to retrieve the Connector Namespace Poll Trigger Config." -ForegroundColor Red
        exit 1
    }
    $triggerConfig = ($triggerConfigJson -join [Environment]::NewLine) | ConvertFrom-Json
}

$receiveUriText = [string]$triggerConfig.properties.pollingEndpoints.receiveUri
if ($triggerConfig.properties.state -ne 'Enabled' -or [string]::IsNullOrWhiteSpace($receiveUriText)) {
    Write-Host "ERROR: the Poll Trigger Config was not enabled with a receive endpoint." -ForegroundColor Red
    exit 1
}

$receiveUri = [Uri]$receiveUriText
$pathMatch = [regex]::Match(
    $receiveUri.AbsolutePath,
    '^(?<base>/api(?:/connectorGateways/[^/]+)?)/triggerConfigs/[^/]+/receive$',
    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
if (-not $pathMatch.Success) {
    Write-Host "ERROR: the service-generated receive endpoint has an unsupported path." -ForegroundColor Red
    exit 1
}

$endpointBuilder = [UriBuilder]$receiveUri
$endpointBuilder.Path = if ($pathMatch.Groups['base'].Value -eq '/api') {
    ''
} else {
    $pathMatch.Groups['base'].Value
}
$endpointBuilder.Query = ''
$endpointBuilder.Fragment = ''
$pollingEndpoint = $endpointBuilder.Uri.AbsoluteUri

az functionapp config appsettings set `
    -g $resourceGroupName `
    -n $functionAppName `
    --settings `
        "ConnectorNamespace__pollingEndpoint=$pollingEndpoint" `
        'ConnectorNamespace__credential=managedidentity' `
        "ConnectorTriggerConfigName=$($descriptor.triggerConfigName)" `
    -o none
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: failed to configure the Function App Poll settings." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Poll Trigger Config '$($descriptor.triggerConfigName)' is Enabled for $($descriptor.functionName)." -ForegroundColor Green
Write-Host "Post-deployment configuration complete." -ForegroundColor Green
