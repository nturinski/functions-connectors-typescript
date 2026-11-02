#!/bin/bash
# Post-deployment configuration for the sharepointApp.
#
# The project template retains one operation descriptor under infra/operations.
# This script resolves the selected SharePoint site/library parameters,
# authorizes the shared connection, creates one Poll Trigger Config, and
# configures the Function App.

set -euo pipefail

GREEN='\033[0;32m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
RED='\033[0;31m'
NC='\033[0m'

echo -e "${YELLOW}Post-deployment configuration...${NC}"

# --- Resolve the retained operation -----------------------------------------
scriptDir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
operationsRoot="${scriptDir}/../operations"
descriptorFiles=()
if [[ -d "$operationsRoot" ]]; then
    while IFS= read -r -d '' descriptorFile; do
        descriptorFiles+=("$descriptorFile")
    done < <(find "$operationsRoot" -type f -name 'trigger-config.json' -print0)
fi

if [[ ${#descriptorFiles[@]} -ne 1 ]]; then
    echo -e "${RED}ERROR: expected exactly one retained operation descriptor under '${operationsRoot}', found ${#descriptorFiles[@]}.${NC}" >&2
    exit 1
fi

if ! command -v jq >/dev/null 2>&1; then
    echo -e "${RED}ERROR: jq is required for this script.${NC}" >&2
    exit 1
fi

descriptorFile="${descriptorFiles[0]}"
functionName=$(jq -er '.functionName | select(type == "string" and length > 0)' "$descriptorFile")
triggerConfigName=$(jq -er '.triggerConfigName | select(type == "string" and length > 0)' "$descriptorFile")
operationName=$(jq -er '.operationName | select(type == "string" and length > 0)' "$descriptorFile")
description=$(jq -er '.description | select(type == "string" and length > 0)' "$descriptorFile")
jq -e '.parameters | type == "array"' "$descriptorFile" >/dev/null

resolve_operation_parameters() {
    jq -ce '
        [.parameters[] |
            . as $parameter |
            if has("value") then
                {name: $parameter.name, value: $parameter.value}
            elif has("valueFromEnvironment") then
                (env[$parameter.valueFromEnvironment] // "") as $value |
                if $value == "" then
                    if ($parameter.required // false) then
                        error("Required operation parameter \($parameter.name) has no value in \($parameter.valueFromEnvironment).")
                    else
                        empty
                    end
                else
                    {name: $parameter.name, value: $value}
                end
            else
                error("Operation parameter \($parameter.name) must define value or valueFromEnvironment.")
            end
        ]' "$descriptorFile"
}

# --- Read azd outputs --------------------------------------------------------
outputs=$(azd env get-values --output json)

resourceGroupName=$(echo "$outputs" | jq -er '.resourceGroupName')
connectorNamespaceName=$(echo "$outputs" | jq -er '.connectorNamespaceName')
sharepointConnectionName=$(echo "$outputs" | jq -er '.sharepointConnectionName')
functionAppName=$(echo "$outputs" | jq -er '.functionAppName')
subscriptionId=$(echo "$outputs" | jq -r '.AZURE_SUBSCRIPTION_ID // empty')
if [[ -z "$subscriptionId" ]]; then
    subscriptionId=$(az account show --query id -o tsv)
fi

apiVersion='2026-05-01-preview'
namespaceResourceId="/subscriptions/${subscriptionId}/resourceGroups/${resourceGroupName}/providers/Microsoft.Web/connectorGateways/${connectorNamespaceName}"
triggerConfigResourceId="${namespaceResourceId}/triggerConfigs/${triggerConfigName}"

# --- Resolve canonical SharePoint parameters -------------------------------
sharepointSiteAddress=$(echo "$outputs" | jq -r '.SHAREPOINT_SITE_ADDRESS // empty')
sharepointTable=$(echo "$outputs" | jq -r '.SHAREPOINT_TABLE // empty')
sharepointFolderPath=$(echo "$outputs" | jq -r '.SHAREPOINT_FOLDER_PATH // empty')
sharepointTableWasConfigured=false
if [[ -n "$sharepointTable" ]]; then
    sharepointTableWasConfigured=true
fi

graph_get() {
    local output
    if ! output=$(az rest \
        --method get \
        --url "$1" \
        --resource 'https://graph.microsoft.com' \
        --only-show-errors 2>&1); then
        return 1
    fi
    if [[ -z "$output" ]]; then
        return 1
    fi
    printf '%s' "$output"
}

select_from_list() {
    local title="$1"
    shift
    local entries=("$@")

    echo "" >&2
    echo -e "${CYAN}${title}${NC}" >&2
    local index=1
    local entry label subLabel
    for entry in "${entries[@]}"; do
        label="${entry%%|*}"
        subLabel=''
        if [[ "$entry" == *'|'* ]]; then
            subLabel="${entry#*|}"
        fi
        if [[ -n "$subLabel" ]]; then
            echo "  [${index}] ${label}  (${subLabel})" >&2
        else
            echo "  [${index}] ${label}" >&2
        fi
        index=$((index + 1))
    done

    local answer
    while true; do
        read -r -p "Enter number (1-${#entries[@]}): " answer
        if [[ "$answer" =~ ^[0-9]+$ ]] && ((answer >= 1 && answer <= ${#entries[@]})); then
            echo $((answer - 1))
            return 0
        fi
        echo -e "${YELLOW}Invalid selection.${NC}" >&2
    done
}

selectedSiteId=''
if [[ -z "$sharepointSiteAddress" ]]; then
    echo ""
    echo -e "${YELLOW}Fetching SharePoint sites accessible to your account from Microsoft Graph...${NC}"
    sitesJson=$(graph_get 'https://graph.microsoft.com/v1.0/sites?search=*&$select=id,displayName,webUrl&$top=100' || true)
    sitesCount=$(echo "$sitesJson" | jq -r '[.value[]? | select(.displayName != null and .displayName != "")] | length' 2>/dev/null || echo 0)

    if [[ "$sitesCount" -gt 0 ]]; then
        siteIds=()
        while IFS= read -r value; do siteIds+=("$value"); done < <(
            echo "$sitesJson" | jq -r '[.value[] | select(.displayName != null and .displayName != "")] | sort_by(.displayName) | .[].id'
        )
        siteNames=()
        while IFS= read -r value; do siteNames+=("$value"); done < <(
            echo "$sitesJson" | jq -r '[.value[] | select(.displayName != null and .displayName != "")] | sort_by(.displayName) | .[].displayName'
        )
        siteUrls=()
        while IFS= read -r value; do siteUrls+=("$value"); done < <(
            echo "$sitesJson" | jq -r '[.value[] | select(.displayName != null and .displayName != "")] | sort_by(.displayName) | .[].webUrl'
        )

        siteEntries=()
        for ((index = 0; index < ${#siteNames[@]}; index++)); do
            siteEntries+=("${siteNames[$index]}|${siteUrls[$index]}")
        done
        selectedIndex=$(select_from_list 'Select a SharePoint site:' "${siteEntries[@]}")
        selectedSiteId="${siteIds[$selectedIndex]}"
        sharepointSiteAddress="${siteUrls[$selectedIndex]}"
    else
        read -r -p "SharePoint site URL (for example, https://contoso.sharepoint.com/sites/MySite): " sharepointSiteAddress
    fi

    if [[ -z "$sharepointSiteAddress" ]]; then
        echo -e "${RED}ERROR: SHAREPOINT_SITE_ADDRESS is required.${NC}" >&2
        exit 1
    fi
    azd env set SHAREPOINT_SITE_ADDRESS "$sharepointSiteAddress" >/dev/null
fi

if [[ -z "$selectedSiteId" ]]; then
    siteHost=$(node -e 'process.stdout.write(new URL(process.argv[1]).host)' "$sharepointSiteAddress" 2>/dev/null || true)
    sitePath=$(node -e 'process.stdout.write(new URL(process.argv[1]).pathname.replace(/\/$/, ""))' "$sharepointSiteAddress" 2>/dev/null || true)
    if [[ -n "$siteHost" ]]; then
        resolvedSiteJson=$(graph_get "https://graph.microsoft.com/v1.0/sites/${siteHost}:${sitePath}?\$select=id" || true)
        selectedSiteId=$(echo "$resolvedSiteJson" | jq -r '.id // empty' 2>/dev/null || true)
    fi
fi

selectedListId=''
if [[ -z "$sharepointTable" ]]; then
    if [[ -n "$selectedSiteId" ]]; then
        listsJson=$(graph_get "https://graph.microsoft.com/v1.0/sites/${selectedSiteId}/lists?\$select=id,displayName,list,webUrl&\$top=200" || true)
        librariesCount=$(echo "$listsJson" | jq -r '[.value[]? | select(.list.template == "documentLibrary")] | length' 2>/dev/null || echo 0)
        if [[ "$librariesCount" -gt 0 ]]; then
            libraryIds=()
            while IFS= read -r value; do libraryIds+=("$value"); done < <(
                echo "$listsJson" | jq -r '[.value[] | select(.list.template == "documentLibrary")] | sort_by(.displayName) | .[].id'
            )
            libraryNames=()
            while IFS= read -r value; do libraryNames+=("$value"); done < <(
                echo "$listsJson" | jq -r '[.value[] | select(.list.template == "documentLibrary")] | sort_by(.displayName) | .[].displayName'
            )
            libraryUrls=()
            while IFS= read -r value; do libraryUrls+=("$value"); done < <(
                echo "$listsJson" | jq -r '[.value[] | select(.list.template == "documentLibrary")] | sort_by(.displayName) | .[].webUrl'
            )

            libraryEntries=()
            for ((index = 0; index < ${#libraryNames[@]}; index++)); do
                libraryEntries+=("${libraryNames[$index]}|${libraryUrls[$index]}")
            done
            selectedIndex=$(select_from_list 'Select a SharePoint document library:' "${libraryEntries[@]}")
            selectedListId="${libraryIds[$selectedIndex]}"
            sharepointTable="$selectedListId"
        fi
    fi

    if [[ -z "$sharepointTable" ]]; then
        read -r -p "SharePoint document library ID or name: " sharepointTable
    fi
    if [[ -z "$sharepointTable" ]]; then
        echo -e "${RED}ERROR: SHAREPOINT_TABLE is required.${NC}" >&2
        exit 1
    fi
    azd env set SHAREPOINT_TABLE "$sharepointTable" >/dev/null
fi

if [[ -z "$sharepointFolderPath" && "$sharepointTableWasConfigured" == false ]]; then
    listIdForFolders="${selectedListId:-$sharepointTable}"
    folderEntries=()
    folderValues=()
    if [[ -n "$selectedSiteId" && -n "$listIdForFolders" ]]; then
        childrenJson=$(graph_get "https://graph.microsoft.com/v1.0/sites/${selectedSiteId}/lists/${listIdForFolders}/drive/root/children?\$select=name,folder&\$top=200" || true)
        folderEntries+=('(whole library)')
        folderValues+=('')
        while IFS= read -r childName; do
            folderEntries+=("$childName")
            folderValues+=("/$childName")
        done < <(echo "$childrenJson" | jq -r '[.value[]? | select(.folder != null)] | sort_by(.name) | .[].name' 2>/dev/null || true)
    fi

    if [[ ${#folderEntries[@]} -gt 1 ]]; then
        selectedIndex=$(select_from_list 'Select a folder, or use the whole library:' "${folderEntries[@]}")
        sharepointFolderPath="${folderValues[$selectedIndex]}"
    else
        read -r -p "Optional folder path (press Enter for the whole library): " sharepointFolderPath
    fi

    if [[ -n "$sharepointFolderPath" ]]; then
        azd env set SHAREPOINT_FOLDER_PATH "$sharepointFolderPath" >/dev/null
    fi
fi

export SHAREPOINT_SITE_ADDRESS="$sharepointSiteAddress"
export SHAREPOINT_TABLE="$sharepointTable"
export SHAREPOINT_FOLDER_PATH="$sharepointFolderPath"

# --- Verify the connector-namespace az CLI extension is installed ----------
extInstalled=$(az extension show --name connector-namespace --query name -o tsv 2>/dev/null || true)
if [[ -z "$extInstalled" ]]; then
    echo -e "${RED}ERROR: 'connector-namespace' Azure CLI extension is not installed.${NC}" >&2
    echo -e "${RED}       Download the latest connector_namespace wheel from:${NC}" >&2
    echo -e "${RED}       https://github.com/Azure/Connectors/releases${NC}" >&2
    exit 2
fi

# --- Authorize the SharePoint connection (OAuth consent) --------------------
connection_exists() {
    az connector-namespace connection show \
        -g "$resourceGroupName" \
        --namespace "$connectorNamespaceName" \
        -n "$1" \
        -o none 2>/dev/null
}

add_connection_access_policy() {
    local connectionName="$1"
    local policyName="$2"
    local principalObjectId="$3"
    local tenantId="$4"
    local required="$5"

    if ! az connector-namespace connection access-policy create \
        -g "$resourceGroupName" \
        --namespace "$connectorNamespaceName" \
        --connection-name "$connectionName" \
        -n "$policyName" \
        --principal-type 'ActiveDirectory' \
        --principal-object-id "$principalObjectId" \
        --principal-tenant-id "$tenantId" \
        -o none 2>/dev/null; then
        if [[ "$required" == 'true' ]]; then
            echo -e "${RED}ERROR: failed to create required access policy '${policyName}'.${NC}" >&2
            return 1
        fi
        echo -e "${YELLOW}WARNING: could not create optional access policy '${policyName}'.${NC}" >&2
    fi
}

create_sharepoint_connection() {
    local connectionName="$1"
    echo -e "${YELLOW}-> Creating SharePoint connection '${connectionName}' on namespace '${connectorNamespaceName}'...${NC}"

    az connector-namespace connection create \
        -g "$resourceGroupName" \
        --namespace "$connectorNamespaceName" \
        -n "$connectionName" \
        --connector-name 'sharepointonline' \
        -o none

    local tenantId functionAppPrincipalId namespacePrincipalId userPrincipalId
    tenantId=$(az account show --query tenantId -o tsv)
    functionAppPrincipalId=$(az functionapp identity show \
        -g "$resourceGroupName" \
        -n "$functionAppName" \
        --query principalId \
        -o tsv)
    namespacePrincipalId=$(az resource show \
        --ids "$namespaceResourceId" \
        --api-version "$apiVersion" \
        --query identity.principalId \
        -o tsv)

    if [[ -z "$functionAppPrincipalId" || -z "$namespacePrincipalId" ]]; then
        echo -e "${RED}ERROR: required managed identity principal IDs were not available.${NC}" >&2
        return 1
    fi

    add_connection_access_policy "$connectionName" 'functionapp-msi' "$functionAppPrincipalId" "$tenantId" true
    add_connection_access_policy "$connectionName" 'namespace-trigger-runtime' "$namespacePrincipalId" "$tenantId" true

    userPrincipalId=$(az ad signed-in-user show --query id -o tsv 2>/dev/null || true)
    if [[ -n "$userPrincipalId" ]]; then
        add_connection_access_policy "$connectionName" 'dev-user' "$userPrincipalId" "$tenantId" false
    fi
}

open_url() {
    local url="$1"
    if command -v xdg-open >/dev/null 2>&1; then
        xdg-open "$url" >/dev/null 2>&1
    elif command -v open >/dev/null 2>&1; then
        open "$url" >/dev/null 2>&1
    elif command -v wslview >/dev/null 2>&1; then
        wslview "$url" >/dev/null 2>&1
    else
        return 1
    fi
}

authorize_connection() {
    local connectionName="$1"
    echo -e "${CYAN}-> Authorizing SharePoint Online connection: ${connectionName}${NC}"

    if ! connection_exists "$connectionName"; then
        create_sharepoint_connection "$connectionName"
    fi

    local currentStatus
    currentStatus=$(az connector-namespace connection show \
        -g "$resourceGroupName" \
        --namespace "$connectorNamespaceName" \
        -n "$connectionName" \
        --query 'properties.overallStatus' \
        -o tsv 2>/dev/null || true)
    if [[ "$(printf '%s' "$currentStatus" | tr '[:upper:]' '[:lower:]')" == 'connected' ]]; then
        echo -e "${GREEN}   already Connected; skipping consent flow${NC}"
        return 0
    fi

    local parameters consentJson link
    parameters='[{"parameterName":"token","redirectUrl":"https://portal.azure.com"}]'
    consentJson=''
    for attempt in 1 2 3 4 5; do
        consentJson=$(az connector-namespace connection list-consent-links \
            -g "$resourceGroupName" \
            --namespace "$connectorNamespaceName" \
            --connection-name "$connectionName" \
            --parameters "$parameters" \
            -o json 2>/dev/null || true)
        if [[ -n "$consentJson" ]]; then
            break
        fi
        echo -e "${YELLOW}   list-consent-links attempt ${attempt} failed; retrying in 5s...${NC}"
        sleep 5
    done

    link=$(echo "$consentJson" | jq -r '.value[0].link // .link // empty')
    if [[ -z "$link" ]]; then
        return 1
    fi

    echo -e "${CYAN}   opening browser for OAuth consent...${NC}"
    if ! open_url "$link"; then
        echo -e "${RED}ERROR: no supported browser opener was found for the OAuth consent flow.${NC}" >&2
        return 1
    fi

    local deadline lastStatus status
    deadline=$(($(date +%s) + 300))
    lastStatus=''
    while [[ $(date +%s) -lt $deadline ]]; do
        status=$(az connector-namespace connection show \
            -g "$resourceGroupName" \
            --namespace "$connectorNamespaceName" \
            -n "$connectionName" \
            --query 'properties.overallStatus' \
            -o tsv 2>/dev/null || true)
        if [[ "$status" != "$lastStatus" ]]; then
            echo -e "${CYAN}   status: ${status:-?}${NC}"
            lastStatus="$status"
        fi
        if [[ "$(printf '%s' "$status" | tr '[:upper:]' '[:lower:]')" == 'connected' ]]; then
            echo -e "${GREEN}   connection authenticated${NC}"
            return 0
        fi
        sleep 3
    done

    return 1
}

echo ""
echo -e "${YELLOW}Authorizing SharePoint Online connection via Azure CLI...${NC}"
if ! authorize_connection "$sharepointConnectionName"; then
    echo -e "${RED}ERROR: SharePoint connection is not Connected. Complete OAuth consent and rerun the postdeploy hook.${NC}" >&2
    exit 1
fi

# --- Create the retained Poll Trigger Config --------------------------------
parameters=$(resolve_operation_parameters)
requestBody=$(jq -nc \
    --arg connectionName "$sharepointConnectionName" \
    --arg description "$description" \
    --arg operationName "$operationName" \
    --argjson parameters "$parameters" \
    '{
        properties: {
            connectionDetails: {
                connectionName: $connectionName,
                connectorName: "sharepointonline"
            },
            deliveryMode: "Poll",
            description: $description,
            operationName: $operationName,
            parameters: $parameters,
            state: "Enabled",
            type: "NotSpecified"
        }
    }')

triggerConfigJson=$(az rest \
    --method put \
    --uri "${triggerConfigResourceId}?api-version=${apiVersion}" \
    --body "$requestBody" \
    --only-show-errors \
    -o json)

if [[ "$(echo "$triggerConfigJson" | jq -r '.properties.pollingEndpoints.receiveUri // empty')" == '' ]]; then
    triggerConfigJson=$(az rest \
        --method get \
        --uri "${triggerConfigResourceId}?api-version=${apiVersion}" \
        --only-show-errors \
        -o json)
fi

state=$(echo "$triggerConfigJson" | jq -r '.properties.state // empty')
receiveUri=$(echo "$triggerConfigJson" | jq -r '.properties.pollingEndpoints.receiveUri // empty')
if [[ "$state" != 'Enabled' || -z "$receiveUri" ]]; then
    echo -e "${RED}ERROR: the Poll Trigger Config was not enabled with a receive endpoint.${NC}" >&2
    exit 1
fi

pollingEndpoint=$(node -e '
    const receive = new URL(process.argv[1]);
    const match = receive.pathname.match(/^(\/api(?:\/connectorGateways\/[^/]+)?)\/triggerConfigs\/[^/]+\/receive$/i);
    if (!match) {
        throw new Error("Unsupported Connector Namespace Poll receive path.");
    }
    receive.pathname = match[1].toLowerCase() === "/api" ? "/" : match[1];
    receive.search = "";
    receive.hash = "";
    process.stdout.write(receive.toString());
' "$receiveUri")

az functionapp config appsettings set \
    -g "$resourceGroupName" \
    -n "$functionAppName" \
    --settings \
        "ConnectorNamespace__pollingEndpoint=${pollingEndpoint}" \
        'ConnectorNamespace__credential=managedidentity' \
        "ConnectorTriggerConfigName=${triggerConfigName}" \
    -o none

echo ""
echo -e "${GREEN}Poll Trigger Config '${triggerConfigName}' is Enabled for ${functionName}.${NC}"
echo -e "${GREEN}Post-deployment configuration complete.${NC}"
