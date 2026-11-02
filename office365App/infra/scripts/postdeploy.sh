#!/bin/bash
# Post-deployment configuration for the office365App.
#
# The project template retains one operation descriptor under infra/operations.
# This script authorizes the shared Office 365 connection, creates that one
# Connector Namespace Poll Trigger Config, and configures the Function App.

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
office365ConnectionName=$(echo "$outputs" | jq -er '.office365ConnectionName')
functionAppName=$(echo "$outputs" | jq -er '.functionAppName')
subscriptionId=$(echo "$outputs" | jq -r '.AZURE_SUBSCRIPTION_ID // empty')
if [[ -z "$subscriptionId" ]]; then
    subscriptionId=$(az account show --query id -o tsv)
fi

apiVersion='2026-05-01-preview'
namespaceResourceId="/subscriptions/${subscriptionId}/resourceGroups/${resourceGroupName}/providers/Microsoft.Web/connectorGateways/${connectorNamespaceName}"
triggerConfigResourceId="${namespaceResourceId}/triggerConfigs/${triggerConfigName}"

# --- Verify the connector-namespace az CLI extension is installed ----------
extInstalled=$(az extension show --name connector-namespace --query name -o tsv 2>/dev/null || true)
if [[ -z "$extInstalled" ]]; then
    echo -e "${RED}ERROR: 'connector-namespace' Azure CLI extension is not installed.${NC}" >&2
    echo -e "${RED}       Download the latest connector_namespace wheel from:${NC}" >&2
    echo -e "${RED}       https://github.com/Azure/Connectors/releases${NC}" >&2
    exit 2
fi

# --- Authorize the Office 365 connection (OAuth consent) --------------------
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

create_office365_connection() {
    local connectionName="$1"
    echo -e "${YELLOW}-> Creating Office 365 connection '${connectionName}' on namespace '${connectorNamespaceName}'...${NC}"

    az connector-namespace connection create \
        -g "$resourceGroupName" \
        --namespace "$connectorNamespaceName" \
        -n "$connectionName" \
        --connector-name 'office365' \
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
    echo -e "${CYAN}-> Authorizing Office 365 Outlook connection: ${connectionName}${NC}"

    if ! connection_exists "$connectionName"; then
        create_office365_connection "$connectionName"
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
echo -e "${YELLOW}Authorizing Office 365 connection via Azure CLI...${NC}"
if ! authorize_connection "$office365ConnectionName"; then
    echo -e "${RED}ERROR: Office 365 connection is not Connected. Complete OAuth consent and rerun the postdeploy hook.${NC}" >&2
    exit 1
fi

# --- Create the retained Poll Trigger Config --------------------------------
parameters=$(resolve_operation_parameters)
requestBody=$(jq -nc \
    --arg connectionName "$office365ConnectionName" \
    --arg description "$description" \
    --arg operationName "$operationName" \
    --argjson parameters "$parameters" \
    '{
        properties: {
            connectionDetails: {
                connectionName: $connectionName,
                connectorName: "office365"
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
