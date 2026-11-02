# Azure Functions Connector Samples (TypeScript)

TypeScript samples for the [Azure Functions Connector Samples](https://github.com/Azure-Samples/functions-connectors/) repo, using [`@azure/functions-extensions-connectors`](https://www.npmjs.com/package/@azure/functions-extensions-connectors) to receive events from Microsoft 365, SharePoint, Teams, and other connectors.

> Azure Functions connectors and the packages used here are in preview.

## Samples

| Folder | Connector | Triggers |
| ------ | --------- | -------- |
| [azureblobApp](azureblobApp/) | [Azure Blob Storage](https://learn.microsoft.com/connectors/azureblob/) | Blob added or modified |
| [genericApp](genericApp/) | Any connector | Generic `connectorTrigger<TItem>` examples with custom item types |
| [office365App](office365App/) | [Office 365 Outlook](https://learn.microsoft.com/connectors/office365/) | New, flagged, and mentioning-you emails; new and upcoming calendar events |
| [onedriveApp](onedriveApp/) | [OneDrive](https://learn.microsoft.com/connectors/onedrive/) | New and updated files |
| [sharepointApp](sharepointApp/) | [SharePoint Online](https://learn.microsoft.com/connectors/sharepointonline/) | New and updated files |
| [teamsApp](teamsApp/) | [Microsoft Teams](https://learn.microsoft.com/connectors/teams/) | Channel messages, mentions, and team membership changes |

Each folder is a self-contained Functions app with its own `azure.yaml`, infrastructure, and source code. Open a sample README for deployment, connector configuration, and verification steps. The five connector-specific samples configure connections and triggers through deployment hooks; `genericApp` requires manual connector setup.

The Office 365 and SharePoint folders are also project-template catalogs. Each
operation owns one function source file and one directory under
`infra/operations`. Template materialization removes every nonselected owned
path, leaving shared authentication, deployment, and infrastructure unchanged.

## More triggers and operations

See the [Operations to Azure Functions Signature Mapping](https://github.com/Azure/azure-functions-connector-extension/blob/main/docs/operations-functions-match.md) for supported connector triggers and their TypeScript signatures. The [Azure Connectors Node.js SDK](https://github.com/Azure/Connectors-NodeJS-SDK) provides typed clients for calling connector actions. These samples pin the SDK to [`@azure/connectors@0.3.0`](https://www.npmjs.com/package/@azure/connectors/v/0.3.0).

## Tests

Run `npm test` from a sample directory to build its TypeScript source and run the offline smoke tests. Office 365 and SharePoint also run deterministic operation-ownership tests. See [tests/README.md](tests/README.md) for coverage, running the repository tests together, and the local SDK archive fallback.

## Prerequisites

- An Azure subscription with permission to provision resources and configure access policies.
- [Azure Developer CLI (`azd`)](https://learn.microsoft.com/azure/developer/azure-developer-cli/install-azd) and [Azure CLI (`az`)](https://learn.microsoft.com/cli/azure/install-azure-cli).
- [Node.js](https://nodejs.org/) 20+ and npm. The deployment templates currently select Node.js 20.
- [`connector-namespace` Azure CLI extension](https://github.com/Azure/Connectors/tree/main/public-preview/connector-namespace-cli).
- PowerShell 7+ (`pwsh`) on Windows, or Bash and `jq` on macOS/Linux, for deployment hooks.
- For local development: [Azure Functions Core Tools](https://learn.microsoft.com/azure/azure-functions/functions-run-local) v4 and [Azurite](https://learn.microsoft.com/azure/storage/common/storage-use-azurite) when using `UseDevelopmentStorage=true`.

The connector-specific templates default the Connector Namespace location to `brazilsouth`. To use another supported region, set `CONNECTOR_NAMESPACE_LOCATION` on the azd environment before provisioning.

## Related repos

- [Azure Functions Connector Extension](https://github.com/Azure/azure-functions-connector-extension) - Connector trigger binding and operation mappings.
- [Azure Functions Node.js Extensions](https://github.com/Azure/azure-functions-nodejs-extensions) - Source for `@azure/functions-extensions-connectors`.
- [Azure Connectors Node.js SDK](https://github.com/Azure/Connectors-NodeJS-SDK) - Typed clients and connector models.
- [.NET samples](https://github.com/Azure-Samples/functions-connectors-net) and [Python samples](https://github.com/Azure-Samples/functions-connectors-python).