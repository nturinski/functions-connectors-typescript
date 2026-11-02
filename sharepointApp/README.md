# SharePoint Online Connector Sample

Azure Functions app demonstrating typed SharePoint Online triggers from [`@azure/functions-extensions-connectors`](https://www.npmjs.com/package/@azure/functions-extensions-connectors). Handlers are in [src/functions](src/functions/).

| Function | Connector operation | Description |
| -------- | ------------------- | ----------- |
| `OnSharepointNewFile` | `GetOnNewFileItems` | New file is created in the configured library or folder |
| `OnSharepointUpdatedFile` | `GetOnUpdatedFileItems` | A file is created or modified in the configured library or folder |

Each operation owns its listed source file and the matching
`infra/operations/<operationId>` directory. A generated project retains one
pair. If you deploy this repository checkout directly, first remove the sibling
source and operation directory; the postdeploy hook intentionally fails unless
exactly one operation descriptor remains.

## Prerequisites

Follow the [shared prerequisites](../README.md#prerequisites). You also need a SharePoint Online account with access to the site and document library, and permission to consent to the connector.

## Deploy to Azure

Run from `sharepointApp`:

```sh
az login
azd auth login
azd up
```

The deployment builds the TypeScript app and runs the [postdeploy hook](azure.yaml). Select the SharePoint site, document library, and optional folder when prompted, then complete OAuth consent in your browser. The hook creates the one retained Trigger Config in Poll mode, derives its polling endpoint, and configures the Function App.

## Resources provisioned

- Flex Consumption Function App with a system-assigned managed identity.
- Storage account, Application Insights, and Log Analytics workspace.
- Connector Namespace and SharePoint Online connection.

The Preview Functions Extension Bundle is already configured in [host.json](host.json).

## Connector configuration

The descriptors use the canonical SharePoint trigger parameters: `dataset`,
`table`, and optional `folderPath`. The hook saves those values as
`SHAREPOINT_SITE_ADDRESS`, `SHAREPOINT_TABLE`, and
`SHAREPOINT_FOLDER_PATH` on the azd environment. To skip the prompts, set them
in advance:

```sh
azd env set SHAREPOINT_SITE_ADDRESS "https://contoso.sharepoint.com/sites/MySite"
azd env set SHAREPOINT_TABLE "<document-library-id-or-name>"
azd env set SHAREPOINT_FOLDER_PATH "/Optional/Subfolder"
```

These settings do not skip any required OAuth consent. To repeat connection and trigger setup without redeploying:

```sh
azd hooks run postdeploy
```

## Verify

Open [Connector Namespaces](https://connectors.azure.com/) and select the namespace created by azd. Expect one SharePoint Online connection in `Connected` state and one retained Trigger Config in `Enabled` state.

Create or modify a file in the selected location, then inspect the Function App logs in Application Insights. If the connection is unauthenticated or triggers are missing, rerun the hook and complete consent.

## Run locally

Configure [local.settings.json](local.settings.json) for your environment and start Azurite if using the default development storage setting:

```sh
npm install
npm start
```

This builds the app and starts the Functions host. For local Poll delivery,
replace the placeholders in `local.settings.json` with the gateway-level
polling endpoint and retained Trigger Config name. Local development uses the
standard Functions developer credential chain.

## More

- [SharePoint Online connector operations](https://learn.microsoft.com/connectors/sharepointonline/)
- [All TypeScript samples](../README.md#samples)
