# Office 365 Outlook Connector Sample

Azure Functions app demonstrating typed Office 365 Outlook triggers from [`@azure/functions-extensions-connectors`](https://www.npmjs.com/package/@azure/functions-extensions-connectors). Handlers are in [src/functions](src/functions/).

| Function | Connector operation | Description |
| -------- | ------------------- | ----------- |
| `OnNewEmail` | `OnNewEmailV3` | New email arrives in Inbox |
| `OnFlaggedEmail` | `OnFlaggedEmailV4` | An email is flagged in Inbox |
| `OnNewMentionMeEmail` | `OnNewMentionMeEmailV3` | New email mentioning you arrives in Inbox |
| `OnNewCalendarEvent` | `CalendarGetOnNewItemsV3` | New event is created in Calendar |
| `OnUpcomingEvent` | `OnUpcomingEventsV3` | Calendar event starts within 15 minutes |

> These handlers log email and calendar metadata. Use a test account and avoid sensitive data.

Each operation owns its listed source file and the matching
`infra/operations/<operationId>` directory. A generated project retains one
pair. If you deploy this repository checkout directly, first remove every
unwanted source file and operation directory; the postdeploy hook intentionally
fails unless exactly one operation descriptor remains.

## Prerequisites

Follow the [shared prerequisites](../README.md#prerequisites). You also need an Office 365 account with access to Outlook email and calendar, and permission to consent to the connector.

## Deploy to Azure

Run from `office365App`:

```sh
az login
azd auth login
azd up
```

The deployment builds the TypeScript app and runs the [postdeploy hook](azure.yaml). Complete the OAuth consent flow in your browser when prompted. The hook waits for the connection to become `Connected`, creates the one retained trigger configuration in Poll mode, derives its polling endpoint, and configures the Function App.

## Resources provisioned

- Flex Consumption Function App with a system-assigned managed identity.
- Storage account, Application Insights, and Log Analytics workspace.
- Connector Namespace and Office 365 Outlook connection.

The Preview Functions Extension Bundle is already configured in [host.json](host.json).

## Connector configuration

The retained operation descriptor supplies its function name, canonical
operation ID, and parameters. Email descriptors use `Inbox`, calendar
descriptors use `Calendar`, and the upcoming-event descriptor uses a 15-minute
look-ahead. Shared scripts do not contain a catalog of operation names. To
repeat connection consent and trigger setup without redeploying:

```sh
azd hooks run postdeploy
```

## Verify

Open [Connector Namespaces](https://connectors.azure.com/) and select the namespace created by azd. Expect one Office 365 Outlook connection in `Connected` state and one retained Trigger Config in `Enabled` state.

Send an email to the connected account or create a calendar event, then inspect the Function App logs in Application Insights. If the connection is unauthenticated or triggers are missing, rerun the hook and complete consent.

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

- [Office 365 Outlook connector operations](https://learn.microsoft.com/connectors/office365/)
- [All TypeScript samples](../README.md#samples)
