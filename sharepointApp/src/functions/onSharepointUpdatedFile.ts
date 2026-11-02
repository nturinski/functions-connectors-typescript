// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License.

import { InvocationContext } from '@azure/functions';
import { connectors, FileTriggerContext } from '@azure/functions-extensions-connectors';
import type { PollConnectorTriggerOptions } from '../pollConnectorTriggerOptions.js';

const onSharepointUpdatedFileOptions = {
    deliveryMode: 'Poll',
    connection: 'ConnectorNamespace',
    triggerConfigName: '%ConnectorTriggerConfigName%',
    handler: async (context: FileTriggerContext, invocationContext: InvocationContext) => {
        invocationContext.log('OnSharepointUpdatedFile trigger received.');

        for (const file of context.files) {
            invocationContext.log(`Name: '${file.Name}'.`);
            invocationContext.log(`Path: '${file.Path}'.`);
            invocationContext.log(`LastModified: '${file.LastModified}'.`);
        }

        return context.rawPayload;
    },
} satisfies Parameters<typeof connectors.sharepoint.onUpdatedFile>[1] & PollConnectorTriggerOptions;

connectors.sharepoint.onUpdatedFile('OnSharepointUpdatedFile', onSharepointUpdatedFileOptions);
