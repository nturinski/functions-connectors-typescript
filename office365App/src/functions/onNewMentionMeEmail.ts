// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License.

import { InvocationContext } from '@azure/functions';
import { connectors, EmailTriggerContext } from '@azure/functions-extensions-connectors';
import type { PollConnectorTriggerOptions } from '../pollConnectorTriggerOptions.js';

const onNewMentionMeEmailOptions = {
    deliveryMode: 'Poll',
    connection: 'ConnectorNamespace',
    triggerConfigName: '%ConnectorTriggerConfigName%',
    handler: async (context: EmailTriggerContext, invocationContext: InvocationContext) => {
        invocationContext.log('OnNewMentionMeEmail trigger received.');

        for (const email of context.emails) {
            invocationContext.log(`Subject: '${email.subject}'.`);
            invocationContext.log(`From: '${email.from}'.`);
        }

        return context.rawPayload;
    },
} satisfies Parameters<typeof connectors.office365.onNewMentionMeEmail>[1] & PollConnectorTriggerOptions;

connectors.office365.onNewMentionMeEmail('OnNewMentionMeEmail', onNewMentionMeEmailOptions);
