// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License.

import { InvocationContext } from '@azure/functions';
import { CalendarEventTriggerContext, connectors } from '@azure/functions-extensions-connectors';
import type { PollConnectorTriggerOptions } from '../pollConnectorTriggerOptions.js';

const onUpcomingEventOptions = {
    deliveryMode: 'Poll',
    connection: 'ConnectorNamespace',
    triggerConfigName: '%ConnectorTriggerConfigName%',
    handler: async (context: CalendarEventTriggerContext, invocationContext: InvocationContext) => {
        invocationContext.log('OnUpcomingEvent trigger received.');

        for (const event of context.calendarEvents) {
            invocationContext.log(`Subject: '${event.subject}'.`);
            invocationContext.log(`Start: '${event.start}'.`);
        }

        return context.rawPayload;
    },
} satisfies Parameters<typeof connectors.office365.onUpcomingEvent>[1] & PollConnectorTriggerOptions;

connectors.office365.onUpcomingEvent('OnUpcomingEvent', onUpcomingEventOptions);
