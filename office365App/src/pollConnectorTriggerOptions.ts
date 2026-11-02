// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License.

export type PollConnectorTriggerOptions = {
    deliveryMode: 'Poll';
    connection: 'ConnectorNamespace';
    triggerConfigName: '%ConnectorTriggerConfigName%';
};
