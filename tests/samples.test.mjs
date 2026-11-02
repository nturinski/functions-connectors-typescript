// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import { readdir, readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import test from 'node:test';

const root = fileURLToPath(new URL('../', import.meta.url));
const samples = {
    azureblobApp: ['OnAzureBlobUpdatedFile'],
    genericApp: ['OnGenericAzureBlobUpdated', 'OnGenericCustomConnectorEvent', 'OnGenericOffice365NewEmail', 'OnGenericSharepointNewFile', 'OnGenericTeamsChannelMessage'],
    office365App: ['OnFlaggedEmail', 'OnNewCalendarEvent', 'OnNewEmail', 'OnNewMentionMeEmail', 'OnUpcomingEvent'],
    onedriveApp: ['OnOneDriveNewFile', 'OnOneDriveUpdatedFile'],
    sharepointApp: ['OnSharepointNewFile', 'OnSharepointUpdatedFile'],
    teamsApp: ['OnGroupMembershipAdd', 'OnGroupMembershipRemoval', 'OnNewChannelMessage', 'OnNewChannelMessageMentioningMe'],
};

function itemFor(name, index) {
    const marker = `sample-item-${index}`;
    if (/Email|Calendar|Upcoming/.test(name)) {
        return { id: marker, subject: marker, from: 'sender@example.com', importance: 'normal', hasAttachments: false, organizer: 'organizer@example.com', start: '2026-10-23T10:00:00Z', end: '2026-10-23T11:00:00Z' };
    }
    if (/Message/.test(name)) {
        return { id: marker, from: { user: { displayName: marker } }, body: { content: marker } };
    }
    return { id: marker, name: marker, Name: marker, Path: `/files/${marker}`, Size: 42, LastModified: '2026-10-23T10:00:00Z' };
}

// Running from an app via npm test selects that app. Running from the repository
// root exercises all six apps, each with its own installed dependencies.
const cwd = path.resolve(process.cwd());
const selected = Object.entries(samples).filter(([appName]) => cwd === path.resolve(root) || cwd === path.resolve(root, appName));
assert.ok(selected.length, 'Run tests from the repository root or a sample app directory.');

for (const [appName, expectedNames] of selected) {
    await test(appName, async (t) => {
        const appRoot = path.join(root, appName);
        const require = createRequire(path.join(appRoot, 'package.json'));

        await t.test('installed connector SDK matches the requested version', async () => {
            const manifest = JSON.parse(await readFile(path.join(appRoot, 'package.json'), 'utf8'));
            const sdk = JSON.parse(await readFile(path.join(appRoot, 'node_modules/@azure/connectors/package.json'), 'utf8'));
            assert.equal(sdk.version, manifest.dependencies['@azure/connectors']);
            // Check both published runtime entry points, not just TypeScript types.
            assert.equal(typeof require('@azure/connectors').ConnectorError, 'function');
            const esm = await import(pathToFileURL(path.join(appRoot, 'node_modules/@azure/connectors/dist/esm/azureConnectors/index.js')).href);
            assert.equal(typeof esm.ConnectorError, 'function');
        });

        const { app } = require('@azure/functions');
        const original = app.connectorTrigger;
        const registrations = new Map();
        // Keep the real extension's normalization and named-context wrappers.
        // Replace only host registration; no Azure service or credentials are used.
        app.connectorTrigger = (name, options) => {
            assert.ok(!registrations.has(name), `Duplicate function: ${name}`);
            registrations.set(name, options);
        };
        try {
            const functionsDir = path.join(appRoot, 'dist/functions');
            for (const file of (await readdir(functionsDir)).filter((file) => file.endsWith('.js')).sort()) {
                await import(pathToFileURL(path.join(functionsDir, file)).href);
            }
        } finally {
            app.connectorTrigger = original;
        }

        await t.test('all expected functions register', () => {
            assert.deepEqual([...registrations.keys()].sort(), [...expectedNames].sort());
        });

        if (appName === 'office365App' || appName === 'sharepointApp') {
            await t.test('all connector-specific wrappers forward Poll binding options', () => {
                for (const name of expectedNames) {
                    const options = registrations.get(name);
                    assert.equal(options.deliveryMode, 'Poll');
                    assert.equal(options.connection, 'ConnectorNamespace');
                    assert.equal(options.triggerConfigName, '%ConnectorTriggerConfigName%');
                }
            });
        }

        for (const [name, { handler }] of registrations) {
            const items = [itemFor(name, 1), itemFor(name, 2)];
            const batch = { body: { value: items }, headers: { 'x-test': 'preserved' } };
            const single = { body: items[0] };
            const cases = [
                ['batch object', batch, batch, 2],
                ['batch JSON', JSON.stringify(batch), batch, 2],
                ['single object', single, single, 1],
                ['single JSON', JSON.stringify(single), single, 1],
                ['empty batch', { body: { value: [] } }, { body: { value: [] } }, 0],
                ['missing body', {}, {}, 0],
                ['null payload', null, {}, 0],
                ['malformed JSON', '{invalid', '{invalid', 0],
            ];

            for (const [label, input, expected, count] of cases) {
                await t.test(`${name}: ${label}`, async () => {
                    const logs = [];
                    const warnings = [];
                    const result = await handler(input, {
                        log: (...args) => logs.push(args.join(' ')),
                        warn: (...args) => warnings.push(args.join(' ')),
                    });
                    assert.deepEqual(result, expected, 'The original payload must be returned unchanged.');
                    if (input && typeof input === 'object') {
                        assert.strictEqual(result, input);
                    }
                    assert.ok(logs.some((line) => line.includes(`${name} `)), 'The sample handler must run.');
                    for (let index = 1; index <= count; index++) {
                        assert.ok(logs.some((line) => line.includes(`sample-item-${index}`)), `Item ${index} must reach the handler.`);
                    }
                    assert.equal(warnings.length, label === 'malformed JSON' ? 1 : 0);
                });
            }
        }
    });
}