// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import { cp, mkdir, readdir, readFile, rm, stat } from 'node:fs/promises';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

const root = fileURLToPath(new URL('../', import.meta.url));
const sharepointParameters = [
    { name: 'dataset', valueFromEnvironment: 'SHAREPOINT_SITE_ADDRESS', required: true },
    { name: 'table', valueFromEnvironment: 'SHAREPOINT_TABLE', required: true },
    { name: 'folderPath', valueFromEnvironment: 'SHAREPOINT_FOLDER_PATH', required: false },
];
const catalogs = {
    office365App: [
        {
            operationId: 'OnNewEmailV3',
            functionName: 'OnNewEmail',
            sourcePath: 'src/functions/onNewEmail.ts',
            registration: 'connectors.office365.onNewEmail',
            parameters: [{ name: 'folderPath', value: 'Inbox' }],
        },
        {
            operationId: 'OnFlaggedEmailV4',
            functionName: 'OnFlaggedEmail',
            sourcePath: 'src/functions/onFlaggedEmail.ts',
            registration: 'connectors.office365.onFlaggedEmail',
            parameters: [{ name: 'folderPath', value: 'Inbox' }],
        },
        {
            operationId: 'OnNewMentionMeEmailV3',
            functionName: 'OnNewMentionMeEmail',
            sourcePath: 'src/functions/onNewMentionMeEmail.ts',
            registration: 'connectors.office365.onNewMentionMeEmail',
            parameters: [{ name: 'folderPath', value: 'Inbox' }],
        },
        {
            operationId: 'CalendarGetOnNewItemsV3',
            functionName: 'OnNewCalendarEvent',
            sourcePath: 'src/functions/onNewCalendarEvent.ts',
            registration: 'connectors.office365.onNewCalendarEvent',
            parameters: [{ name: 'table', value: 'Calendar' }],
        },
        {
            operationId: 'OnUpcomingEventsV3',
            functionName: 'OnUpcomingEvent',
            sourcePath: 'src/functions/onUpcomingEvent.ts',
            registration: 'connectors.office365.onUpcomingEvent',
            parameters: [
                { name: 'table', value: 'Calendar' },
                { name: 'lookAheadTimeInMinutes', value: 15 },
            ],
        },
    ],
    sharepointApp: [
        {
            operationId: 'GetOnNewFileItems',
            functionName: 'OnSharepointNewFile',
            sourcePath: 'src/functions/onSharepointNewFile.ts',
            registration: 'connectors.sharepoint.onNewFile',
            parameters: sharepointParameters,
        },
        {
            operationId: 'GetOnUpdatedFileItems',
            functionName: 'OnSharepointUpdatedFile',
            sourcePath: 'src/functions/onSharepointUpdatedFile.ts',
            registration: 'connectors.sharepoint.onUpdatedFile',
            parameters: sharepointParameters,
        },
    ],
};
async function exists(filePath) {
    try {
        await stat(filePath);
        return true;
    } catch (error) {
        if (error.code === 'ENOENT') {
            return false;
        }
        throw error;
    }
}

async function findDescriptors(directory) {
    const descriptors = [];
    for (const entry of await readdir(directory, { withFileTypes: true })) {
        const entryPath = path.join(directory, entry.name);
        if (entry.isDirectory()) {
            descriptors.push(...await findDescriptors(entryPath));
        } else if (entry.name === 'trigger-config.json') {
            descriptors.push(entryPath);
        }
    }
    return descriptors;
}

async function listFiles(directory, relativeDirectory = '') {
    const files = [];
    for (const entry of await readdir(path.join(directory, relativeDirectory), { withFileTypes: true })) {
        if (entry.isDirectory() && ['.azure', 'dist', 'node_modules'].includes(entry.name)) {
            continue;
        }
        const relativePath = path.join(relativeDirectory, entry.name);
        if (entry.isDirectory()) {
            files.push(...await listFiles(directory, relativePath));
        } else {
            files.push(relativePath);
        }
    }
    return files;
}

const cwd = path.resolve(process.cwd());
const selected = Object.entries(catalogs).filter(([appName]) => cwd === path.resolve(root) || cwd === path.resolve(root, appName));
assert.ok(selected.length, 'Run tests from the repository root or an ownership-enabled sample app directory.');

for (const [appName, operations] of selected) {
    await test(`${appName} operation ownership`, async (t) => {
        const appRoot = path.join(root, appName);
        const ownedRoots = operations.flatMap(({ operationId, sourcePath }) => [
            sourcePath,
            `infra/operations/${operationId}`,
        ]);
        assert.equal(new Set(ownedRoots.map((ownedPath) => ownedPath.toLowerCase())).size, ownedRoots.length);
        for (const ownedPath of ownedRoots) {
            assert.equal(ownedPath, ownedPath.split(path.sep).join('/'));
            assert.equal(path.posix.isAbsolute(ownedPath), false);
            assert.ok(ownedPath.split('/').every((segment) => segment && segment !== '.' && segment !== '..'));
        }
        const sharedFiles = (await listFiles(appRoot)).filter((file) => {
            const posixPath = file.split(path.sep).join('/');
            return !ownedRoots.some((ownedRoot) => posixPath === ownedRoot || posixPath.startsWith(`${ownedRoot}/`));
        });
        const expectedOperationIds = operations.map(({ operationId }) => operationId).sort();
        const actualOperationIds = (await readdir(path.join(appRoot, 'infra/operations'), { withFileTypes: true }))
            .filter((entry) => entry.isDirectory())
            .map((entry) => entry.name)
            .sort();
        assert.deepEqual(actualOperationIds, expectedOperationIds);

        const expectedSourceFiles = operations.map(({ sourcePath }) => path.basename(sourcePath)).sort();
        const actualSourceFiles = (await readdir(path.join(appRoot, 'src/functions')))
            .filter((file) => file.endsWith('.ts'))
            .sort();
        assert.deepEqual(actualSourceFiles, expectedSourceFiles);

        for (const operation of operations) {
            await t.test(`${operation.operationId} owns its source and descriptor`, async () => {
                const operationPath = `infra/operations/${operation.operationId}`;
                const ownedPaths = [operation.sourcePath, operationPath];
                assert.deepEqual(ownedPaths, [
                    operation.sourcePath,
                    `infra/operations/${operation.operationId}`,
                ]);

                const operationFiles = await readdir(path.join(appRoot, operationPath));
                assert.deepEqual(operationFiles, ['trigger-config.json']);

                const descriptor = JSON.parse(await readFile(path.join(appRoot, operationPath, 'trigger-config.json'), 'utf8'));
                assert.equal(descriptor.functionName, operation.functionName);
                assert.equal(descriptor.operationName, operation.operationId);
                assert.equal(descriptor.triggerConfigName, operation.operationId);
                assert.equal(typeof descriptor.description, 'string');
                assert.ok(descriptor.description.length > 0);
                assert.deepEqual(descriptor.parameters, operation.parameters);

                const source = await readFile(path.join(appRoot, operation.sourcePath), 'utf8');
                const optionsName = `${operation.functionName[0].toLowerCase()}${operation.functionName.slice(1)}Options`;
                assert.match(source, new RegExp(`const ${optionsName} = \\{`));
                assert.match(
                    source,
                    new RegExp(`${operation.registration.replaceAll('.', '\\.')}\\('${operation.functionName}', ${optionsName}\\);`),
                );
                assert.match(source, /deliveryMode: 'Poll'/);
                assert.match(source, /connection: 'ConnectorNamespace'/);
                assert.match(source, /triggerConfigName: '%ConnectorTriggerConfigName%'/);
                assert.match(
                    source,
                    new RegExp(`satisfies Parameters<typeof ${operation.registration.replaceAll('.', '\\.')}>\\[1\\] & PollConnectorTriggerOptions`),
                );
            });
        }

        const sharedInfra = await Promise.all([
            readFile(path.join(appRoot, 'infra/connectorNamespace.bicep'), 'utf8'),
            readFile(path.join(appRoot, 'infra/main.bicep'), 'utf8'),
            readFile(path.join(appRoot, 'infra/resources.bicep'), 'utf8'),
            readFile(path.join(appRoot, 'infra/scripts/postdeploy.ps1'), 'utf8'),
            readFile(path.join(appRoot, 'infra/scripts/postdeploy.sh'), 'utf8'),
        ]);
        for (const { operationId, functionName } of operations) {
            for (const contents of sharedInfra) {
                assert.ok(!contents.includes(operationId), `Shared infrastructure must not name ${operationId}.`);
                assert.ok(!contents.includes(functionName), `Shared infrastructure must not name ${functionName}.`);
            }
        }

        const [powerShellScript, shellScript] = await Promise.all([
            readFile(path.join(appRoot, 'infra/scripts/postdeploy.ps1'), 'utf8'),
            readFile(path.join(appRoot, 'infra/scripts/postdeploy.sh'), 'utf8'),
        ]);
        assert.match(powerShellScript, /\$descriptorFiles\.Count -ne 1/);
        assert.match(shellScript, /\$\{#descriptorFiles\[@\]\} -ne 1/);
        for (const script of [powerShellScript, shellScript]) {
            assert.match(script, /2026-05-01-preview/);
            assert.match(script, /\/triggerConfigs\//);
            assert.match(script, /az rest/);
            assert.match(script, /deliveryMode/);
            assert.match(script, /['"]Poll['"]/);
            assert.match(script, /['"]Enabled['"]/);
            assert.match(script, /['"]NotSpecified['"]/);
            assert.match(script, /pollingEndpoints\.receiveUri/);
            assert.match(script, /ConnectorNamespace__pollingEndpoint/);
            assert.match(script, /ConnectorNamespace__credential=managedidentity/);
            assert.match(script, /ConnectorTriggerConfigName/);
            assert.doesNotMatch(script, /notificationDetails|runtime\/webhooks\/connector|connector_extension|connector-namespace trigger create/);
            assert.equal((script.match(/--method put/g) ?? []).length, 1);
        }

        const connectorNamespace = await readFile(path.join(appRoot, 'infra/connectorNamespace.bicep'), 'utf8');
        for (const accessPolicyName of ['functionapp-msi', 'namespace-trigger-runtime', 'dev-user']) {
            assert.match(connectorNamespace, new RegExp(accessPolicyName));
        }
        const localSettings = JSON.parse(await readFile(path.join(appRoot, 'local.settings.json'), 'utf8'));
        assert.deepEqual(
            Object.keys(localSettings.Values).sort(),
            ['AzureWebJobsStorage', 'ConnectorNamespace__pollingEndpoint', 'ConnectorTriggerConfigName', 'FUNCTIONS_WORKER_RUNTIME'].sort(),
        );

        const scratchRoot = path.join(root, 'tests', `.operation-ownership-${process.pid}-${randomUUID()}`);
        await mkdir(scratchRoot, { recursive: true });
        try {
            for (const selectedOperation of operations) {
                await t.test(`${selectedOperation.operationId} materializes without sibling artifacts`, async () => {
                    const generatedRoot = path.join(scratchRoot, selectedOperation.operationId);
                    await cp(appRoot, generatedRoot, {
                        recursive: true,
                        filter: (source) => {
                            const relative = path.relative(appRoot, source);
                            return !relative.split(path.sep).some((segment) => ['.azure', 'dist', 'node_modules'].includes(segment));
                        },
                    });

                    for (const sibling of operations) {
                        if (sibling !== selectedOperation) {
                            await rm(path.join(generatedRoot, sibling.sourcePath), { force: true });
                            await rm(path.join(generatedRoot, 'infra/operations', sibling.operationId), { recursive: true, force: true });
                        }
                    }

                    assert.ok(await exists(path.join(generatedRoot, selectedOperation.sourcePath)));
                    assert.ok(await exists(path.join(generatedRoot, 'infra/operations', selectedOperation.operationId)));
                    for (const sibling of operations) {
                        if (sibling !== selectedOperation) {
                            assert.equal(await exists(path.join(generatedRoot, sibling.sourcePath)), false);
                            assert.equal(await exists(path.join(generatedRoot, 'infra/operations', sibling.operationId)), false);
                        }
                    }

                    const generatedSources = (await readdir(path.join(generatedRoot, 'src/functions')))
                        .filter((file) => file.endsWith('.ts'));
                    assert.deepEqual(generatedSources, [path.basename(selectedOperation.sourcePath)]);

                    const descriptors = await findDescriptors(path.join(generatedRoot, 'infra/operations'));
                    assert.equal(descriptors.length, 1);
                    const descriptor = JSON.parse(await readFile(descriptors[0], 'utf8'));
                    assert.equal(descriptor.operationName, selectedOperation.operationId);
                    assert.equal(descriptor.functionName, selectedOperation.functionName);

                    for (const sharedPath of sharedFiles) {
                        const [original, generated] = await Promise.all([
                            readFile(path.join(appRoot, sharedPath)),
                            readFile(path.join(generatedRoot, sharedPath)),
                        ]);
                        assert.deepEqual(generated, original, `${sharedPath} must be preserved exactly.`);
                    }
                });
            }
        } finally {
            await rm(scratchRoot, { recursive: true, force: true });
        }
    });
}
