# Sample smoke tests

Run `npm install` followed by `npm test` from any sample app directory. Each app's test command builds its TypeScript source and runs the shared suite against that app's own dependencies. Office 365 and SharePoint also run the ownership suite. After installing and building all six apps, run the following from the repository root:

```sh
node --test tests/samples.test.mjs tests/operation-ownership.test.mjs
```

The suite checks:

- The installed `@azure/connectors` version matches the exact version requested by the app, plus CommonJS and ESM SDK entry-point loading.
- All 19 expected sample functions register through the actual connector extension.
- Office 365 and SharePoint wrappers forward explicit Poll delivery, connection, and Trigger Config settings through the real connector-specific wrappers.
- Eight payload scenarios per handler: object and JSON batches, object and JSON single items, empty batches, missing bodies, null payloads, and malformed JSON.
- Items reach each handler, original payloads are preserved, and malformed JSON produces a warning.
- Every Office 365 and SharePoint operation has the exact source/descriptor ownership mapping expected by project templates.
- Materializing each operation removes all sibling source/config artifacts while preserving shared deployment, authentication, infrastructure, and project files byte-for-byte.
- Shared deployment scripts contain no sibling operation catalog, require exactly one retained descriptor, create an enabled Poll Trigger Config through ARM `2026-05-01-preview`, and configure the derived Function settings.

Only Azure Functions host registration is intercepted. The connector extension's normalization and typed-context wrappers run unchanged. These offline tests do not load local settings, need credentials, start a Functions host, call connector services, or provision resources. They do not replace end-to-end validation with authorized connections and live events.

If package installation fails, tests may still run against previously installed dependencies. The version check intentionally fails in that situation so a passing handler test cannot be mistaken for validation of the requested SDK upgrade. Refresh and commit each app's lockfile only after installation succeeds.

## Local SDK archive fallback

If the configured registry cannot resolve the requested release, a trusted local SDK archive can be used for validation. From each app directory, run `npm install <absolute-path-to-sdk-archive> --no-save --package-lock=false --ignore-scripts`, followed by `npm test`.

This installs the archive without saving a machine-specific dependency path or changing the lockfile. The suite still checks the installed SDK version against the app's requested version. Registry-based clean installs and lockfile regeneration remain dependent on the registry making that release available; a successful archive-based test does not resolve that feed issue.

The connector extension currently declares its own `@azure/connectors@^0.2.0-preview` dependency. npm therefore installs that version separately beneath the extension alongside the app's direct `0.3.0` SDK dependency. The local fallback does not override the extension's declared dependency range.