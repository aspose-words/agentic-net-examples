# security-and-protection Examples for Aspose.Words for .NET

This folder contains the live, publish-ready C# examples for the security-and-protection category. Each file is a standalone console example selected from the verified 26.5.0 run.

## Snapshot

- Category: security-and-protection
- Slug: security-and-protection
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20260619_131835_59df5f
- Security Workflow examples: 30

## Category rules that shaped these examples

- Use native Aspose.Words APIs directly.
- Create local sample source documents when a task refers to an existing file, folder, stream, or template.
- Do not assume external files or folders already exist.
- Prefer documented protection and password workflows using `Document.Protect`, `Document.Unprotect`, `LoadOptions`, and format-specific save options.
- Keep validation narrow and task-specific.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words 26.5.0

## Running Examples

Each file in this folder is a single, standalone `.cs` console example. To run one example:

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.5.0

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\security-and-protection\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `security-and-protection/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.5.0

# PowerShell example
Copy-Item ..\security-and-protection\create-a-cancellationtokensource-and-pass-its-token-to-document-load-to-enable-interruptio.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `create-a-cancellationtokensource-and-pass-its-token-to-document-load-to-enable-interruptio.cs` | Create a CancellationTokenSource and pass its token to Document.Load to enable interruption. | Security Workflow | doc | mcp |
| 2 | `attach-a-callback-to-documentloadingargs-throwifcancellationrequested-during-loading-to-en.cs` | Attach a callback to DocumentLoadingArgs.ThrowIfCancellationRequested during loading to ensure proper response. | Security Workflow | doc | mcp |
| 3 | `use-token-throwifcancellationrequested-inside-a-field-update-loop-to-abort-long-running-op.cs` | Use token.ThrowIfCancellationRequested inside a field update loop to abort long-running operations efficiently. | Security Workflow | docx | mcp |
| 4 | `implement-a-low-code-ui-button-that-calls-cancellationtokensource-cancel-to-stop-processin.cs` | Implement a low-code UI button that calls CancellationTokenSource.Cancel to stop processing immediately. | Security Workflow | docx | mcp |
| 5 | `combine-a-cancellationtoken-with-document-saveasync-to-allow-asynchronous-saving-cancellat.cs` | Combine a CancellationToken with Document.SaveAsync to allow asynchronous saving cancellation by checking token status. | Security Workflow | doc | mcp |
| 6 | `monitor-token-iscancellationrequested-during-layout-building-and-exit-the-layout-routine-p.cs` | Monitor token.IsCancellationRequested during layout building and exit the layout routine promptly if needed. | Security Workflow | docx | mcp |
| 7 | `handle-operationcanceledexception-after-a-cancelled-document-load-to-perform-necessary-cle.cs` | Handle OperationCanceledException after a cancelled Document.Load to perform necessary cleanup properly. | Security Workflow | doc | mcp |
| 8 | `log-cancellation-events-with-timestamps-to-an-audit-file-for-compliance-tracking.cs` | Log cancellation events with timestamps to an audit file for compliance tracking. | Security Workflow | docx | mcp |
| 9 | `pass-the-same-cancellationtoken-to-both-loading-and-saving-methods-for-consistent-interrup.cs` | Pass the same CancellationToken to both loading and saving methods for consistent interruption control. | Security Workflow | docx | mcp |
| 10 | `use-token-iscancellationrequested-in-a-custom-linq-reporting-engine-query-to-abort-large-r.cs` | Use token.IsCancellationRequested in a custom LINQ Reporting Engine query to abort large reports. | Security Workflow | docx | mcp |
| 11 | `wrap-a-batch-of-document-load-calls-in-a-foreach-loop-that-checks-token-cancellation-befor.cs` | Wrap a batch of Document.Load calls in a foreach loop that checks token cancellation before each load. | Security Workflow | doc | mcp |
| 12 | `configure-documentloadingargs-to-invoke-a-user-defined-method-when-cancellation-is-request.cs` | Configure DocumentLoadingArgs to invoke a user-defined method when cancellation is requested during loading. | Security Workflow | doc | mcp |
| 13 | `integrate-cancellationtoken-into-a-background-worker-that-processes-documents-without-bloc.cs` | Integrate CancellationToken into a background worker that processes documents without blocking the UI. | Security Workflow | doc | mcp |
| 14 | `validate-that-long-running-field-updates-respect-the-token-by-inserting-periodic-cancellat.cs` | Validate that long-running field updates respect the token by inserting periodic cancellation checks. | Security Workflow | docx | mcp |
| 15 | `create-a-reusable-helper-method-accepting-a-cancellationtoken-to-perform-safe-document-pro.cs` | Create a reusable helper method accepting a CancellationToken to perform safe document processing. | Security Workflow | doc | mcp |
| 16 | `ensure-resource-leaks-are-prevented-by-disposing-the-document-object-after-catching-operat.cs` | Ensure resource leaks are prevented by disposing the Document object after catching OperationCanceledException. | Security Workflow | doc | mcp |
| 17 | `use-token-throwifcancellationrequested-inside-a-custom-image-extraction-routine-to-stop-ea.cs` | Use token.ThrowIfCancellationRequested inside a custom image extraction routine to stop early if needed. | Security Workflow | docx | mcp |
| 18 | `combine-token-monitoring-with-progress-reporting-to-inform-users-when-cancellation-occurs.cs` | Combine token monitoring with progress reporting to inform users when cancellation occurs during processing. | Security Workflow | docx | mcp |
| 19 | `implement-a-timeout-mechanism-that-triggers-cancellationtokensource-cancel-after-a-predefi.cs` | Implement a timeout mechanism that triggers CancellationTokenSource.Cancel after a predefined duration automatically. | Security Workflow | docx | mcp |
| 20 | `pass-the-token-to-documentbuilder-operations-to-allow-interruption-while-constructing-comp.cs` | Pass the token to DocumentBuilder operations to allow interruption while constructing complex documents. | Security Workflow | doc | llm |
| 21 | `test-cancellation-behavior-by-simulating-user-aborts-during-document-layout-generation-in.cs` | Test cancellation behavior by simulating user aborts during document layout generation in unit tests. | Security Workflow | doc | mcp |
| 22 | `document-the-required-net-version-for-cancellationtoken-support-in-the-project-s-readme-fi.cs` | Document the required .NET version for CancellationToken support in the project's README file. | Security Workflow | doc | mcp |
| 23 | `use-token-iscancellationrequested-within-a-while-loop-processing-document-nodes-to-enable.cs` | Use token.IsCancellationRequested within a while loop processing document nodes to enable graceful exit. | Security Workflow | doc | mcp |
| 24 | `provide-a-configuration-setting-that-enables-or-disables-cancellation-support-for-specific.cs` | Provide a configuration setting that enables or disables cancellation support for specific processing stages. | Security Workflow | docx | mcp |
| 25 | `capture-and-log-the-stack-trace-when-operationcanceledexception-is-thrown-for-debugging-pu.cs` | Capture and log the stack trace when OperationCanceledException is thrown for debugging purposes. | Security Workflow | docx | mcp |
| 26 | `apply-token-checks-before-invoking-external-resource-retrieval-to-avoid-unnecessary-networ.cs` | Apply token checks before invoking external resource retrieval to avoid unnecessary network calls. | Security Workflow | docx | mcp |
| 27 | `create-an-extension-method-that-adds-cancellation-support-to-existing-synchronous-document.cs` | Create an extension method that adds cancellation support to existing synchronous Document.Save calls. | Security Workflow | doc | mcp |
| 28 | `ensure-that-the-cancellationtokensource-is-disposed-after-processing-completes-to-free-sys.cs` | Ensure that the CancellationTokenSource is disposed after processing completes to free system resources. | Security Workflow | docx | mcp |
| 29 | `demonstrate-chaining-multiple-cancellationtokens-using-cancellationtokensource-createlinke.cs` | Demonstrate chaining multiple CancellationTokens using CancellationTokenSource.CreateLinkedTokenSource for complex workflows in applications. | Security Workflow | docx | mcp |
| 30 | `verify-that-document-processing-pipelines-respect-cancellation-when-integrated-with-third.cs` | Verify that document processing pipelines respect cancellation when integrated with third-party reporting tools. | Security Workflow | doc | mcp |

## Common failure patterns seen during generation and how they were corrected

### Unsupported API invention

- Symptom: Generated code references members that do not exist in the selected package version.
- Fix: Replace invented members with documented Aspose.Words APIs already proven in this category.

### Missing local bootstrap inputs

- Symptom: The example assumes source files, folders, images, or data already exist.
- Fix: Create deterministic local inputs before loading, processing, or validating them.

### Over-broad validation

- Symptom: The example fails at runtime while checking unrelated document internals.
- Fix: Validate only the requested behavior and the existence of expected outputs.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.5.0 run.
- Preserve file-to-task traceability when updating this folder.
- Keep examples standalone and bootstrap local inputs inside the example whenever external sources are mentioned.
