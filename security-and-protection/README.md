# Security and Protection Examples for Aspose.Words for .NET

This folder contains the **live, publish-ready** C# examples for the **Security and Protection** category. Each file is a standalone example selected from the latest verified generation run and aligned with the active category rules.

## Snapshot

- Category: **Security and Protection**
- Slug: **security-and-protection**
- Total examples: **30**
- Workflow examples: **30 / 30** use the standard security workflow

## Category rules that shaped these examples

- Use native Aspose.Words APIs directly.
- Create local sample source documents when a task refers to an existing file, folder, stream, or template.
- Do not assume external files or folders already exist.
- Prefer documented protection and password workflows using `Document.Protect`, `Document.Unprotect`, `LoadOptions`, and format-specific save options.
- Keep validation narrow and task-specific.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.3.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\security-and-protection\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0

# PowerShell example
Copy-Item ..\security-and-protection\create-a-cancellationtokensource-and-pass-its-token-to-document-load-to-enable-interruptio.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `create-a-cancellationtokensource-and-pass-its-token-to-document-load-to-enable-interruptio.cs` | Create a CancellationTokenSource and pass its token to Document.Load to enable interruption. | security-workflow | doc | mcp |
| 2 | `attach-a-callback-to-documentloadingargs-throwifcancellationrequested-during-loading-to-en.cs` | Attach a callback to DocumentLoadingArgs.ThrowIfCancellationRequested during loading to ensure proper response. | security-workflow | doc | mcp |
| 3 | `use-token-throwifcancellationrequested-inside-a-field-update-loop-to-abort-long-running-op.cs` | Use token.ThrowIfCancellationRequested inside a field update loop to abort long‑running operations efficiently. | security-workflow | docx | mcp |
| 4 | `implement-a-low-code-ui-button-that-calls-cancellationtokensource-cancel-to-stop-processin.cs` | Implement a low‑code UI button that calls CancellationTokenSource.Cancel to stop processing immediately. | security-workflow | docx | mcp |
| 5 | `combine-a-cancellationtoken-with-document-saveasync-to-allow-asynchronous-saving-cancellat.cs` | Combine a CancellationToken with Document.SaveAsync to allow asynchronous saving cancellation by checking token status. | security-workflow | doc | mcp |
| 6 | `monitor-token-iscancellationrequested-during-layout-building-and-exit-the-layout-routine-p.cs` | Monitor token.IsCancellationRequested during layout building and exit the layout routine promptly if needed. | security-workflow | docx | mcp |
| 7 | `handle-operationcanceledexception-after-a-cancelled-document-load-to-perform-necessary-cle.cs` | Handle OperationCanceledException after a cancelled Document.Load to perform necessary cleanup properly. | security-workflow | doc | mcp |
| 8 | `log-cancellation-events-with-timestamps-to-an-audit-file-for-compliance-tracking.cs` | Log cancellation events with timestamps to an audit file for compliance tracking. | security-workflow | docx | mcp |
| 9 | `pass-the-same-cancellationtoken-to-both-loading-and-saving-methods-for-consistent-interrup.cs` | Pass the same CancellationToken to both loading and saving methods for consistent interruption control. | security-workflow | docx | mcp |
| 10 | `use-token-iscancellationrequested-in-a-custom-linq-reporting-engine-query-to-abort-large-r.cs` | Use token.IsCancellationRequested in a custom LINQ Reporting Engine query to abort large reports. | security-workflow | docx | mcp |
| 11 | `wrap-a-batch-of-document-load-calls-in-a-foreach-loop-that-checks-token-cancellation-befor.cs` | Wrap a batch of Document.Load calls in a foreach loop that checks token cancellation before each load. | security-workflow | doc | mcp |
| 12 | `configure-documentloadingargs-to-invoke-a-user-defined-method-when-cancellation-is-request.cs` | Configure DocumentLoadingArgs to invoke a user‑defined method when cancellation is requested during loading. | security-workflow | doc | mcp |
| 13 | `integrate-cancellationtoken-into-a-background-worker-that-processes-documents-without-bloc.cs` | Integrate CancellationToken into a background worker that processes documents without blocking the UI. | security-workflow | doc | mcp |
| 14 | `validate-that-long-running-field-updates-respect-the-token-by-inserting-periodic-cancellat.cs` | Validate that long‑running field updates respect the token by inserting periodic cancellation checks. | security-workflow | docx | mcp |
| 15 | `create-a-reusable-helper-method-accepting-a-cancellationtoken-to-perform-safe-document-pro.cs` | Create a reusable helper method accepting a CancellationToken to perform safe document processing. | security-workflow | doc | mcp |
| 16 | `ensure-resource-leaks-are-prevented-by-disposing-the-document-object-after-catching-operat.cs` | Ensure resource leaks are prevented by disposing the Document object after catching OperationCanceledException. | security-workflow | doc | mcp |
| 17 | `use-token-throwifcancellationrequested-inside-a-custom-image-extraction-routine-to-stop-ea.cs` | Use token.ThrowIfCancellationRequested inside a custom image extraction routine to stop early if needed. | security-workflow | docx | mcp |
| 18 | `combine-token-monitoring-with-progress-reporting-to-inform-users-when-cancellation-occurs.cs` | Combine token monitoring with progress reporting to inform users when cancellation occurs during processing. | security-workflow | docx | mcp |
| 19 | `implement-a-timeout-mechanism-that-triggers-cancellationtokensource-cancel-after-a-predefi.cs` | Implement a timeout mechanism that triggers CancellationTokenSource.Cancel after a predefined duration automatically. | security-workflow | docx | mcp |
| 20 | `pass-the-token-to-documentbuilder-operations-to-allow-interruption-while-constructing-comp.cs` | Pass the token to DocumentBuilder operations to allow interruption while constructing complex documents. | security-workflow | doc | mcp |
| 21 | `test-cancellation-behavior-by-simulating-user-aborts-during-document-layout-generation-in.cs` | Test cancellation behavior by simulating user aborts during document layout generation in unit tests. | security-workflow | doc | mcp |
| 22 | `document-the-required-net-version-for-cancellationtoken-support-in-the-project-s-readme-fi.cs` | Document the required .NET version for CancellationToken support in the project’s README file. | security-workflow | doc | mcp |
| 23 | `use-token-iscancellationrequested-within-a-while-loop-processing-document-nodes-to-enable.cs` | Use token.IsCancellationRequested within a while loop processing document nodes to enable graceful exit. | security-workflow | doc | mcp |
| 24 | `provide-a-configuration-setting-that-enables-or-disables-cancellation-support-for-specific.cs` | Provide a configuration setting that enables or disables cancellation support for specific processing stages. | security-workflow | docx | mcp |
| 25 | `capture-and-log-the-stack-trace-when-operationcanceledexception-is-thrown-for-debugging-pu.cs` | Capture and log the stack trace when OperationCanceledException is thrown for debugging purposes. | security-workflow | docx | mcp |
| 26 | `apply-token-checks-before-invoking-external-resource-retrieval-to-avoid-unnecessary-networ.cs` | Apply token checks before invoking external resource retrieval to avoid unnecessary network calls. | security-workflow | docx | mcp |
| 27 | `create-an-extension-method-that-adds-cancellation-support-to-existing-synchronous-document.cs` | Create an extension method that adds cancellation support to existing synchronous Document.Save calls. | security-workflow | doc | mcp |
| 28 | `ensure-that-the-cancellationtokensource-is-disposed-after-processing-completes-to-free-sys.cs` | Ensure that the CancellationTokenSource is disposed after processing completes to free system resources. | security-workflow | docx | mcp |
| 29 | `demonstrate-chaining-multiple-cancellationtokens-using-cancellationtokensource-createlinke.cs` | Demonstrate chaining multiple CancellationTokens using CancellationTokenSource.CreateLinkedTokenSource for complex workflows in applications. | security-workflow | docx | mcp |
| 30 | `verify-that-document-processing-pipelines-respect-cancellation-when-integrated-with-third.cs` | Verify that document processing pipelines respect cancellation when integrated with third‑party reporting tools. | security-workflow | doc | mcp |

## Common failure patterns seen during generation and how they were corrected

### Assuming external protected input files already exist

- Symptom: Runtime failures when loading a protected DOC or DOCX that was never created in the example.
- Fix: Create the sample protected input locally first, then reopen it with the correct `LoadOptions` password.

### Using non-existent or format-incompatible protection APIs

- Symptom: Build failures or no-op behavior caused by invented permission, encryption, or save-option members.
- Fix: Use only documented `Document.Protect`, `Document.Unprotect`, `ProtectionType`, `LoadOptions`, and format-specific save options.

### Over-validating unrelated document structure

- Symptom: The requested protection or unprotection succeeds, but the example fails because of unnecessary structural checks.
- Fix: Validate only the intended protection state, password behavior, or output file existence.

## Notes for maintainers

- The selected file for each task is the verified winner recorded in the batch run.
- This category performed best with light primary rules.
- Preserve exact file-to-task traceability when updating the category.
- Bootstrap all sample input files locally inside the example when the task refers to an existing asset.
