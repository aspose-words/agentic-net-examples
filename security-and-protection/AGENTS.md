---
name: security-and-protection
description: Verified C# examples for security and protection workflows in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Security and Protection

## Purpose

This folder is a **live, curated example set** for security and protection scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free use of documented Aspose.Words APIs for passwords, document protection, editing restrictions, protection removal, and protected load/save workflows.

## Non-negotiable conventions

- Always use documented Aspose.Words APIs directly.
- Always create local sample source documents when a task refers to an existing file, folder, stream, template, or input asset.
- Prefer `Document.Protect`, `Document.Unprotect`, `LoadOptions`, and documented format-specific save options.
- Keep validation narrow and task-specific.
- Do not invent encryption, permission, or protection helper APIs.

## Recommended workflow selection

- **Security workflow**: 30 examples

This category performed best with light primary rules and direct documented protection workflows.

## Validation priorities

1. The code must compile and run without manual input.
2. Required sample inputs must be bootstrapped locally inside the example.
3. Requested protected or unprotected output files must be created successfully.
4. Validation should focus only on the exact protection or password outcome requested by the task.

## File-to-task reference

- `create-a-cancellationtokensource-and-pass-its-token-to-document-load-to-enable-interruptio.cs`
  - Task: Create a CancellationTokenSource and pass its token to Document.Load to enable interruption.
  - Workflow: security-workflow
  - Outputs: doc
  - Selected engine: mcp
- `attach-a-callback-to-documentloadingargs-throwifcancellationrequested-during-loading-to-en.cs`
  - Task: Attach a callback to DocumentLoadingArgs.ThrowIfCancellationRequested during loading to ensure proper response.
  - Workflow: security-workflow
  - Outputs: doc
  - Selected engine: mcp
- `use-token-throwifcancellationrequested-inside-a-field-update-loop-to-abort-long-running-op.cs`
  - Task: Use token.ThrowIfCancellationRequested inside a field update loop to abort long‑running operations efficiently.
  - Workflow: security-workflow
  - Outputs: docx
  - Selected engine: mcp
- `implement-a-low-code-ui-button-that-calls-cancellationtokensource-cancel-to-stop-processin.cs`
  - Task: Implement a low‑code UI button that calls CancellationTokenSource.Cancel to stop processing immediately.
  - Workflow: security-workflow
  - Outputs: docx
  - Selected engine: mcp
- `combine-a-cancellationtoken-with-document-saveasync-to-allow-asynchronous-saving-cancellat.cs`
  - Task: Combine a CancellationToken with Document.SaveAsync to allow asynchronous saving cancellation by checking token status.
  - Workflow: security-workflow
  - Outputs: doc
  - Selected engine: mcp
- `monitor-token-iscancellationrequested-during-layout-building-and-exit-the-layout-routine-p.cs`
  - Task: Monitor token.IsCancellationRequested during layout building and exit the layout routine promptly if needed.
  - Workflow: security-workflow
  - Outputs: docx
  - Selected engine: mcp
- `handle-operationcanceledexception-after-a-cancelled-document-load-to-perform-necessary-cle.cs`
  - Task: Handle OperationCanceledException after a cancelled Document.Load to perform necessary cleanup properly.
  - Workflow: security-workflow
  - Outputs: doc
  - Selected engine: mcp
- `log-cancellation-events-with-timestamps-to-an-audit-file-for-compliance-tracking.cs`
  - Task: Log cancellation events with timestamps to an audit file for compliance tracking.
  - Workflow: security-workflow
  - Outputs: docx
  - Selected engine: mcp
- `pass-the-same-cancellationtoken-to-both-loading-and-saving-methods-for-consistent-interrup.cs`
  - Task: Pass the same CancellationToken to both loading and saving methods for consistent interruption control.
  - Workflow: security-workflow
  - Outputs: docx
  - Selected engine: mcp
- `use-token-iscancellationrequested-in-a-custom-linq-reporting-engine-query-to-abort-large-r.cs`
  - Task: Use token.IsCancellationRequested in a custom LINQ Reporting Engine query to abort large reports.
  - Workflow: security-workflow
  - Outputs: docx
  - Selected engine: mcp
- `wrap-a-batch-of-document-load-calls-in-a-foreach-loop-that-checks-token-cancellation-befor.cs`
  - Task: Wrap a batch of Document.Load calls in a foreach loop that checks token cancellation before each load.
  - Workflow: security-workflow
  - Outputs: doc
  - Selected engine: mcp
- `configure-documentloadingargs-to-invoke-a-user-defined-method-when-cancellation-is-request.cs`
  - Task: Configure DocumentLoadingArgs to invoke a user‑defined method when cancellation is requested during loading.
  - Workflow: security-workflow
  - Outputs: doc
  - Selected engine: mcp
- `integrate-cancellationtoken-into-a-background-worker-that-processes-documents-without-bloc.cs`
  - Task: Integrate CancellationToken into a background worker that processes documents without blocking the UI.
  - Workflow: security-workflow
  - Outputs: doc
  - Selected engine: mcp
- `validate-that-long-running-field-updates-respect-the-token-by-inserting-periodic-cancellat.cs`
  - Task: Validate that long‑running field updates respect the token by inserting periodic cancellation checks.
  - Workflow: security-workflow
  - Outputs: docx
  - Selected engine: mcp
- `create-a-reusable-helper-method-accepting-a-cancellationtoken-to-perform-safe-document-pro.cs`
  - Task: Create a reusable helper method accepting a CancellationToken to perform safe document processing.
  - Workflow: security-workflow
  - Outputs: doc
  - Selected engine: mcp
- `ensure-resource-leaks-are-prevented-by-disposing-the-document-object-after-catching-operat.cs`
  - Task: Ensure resource leaks are prevented by disposing the Document object after catching OperationCanceledException.
  - Workflow: security-workflow
  - Outputs: doc
  - Selected engine: mcp
- `use-token-throwifcancellationrequested-inside-a-custom-image-extraction-routine-to-stop-ea.cs`
  - Task: Use token.ThrowIfCancellationRequested inside a custom image extraction routine to stop early if needed.
  - Workflow: security-workflow
  - Outputs: docx
  - Selected engine: mcp
- `combine-token-monitoring-with-progress-reporting-to-inform-users-when-cancellation-occurs.cs`
  - Task: Combine token monitoring with progress reporting to inform users when cancellation occurs during processing.
  - Workflow: security-workflow
  - Outputs: docx
  - Selected engine: mcp
- `implement-a-timeout-mechanism-that-triggers-cancellationtokensource-cancel-after-a-predefi.cs`
  - Task: Implement a timeout mechanism that triggers CancellationTokenSource.Cancel after a predefined duration automatically.
  - Workflow: security-workflow
  - Outputs: docx
  - Selected engine: mcp
- `pass-the-token-to-documentbuilder-operations-to-allow-interruption-while-constructing-comp.cs`
  - Task: Pass the token to DocumentBuilder operations to allow interruption while constructing complex documents.
  - Workflow: security-workflow
  - Outputs: doc
  - Selected engine: mcp
- `test-cancellation-behavior-by-simulating-user-aborts-during-document-layout-generation-in.cs`
  - Task: Test cancellation behavior by simulating user aborts during document layout generation in unit tests.
  - Workflow: security-workflow
  - Outputs: doc
  - Selected engine: mcp
- `document-the-required-net-version-for-cancellationtoken-support-in-the-project-s-readme-fi.cs`
  - Task: Document the required .NET version for CancellationToken support in the project’s README file.
  - Workflow: security-workflow
  - Outputs: doc
  - Selected engine: mcp
- `use-token-iscancellationrequested-within-a-while-loop-processing-document-nodes-to-enable.cs`
  - Task: Use token.IsCancellationRequested within a while loop processing document nodes to enable graceful exit.
  - Workflow: security-workflow
  - Outputs: doc
  - Selected engine: mcp
- `provide-a-configuration-setting-that-enables-or-disables-cancellation-support-for-specific.cs`
  - Task: Provide a configuration setting that enables or disables cancellation support for specific processing stages.
  - Workflow: security-workflow
  - Outputs: docx
  - Selected engine: mcp
- `capture-and-log-the-stack-trace-when-operationcanceledexception-is-thrown-for-debugging-pu.cs`
  - Task: Capture and log the stack trace when OperationCanceledException is thrown for debugging purposes.
  - Workflow: security-workflow
  - Outputs: docx
  - Selected engine: mcp
- `apply-token-checks-before-invoking-external-resource-retrieval-to-avoid-unnecessary-networ.cs`
  - Task: Apply token checks before invoking external resource retrieval to avoid unnecessary network calls.
  - Workflow: security-workflow
  - Outputs: docx
  - Selected engine: mcp
- `create-an-extension-method-that-adds-cancellation-support-to-existing-synchronous-document.cs`
  - Task: Create an extension method that adds cancellation support to existing synchronous Document.Save calls.
  - Workflow: security-workflow
  - Outputs: doc
  - Selected engine: mcp
- `ensure-that-the-cancellationtokensource-is-disposed-after-processing-completes-to-free-sys.cs`
  - Task: Ensure that the CancellationTokenSource is disposed after processing completes to free system resources.
  - Workflow: security-workflow
  - Outputs: docx
  - Selected engine: mcp
- `demonstrate-chaining-multiple-cancellationtokens-using-cancellationtokensource-createlinke.cs`
  - Task: Demonstrate chaining multiple CancellationTokens using CancellationTokenSource.CreateLinkedTokenSource for complex workflows in applications.
  - Workflow: security-workflow
  - Outputs: docx
  - Selected engine: mcp
- `verify-that-document-processing-pipelines-respect-cancellation-when-integrated-with-third.cs`
  - Task: Verify that document processing pipelines respect cancellation when integrated with third‑party reporting tools.
  - Workflow: security-workflow
  - Outputs: doc
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Missing local protected sample inputs**
  - Symptom: Runtime failures when the example tries to open a protected file that was never created.
  - Preferred fix: Create the sample file locally first, apply protection, save it, and then reopen it with the correct password or load options.
- **Invented protection or encryption APIs**
  - Symptom: Build failures caused by non-existent protection helpers or unsupported save-option members.
  - Preferred fix: Use only documented `Document.Protect`, `Document.Unprotect`, `ProtectionType`, `LoadOptions`, and format-specific save options.
- **Over-validating unrelated structure**
  - Symptom: The requested security action succeeds, but the example fails because of unnecessary structural checks.
  - Preferred fix: Validate only the exact requested protection state, password behavior, or output file existence.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.3.0`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required package

```bash
dotnet add package Aspose.Words --version 26.3.0
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\security-and-protection\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve exact file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the verified winner from the latest batch report rather than a merely compiling draft.
- Bootstrap file-based inputs locally instead of depending on machine-specific paths.
