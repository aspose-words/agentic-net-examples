---
name: security-and-protection
description: Verified C# examples for security-and-protection scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Security and Protection

## Purpose

Provide concise, task-focused code snippets that enable developers to protect, unprotect, and safely load or save Word documents while handling cancellation without relying on undocumented or unavailable APIs.

## Non-negotiable conventions

- Do not reference Aspose.Words members that are not present in version 26.10.0.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Security Workflow workflow**: 30 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Confirm that the output file is created and that its protection status matches the requested ProtectionType.

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
  - Selected engine: llm
- `use-token-throwifcancellationrequested-inside-a-field-update-loop-to-abort-long-running-op.cs`
  - Task: Use token.ThrowIfCancellationRequested inside a field update loop to abort long-running operations efficiently.
  - Workflow: security-workflow
  - Outputs: docx
  - Selected engine: mcp
- `implement-a-low-code-ui-button-that-calls-cancellationtokensource-cancel-to-stop-processin.cs`
  - Task: Implement a low-code UI button that calls CancellationTokenSource.Cancel to stop processing immediately.
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
  - Selected engine: llm
- `use-token-iscancellationrequested-in-a-custom-linq-reporting-engine-query-to-abort-large-r.cs`
  - Task: Use token.IsCancellationRequested in a custom LINQ Reporting Engine query to abort large reports.
  - Workflow: security-workflow
  - Outputs: docx
  - Selected engine: existing_repo
- `wrap-a-batch-of-document-load-calls-in-a-foreach-loop-that-checks-token-cancellation-befor.cs`
  - Task: Wrap a batch of Document.Load calls in a foreach loop that checks token cancellation before each load.
  - Workflow: security-workflow
  - Outputs: doc
  - Selected engine: mcp
- `configure-documentloadingargs-to-invoke-a-user-defined-method-when-cancellation-is-request.cs`
  - Task: Configure DocumentLoadingArgs to invoke a user-defined method when cancellation is requested during loading.
  - Workflow: security-workflow
  - Outputs: doc
  - Selected engine: existing_repo
- `integrate-cancellationtoken-into-a-background-worker-that-processes-documents-without-bloc.cs`
  - Task: Integrate CancellationToken into a background worker that processes documents without blocking the UI.
  - Workflow: security-workflow
  - Outputs: doc
  - Selected engine: mcp
- `validate-that-long-running-field-updates-respect-the-token-by-inserting-periodic-cancellat.cs`
  - Task: Validate that long-running field updates respect the token by inserting periodic cancellation checks.
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
  - Task: Document the required .NET version for CancellationToken support in the project's README file.
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
  - Task: Verify that document processing pipelines respect cancellation when integrated with third-party reporting tools.
  - Workflow: security-workflow
  - Outputs: doc
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Unsupported API invention**
  - Symptom: Compilation fails because members such as LoadOptions.LoadingCallback, Document.LoadAsync, or Document.Dispose are not defined.
  - Root cause: The code references APIs that do not exist in the shipped Aspose.Words package.
  - Preferred fix: Replace invented members with the documented Document constructor, LoadOptions, and explicit token-aware wrappers using Task.Run.
  - Prevention: Consult the official Aspose.Words 26.10.0 reference and limit code to APIs listed in the key_apis set.

- **Missing CancellationToken support**
  - Symptom: Build errors occur when passing a CancellationToken directly to Document.Load or Document.Save, and runtime checks cannot locate cancellation callbacks.
  - Root cause: Aspose.Words does not expose overloads that accept a CancellationToken in this version.
  - Preferred fix: Wrap the load or save call inside Task.Run and invoke token.ThrowIfCancellationRequested before and after the operation.
  - Prevention: Design cancellation logic around token checks surrounding synchronous Aspose.Words calls rather than expecting built-in overloads.

- **Missing local bootstrap inputs**
  - Symptom: Code assumes that source documents, templates, or image files already exist, leading to FileNotFound exceptions or invalid test results.
  - Root cause: Samples do not create the required files before processing.
  - Preferred fix: Generate deterministic sample documents, templates, or images at the start of each example and clean them up afterward.
  - Prevention: Always include a self-contained setup block that creates any files the task needs.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.10.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0
```

```powershell
Copy-Item ..\security-and-protection\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Update the examples whenever a new Aspose.Words version adds native CancellationToken overloads to simplify the wrappers.
