---
name: security-and-protection
description: Verified C# examples for security-and-protection scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Security and Protection

## Purpose

Provide an AI coding assistant with clear guidance to implement cancellation-aware document loading, saving, layout, field updates, and protection workflows while adhering to Aspose.Words 26.11.0's documented API surface.

## Non-negotiable conventions

- Do not reference members that are not present in Aspose.Words 26.11.0, such as LoadOptions.CancellationToken, Document.LoadAsync, or undocumented callback interfaces.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Security Workflow workflow**: 30 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Prioritize confirming that the document was saved (or not saved) as intended and that any protection or cancellation behavior matches the token state.

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
  - Selected engine: existing_repo
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
  - Selected engine: llm
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
  - Selected engine: llm
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
  - Selected engine: llm
- `implement-a-timeout-mechanism-that-triggers-cancellationtokensource-cancel-after-a-predefi.cs`
  - Task: Implement a timeout mechanism that triggers CancellationTokenSource.Cancel after a predefined duration automatically.
  - Workflow: security-workflow
  - Outputs: docx
  - Selected engine: mcp
- `pass-the-token-to-documentbuilder-operations-to-allow-interruption-while-constructing-comp.cs`
  - Task: Pass the token to DocumentBuilder operations to allow interruption while constructing complex documents.
  - Workflow: security-workflow
  - Outputs: doc
  - Selected engine: llm
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
  - Symptom: Compilation fails because the code references members such as LoadOptions.CancellationToken or Document.LoadAsync that do not exist in the referenced package.
  - Root cause: The author assumed newer or custom APIs that are not part of Aspose.Words 26.11.0.
  - Preferred fix: Replace the invented members with the documented constructors and methods, performing explicit token checks before calling the supported overloads.
  - Prevention: Consult the official Aspose.Words 26.11.0 reference and limit code to the APIs listed in the key_apis set.

- **Missing local bootstrap inputs**
  - Symptom: Build succeeds but runtime errors occur because the sample expects external files, folders, or streams that were never created.
  - Root cause: The example relied on pre-existing resources instead of generating deterministic test inputs.
  - Preferred fix: Create temporary documents, directories, or streams within the sample before loading or processing them.
  - Prevention: Always include deterministic bootstrap code that establishes required inputs in the same method.

- **Over-broad validation**
  - Symptom: The program throws exceptions while checking node types or fields unrelated to the requested protection or cancellation scenario.
  - Root cause: The validation logic inspects the entire document tree instead of the specific outcome being demonstrated.
  - Preferred fix: Validate only the presence of the output file and, when relevant, the ProtectionType or password state.
  - Prevention: Keep validation narrow and task-specific as described in the core_rules.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.11.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0
```

```powershell
Copy-Item ..\security-and-protection\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Update examples whenever a new Aspose.Words version adds native cancellation support, and revise the token-checking pattern to use any newly introduced overloads.
