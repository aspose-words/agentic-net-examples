---
name: security-and-protection
description: C# examples for security-and-protection using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - security-and-protection

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **security-and-protection** category.
This folder contains standalone C# examples for security-and-protection operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **security-and-protection**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using System;` (27/27 files) ← category-specific
- `using Aspose.Words;` (26/27 files)
- `using System.Threading;` (18/27 files)
- `using Aspose.Words.Loading;` (16/27 files)
- `using System.IO;` (14/27 files)
- `using Aspose.Words.Saving;` (14/27 files)
- `using System.Collections.Generic;` (3/27 files)
- `using Aspose.Words.Fields;` (3/27 files)
- `using System.Threading.Tasks;` (2/27 files)
- `using System.Linq;` (2/27 files)
- `using Aspose.Words.Reporting;` (2/27 files)
- `using System.Net.Http;` (1/27 files)

## Common Code Pattern

Most files follow this pattern:

```csharp
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);
// ... operations ...
doc.Save("output.docx");
```

## Files in this folder

| File | Key APIs | Description |
|------|----------|-------------|
| [apply-token-checks-before-invoking-external-resource-re...](./apply-token-checks-before-invoking-external-resource-retrieval-avoid-unnecessary.cs) | `Aspose`, `System`, `ResourceLoadingAction` | Apply token checks before invoking external resource retrieval avoid unnecessary |
| [attach-callback-documentloadingargs-throwifcancellation...](./attach-callback-documentloadingargs-throwifcancellationrequested-during-ensure-proper.cs) | `OperationCanceledException`, `Aspose`, `CancellationTokenSource` | Attach callback documentloadingargs throwifcancellationrequested during ensur... |
| [cancellationtokensource-pass-its-token-document-enable-...](./cancellationtokensource-pass-its-token-document-enable-interruption.cs) | `Aspose`, `CancellationTokenSource`, `LoadingProgressCallback` | Cancellationtokensource pass its token document enable interruption |
| [capture-log-stack-trace-when-operationcanceledexception...](./capture-log-stack-trace-when-operationcanceledexception-is-thrown-debugging-purposes.cs) | `Aspose`, `OperationCanceledException`, `Words` | Capture log stack trace when operationcanceledexception is thrown debugging p... |
| [chaining-multiple-cancellationtokens-cancellationtokens...](./chaining-multiple-cancellationtokens-cancellationtokensource-createlinkedtokensource.cs) | `OperationCanceledException`, `CancellationTokenSource`, `Aspose` | Chaining multiple cancellationtokens cancellationtokensource createlinkedtoke... |
| [combine-cancellationtoken-document-saveasync-allow-asyn...](./combine-cancellationtoken-document-saveasync-allow-asynchronous-cancellation-checking.cs) | `System`, `Aspose`, `Document` | Combine cancellationtoken document saveasync allow asynchronous cancellation... |
| [combine-token-monitoring-progress-reporting-inform-user...](./combine-token-monitoring-progress-reporting-inform-users-when-cancellation-occurs.cs) | `Aspose`, `OperationCanceledException`, `Words` | Combine token monitoring progress reporting inform users when cancellation oc... |
| [configure-documentloadingargs-invoke-user-defined-metho...](./configure-documentloadingargs-invoke-user-defined-method-when-cancellation-is.cs) | `Document`, `Aspose`, `LoadingProgressCallback` | Configure documentloadingargs invoke user defined method when cancellation is |
| [document-required-net-version-cancellationtoken-support...](./document-required-net-version-cancellationtoken-support-project-s-readme-file.cs) | `System`, `README`, `AppDomain` | Document required net version cancellationtoken support project s readme file |
| [ensure-resource-leaks-are-prevented-disposing-document-...](./ensure-resource-leaks-are-prevented-disposing-document-object-after-catching.cs) | `Aspose`, `HtmlSaveOptions`, `Document` | Ensure resource leaks are prevented disposing document object after catching |
| [ensure-that-cancellationtokensource-is-disposed-after-p...](./ensure-that-cancellationtokensource-is-disposed-after-processing-completes-free.cs) | `OperationCanceledException`, `CancellationTokenSource`, `Aspose` | Ensure that cancellationtokensource is disposed after processing completes free |
| [extension-method-that-adds-cancellation-support-existin...](./extension-method-that-adds-cancellation-support-existing-synchronous-document-calls.cs) | `ArgumentNullException`, `Aspose`, `Document` | Extension method that adds cancellation support existing synchronous document... |
| [handle-operationcanceledexception-after-cancelled-docum...](./handle-operationcanceledexception-after-cancelled-document-perform-necessary-cleanup.cs) | `Document`, `OperationCanceledException`, `Aspose` | Handle operationcanceledexception after cancelled document perform necessary... |
| [implement-timeout-mechanism-that-triggers-cancellationt...](./implement-timeout-mechanism-that-triggers-cancellationtokensource-cancel-after.cs) | `Aspose`, `OperationCanceledException`, `Words` | Implement timeout mechanism that triggers cancellationtokensource cancel after |
| [integrate-cancellationtoken-background-worker-that-proc...](./integrate-cancellationtoken-background-worker-that-processes-documents-without.cs) | `OperationCanceledException`, `System`, `Aspose` | Integrate cancellationtoken background worker that processes documents without |
| [log-cancellation-events-timestamps-audit-file-complianc...](./log-cancellation-events-timestamps-audit-file-compliance-tracking.cs) | `Aspose`, `OperationCanceledException`, `AuditLogger` | Log cancellation events timestamps audit file compliance tracking |
| [monitor-token-iscancellationrequested-during-layout-bui...](./monitor-token-iscancellationrequested-during-layout-building-exit-layout-routine.cs) | `Document`, `DocumentBuilder`, `System` | Monitor token iscancellationrequested during layout building exit layout routine |
| [pass-same-cancellationtoken-both-methods-consistent-int...](./pass-same-cancellationtoken-both-methods-consistent-interruption-control.cs) | `Aspose`, `OperationCanceledException`, `System` | Pass same cancellationtoken both methods consistent interruption control |
| [provide-configuration-setting-that-enables-disables-can...](./provide-configuration-setting-that-enables-disables-cancellation-support-specific.cs) | `Aspose`, `OperationCanceledException`, `Words` | Provide configuration setting that enables disables cancellation support spec... |
| [reusable-helper-method-accepting-cancellationtoken-perf...](./reusable-helper-method-accepting-cancellationtoken-perform-safe-document-processing.cs) | `Aspose`, `OperationCanceledException`, `System` | Reusable helper method accepting cancellationtoken perform safe document proc... |
| [token-iscancellationrequested-custom-linq-reporting-eng...](./token-iscancellationrequested-custom-linq-reporting-engine-query-abort-large-reports.cs) | `System`, `Document`, `DocumentBuilder` | Token iscancellationrequested custom linq reporting engine query abort large... |
| [token-iscancellationrequested-within-while-loop-process...](./token-iscancellationrequested-within-while-loop-processing-document-nodes-enable.cs) | `Document`, `CancellationTokenSource`, `DocumentProcessor` | Token iscancellationrequested within while loop processing document nodes enable |
| [token-throwifcancellationrequested-inside-custom-image-...](./token-throwifcancellationrequested-inside-custom-image-extraction-routine-stop-early.cs) | `Aspose`, `System`, `Timer` | Token throwifcancellationrequested inside custom image extraction routine sto... |
| [token-throwifcancellationrequested-inside-field-update-...](./token-throwifcancellationrequested-inside-field-update-loop-abort-long-running.cs) | `Aspose`, `Document`, `DocumentBuilder` | Token throwifcancellationrequested inside field update loop abort long running |
| [validate-that-long-running-field-updates-respect-token-...](./validate-that-long-running-field-updates-respect-token-inserting-periodic.cs) | `Aspose`, `Document`, `DocumentBuilder` | Validate that long running field updates respect token inserting periodic |
| [verify-that-document-processing-pipelines-respect-cance...](./verify-that-document-processing-pipelines-respect-cancellation-when-integrated-third.cs) | `Aspose`, `Words`, `OperationCanceledException` | Verify that document processing pipelines respect cancellation when integrate... |
| [wrap-batch-document-calls-foreach-loop-that-checks-toke...](./wrap-batch-document-calls-foreach-loop-that-checks-token-cancellation-before-each.cs) | `System`, `Aspose`, `OperationCanceledException` | Wrap batch document calls foreach loop that checks token cancellation before... |

## Category Statistics
- Total examples: 27

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for:
  - **Boundaries** — Always / Ask First / Never rules for all examples
  - **Common Mistakes** — verified anti-patterns that cause build failures
  - **Domain Knowledge** — cross-cutting API-specific gotchas
  - **Testing Guide** — build and run verification steps
- Review code examples in this folder for security-and-protection patterns

<!-- AUTOGENERATED:START -->
Updated: 2026-03-16 | Run: `20260316_082602`
<!-- AUTOGENERATED:END -->
