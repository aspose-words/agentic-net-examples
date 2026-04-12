---
name: aspose-words-examples
description: AI-friendly C# code examples for Aspose.Words for .NET
language: csharp
framework: net8.0
package: Aspose.Words 26.3.0
---

# Aspose.Words for .NET Examples

AI-friendly repository containing validated C# examples for Aspose.Words for .NET API.

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET.
When working in this repository:
- Each `.cs` file is a **standalone Console Application** — do not create multi-file projects
- All examples must **compile and run** without errors using `dotnet build` and `dotnet run`
- Follow the conventions, boundaries, and anti-patterns documented below exactly
- Use the **Command Reference** section for build/run commands

## Repository Overview

This repository contains **1746** verified code examples demonstrating Aspose.Words for .NET capabilities.

**Statistics** (as of 2026-04-12):
- Total Examples: 1746
- Categories: 29

## Category Details

### barcode-image
- Examples: 30
- Guide: [AGENTS.md](./barcode-image/AGENTS.md)

### charts
- Examples: 39
- Guide: [AGENTS.md](./charts/AGENTS.md)

### comments
- Examples: 30
- Guide: [AGENTS.md](./comments/AGENTS.md)

### comparison
- Examples: 30
- Guide: [AGENTS.md](./comparison/AGENTS.md)

### content-control
- Examples: 35
- Guide: [AGENTS.md](./content-control/AGENTS.md)

### conversion
- Examples: 108
- Guide: [AGENTS.md](./conversion/AGENTS.md)

### extraction
- Examples: 30
- Guide: [AGENTS.md](./extraction/AGENTS.md)

### find-and-replace
- Examples: 63
- Guide: [AGENTS.md](./find-and-replace/AGENTS.md)

### fonts
- Examples: 35
- Guide: [AGENTS.md](./fonts/AGENTS.md)

### form-fields
- Examples: 28
- Guide: [AGENTS.md](./form-fields/AGENTS.md)

### hyphenation
- Examples: 37
- Guide: [AGENTS.md](./hyphenation/AGENTS.md)

### images
- Examples: 75
- Guide: [AGENTS.md](./images/AGENTS.md)

### join-documents
- Examples: 34
- Guide: [AGENTS.md](./join-documents/AGENTS.md)

### linq-reporting
- Examples: 404
- Guide: [AGENTS.md](./linq-reporting/AGENTS.md)

### lists
- Examples: 30
- Guide: [AGENTS.md](./lists/AGENTS.md)

### mail-merge
- Examples: 29
- Guide: [AGENTS.md](./mail-merge/AGENTS.md)

### officemath
- Examples: 18
- Guide: [AGENTS.md](./officemath/AGENTS.md)

### ole-objects
- Examples: 27
- Guide: [AGENTS.md](./ole-objects/AGENTS.md)

### paragraphs
- Examples: 33
- Guide: [AGENTS.md](./paragraphs/AGENTS.md)

### programming-with-documents
- Examples: 106
- Guide: [AGENTS.md](./programming-with-documents/AGENTS.md)

### range
- Examples: 29
- Guide: [AGENTS.md](./range/AGENTS.md)

### rendering
- Examples: 57
- Guide: [AGENTS.md](./rendering/AGENTS.md)

### security-and-protection
- Examples: 27
- Guide: [AGENTS.md](./security-and-protection/AGENTS.md)

### shapes
- Examples: 35
- Guide: [AGENTS.md](./shapes/AGENTS.md)

### split-document
- Examples: 23
- Guide: [AGENTS.md](./split-document/AGENTS.md)

### track-changes
- Examples: 37
- Guide: [AGENTS.md](./track-changes/AGENTS.md)

### vba-macros
- Examples: 30
- Guide: [AGENTS.md](./vba-macros/AGENTS.md)

### watermark
- Examples: 28
- Guide: [AGENTS.md](./watermark/AGENTS.md)

### working-with-table
- Examples: 189
- Guide: [AGENTS.md](./working-with-table/AGENTS.md)

## Boundaries

### ✅ Always

These rules are mandatory for every example.

- Use public `Aspose.Words` APIs and keep examples self-contained.
- Build with the configured .NET target framework and package version.
- Save the resulting document after completing modifications.
- Keep input and output paths explicit and easy to follow.

### ⚠️ Ask First

- Adding extra NuGet packages beyond `Aspose.Words`.
- Changing shared repository infrastructure or project templates.
- Introducing platform-specific behavior not already present in the examples.

### 🚫 Never

- Never use excluded namespaces in generated code.
- Never create multi-file example projects for a single example.
- Never modify generated `AGENTS.md` or `index.json` by hand.

## Common Mistakes (Anti-Patterns)

### CS0246: Shape type not found
Add the missing using directive for the Aspose.Words.Drawing namespace, which defines the Shape class used by DocumentBuilder.InsertChart. The Chart class lives in Aspose.Words.Drawing.Charts, which was already imported.

```csharp
using Aspose.Words.Drawing; // required for Shape
// existing using Aspose.Words.Drawing.Charts; stays unchanged
```

Shape is part of Aspose.Words.Drawing, not Aspose.Words.Drawing.Charts. Without the directive the compiler cannot resolve the type. This fix works with Aspose.Words for .NET 23.10+ on all supported platforms.

### Missing static Main entry point
Added a Program class with a static Main method that serves as the application entry point. The Main method demonstrates how to load two DOCX files into byte arrays, invoke the RevisionReportGenerator, and output the generated report to the console. This satisfies the compiler requirement for an executable project.

```csharp
class Program
{
    static void Main(string[] args)
    {
        // usage example
    }
}
```

The original file only defined a helper class; a console application must expose a static Main method. No changes to Aspose.Words APIs were required.

### CS0246: Missing Shape type
Add a reference to the Aspose.Words.Drawing namespace, which contains the Shape class used for chart manipulation. If you do not need to manipulate the chart programmatically, you can simply remove the Shape‑related code.

```csharp
using Aspose.Words.Drawing; // <-- added
// ... later, if needed:
Shape chartShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
```

Shape is part of Aspose.Words.Drawing. The project must reference the Aspose.Words NuGet package (which includes this namespace).

### CS5001: Missing static Main method
Added a public static Main(string[] args) method. This satisfies the compiler requirement for an executable entry point.

```csharp
public static void Main(string[] args)
{
    if (args.Length != 2)
    {
        Console.WriteLine("Usage: DocxToHtmlConverter <input.docx> <output.html>");
        return;
    }
    ConvertComplexTablesToHtml(args[0], args[1]);
}
```

Console applications compiled with `dotnet build` need a static Main method. Adding it makes the example runnable from the command line. No change to Aspose.Words API is required.

### BookmarkCollection.Exists method not found
Replace the non‑existent Exists call with the indexer accessor. The BookmarkCollection indexer returns null when a bookmark is missing, so a simple null‑check is sufficient.

```csharp
Bookmark startBookmark = sourceDoc.Range.Bookmarks[startBookmarkName];
Bookmark endBookmark = sourceDoc.Range.Bookmarks[endBookmarkName];
if (startBookmark == null || endBookmark == null) continue;
```

Aspose.Words for .NET removed the Exists method in recent versions. Using the indexer is the recommended, version‑agnostic way to test for a bookmark's presence.

### CS0246: Missing using directive or assembly reference for Aspose.Words types
Add the Aspose.Words namespace with a using directive and ensure the Aspose.Words NuGet package (or DLL) is referenced by the project. The types Document and DocumentBuilder are defined in the Aspose.Words assembly.

```csharp
using Aspose.Words; // add this at the top of the file

// In the .csproj file add:
<ItemGroup>
  <PackageReference Include="Aspose.Words" Version="26.3.0" />
</ItemGroup>
```

Without the using directive the compiler cannot resolve the types. Adding the NuGet package also supplies the required assembly at runtime. 

## Domain Knowledge

Cross-cutting rules and API-specific gotchas.

- Keep examples minimal, strongly typed, and easy to compile.

## Command Reference

### Build and Run

Each repository example is a standalone `.cs` file. Run examples by creating a temporary console project and copying the desired example to `Program.cs`.

```bash
# Create a new project
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject

# Add Aspose.Words
dotnet add package Aspose.Words --version 26.3.0

# Copy a repository example into the project as Program.cs
# PowerShell:
Copy-Item ..\<category-folder>\<example-file>.cs .\Program.cs

# Build
dotnet build --configuration Release --verbosity minimal

# Run
dotnet run --configuration Release --no-build
```

### Project File (.csproj)

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Aspose.Words" Version="26.3.0" />
  </ItemGroup>
</Project>
```

### Environment

- .NET SDK: net8.0 or higher
- NuGet: Aspose.Words 26.3.0
- All examples are standalone Console Applications
- Each `.cs` file must be compiled and run independently inside a temporary console project

## Testing Guide

Every example must pass these verification steps.

### Build Verification

```bash
dotnet build --configuration Release --verbosity minimal
```

- Success: exit code `0` and no `CS####` compiler errors
- Failure: any `error CS####` line indicates a build failure

### Run Verification

```bash
dotnet run --configuration Release --no-build
```

- Success: exit code `0`, no unhandled exception, and expected output file(s) created where applicable
- Failure: non-zero exit code, `Unhandled exception`, or missing expected output

## How to Use These Examples

### Prerequisites

- .NET SDK (net8.0 or higher)
- Aspose.Words for .NET (26.3.0 or higher)
- NuGet package restore enabled
- Any required input files available in the temporary project working directory

### Running an Example

1. Choose a category folder and a single `.cs` example file
2. Create a temporary console project targeting `net8.0`
3. Add `Aspose.Words` version `26.3.0`
4. Copy the selected example file into the project as `Program.cs`
5. Build and run the project

Updated: 2026-03-26 | Run: `20260315_203408_65134f` | Examples: 1534 | Categories: 29
