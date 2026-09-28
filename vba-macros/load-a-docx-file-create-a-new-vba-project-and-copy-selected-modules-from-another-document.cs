using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Prepare temporary directory and file paths
        string tempDir = Path.Combine(Path.GetTempPath(), "AsposeVbaDemo");
        Directory.CreateDirectory(tempDir);
        string sourcePath = Path.Combine(tempDir, "source.docm");
        string targetPath = Path.Combine(tempDir, "target.docm");

        // ---------- Create source document with a VBA project ----------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the source document containing VBA modules.");

        // Ensure the source document has a VBA project
        sourceDoc.VbaProject = new VbaProject();

        // Create and add a VBA module to the source document
        string moduleName = "TestModule";
        string moduleCode = @"Sub Hello()
    MsgBox ""Hello from VBA!""
End Sub";

        VbaModule sourceModule = new VbaModule
        {
            Name = moduleName,
            SourceCode = moduleCode
        };
        sourceDoc.VbaProject.Modules.Add(sourceModule);

        // Save the source document as a macro‑enabled file
        sourceDoc.Save(sourcePath, SaveFormat.Docm);

        // ---------- Create target document ----------
        Document targetDoc = new Document(); // blank document
        targetDoc.VbaProject = new VbaProject(); // ensure a VBA project exists

        // Load the source document to access its VBA modules
        Document loadedSource = new Document(sourcePath);

        // Copy the selected module(s) to the target document
        foreach (VbaModule srcModule in loadedSource.VbaProject.Modules)
        {
            if (srcModule.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase))
            {
                // Guard against null source code
                string sourceCode = srcModule.SourceCode ?? string.Empty;

                VbaModule newModule = new VbaModule
                {
                    Name = srcModule.Name,
                    SourceCode = sourceCode
                };
                targetDoc.VbaProject.Modules.Add(newModule);
            }
        }

        // Save the target document as a macro‑enabled file
        targetDoc.Save(targetPath, SaveFormat.Docm);

        // ---------- Validation ----------
        bool moduleCopied = false;
        VbaModule copied = targetDoc.VbaProject?.Modules[moduleName];
        if (copied != null && !string.IsNullOrEmpty(copied.SourceCode))
        {
            moduleCopied = true;
        }

        Console.WriteLine($"VBA module '{moduleName}' copied to target document: {moduleCopied}");
    }
}
