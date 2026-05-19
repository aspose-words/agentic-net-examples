using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a source macro‑enabled document with a VBA project.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();

        // Create a new VBA project and add a simple module.
        VbaProject sourceProject = new VbaProject
        {
            Name = "SourceProject"
        };

        VbaModule module1 = new VbaModule
        {
            Name = "Module1",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub Hello()\n    MsgBox \"Hello from source\"\nEnd Sub"
        };

        sourceProject.Modules.Add(module1);

        // Assign the VBA project to the document.
        sourceDoc.VbaProject = sourceProject;

        // Save the source document as a macro‑enabled file.
        string sourcePath = Path.Combine(outputDir, "Source.docm");
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the source document and clone its VBA project.
        // -----------------------------------------------------------------
        Document loadedSource = new Document(sourcePath);

        // Ensure the source document actually contains macros.
        if (!loadedSource.HasMacros)
            throw new InvalidOperationException("Source document does not contain a VBA project.");

        // Perform a deep clone of the VBA project.
        VbaProject clonedProject = loadedSource.VbaProject.Clone();

        // -----------------------------------------------------------------
        // 3. Create a new target document and assign the cloned VBA project.
        // -----------------------------------------------------------------
        Document targetDoc = new Document();

        targetDoc.VbaProject = clonedProject;

        // Save the target document.
        string targetPath = Path.Combine(outputDir, "Target.docm");
        targetDoc.Save(targetPath);
    }
}
