using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Folder for generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a source macro‑enabled document with a VBA project.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();

        // Create a new VBA project and give it a name.
        VbaProject sourceProject = new VbaProject
        {
            Name = "SourceProject"
        };

        // Create a procedural module with some simple VBA code.
        VbaModule sourceModule = new VbaModule
        {
            Name = "Module1",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from source document!""
End Sub"
        };

        // Add the module to the project.
        sourceProject.Modules.Add(sourceModule);

        // Attach the VBA project to the document.
        sourceDoc.VbaProject = sourceProject;

        // Save the source document as a macro‑enabled file.
        string sourcePath = Path.Combine(outputDir, "Source.docm");
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the source document (optional – we already have it in memory).
        // -----------------------------------------------------------------
        Document loadedSource = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 3. Clone the entire VBA project from the source document.
        // -----------------------------------------------------------------
        VbaProject clonedProject = loadedSource.VbaProject.Clone();

        // -----------------------------------------------------------------
        // 4. Create a new target document and assign the cloned VBA project.
        // -----------------------------------------------------------------
        Document targetDoc = new Document();
        targetDoc.VbaProject = clonedProject;

        // Save the target document; it now contains the same VBA macros.
        string targetPath = Path.Combine(outputDir, "Target.docm");
        targetDoc.Save(targetPath);

        // Simple validation – the target document should have macros.
        if (targetDoc.HasMacros)
        {
            Console.WriteLine("VBA project successfully cloned to target document.");
        }
        else
        {
            Console.WriteLine("Cloning failed – target document has no macros.");
        }
    }
}
