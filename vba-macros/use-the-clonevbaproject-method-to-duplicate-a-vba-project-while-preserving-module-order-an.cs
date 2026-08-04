using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class CloneVbaProjectExample
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docm");
        string clonedPath = Path.Combine(Directory.GetCurrentDirectory(), "Cloned.docm");

        // -------------------------------------------------
        // 1. Create a macro‑enabled source document with a VBA project.
        // -------------------------------------------------
        Document sourceDoc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject sourceProject = new VbaProject
        {
            Name = "SourceProject"
        };
        sourceDoc.VbaProject = sourceProject;

        // Create first VBA module.
        VbaModule module1 = new VbaModule
        {
            Name = "Module1",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub Hello()\n    MsgBox \"Hello from Module1\"\nEnd Sub"
        };
        sourceProject.Modules.Add(module1);

        // Create second VBA module.
        VbaModule module2 = new VbaModule
        {
            Name = "Module2",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub Goodbye()\n    MsgBox \"Goodbye from Module2\"\nEnd Sub"
        };
        sourceProject.Modules.Add(module2);

        // Save the source document as a macro‑enabled file.
        sourceDoc.Save(sourcePath, SaveFormat.Docm);

        // -------------------------------------------------
        // 2. Clone the VBA project from the source document.
        // -------------------------------------------------
        // Load the source document (optional, we already have it in memory).
        Document loadedSource = new Document(sourcePath);

        // Perform a deep clone of the VBA project.
        VbaProject clonedProject = loadedSource.VbaProject.Clone();

        // -------------------------------------------------
        // 3. Create a new destination document and assign the cloned project.
        // -------------------------------------------------
        Document destDoc = new Document();

        // Assign the cloned VBA project to the destination document.
        destDoc.VbaProject = clonedProject;

        // Save the destination document. The cloned project retains module order and references.
        destDoc.Save(clonedPath, SaveFormat.Docm);

        // -------------------------------------------------
        // 4. Simple validation: output module names and source code.
        // -------------------------------------------------
        Console.WriteLine("Cloned VBA project modules:");
        foreach (VbaModule mod in destDoc.VbaProject.Modules)
        {
            Console.WriteLine($"- {mod.Name}");
            Console.WriteLine(mod.SourceCode);
        }
    }
}
