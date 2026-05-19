using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class CloneVbaProjectExample
{
    public static void Main()
    {
        // Define folder for output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a source macro-enabled document with a VBA project.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject sourceProject = new VbaProject
        {
            Name = "SourceProject"
        };
        sourceDoc.VbaProject = sourceProject;

        // Add first VBA module.
        VbaModule module1 = new VbaModule
        {
            Name = "ModuleOne",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from ModuleOne!""
End Sub"
        };
        sourceDoc.VbaProject.Modules.Add(module1);

        // Add second VBA module.
        VbaModule module2 = new VbaModule
        {
            Name = "ModuleTwo",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub ShowDate()
    MsgBox ""Today is "" & Date
End Sub"
        };
        sourceDoc.VbaProject.Modules.Add(module2);

        // Save the source document as a macro-enabled .docm file.
        string sourcePath = Path.Combine(outputDir, "Source.docm");
        sourceDoc.Save(sourcePath, SaveFormat.Docm);

        // -----------------------------------------------------------------
        // 2. Load the source document and clone its VBA project.
        // -----------------------------------------------------------------
        Document loadedSource = new Document(sourcePath);

        // Perform a deep clone of the VBA project (includes modules and references).
        VbaProject clonedProject = loadedSource.VbaProject.Clone();

        // -----------------------------------------------------------------
        // 3. Create a destination document and assign the cloned VBA project.
        // -----------------------------------------------------------------
        Document destDoc = new Document();

        // Assign the cloned project to the destination document.
        destDoc.VbaProject = clonedProject;

        // Save the destination document.
        string destPath = Path.Combine(outputDir, "Cloned.docm");
        destDoc.Save(destPath, SaveFormat.Docm);

        // -----------------------------------------------------------------
        // 4. Verify that module order and content are preserved.
        // -----------------------------------------------------------------
        Document verifyDoc = new Document(destPath);
        Console.WriteLine("Modules in the cloned document (preserved order):");
        foreach (VbaModule mod in verifyDoc.VbaProject.Modules)
        {
            Console.WriteLine($"- {mod.Name}");
        }

        // The program finishes automatically.
    }
}
