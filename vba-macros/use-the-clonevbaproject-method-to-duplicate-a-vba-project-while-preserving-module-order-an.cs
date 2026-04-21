using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Create a source macro-enabled document with a VBA project.
        Document sourceDoc = new Document();

        // Initialize a new VBA project.
        VbaProject sourceProject = new VbaProject
        {
            Name = "SourceProject"
        };
        sourceDoc.VbaProject = sourceProject;

        // Add first VBA module.
        VbaModule module1 = new VbaModule
        {
            Name = "Module1",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub Hello()\n    MsgBox \"Hello\"\nEnd Sub"
        };
        sourceDoc.VbaProject.Modules.Add(module1);

        // Add second VBA module.
        VbaModule module2 = new VbaModule
        {
            Name = "Module2",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub Bye()\n    MsgBox \"Bye\"\nEnd Sub"
        };
        sourceDoc.VbaProject.Modules.Add(module2);

        // Save the source document.
        string sourcePath = Path.Combine(outputDir, "Source.docm");
        sourceDoc.Save(sourcePath);

        // Clone the VBA project from the source document.
        VbaProject clonedProject = sourceDoc.VbaProject.Clone();

        // Create a destination document and assign the cloned project.
        Document destDoc = new Document();
        destDoc.VbaProject = clonedProject;

        // Ensure the destination document contains the same modules in the same order.
        // Remove any existing modules that may have been added automatically.
        while (destDoc.VbaProject.Modules.Count > 0)
        {
            VbaModule toRemove = destDoc.VbaProject.Modules[0];
            destDoc.VbaProject.Modules.Remove(toRemove);
        }

        // Add cloned modules preserving the original order.
        foreach (VbaModule srcModule in sourceDoc.VbaProject.Modules)
        {
            VbaModule clonedModule = srcModule.Clone();
            destDoc.VbaProject.Modules.Add(clonedModule);
        }

        // Optional: display module names to verify order.
        Console.WriteLine("Modules in the cloned document:");
        foreach (VbaModule mod in destDoc.VbaProject.Modules)
        {
            Console.WriteLine($"- {mod.Name}");
        }

        // Save the cloned document.
        string destPath = Path.Combine(outputDir, "Cloned.docm");
        destDoc.Save(destPath);
    }
}
