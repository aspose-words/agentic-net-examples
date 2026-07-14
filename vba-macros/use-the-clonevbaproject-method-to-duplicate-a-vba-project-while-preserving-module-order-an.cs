using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a source macro-enabled document with a VBA project and two modules.
        Document sourceDoc = new Document();
        VbaProject sourceProject = new VbaProject();
        sourceProject.Name = "SourceProject";
        sourceDoc.VbaProject = sourceProject;

        VbaModule module1 = new VbaModule();
        module1.Name = "Module1";
        module1.Type = VbaModuleType.ProceduralModule;
        module1.SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello from Module1\"\nEnd Sub";
        sourceDoc.VbaProject.Modules.Add(module1);

        VbaModule module2 = new VbaModule();
        module2.Name = "Module2";
        module2.Type = VbaModuleType.ProceduralModule;
        module2.SourceCode = "Sub GoodbyeWorld()\n    MsgBox \"Goodbye from Module2\"\nEnd Sub";
        sourceDoc.VbaProject.Modules.Add(module2);

        // Save the source document as a macro-enabled file.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);
        string sourcePath = Path.Combine(artifactsDir, "Source.docm");
        sourceDoc.Save(sourcePath);

        // Load the source document (optional, we already have it in memory).
        Document src = new Document(sourcePath);

        // Create a destination document.
        Document destDoc = new Document();

        // Clone the VBA project from the source and assign it to the destination.
        VbaProject clonedProject = src.VbaProject.Clone();
        destDoc.VbaProject = clonedProject;

        // Output the names of the cloned modules to verify order preservation.
        Console.WriteLine("Cloned VBA project modules order:");
        foreach (VbaModule mod in destDoc.VbaProject.Modules)
        {
            Console.WriteLine(mod.Name);
        }

        // Save the destination document, which now contains the cloned VBA project.
        string destPath = Path.Combine(artifactsDir, "Cloned.docm");
        destDoc.Save(destPath);
    }
}
