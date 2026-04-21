using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a source document with a VBA project.
        Document sourceDoc = new Document();

        // Initialize a new VBA project.
        VbaProject sourceProject = new VbaProject
        {
            Name = "SourceProject"
        };

        // Create a VBA module with simple macro code.
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

        // Assign the VBA project to the source document.
        sourceDoc.VbaProject = sourceProject;

        // Save the source document as a macro-enabled file (optional, for inspection).
        string sourcePath = Path.Combine(Environment.CurrentDirectory, "Source.docm");
        sourceDoc.Save(sourcePath);

        // Create a destination (blank) document.
        Document destDoc = new Document();

        // Clone the entire VBA project from the source document.
        VbaProject clonedProject = sourceDoc.VbaProject.Clone();

        // Assign the cloned project to the destination document.
        destDoc.VbaProject = clonedProject;

        // Save the destination document, which now contains the cloned VBA project.
        string destPath = Path.Combine(Environment.CurrentDirectory, "Cloned.docm");
        destDoc.Save(destPath);
    }
}
