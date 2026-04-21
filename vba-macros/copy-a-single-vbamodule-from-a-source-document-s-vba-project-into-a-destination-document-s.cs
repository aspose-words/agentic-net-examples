using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Paths for the sample documents.
        const string sourcePath = "Source.docm";
        const string destinationPath = "Destination.docm";

        // ---------- Create a source document with a VBA project and a single module ----------
        Document sourceDoc = new Document();

        // Ensure the source document has a VBA project.
        if (sourceDoc.VbaProject == null)
        {
            sourceDoc.VbaProject = new VbaProject();
            sourceDoc.VbaProject.Name = "SourceProject";
        }

        // Create a VBA module with simple macro code.
        VbaModule sourceModule = new VbaModule
        {
            Name = "MyMacro",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello from source\"\nEnd Sub"
        };

        // Add the module to the source document's VBA project.
        sourceDoc.VbaProject.Modules.Add(sourceModule);

        // Save the source document in macro-enabled format.
        sourceDoc.Save(sourcePath);

        // ---------- Load the source document ----------
        Document src = new Document(sourcePath);

        // ---------- Create a destination document ----------
        Document destDoc = new Document();

        // Ensure the destination document has a VBA project.
        if (destDoc.VbaProject == null)
        {
            destDoc.VbaProject = new VbaProject();
            destDoc.VbaProject.Name = "DestinationProject";
        }

        // Retrieve the module to copy by name.
        VbaModule moduleToCopy = src.VbaProject?.Modules["MyMacro"];
        if (moduleToCopy != null)
        {
            // Clone the module to create an independent copy.
            VbaModule clonedModule = moduleToCopy.Clone();

            // Add the cloned module to the destination document's VBA project.
            destDoc.VbaProject.Modules.Add(clonedModule);
        }

        // Save the destination document in macro-enabled format.
        destDoc.Save(destinationPath);

        // Simple verification output.
        Console.WriteLine($"Module '{moduleToCopy?.Name}' copied to '{destinationPath}'.");
    }
}
