using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        if (doc.VbaProject == null)
        {
            VbaProject project = new VbaProject();
            project.Name = "SampleProject";
            doc.VbaProject = project;
        }

        // Add a simple VBA module (required for a VBA project to be valid).
        VbaModule module = new VbaModule
        {
            Name = "SampleModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello, World!\"\nEnd Sub"
        };
        doc.VbaProject.Modules.Add(module);

        // Save the document in a macro‑enabled format.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string docPath = Path.Combine(outputDir, "VbaReferenceDemo.docm");
        doc.Save(docPath);

        // Reload the document to work with the saved VBA project.
        Document loadedDoc = new Document(docPath);

        // Access the references collection.
        VbaReferenceCollection references = loadedDoc.VbaProject.References;

        // Record the initial count.
        int initialCount = references.Count;
        Console.WriteLine($"Initial reference count: {initialCount}");

        // If there is at least one reference, remove the first one.
        if (initialCount > 0)
        {
            // Remove the reference at index 0.
            references.RemoveAt(0);

            // Record the new count.
            int newCount = references.Count;
            Console.WriteLine($"Reference count after removal: {newCount}");

            // Verify that the count decreased.
            Console.WriteLine(newCount < initialCount
                ? "Reference removal confirmed."
                : "Reference count did not decrease as expected.");
        }
        else
        {
            Console.WriteLine("No references found to remove.");
        }

        // Save the modified document.
        string modifiedPath = Path.Combine(outputDir, "VbaReferenceDemo_Modified.docm");
        loadedDoc.Save(modifiedPath);
    }
}
