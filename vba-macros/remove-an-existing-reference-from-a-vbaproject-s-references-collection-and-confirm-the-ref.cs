using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class RemoveVbaReferenceExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "SampleProject";
        doc.VbaProject = vbaProject;

        // (Optional) Add a simple VBA module so the project looks realistic.
        VbaModule module = new VbaModule
        {
            Name = "SampleModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello, World!\"\nEnd Sub"
        };
        doc.VbaProject.Modules.Add(module);

        // Save the initial document (contains the VBA project with its default references).
        const string originalPath = "Original.docm";
        doc.Save(originalPath);

        // Access the references collection.
        VbaReferenceCollection references = doc.VbaProject.References;
        int initialCount = references.Count;

        // Ensure there is at least one reference to remove.
        if (initialCount > 0)
        {
            // Remove the first reference.
            references.RemoveAt(0);
        }

        int afterCount = references.Count;

        // Output the reference counts to verify the removal.
        Console.WriteLine($"Initial reference count: {initialCount}");
        Console.WriteLine($"Reference count after removal: {afterCount}");

        // Save the modified document.
        const string modifiedPath = "RemovedReference.docm";
        doc.Save(modifiedPath);
    }
}
