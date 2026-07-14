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

        // Add a simple procedural module so the document is truly macro‑enabled.
        VbaModule module = new VbaModule
        {
            Name = "SampleModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello, World!\"\nEnd Sub"
        };
        doc.VbaProject.Modules.Add(module);

        // Access the references collection.
        VbaReferenceCollection references = doc.VbaProject.References;

        // Record the initial count of references.
        int initialCount = references.Count;
        Console.WriteLine($"Initial reference count: {initialCount}");

        // If there is at least one reference, remove the last one and confirm the count decreases.
        if (initialCount > 0)
        {
            // Remove the reference at the last index.
            references.RemoveAt(initialCount - 1);

            int afterRemovalCount = references.Count;
            Console.WriteLine($"Reference count after removal: {afterRemovalCount}");

            // Simple validation that the count decreased by one.
            Console.WriteLine(afterRemovalCount == initialCount - 1
                ? "Reference removal confirmed."
                : "Reference count did not decrease as expected.");
        }
        else
        {
            Console.WriteLine("No references were present to remove.");
        }

        // Save the document in a macro‑enabled format.
        doc.Save("RemoveVbaReferenceResult.docm");
    }
}
