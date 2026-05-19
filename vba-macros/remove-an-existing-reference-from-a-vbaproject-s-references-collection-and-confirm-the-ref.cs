using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "SampleProject";
        doc.VbaProject = vbaProject;

        // Add a simple VBA module so the project is not empty.
        VbaModule module = new VbaModule
        {
            Name = "SampleModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello, World!\"\nEnd Sub"
        };
        doc.VbaProject.Modules.Add(module);

        // Access the references collection.
        VbaReferenceCollection references = doc.VbaProject.References;

        // Record the initial reference count.
        int initialCount = references.Count;
        Console.WriteLine($"Initial reference count: {initialCount}");

        // If there is at least one reference, remove the first one.
        if (initialCount > 0)
        {
            references.RemoveAt(0);
            int afterRemovalCount = references.Count;
            Console.WriteLine($"Reference count after removal: {afterRemovalCount}");
        }
        else
        {
            Console.WriteLine("No references found to remove.");
        }

        // Save the document in a macro‑enabled format.
        doc.Save("Output.docm");
    }
}
