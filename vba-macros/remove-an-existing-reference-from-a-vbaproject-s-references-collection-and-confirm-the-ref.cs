using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Path for the temporary macro-enabled document.
        const string filePath = "Sample.docm";

        // 1. Create a new blank document.
        Document doc = new Document();

        // 2. Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "SampleProject";
        doc.VbaProject = vbaProject;

        // 3. Add a simple VBA module so the project is not empty.
        VbaModule module = new VbaModule();
        module.Name = "Module1";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello, World!\"\nEnd Sub";
        doc.VbaProject.Modules.Add(module);

        // 4. Save the document in macro-enabled format.
        doc.Save(filePath);

        // 5. Reload the document to work with the persisted VBA project.
        Document loadedDoc = new Document(filePath);

        // 6. Access the references collection.
        VbaReferenceCollection references = loadedDoc.VbaProject.References;

        // 7. Record the initial reference count.
        int initialCount = references.Count;
        Console.WriteLine($"Initial reference count: {initialCount}");

        // 8. If there is at least one reference, remove the first one.
        if (initialCount > 0)
        {
            references.RemoveAt(0);
            int newCount = references.Count;
            Console.WriteLine($"Reference count after removal: {newCount}");

            // 9. Verify that the count has decreased.
            bool decreased = newCount == initialCount - 1;
            Console.WriteLine($"Reference count decreased: {decreased}");
        }
        else
        {
            Console.WriteLine("No references found to remove.");
        }

        // 10. Save the modified document (optional).
        loadedDoc.Save("Sample_Modified.docm");
    }
}
