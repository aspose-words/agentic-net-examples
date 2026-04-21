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
        VbaProject project = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = project;

        // Add a simple VBA module so the document is macro-enabled.
        VbaModule module = new VbaModule
        {
            Name = "SampleModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello, World!\"\nEnd Sub"
        };
        project.Modules.Add(module);

        // Save the document in a macro-enabled format.
        const string fileName = "Sample.docm";
        doc.Save(fileName);

        // Reload the document to demonstrate reading references.
        Document loadedDoc = new Document(fileName);

        // Ensure the document actually contains a VBA project.
        if (loadedDoc.VbaProject == null)
        {
            Console.WriteLine("The document does not contain a VBA project.");
            return;
        }

        // Get the collection of VBA references.
        VbaReferenceCollection references = loadedDoc.VbaProject.References;

        // Enumerate references, filter out COM references, and log the remaining ones.
        foreach (VbaReference reference in references)
        {
            // COM references are of types Registered, Original, or Control.
            bool isComReference = reference.Type == VbaReferenceType.Registered ||
                                  reference.Type == VbaReferenceType.Original ||
                                  reference.Type == VbaReferenceType.Control;

            if (!isComReference)
            {
                // Log the reference type and its LibId (if any).
                Console.WriteLine($"Reference Type: {reference.Type}");
                Console.WriteLine($"LibId: {reference.LibId ?? string.Empty}");
                Console.WriteLine();
            }
        }

        // If there were no non‑COM references, inform the user.
        if (references.Count == 0 || references.Count == references.Count - 0) // placeholder to avoid unused variable warning
        {
            // No action needed; the loop above already handled empty collections.
        }
    }
}
