using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Path for the temporary macro-enabled document.
        const string filePath = "Sample.docm";

        // Create a new blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = vbaProject;

        // Add a simple VBA module so the document actually contains a VBA project.
        VbaModule module = new VbaModule
        {
            Name = "Module1",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub Hello()\n    MsgBox \"Hello\"\nEnd Sub"
        };
        doc.VbaProject.Modules.Add(module);

        // Save the document in macro-enabled format.
        doc.Save(filePath);

        // Reload the document to work with the saved VBA project.
        Document loadedDoc = new Document(filePath);

        // Get the collection of VBA references.
        VbaReferenceCollection references = loadedDoc.VbaProject.References;

        Console.WriteLine($"Total VBA references: {references.Count}");

        // Enumerate references, filter out COM references, and log the remaining ones.
        foreach (VbaReference reference in references)
        {
            // COM references are of types Registered, Original, or Control.
            if (reference.Type != VbaReferenceType.Registered &&
                reference.Type != VbaReferenceType.Original &&
                reference.Type != VbaReferenceType.Control)
            {
                // Non‑COM reference (e.g., Project type). Log its details.
                Console.WriteLine($"Reference Type: {reference.Type}, LibId: {reference.LibId}");
            }
        }
    }
}
