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
        VbaProject vbaProject = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = vbaProject;

        // Add a simple procedural module with some VBA code.
        VbaModule module = new VbaModule
        {
            Name = "Module1",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub Hello()\n    MsgBox \"Hello\"\nEnd Sub"
        };
        doc.VbaProject.Modules.Add(module);

        // Save the document as a macro‑enabled file.
        const string filePath = "Sample.docm";
        doc.Save(filePath);

        // Reload the document to work with the saved VBA project.
        Document loadedDoc = new Document(filePath);

        // Get the collection of VBA references (if any).
        VbaReferenceCollection references = loadedDoc.VbaProject?.References;

        if (references != null && references.Count > 0)
        {
            // Enumerate references, filter out COM references, and log the rest.
            foreach (VbaReference reference in references)
            {
                // COM references are of types Registered, Original, or Control.
                if (reference.Type != VbaReferenceType.Registered &&
                    reference.Type != VbaReferenceType.Original &&
                    reference.Type != VbaReferenceType.Control)
                {
                    Console.WriteLine($"Reference Type: {reference.Type}, LibId: {reference.LibId}");
                }
            }
        }
        else
        {
            Console.WriteLine("No VBA references found in the document.");
        }
    }
}
