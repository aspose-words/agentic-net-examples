using System;
using Aspose.Words;
using Aspose.Words.Saving;
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
            doc.VbaProject = new VbaProject();
        }

        // Save the document as a macro‑enabled .docm file.
        const string originalPath = "SampleWithReference.docm";
        doc.Save(originalPath, SaveFormat.Docm);

        // Reload the document.
        Document loadedDoc = new Document(originalPath);

        // Count references before removal.
        int countBefore = loadedDoc.VbaProject?.References?.Count ?? 0;

        // Remove the first reference if any exist.
        if (countBefore > 0)
        {
            loadedDoc.VbaProject.References.RemoveAt(0);
        }

        // Count references after removal.
        int countAfter = loadedDoc.VbaProject?.References?.Count ?? 0;

        // Save the modified document.
        const string modifiedPath = "SampleWithoutReference.docm";
        loadedDoc.Save(modifiedPath, SaveFormat.Docm);

        // Output the results.
        Console.WriteLine($"References before removal: {countBefore}");
        Console.WriteLine($"References after removal: {countAfter}");
    }
}
