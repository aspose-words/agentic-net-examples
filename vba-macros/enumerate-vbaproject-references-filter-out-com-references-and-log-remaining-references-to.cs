using System;
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
            doc.VbaProject = new VbaProject();
        }

        // Save the document in a macro‑enabled format (DOCM).
        const string fileName = "Sample.docm";
        doc.Save(fileName, SaveFormat.Docm);

        // Reload the document to work with the saved VBA project.
        Document loadedDoc = new Document(fileName);

        // Enumerate VBA references, filter out COM (type‑library) references,
        // and log the remaining (project) references.
        foreach (VbaReference reference in loadedDoc.VbaProject.References)
        {
            // Keep only non‑COM references (e.g., project references).
            if (reference.Type != VbaReferenceType.Project)
                continue;

            // Log reference details to the console.
            Console.WriteLine($"Reference LibId: {reference.LibId}, Type: {reference.Type}");
        }
    }
}
