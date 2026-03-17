using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load a Word document that contains a VBA project.
        Document doc = new Document("Input.docm");

        // Get the collection of VBA references.
        VbaReferenceCollection references = doc.VbaProject.References;

        // Enumerate the references, filter out COM references (Registered, Original, Control),
        // and log the remaining (Project) references to the console.
        for (int i = 0; i < references.Count; i++)
        {
            VbaReference reference = references[i];

            // Keep only non‑COM references (i.e., external VBA project references).
            if (reference.Type == VbaReferenceType.Project)
            {
                // Log useful information about the reference.
                Console.WriteLine($"Reference {i}: Type = {reference.Type}, LibId = {reference.LibId}");
            }
        }
    }
}
