using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        const string fileName = "Input.docm";

        // Ensure the file exists. If not, create an empty .docm document.
        if (!File.Exists(fileName))
        {
            var emptyDoc = new Document();
            emptyDoc.Save(fileName, SaveFormat.Docm);
        }

        // Load the Word document.
        Document doc = new Document(fileName);

        // The document may not contain a VBA project.
        if (doc.VbaProject == null)
        {
            Console.WriteLine("The document does not contain a VBA project.");
            return;
        }

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
                Console.WriteLine($"Reference {i}: Type = {reference.Type}, LibId = {reference.LibId}");
            }
        }

        // If there were no references, inform the user.
        if (references.Count == 0)
        {
            Console.WriteLine("No VBA references found in the document.");
        }
    }
}
