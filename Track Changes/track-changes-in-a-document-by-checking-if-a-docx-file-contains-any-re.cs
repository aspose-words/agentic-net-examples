using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX file.
        Document doc = new Document("input.docx");

        // Determine whether the document has any tracked changes.
        bool hasRevisions = doc.HasRevisions;

        // Output the result.
        Console.WriteLine(hasRevisions
            ? "The document contains tracked changes."
            : "The document does not contain any tracked changes.");
    }
}
