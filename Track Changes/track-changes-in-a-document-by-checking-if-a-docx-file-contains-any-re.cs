using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX file.
        Document doc = new Document("input.docx");

        // Check whether the document contains any tracked changes.
        bool hasRevisions = doc.HasRevisions;

        // Output the result.
        Console.WriteLine(hasRevisions
            ? "The document contains revisions."
            : "The document does not contain any revisions.");
    }
}
