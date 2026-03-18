using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the document from a file.
        Document doc = new Document("input.docx");

        // Determine whether the document contains any tracked changes.
        if (doc.HasRevisions)
        {
            // The document has revisions; you can process or reject them here.
            Console.WriteLine("Document contains revisions.");
        }
        else
        {
            // No revisions are present; safe to continue processing.
            Console.WriteLine("Document does not contain revisions.");
        }

        // Optionally save the document after inspection.
        doc.Save("output.docx");
    }
}
