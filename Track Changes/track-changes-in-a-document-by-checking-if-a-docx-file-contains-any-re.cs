using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX file.
        Document doc = new Document("InputDocument.docx");

        // Check whether the document contains any tracked changes (revisions).
        bool hasRevisions = doc.HasRevisions;

        // Output the result.
        Console.WriteLine(hasRevisions
            ? "The document contains revisions."
            : "The document does not contain any revisions.");

        // (Optional) Save a copy of the document if needed.
        // doc.Save("OutputDocument.docx");
    }
}
