using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX document from disk.
        Document doc = new Document("input.docx");

        // Delete all characters within the document's main range.
        doc.Range.Delete();

        // Save the modified document back to disk.
        doc.Save("output.docx");
    }
}
