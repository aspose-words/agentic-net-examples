using System;
using Aspose.Words;
// Optional: create an alias to avoid ambiguity with System.Range
using AsposeRange = Aspose.Words.Range;

class Program
{
    static void Main()
    {
        // Load the DOCX document from disk (lifecycle: load)
        Document doc = new Document("input.docx");

        // The Range property of the Document represents the entire document.
        // Use the alias (or 'var') to avoid the ambiguous reference with System.Range.
        AsposeRange wholeDocRange = doc.Range;
        // Alternatively you could write: var wholeDocRange = doc.Range;

        // Example usage: output the text of the entire document to the console.
        Console.WriteLine("Document text:");
        Console.WriteLine(wholeDocRange.Text);

        // Optionally, save the document (lifecycle: save) to verify that no changes were made.
        doc.Save("output.docx");
    }
}
