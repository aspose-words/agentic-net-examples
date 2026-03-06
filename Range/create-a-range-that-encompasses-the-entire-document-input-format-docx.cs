using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("input.docx");

        // The Range property of the Document represents the entire document.
        // Use 'var' (or a fully‑qualified type) to avoid the ambiguous reference with System.Range.
        var fullRange = doc.Range;

        // Example usage: output the text of the whole document to the console.
        Console.WriteLine(fullRange.Text.Trim());

        // Save the document (optional, demonstrates the save rule).
        doc.Save("output.docx");
    }
}
