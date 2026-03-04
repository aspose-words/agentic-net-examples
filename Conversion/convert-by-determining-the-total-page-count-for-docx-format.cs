using System;
using Aspose.Words;
using Aspose.Words.Properties;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("input.docx");

        // Get the total number of pages. This triggers a layout calculation if needed.
        int totalPages = doc.PageCount;

        // Optionally store the page count in the built‑in document properties.
        doc.BuiltInDocumentProperties.Pages = totalPages;

        // Output the page count.
        Console.WriteLine($"Total pages: {totalPages}");

        // Save the document (optional, demonstrates the required save rule).
        doc.Save("output.docx");
    }
}
