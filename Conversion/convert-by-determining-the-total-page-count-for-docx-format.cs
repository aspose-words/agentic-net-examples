using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        // Replace "input.docx" with the path to your document.
        Document doc = new Document("input.docx");

        // The PageCount property triggers a layout operation if needed
        // and returns the total number of pages in the document.
        int totalPages = doc.PageCount;

        // Output the page count.
        Console.WriteLine($"Total pages: {totalPages}");
    }
}
