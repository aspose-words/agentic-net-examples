using Aspose.Words;
using System;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("input.docx");

        // Determine the total number of pages.
        int totalPages = doc.PageCount;

        // Output the page count.
        Console.WriteLine($"Total pages: {totalPages}");
    }
}
