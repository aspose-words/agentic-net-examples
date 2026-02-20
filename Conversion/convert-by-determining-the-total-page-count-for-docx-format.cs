using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX document from file.
        Document doc = new Document("input.docx");

        // Retrieve the total number of pages after layout.
        int totalPages = doc.PageCount;

        // Output the page count.
        Console.WriteLine($"Total pages: {totalPages}");
    }
}
