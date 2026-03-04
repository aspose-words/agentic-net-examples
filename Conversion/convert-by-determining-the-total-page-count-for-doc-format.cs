using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOC document from disk.
        Document doc = new Document("InputDocument.doc");

        // Retrieve the total number of pages after layout.
        int totalPages = doc.PageCount;

        // Output the page count.
        Console.WriteLine($"Total page count: {totalPages}");
    }
}
