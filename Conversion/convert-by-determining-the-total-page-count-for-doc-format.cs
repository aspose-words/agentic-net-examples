using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOC document from file.
        Document doc = new Document("input.doc");

        // Retrieve the total number of pages after layout.
        int totalPages = doc.PageCount;

        // Display the page count.
        Console.WriteLine($"Total pages: {totalPages}");
    }
}
