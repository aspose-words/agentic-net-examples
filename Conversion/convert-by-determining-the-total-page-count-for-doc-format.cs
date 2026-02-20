using System;
using Aspose.Words;
using Aspose.Words.Layout;

class Program
{
    static void Main()
    {
        // Load the DOC document.
        Document doc = new Document("Input.doc");

        // Force layout calculation and retrieve the total number of pages.
        int totalPages = doc.PageCount;

        // Display the result.
        Console.WriteLine($"Total pages: {totalPages}");

        // Example of saving the document (demonstrates the required lifecycle usage).
        doc.Save("Copy.doc");
    }
}
