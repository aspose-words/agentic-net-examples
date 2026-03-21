using Aspose.Words;
using System;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Add a paragraph with some text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");

        // Output the text of the first paragraph.
        Console.WriteLine("Document created successfully.");
        Console.WriteLine("First paragraph text:");
        Console.WriteLine(doc.FirstSection.Body.FirstParagraph.GetText().Trim());
    }
}
