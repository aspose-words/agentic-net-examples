using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document sourceDocument = new Document();

        // Use DocumentBuilder to add some content to the document.
        DocumentBuilder builder = new DocumentBuilder(sourceDocument);
        builder.Writeln("This is the source document created programmatically.");
        builder.InsertParagraph();
        builder.Writeln("It contains two paragraphs.");

        // Example verification – output the number of sections in the loaded document.
        Console.WriteLine($"Document created successfully. Sections count: {sourceDocument.Sections.Count}");
    }
}
