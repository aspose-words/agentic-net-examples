using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add some content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello PDF! This document will be saved as a PDF file.");

        // Define the output PDF file path.
        string outputPath = "Hello.pdf";

        // Save the document directly to PDF format.
        doc.Save(outputPath, SaveFormat.Pdf);

        Console.WriteLine($"Document saved to {outputPath}");
    }
}
