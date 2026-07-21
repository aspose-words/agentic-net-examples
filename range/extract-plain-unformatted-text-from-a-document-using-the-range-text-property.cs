using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");
        builder.Writeln("This is a test document.");

        // Save the document locally.
        string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
        doc.Save(filePath);

        // Load the saved document.
        Document loadedDoc = new Document(filePath);

        // Extract plain unformatted text using the Range.Text property.
        string extractedText = loadedDoc.Range.Text.Trim();

        // Output the extracted text.
        Console.WriteLine("Extracted text:");
        Console.WriteLine(extractedText);
    }
}
