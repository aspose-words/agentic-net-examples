using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some sample text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");

        // Optionally save the document to a file (not required for extraction).
        const string docPath = "SampleDocument.docx";
        doc.Save(docPath);

        // Extract plain, unformatted text from the whole document using the Range.Text property.
        string plainText = doc.Range.Text.Trim();

        // Output the extracted text to the console.
        Console.WriteLine("Extracted text:");
        Console.WriteLine(plainText);
    }
}
