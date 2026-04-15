using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and add sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");
        builder.Writeln("This is a sample document.");

        // Extract plain, unformatted text using the Range.Text property.
        string extractedText = doc.Range.Text;

        // Display the extracted text.
        Console.WriteLine("Extracted text:");
        Console.WriteLine(extractedText.Trim());

        // Save the document locally (optional verification).
        string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
        doc.Save(outputFile);
    }
}
