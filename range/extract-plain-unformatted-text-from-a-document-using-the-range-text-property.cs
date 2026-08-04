using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text to the document.
        builder.Writeln("Hello world! This is a sample document.");

        // Save the document to a local file.
        string filePath = "Sample.docx";
        doc.Save(filePath);

        // Load the document from the saved file.
        Document loadedDoc = new Document(filePath);

        // Extract plain, unformatted text using the Range.Text property.
        string plainText = loadedDoc.Range.Text;

        // Output the extracted text.
        Console.WriteLine("Extracted text:");
        Console.WriteLine(plainText.Trim());
    }
}
