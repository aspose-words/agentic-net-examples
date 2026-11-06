using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");
        builder.Writeln("This is a sample document.");

        // Save the document locally.
        string filePath = "Sample.docx";
        doc.Save(filePath);

        // Load the document from the saved file.
        Document loadedDoc = new Document(filePath);

        // Extract plain, unformatted text using the Range.Text property.
        string extractedText = loadedDoc.Range.Text;

        // Write the extracted text to the console.
        Console.WriteLine("Extracted text:");
        Console.WriteLine(extractedText);
    }
}
