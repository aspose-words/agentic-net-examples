using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a sample document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");
        builder.Writeln("This is a sample document.");

        // Save the document locally.
        string filePath = "Sample.docx";
        doc.Save(filePath);

        // Load the document from the saved file.
        Document loadedDoc = new Document(filePath);

        // Extract plain unformatted text using the Range.Text property.
        string plainText = loadedDoc.Range.Text;

        // Output the extracted text.
        Console.WriteLine(plainText);
    }
}
