using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add initial content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original content.");

        // Append additional text to the end of the document's range.
        builder.MoveToDocumentEnd();
        builder.Writeln("Appended content.");

        // Save the document to a local file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AppendedDocument.docx");
        doc.Save(outputPath);

        // Output the final document text to the console for verification.
        Console.WriteLine("Document text after appending:");
        Console.WriteLine(doc.GetText().Trim());
    }
}
