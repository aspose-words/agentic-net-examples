using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a paragraph with some initial text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World! This is the original text.");

        // Replace the existing text using the Range.Replace method.
        // This updates the document content without needing to assign to the read‑only Text property.
        doc.Range.Replace("Hello World! This is the original text.",
                          "This is the replaced text for the whole document.");

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // Save the modified document.
        doc.Save(outputPath);

        // Load the saved document to verify the replacement and print the text to console.
        Document loadedDoc = new Document(outputPath);
        Console.WriteLine("Document text after replacement:");
        Console.WriteLine(loadedDoc.Range.Text);
    }
}
