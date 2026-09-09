using System;
using System.IO;
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

        // Save the document to the current directory (ensures the source file exists).
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
        doc.Save(docPath);

        // Extract plain, unformatted text from the whole document using Range.Text.
        string extractedText = doc.Range.Text;

        // Output the extracted text to the console.
        Console.WriteLine(extractedText.Trim());
    }
}
