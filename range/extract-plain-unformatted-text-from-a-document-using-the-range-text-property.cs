using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file paths.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string docPath = Path.Combine(artifactsDir, "Sample.docx");
        string txtPath = Path.Combine(artifactsDir, "Extracted.txt");

        // Create a new blank document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");

        // Save the document to the local file system.
        doc.Save(docPath);

        // Extract plain unformatted text using the Range.Text property.
        string extractedText = doc.Range.Text.Trim();

        // Output the extracted text to the console.
        Console.WriteLine("Extracted text:");
        Console.WriteLine(extractedText);

        // Optionally, write the extracted text to a .txt file.
        File.WriteAllText(txtPath, extractedText);
    }
}
