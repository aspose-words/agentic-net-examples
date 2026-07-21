using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the default font size that will be applied to all subsequently inserted text.
        builder.Font.Size = 14; // 14 points

        // Insert several paragraphs; they will inherit the default font size.
        builder.Writeln("First paragraph with default font size.");
        builder.Writeln("Second paragraph with the same default font size.");
        builder.Writeln("Third paragraph continues to use the default size.");

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DefaultFontSize.docx");

        // Save the document.
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
