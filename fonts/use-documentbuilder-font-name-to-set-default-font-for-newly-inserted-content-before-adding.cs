using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define output file path
        string outputPath = "DefaultFont.docx";

        // Create a new blank document
        Document doc = new Document();

        // Initialize DocumentBuilder for the document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the default font name for all subsequently inserted content
        string desiredFont = "Arial";
        builder.Font.Name = desiredFont;

        // Validate that the font name was set correctly
        if (!string.Equals(builder.Font.Name, desiredFont, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException($"Failed to set font name to '{desiredFont}'.");
        }

        // Insert a paragraph using the default font
        builder.Writeln("This paragraph is formatted with the default font set to Arial.");

        // Save the document to disk
        doc.Save(outputPath);

        // Verify that the file was created
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException($"The document was not saved to '{outputPath}'.");
        }

        // Optional: indicate success (no user interaction required)
        Console.WriteLine("Document created successfully at " + Path.GetFullPath(outputPath));
    }
}
