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

        // Apply the built‑in "Heading 2" style to the paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

        // Adjust spacing before and after the paragraph (in points).
        builder.ParagraphFormat.SpaceBefore = 12f; // 12 points before
        builder.ParagraphFormat.SpaceAfter = 12f;  // 12 points after

        // Write a sample paragraph that will use the above formatting.
        builder.Writeln("This paragraph uses the built‑in Heading 2 style with custom spacing.");

        // Define output file path.
        string outputPath = "Result.docx";

        // Save the document to disk.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
