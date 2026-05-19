using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the paragraph line spacing to double (24 points) using a multiple rule.
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 24; // Double of the default 12‑point line spacing.

        // Write a sample paragraph.
        builder.Writeln("This paragraph has double line spacing.");

        // Define the output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DoubleLineSpacing.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        bool fileExists = File.Exists(outputPath);
        Console.WriteLine($"Document saved to '{outputPath}'. File exists: {fileExists}");
    }
}
