using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a regular paragraph.
        builder.Writeln("This is a normal paragraph.");

        // Apply the built‑in 'Quote' style.
        builder.ParagraphFormat.StyleName = "Quote";

        // Increase left indent for emphasis (value in points).
        builder.ParagraphFormat.LeftIndent = 30;

        // Add a paragraph that uses the Quote style with increased indent.
        builder.Writeln("This paragraph is styled as a Quote with increased left indent.");

        // Reset formatting for any subsequent content (optional).
        builder.ParagraphFormat.ClearFormatting();

        // Save the document to disk.
        string outputPath = "QuoteParagraph.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to {outputPath}");
        }
    }
}
