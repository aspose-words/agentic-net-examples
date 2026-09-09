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

        // Set custom spacing before and after each paragraph (in points).
        builder.ParagraphFormat.SpaceBefore = 12; // 12 points before the paragraph.
        builder.ParagraphFormat.SpaceAfter = 12;  // 12 points after the paragraph.

        // Add paragraphs that will inherit the spacing settings.
        builder.Writeln("First paragraph with custom spacing.");
        builder.Writeln("Second paragraph with the same custom spacing.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ParagraphSpacing.docx");
        doc.Save(outputPath);
    }
}
