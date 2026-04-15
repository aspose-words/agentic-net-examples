using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define output folder and ensure it exists.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two paragraphs with default spacing.
        builder.Writeln("First paragraph with default spacing.");
        builder.Writeln("Second paragraph with default spacing.");

        // Modify spacing of the first paragraph.
        Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
        firstParagraph.ParagraphFormat.SpaceBefore = 24; // 24 points before the paragraph.
        firstParagraph.ParagraphFormat.SpaceAfter = 24;  // 24 points after the paragraph.

        // Save the document.
        string outputPath = Path.Combine(artifactsDir, "ParagraphSpacing.docx");
        doc.Save(outputPath);
    }
}
