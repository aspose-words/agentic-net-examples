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

        // Insert a sample paragraph with enough text to demonstrate justification.
        builder.Writeln("This is a sample paragraph that will be justified. It contains enough text to demonstrate how the justification works on a narrow page layout. The text should wrap correctly between words.");

        // Set the paragraph alignment to justified.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;

        // Enable word wrap (wrap by whole words) for the paragraph.
        builder.ParagraphFormat.WordWrap = true;

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document to the output folder.
        string outputPath = Path.Combine(outputDir, "JustifiedParagraph.docx");
        doc.Save(outputPath);
    }
}
