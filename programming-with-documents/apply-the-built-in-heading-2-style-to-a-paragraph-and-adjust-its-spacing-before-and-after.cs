using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in Heading 2 style to the upcoming paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

        // Disable automatic spacing so custom values take effect.
        builder.ParagraphFormat.SpaceBeforeAuto = false;
        builder.ParagraphFormat.SpaceAfterAuto = false;

        // Set custom spacing (points) before and after the paragraph.
        builder.ParagraphFormat.SpaceBefore = 12; // 12 points before
        builder.ParagraphFormat.SpaceAfter = 6;   // 6 points after

        // Insert the paragraph text.
        builder.Writeln("This is a Heading 2 paragraph with custom spacing.");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "Heading2Spacing.docx");
        doc.Save(outputPath);
    }
}
