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

        // Add a normal paragraph before the first heading.
        builder.Writeln("This is an introductory paragraph.");

        // ----- First heading (Heading 1) -----
        // Force a page break before the heading.
        builder.ParagraphFormat.PageBreakBefore = true;
        // Apply the Heading 1 style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");

        // Reset page‑break flag for following normal paragraphs.
        builder.ParagraphFormat.PageBreakBefore = false;
        // Use the Normal style for body text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content under Heading 1.");

        // ----- Second heading (Heading 2) -----
        builder.ParagraphFormat.PageBreakBefore = true;
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 1.1");

        // Normal paragraph after the second heading.
        builder.ParagraphFormat.PageBreakBefore = false;
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("More content under Heading 1.1.");

        // ----- Third heading (Heading 1) -----
        builder.ParagraphFormat.PageBreakBefore = true;
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 2");

        // Save the document to a local file.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "HeadingsWithPageBreaks.docx");
        doc.Save(outputPath);
    }
}
