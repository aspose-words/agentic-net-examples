using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TOC that includes outline levels 1 and 2.
        builder.InsertTableOfContents("\\o \"1-2\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // First heading (outline level 1, using built‑in Heading 1 style).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        // Subheading – set explicit outline level to 2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level2;
        builder.Writeln("Section 1.1");

        // Another top‑level heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText; // reset to default.
        builder.Writeln("Chapter 2");

        // Second subheading with outline level 2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level2;
        builder.Writeln("Section 2.1");

        // Update fields so the TOC reflects the inserted entries.
        doc.UpdateFields();

        // Save the document.
        doc.Save(Path.Combine(artifactsDir, "OutlineLevelExample.docx"));
    }
}
