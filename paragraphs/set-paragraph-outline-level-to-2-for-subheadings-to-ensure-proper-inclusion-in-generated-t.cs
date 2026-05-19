using System;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TOC that will include outline levels 1 and 2.
        builder.InsertTableOfContents("\\o \"1-2\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add a top‑level heading (outline level 1).
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;
        builder.Writeln("Chapter 1");

        // Add subheadings (outline level 2) that should appear in the TOC.
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Section 1.2");

        // Return to normal body text.
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText;
        builder.Writeln("This is a regular paragraph that will not be listed in the TOC.");

        // Update fields so the TOC reflects the inserted headings.
        doc.UpdateFields();

        // Save the document to an output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "OutlineLevelExample.docx");
        doc.Save(outputPath);
    }
}
