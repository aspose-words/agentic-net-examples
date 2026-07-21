using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents that will capture headings of levels 1‑3.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add a main heading using the built‑in Heading 1 style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Main Heading");

        // Add a subheading. Use a normal style and explicitly set its outline level to Level2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level2;
        builder.Writeln("Subheading (Outline Level 2)");

        // Add regular body text.
        builder.Writeln("This is some body text under the subheading.");

        // Update fields so the TOC reflects the inserted headings.
        doc.UpdateFields();

        // Save the document.
        string outputDir = "Output";
        System.IO.Directory.CreateDirectory(outputDir);
        string outputPath = System.IO.Path.Combine(outputDir, "OutlineLevelExample.docx");
        doc.Save(outputPath);
    }
}
