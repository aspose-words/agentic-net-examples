using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents that will include headings up to level 3.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // ----- Main heading (outline level 1) -----
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;
        builder.Writeln("Main Heading");

        // ----- Subheading (outline level 2) -----
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level2; // Set outline level to 2
        builder.Writeln("Subheading Level 2");

        // ----- Normal paragraph (body text) -----
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText;
        builder.Writeln("This is a regular paragraph that will not appear in the TOC.");

        // Update fields so the TOC reflects the inserted headings.
        doc.UpdateFields();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutlineLevelExample.docx");
        doc.Save(outputPath);
    }
}
