using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents field.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add headings that will be captured by the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 1.1");
        builder.Writeln("Heading 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 2");
        builder.Writeln("Heading 3");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 3.1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 3.1.1");
        builder.Writeln("Heading 3.1.2");
        builder.Writeln("Heading 3.1.3");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading4;
        builder.Writeln("Heading 3.1.3.1");
        builder.Writeln("Heading 3.1.3.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 3.2");
        builder.Writeln("Heading 3.3");

        // Update all fields (including the TOC) to generate entries.
        doc.UpdateFields();

        // Rebuild the page layout so that TOC page numbers are refreshed.
        doc.UpdatePageLayout();

        // Save the document.
        string outputPath = Path.Combine(artifactsDir, "RebuiltToc.docx");
        doc.Save(outputPath);
    }
}
