using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define output folder and ensure it exists.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document and a builder attached to it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents that includes heading levels 1‑3.
        // \\o "1-3" – include headings 1‑3,
        // \\h – make entries hyperlinks,
        // \\z – hide page numbers in web layout,
        // \\u – use outline levels.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add sample headings that will appear in the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 1.1");
        builder.Writeln("Heading 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 1.2.1");
        builder.Writeln("Heading 1.2.2");

        // A heading level 4 will NOT be included because the TOC is limited to 1‑3.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading4;
        builder.Writeln("Heading 1.2.2.1");

        // Update all fields (including the TOC) so the document reflects the current content.
        doc.UpdateFields();

        // Save the document.
        string outputPath = Path.Combine(artifactsDir, "CustomToc.docx");
        doc.Save(outputPath);
    }
}
