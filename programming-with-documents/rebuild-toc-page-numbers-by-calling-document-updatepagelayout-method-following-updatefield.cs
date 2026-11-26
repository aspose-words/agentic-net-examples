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

        // Insert a Table of Contents at the beginning of the document.
        // The switches string defines the TOC format (levels 1‑3, hyperlinks, etc.).
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Insert a page break after the TOC.
        builder.InsertBreak(BreakType.PageBreak);

        // Add first heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.Writeln("Content of chapter 1.");

        // Insert a page break before the next heading.
        builder.InsertBreak(BreakType.PageBreak);

        // Add second heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.Writeln("Content of chapter 2.");

        // Update fields (including the TOC) and then rebuild page layout.
        doc.UpdateFields();
        doc.UpdatePageLayout();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);
    }
}
