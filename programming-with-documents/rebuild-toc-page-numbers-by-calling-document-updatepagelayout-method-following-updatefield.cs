using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents at the beginning of the document.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Add a page break after the TOC.
        builder.InsertBreak(BreakType.PageBreak);

        // Insert some headings that will appear in the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");
        builder.Writeln("This is the introduction chapter content.");

        builder.InsertBreak(BreakType.PageBreak);

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2: Getting Started");
        builder.Writeln("Content for the second chapter.");

        builder.InsertBreak(BreakType.PageBreak);

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1: Installation");
        builder.Writeln("Details about installation.");

        // Update all fields (including the TOC) to reflect the current document structure.
        doc.UpdateFields();

        // Rebuild page layout so that TOC page numbers are recalculated.
        doc.UpdatePageLayout();

        // Save the document to disk.
        string outputPath = "RebuiltTOC.docx";
        doc.Save(outputPath);

        // Optional: Verify that the file was saved successfully.
        if (System.IO.File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
    }
}
