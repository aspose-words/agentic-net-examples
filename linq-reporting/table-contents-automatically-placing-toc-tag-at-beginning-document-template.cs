using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class GenerateToc
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample headings to demonstrate the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Subsection 1.1.1");

        // Move the cursor to the very beginning of the document.
        builder.MoveToDocumentStart();

        // Insert a Table of Contents field.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Optionally insert a page break after the TOC.
        builder.InsertBreak(BreakType.PageBreak);

        // Update all fields so the TOC reflects the current headings.
        doc.UpdateFields();

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DocumentWithToc.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
