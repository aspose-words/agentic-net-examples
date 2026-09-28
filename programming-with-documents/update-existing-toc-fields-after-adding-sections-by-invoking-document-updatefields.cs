using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents at the beginning of the document.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Add a page break after the TOC.
        builder.InsertBreak(BreakType.PageBreak);

        // First chapter with Heading 1.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the introduction content.");

        // Second chapter with Heading 1 and a subheading with Heading 2.
        builder.InsertBreak(BreakType.PageBreak);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2: Details");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1: Overview");
        builder.Writeln("Details about the overview.");

        // Third chapter with Heading 1.
        builder.InsertBreak(BreakType.PageBreak);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 3: Conclusion");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Final remarks.");

        // Update all fields in the document, including the TOC.
        doc.UpdateFields();

        // Save the document to disk.
        string outputPath = "UpdatedTOC.docx";
        doc.Save(outputPath);

        // Simple verification that the file was saved.
        if (System.IO.File.Exists(outputPath))
        {
            Console.WriteLine("Document saved and fields updated successfully.");
        }
    }
}
