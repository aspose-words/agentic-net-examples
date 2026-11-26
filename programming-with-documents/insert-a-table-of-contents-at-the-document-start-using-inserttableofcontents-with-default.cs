using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents at the beginning with default switches.
        // Default switches: show heading levels 1-3, hyperlink entries, hide page numbers for web view, and use outline levels.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Add a page break after the TOC so headings start on a new page.
        builder.InsertBreak(BreakType.PageBreak);

        // Insert sample headings that will appear in the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1: Overview");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2: Details");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1: Deep Dive");

        // Save the document to disk.
        string outputPath = "TableOfContents.docx";
        doc.Save(outputPath);
    }
}
