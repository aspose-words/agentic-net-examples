using System;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Associate a DocumentBuilder with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents (TOC) field.
        // The switches configure the TOC to include heading levels 1‑3,
        // make entries clickable hyperlinks, and preserve formatting.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Insert a page break so that the headings start on a new page.
        builder.InsertBreak(BreakType.PageBreak);

        // Add headings of various levels. Each heading that falls within the
        // range specified in the TOC switches (1‑3) will appear as an entry.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1 – Introduction");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("1.1 Overview");
        builder.Writeln("1.2 Scope");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2 – Details");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("2.1 Architecture");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("2.1.1 Components");
        builder.Writeln("2.1.2 Communication");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("2.2 Implementation");

        // Update all fields in the document so the TOC reflects the inserted headings.
        doc.UpdateFields();

        // Save the document to disk.
        doc.Save("Generated_TOC.docx");
    }
}
