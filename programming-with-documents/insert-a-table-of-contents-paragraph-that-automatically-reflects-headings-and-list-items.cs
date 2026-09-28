using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents that includes heading levels 1‑3.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.Writeln(); // Add a blank line after TOC.

        // Separate TOC from the main content.
        builder.InsertBreak(BreakType.PageBreak);

        // ----- Headings -----
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1: Overview");

        // ----- List items (included in TOC as level 3) -----
        builder.ListFormat.ApplyBulletDefault();

        // List item A with TC field for TOC.
        builder.InsertField("TC \"Item A\" \\l 3", null);
        builder.Writeln("Item A");

        // List item B with TC field for TOC.
        builder.InsertField("TC \"Item B\" \\l 3", null);
        builder.Writeln("Item B");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Reset style to normal for any following text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // Update all fields (including the TOC) to reflect the inserted entries.
        doc.UpdateFields();

        // Save the document.
        const string outputPath = "TableOfContents.docx";
        doc.Save(outputPath);
    }
}
