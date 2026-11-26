using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents that includes only Heading 1 and Heading 2 styles.
        // The \\t switch takes a comma‑separated list of style names.
        builder.InsertTableOfContents(@"\\o ""1-3"" \\h \\z \\t ""Heading 1,Heading 2""");

        // Add a page break after the TOC.
        builder.InsertBreak(BreakType.PageBreak);

        // Add content with various styles.
        // Heading 1 – will appear in TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");

        // Normal paragraph – will NOT appear in TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is a normal paragraph that should not be listed in the TOC.");

        // Heading 2 – will appear in TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1: Background");

        // Title style – will NOT appear in TOC because it is not listed in the \\t switch.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
        builder.Writeln("Document Title");

        // Update fields to generate the TOC.
        doc.UpdateFields();

        // Save the document.
        string outputPath = "TOC_Styles.docx";
        doc.Save(outputPath);
    }
}
