using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents at the beginning using default switches.
        // Default switches: include heading levels 1‑3, hyperlink entries, hide page numbers in web view, and use outline levels.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Add headings that will be captured by the TOC.
        builder.MoveToDocumentEnd();

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.Writeln("Content of chapter 1.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Details of section 1.1.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.Writeln("Content of chapter 2.");

        // Update all fields (including the TOC) before saving.
        doc.UpdateFields();

        // Save the document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
