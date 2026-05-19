using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents field that will include headings level 1‑3.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        // Add a page break after the TOC so that headings start on a new page.
        builder.InsertBreak(BreakType.PageBreak);

        // Add headings that will become TOC entries.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Section 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Subsection 2.1.1");

        // Insert a new section and add more headings.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 3");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 3.1");

        // Update all fields in the document, including the TOC.
        doc.UpdateFields();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "UpdatedToc.docx");
        doc.Save(outputPath);
    }
}
