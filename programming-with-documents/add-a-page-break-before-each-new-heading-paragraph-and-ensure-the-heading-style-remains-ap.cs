using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder attached to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an introductory paragraph (normal style).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is an introductory paragraph that precedes the first heading.");

        // Insert a page break before the first heading.
        builder.InsertBreak(BreakType.PageBreak);

        // Add Heading 1. The heading style is applied after the page break.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Getting Started");

        // Add some body text under the heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content for Chapter 1 goes here. It follows the heading on a new page.");

        // Insert a page break before the next heading.
        builder.InsertBreak(BreakType.PageBreak);

        // Add Heading 2. The heading style remains applied.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1: Overview");

        // Add body text under the second heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Details for Section 1.1 are presented here.");

        // Insert a page break before another heading.
        builder.InsertBreak(BreakType.PageBreak);

        // Add Heading 2 again.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.2: Details");

        // Add body text under the third heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Further information for Section 1.2.");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeadingPageBreak.docx");

        // Save the document to the specified file.
        doc.Save(outputPath);
    }
}
