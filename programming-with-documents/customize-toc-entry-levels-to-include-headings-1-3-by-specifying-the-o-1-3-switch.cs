using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents that includes heading levels 1 through 3.
        // The switch \o "1-3" tells Word to include those levels.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Add a blank line after the TOC.
        builder.Writeln();

        // Insert Heading 1.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");

        // Insert Heading 2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1: Overview");

        // Insert Heading 3.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Subsection 1.1.1: Details");

        // Update fields (TOC) so that entries are generated.
        doc.UpdateFields();

        // Save the document to disk.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
