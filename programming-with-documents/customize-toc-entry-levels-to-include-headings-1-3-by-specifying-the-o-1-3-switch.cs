using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents at the beginning with the switch \o "1-3"
        // to include heading levels 1 through 3.
        builder.MoveToDocumentStart();
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Add some headings of levels 1, 2, and 3.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1 – Introduction");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1 – Overview");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Subsection 1.1.1 – Details");

        // Add a normal paragraph after the headings.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is some body text that follows the headings.");

        // Update fields so the TOC reflects the inserted headings.
        doc.UpdateFields();

        // Save the document to disk.
        const string outputPath = "CustomizedTOC.docx";
        doc.Save(outputPath);
    }
}
