using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents field at the beginning of the document.
        // \\o "1-3" – include outline levels 1 to 3.
        // \\h – use hyperlinks.
        // \\z – hide page numbers in web layout.
        // \\u – use paragraph outline levels.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Add a page break after the TOC so content starts on a new page.
        builder.InsertBreak(BreakType.PageBreak);

        // ----- Sample content with headings and list items -----

        // Heading 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");

        // Normal paragraph under heading
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This paragraph introduces the topic.");

        // Heading 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1: Overview");

        // Numbered list
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("First numbered item");
        builder.Writeln("Second numbered item");
        builder.ListFormat.RemoveNumbers();

        // Heading 1 again
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2: Details");

        // Bulleted list
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("First bullet item");
        builder.Writeln("Second bullet item");
        builder.ListFormat.RemoveNumbers();

        // Update all fields (including the TOC) so they reflect the current document structure.
        doc.UpdateFields();

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Optional verification that the file was saved correctly.
        if (System.IO.File.Exists(outputPath))
        {
            Document reopened = new Document(outputPath);
        }
    }
}
