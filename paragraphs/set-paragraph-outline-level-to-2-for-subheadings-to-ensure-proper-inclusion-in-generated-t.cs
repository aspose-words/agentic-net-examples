using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents field. It will include headings with outline levels 1‑3.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.Writeln(); // Add a blank line after TOC.

        // Insert a main heading (outline level 1).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        // Insert a subheading and explicitly set its outline level to 2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level2;
        builder.Writeln("Section 1.1");

        // Add some body text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is some sample content under the subheading.");

        // Update fields so the TOC reflects the headings.
        doc.UpdateFields();

        // Save the document.
        doc.Save("Output.docx");
    }
}
