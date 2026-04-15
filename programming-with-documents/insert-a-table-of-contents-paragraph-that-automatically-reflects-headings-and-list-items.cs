using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents field that picks up headings level 1‑3.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add some headings.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Section 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        // Add a bulleted list that should also appear in the TOC.
        List list = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = list;

        // Set outline level for list items so they are included in the TOC.
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;
        builder.Writeln("First list item");
        builder.Writeln("Second list item");

        // Reset list formatting and outline level.
        builder.ListFormat.RemoveNumbers();
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText;

        // Update fields to generate the TOC entries.
        doc.UpdateFields();

        // Save the document to an output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "TableOfContents.docx");
        doc.Save(outputPath);
    }
}
