using System;
using System.IO;
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

        // Insert a Table of Contents field.
        // \o "1-3" – include heading levels 1 to 3.
        // \h – make entries hyperlinks.
        // \z – hide tab leader and page numbers in web layout.
        // \u – use paragraph outline levels.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add headings that will appear in the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Section 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1");

        // Add a bulleted list. To make list items appear in the TOC,
        // set an outline level for each list paragraph.
        List list = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = list;

        // List item 1
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level3; // Include in TOC
        builder.Writeln("First item");

        // List item 2
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level3;
        builder.Writeln("Second item");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Reset outline level to default for subsequent paragraphs.
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText;

        // Update all fields (including the TOC) so the document shows the correct entries.
        doc.UpdateFields();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableOfContents.docx");
        doc.Save(outputPath);
    }
}
