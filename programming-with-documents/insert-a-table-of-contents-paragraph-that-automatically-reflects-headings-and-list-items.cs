using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents (TOC) field.
        // \o "1-3"  – include heading levels 1 through 3.
        // \h        – make entries clickable hyperlinks.
        // \z        – hide page numbers in web layout.
        // \u        – use outline levels.
        // \t "List Paragraph" – include paragraphs styled as List Paragraph (list items).
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u \\t \"List Paragraph\"");
        builder.InsertBreak(BreakType.PageBreak);

        // ---------- Add headings ----------
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1: Overview");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Subsection 1.1.1: Details");

        // ---------- Add a bulleted list ----------
        // Create a bullet list using a predefined list template.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;
        builder.ListFormat.ListLevelNumber = 0; // top level

        builder.Writeln("First bullet item");
        builder.Writeln("Second bullet item");
        builder.Writeln("Third bullet item");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Add another heading after the list.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.2: After List");

        // Update all fields (including the TOC) so the result is stored in the file.
        doc.UpdateFields();

        // Save the document to the local file system.
        doc.Save("TableOfContents.docx");
    }
}
