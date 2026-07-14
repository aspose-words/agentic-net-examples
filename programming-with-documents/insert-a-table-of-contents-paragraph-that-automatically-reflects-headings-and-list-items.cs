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

        // Insert a Table of Contents field.
        // \o "1-3"  – include heading levels 1 to 3.
        // \h        – make entries hyperlinks.
        // \z        – hide page numbers in web layout.
        // \u        – use outline levels.
        // \t "ListParagraph;5" – also include paragraphs with the ListParagraph style (list items) as level 5 entries.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u \\t \"ListParagraph;5\"");
        builder.InsertBreak(BreakType.PageBreak);

        // Add some headings that will appear in the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Subsection 1.1.1");

        // Insert a bulleted list. List items use the ListParagraph style, which we added to the TOC switches.
        List list = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = list;
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");
        builder.ListFormat.RemoveNumbers(); // End the list.

        // Add another heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        // Update all fields (including the TOC) so that the entries are generated.
        doc.UpdateFields();

        // Save the document to the local file system.
        doc.Save("TableOfContents.docx");
    }
}
