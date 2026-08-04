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

        // Insert a TOC that captures outline levels 1‑9 and creates hyperlinks.
        builder.InsertTableOfContents("\\o \"1-9\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add headings (levels 1‑3).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Section 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        // Add a bulleted list and set its outline level so it appears in the TOC.
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.ListLevelNumber = 0;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level1; // Include in TOC.
        builder.Writeln("First bullet item");
        builder.Writeln("Second bullet item");
        // Reset outline level for subsequent paragraphs.
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText;
        builder.ListFormat.RemoveNumbers();

        // Add another heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1");

        // Update fields so the TOC reflects the added content.
        doc.UpdateFields();

        // Save the document.
        doc.Save("TableOfContents.docx");
    }
}
