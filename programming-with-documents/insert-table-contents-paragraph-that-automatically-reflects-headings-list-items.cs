using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;

class TocExample
{
    static void Main()
    {
        // Create a new empty document and associate a DocumentBuilder with it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents field.
        // The switches configure the TOC to include heading levels 1‑3,
        // create hyperlinks, hide page numbers for web view and use a dash as the entry separator.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u \\p \"-\"");

        // Add a page break so that the following content starts on a new page.
        builder.InsertBreak(BreakType.PageBreak);

        // -----------------------------------------------------------------
        // Populate the document with headings (which the TOC will pick up).
        // -----------------------------------------------------------------

        // Heading 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1 – Introduction");

        // Heading 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1 – Overview");

        // Heading 3
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Subsection 1.1.1 – Details");

        // Reset style to normal for regular text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Some introductory paragraph under the first heading.");

        // -----------------------------------------------------------------
        // Insert a numbered list. List items can also be included in the TOC
        // if we give them an outline level that falls within the TOC range.
        // -----------------------------------------------------------------

        // Create a numbered list template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = list;

        // First list item – give it an outline level of 2 (will appear in TOC).
        builder.ListFormat.ListLevelNumber = 0; // top level of the list
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level2;
        builder.Writeln("First item – appears in TOC as level 2.");

        // Second list item – also level 2.
        builder.Writeln("Second item – appears in TOC as level 2.");

        // Third list item – make it a sub‑item (outline level 3) to be captured as level 3.
        builder.ListFormat.ListLevelNumber = 1; // indent one level
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level3;
        builder.Writeln("Sub‑item – appears in TOC as level 3.");

        // Reset list formatting and outline level for subsequent content.
        builder.ListFormat.List = null;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText;
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // Add another heading to demonstrate multiple entries.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2 – Conclusion");

        // -----------------------------------------------------------------
        // Update all fields (including the TOC) so that the entries are generated.
        // -----------------------------------------------------------------
        doc.UpdateFields();

        // Save the document.
        doc.Save("TableOfContentsWithHeadingsAndList.docx", SaveFormat.Docx);
    }
}
