using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents at the beginning of the document.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Add some initial headings that will appear in the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("First Section");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Subsection A");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Subsection B");

        // Return to normal style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // Insert a page break before adding a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Add a new heading that should be captured by the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("New Section Added Programmatically");

        // Return to normal style for any further content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // Update all fields in the document, including the TOC.
        doc.UpdateFields();

        // Save the modified document.
        doc.Save("OutputWithUpdatedToc.docx");
    }
}
