using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load an existing document that already contains a TOC field.
        Document doc = new Document("InputWithToc.docx");

        // Use DocumentBuilder to add a new section with a heading that the TOC should pick up.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new section on a new page.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Apply a heading style (e.g., Heading 1) so the TOC will include this entry.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("New Section Added Programmatically");

        // Return to normal style for subsequent content if needed.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // Update all fields in the document, including the TOC, to reflect the new heading.
        doc.UpdateFields();

        // Save the modified document.
        doc.Save("OutputWithUpdatedToc.docx");
    }
}
