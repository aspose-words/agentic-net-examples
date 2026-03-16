using System;
using Aspose.Words;
using Aspose.Words.Saving;

class SplitDocumentByPageAndHeading
{
    static void Main()
    {
        // Create a new document and a builder to populate it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Helper to insert a heading that starts on a new page.
        void InsertHeading(string text, int headingLevel)
        {
            // Apply the appropriate built‑in heading style.
            builder.ParagraphFormat.StyleIdentifier = headingLevel switch
            {
                1 => StyleIdentifier.Heading1,
                2 => StyleIdentifier.Heading2,
                3 => StyleIdentifier.Heading3,
                4 => StyleIdentifier.Heading4,
                5 => StyleIdentifier.Heading5,
                _ => StyleIdentifier.Heading1
            };

            // Force the heading to begin on a new page.
            builder.ParagraphFormat.PageBreakBefore = true;

            builder.Writeln(text);

            // Reset formatting for following content.
            builder.ParagraphFormat.ClearFormatting();
        }

        // Insert sample content: headings of various levels interleaved with body text.
        InsertHeading("Chapter 1", 1);
        builder.Writeln("Body text for chapter 1.");
        InsertHeading("Section 1.1", 2);
        builder.Writeln("Details for section 1.1.");
        InsertHeading("Subsection 1.1.1", 3);
        builder.Writeln("More details.");
        InsertHeading("Chapter 2", 1);
        builder.Writeln("Body text for chapter 2.");
        InsertHeading("Section 2.1", 2);
        builder.Writeln("Details for section 2.1.");

        // Configure HtmlSaveOptions to split the document:
        // - Split at explicit page breaks (PageBreak flag).
        // - Split at heading paragraphs (HeadingParagraph flag).
        // - Limit splitting to heading levels 1‑2 (adjust as needed).
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.PageBreak | DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 2
        };

        // Save the document; Aspose.Words will generate separate HTML parts for each split segment.
        doc.Save("SplitDocument.html", saveOptions);
    }
}
