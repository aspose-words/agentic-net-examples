using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a simple document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading (level 1) – this will be used for splitting.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        // Add some regular content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the first chapter.");

        // Insert an explicit page break.
        builder.InsertBreak(BreakType.PageBreak);

        // Add another heading (level 2) – also a split point.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("More content after the page break.");

        // Define split criteria: split at page breaks and heading paragraphs.
        DocumentSplitCriteria splitCriteria = DocumentSplitCriteria.PageBreak | DocumentSplitCriteria.HeadingParagraph;

        // Configure HTML save options with the custom split criteria.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = splitCriteria,
            DocumentSplitHeadingLevel = 3, // split up to heading level 3
            UpdateFields = false           // avoid processing fields that may cause errors
        };

        // Save the document; it will be split into multiple HTML files (output.html, output_1.html, etc.).
        doc.Save("output.html", saveOptions);
    }
}
