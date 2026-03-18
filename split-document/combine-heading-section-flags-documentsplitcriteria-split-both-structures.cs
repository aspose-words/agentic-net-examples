using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

namespace DocumentSplitExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add content with headings and section breaks.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // First heading – will be a split point because of HeadingParagraph flag.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 1");

            // Add some body text.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("This is the first chapter.");

            // Insert a section break – will be a split point because of SectionBreak flag.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Second heading – another split point.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 2");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("This is the second chapter.");

            // Prepare HtmlSaveOptions with combined split criteria.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Combine HeadingParagraph and SectionBreak flags using bitwise OR.
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph | DocumentSplitCriteria.SectionBreak,
                // Split at headings up to level 2 (optional, can be set to 0 to ignore headings).
                DocumentSplitHeadingLevel = 2,
                // Export each part to a separate HTML file.
                ExportDocumentProperties = true
            };

            // Save the document; each split part will be written to a separate file.
            // The base file name is "SplitDocument.html". Additional parts will be named
            // "SplitDocument-01.html", "SplitDocument-02.html", etc.
            doc.Save("SplitDocument.html", saveOptions);
        }
    }
}
