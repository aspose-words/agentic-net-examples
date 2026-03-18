using System;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

class RemoveFootersExample
{
    static void Main()
    {
        // Load an existing document.
        Document doc = new Document("Input.docx");

        // Iterate through each section and delete all footers.
        foreach (Section section in doc.Sections)
        {
            // Remove the "First" footer (appears on the first page of the section).
            HeaderFooter footer = section.HeadersFooters[HeaderFooterType.FooterFirst];
            footer?.Remove();

            // Remove the "Primary" footer (appears on odd pages).
            footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
            footer?.Remove();

            // Remove the "Even" footer (appears on even pages).
            footer = section.HeadersFooters[HeaderFooterType.FooterEven];
            footer?.Remove();
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
