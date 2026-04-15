using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a primary header to the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header of Section 1");

        // Return to the main body and add some content.
        builder.MoveToDocumentEnd();
        builder.Writeln("Content of Section 1.");

        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Copy the header from the previous section into the current section.
        Section previousSection = doc.Sections[0];
        HeaderFooter previousHeader = previousSection.HeadersFooters[HeaderFooterType.HeaderPrimary];

        if (previousHeader != null)
        {
            // Clone the header to detach it from the original section.
            HeaderFooter copiedHeader = (HeaderFooter)previousHeader.Clone(true);
            // Add the cloned header to the current section's HeadersFooters collection.
            doc.Sections[1].HeadersFooters.Add(copiedHeader);
        }

        // Add body content to the second section.
        builder.Writeln("Content of Section 2.");

        // Save the resulting document.
        doc.Save("HeaderCopy.docx");
    }
}
