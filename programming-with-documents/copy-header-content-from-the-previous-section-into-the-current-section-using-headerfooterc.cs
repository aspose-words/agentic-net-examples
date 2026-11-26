using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a header to the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header of Section 1");
        builder.MoveToDocumentEnd();

        // Add body text to the first section.
        builder.Writeln("This is the first section.");

        // Insert a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        // Add body text to the second section.
        builder.Writeln("This is the second section.");

        // Copy the header from the previous section (section 0) to the current section (section 1).
        Section previousSection = doc.Sections[0];
        Section currentSection = doc.Sections[1];

        HeaderFooter previousHeader = previousSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        if (previousHeader != null)
        {
            // Clone the header and add it to the current section's header/footer collection.
            HeaderFooter clonedHeader = (HeaderFooter)previousHeader.Clone(true);
            currentSection.HeadersFooters.Add(clonedHeader);
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
