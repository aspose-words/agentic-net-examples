using System;
using System.IO;
using Aspose.Words; // HeaderFooter, HeaderFooterType, BreakType are in this namespace

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content to the first section.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Content of the first section.");

        // Create a primary header for the first section and add a paragraph to it.
        HeaderFooter firstHeader = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        firstHeader.AppendParagraph("Header for the first section");
        // Add the header to the first section's HeadersFooters collection.
        doc.FirstSection.HeadersFooters.Add(firstHeader);

        // Insert a section break to start a new (second) section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of the second section.");

        // Retrieve the header from the previous (first) section.
        HeaderFooter previousHeader = doc.Sections[0].HeadersFooters[HeaderFooterType.HeaderPrimary];

        if (previousHeader != null)
        {
            // Clone the header so it can be added to another section.
            HeaderFooter clonedHeader = (HeaderFooter)previousHeader.Clone(true);
            // Add the cloned header to the second section.
            doc.Sections[1].HeadersFooters.Add(clonedHeader);
        }

        // Save the document to a file in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeaderCopy.docx");
        doc.Save(outputPath);
    }
}
