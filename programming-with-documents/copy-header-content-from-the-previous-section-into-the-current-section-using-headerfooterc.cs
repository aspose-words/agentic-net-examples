using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- First section ----------
        // Add a primary header to the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header of the first section");

        // Return to the body of the first section and add some content.
        builder.MoveToSection(0);
        builder.Writeln("Content of the first section.");

        // Insert a section break to start a new (second) section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Second section ----------
        // Add some body content to the second section.
        builder.Writeln("Content of the second section.");

        // Copy the header from the previous (first) section into the current (second) section.
        // Retrieve the header from the first section.
        HeaderFooter previousHeader = doc.Sections[0].HeadersFooters[HeaderFooterType.HeaderPrimary];
        if (previousHeader != null)
        {
            // Clone the header (deep clone with its child nodes).
            HeaderFooter clonedHeader = (HeaderFooter)previousHeader.Clone(true);

            // Add the cloned header to the second section's HeadersFooters collection.
            doc.Sections[1].HeadersFooters.Add(clonedHeader);
        }

        // Save the document to the local file system.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "HeaderCopyExample.docx");
        doc.Save(outputPath);
    }
}
