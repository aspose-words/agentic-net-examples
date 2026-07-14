using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace HeaderCopyExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ----- Section 1 -----
            // Create a primary header for the first section.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Header for Section 1");

            // Return to the main body of the document.
            builder.MoveToDocumentEnd();

            // Add some body text to the first section.
            builder.Writeln("Content of Section 1.");

            // Insert a section break to start a new section.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // ----- Section 2 -----
            // Copy the header from the previous (first) section.
            HeaderFooter previousHeader = doc.Sections[0].HeadersFooters[HeaderFooterType.HeaderPrimary];
            if (previousHeader != null)
            {
                // Clone the header node (deep clone) and add it to the current section.
                HeaderFooter copiedHeader = (HeaderFooter)previousHeader.Clone(true);
                doc.Sections[1].HeadersFooters.Add(copiedHeader);
            }

            // Add body text to the second section.
            builder.Writeln("Content of Section 2.");

            // Save the document to a file in the current directory.
            doc.Save("HeaderCopyResult.docx");
        }
    }
}
